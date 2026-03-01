import logging
from pathlib import Path
import cimpy
import multiprocessing
import cmath
import math
import json
import numpy as np
from pyvolt import network
from flask import Flask, request
import networkx as nx

# Flask setup
app = Flask(__name__)

# Logging configuration
logging.basicConfig(filename='bus_priority_score.log', level=logging.INFO)
logging.info("🚀 Volt/VAR Control Module Started...")

# ---------------- Timeout-enforced PF ---------------- #
def _solve_worker(system, return_dict):
    try:
        from pyvolt.nv_powerflow import solve
        results_pf, _ = solve(system)
        return_dict["results"] = results_pf
    except Exception as e:
        return_dict["error"] = str(e)

def safe_powerflow(system, timeout=3.0):
    manager = multiprocessing.Manager()
    return_dict = manager.dict()
    p = multiprocessing.Process(target=_solve_worker, args=(system, return_dict))
    p.start()
    p.join(timeout)

    if p.is_alive():
        p.terminate()
        p.join()
        logging.warning("⏱️ Power flow forcibly terminated after %.1f seconds.", timeout)
        return None

    if "error" in return_dict:
        logging.warning(f"⚠️ Power flow failed: {return_dict['error']}")
        return None

    return return_dict.get("results", None)
# ------------------------------------------------------ #


def optimize_powerflow(base_apparent_power=25):
    try:
        # ---------------- Load system ----------------
        this_file_folder = Path(__file__).resolve().parent
        xml_path = this_file_folder / "network"
        xml_files = [str(xml_path / fname) for fname in [
            "Rootnet_FULL_NE_06J16h_DI.xml",
            "Rootnet_FULL_NE_06J16h_EQ.xml",
            "Rootnet_FULL_NE_06J16h_SV.xml",
            "Rootnet_FULL_NE_06J16h_TP.xml"
        ]]
        res = cimpy.cim_import(xml_files, "cgmes_v2_4_15")
        system = network.System()
        system.load_cim_data(res["topology"], base_apparent_power)
        logging.info("✅ System loaded successfully.")

        # ---------------- Initial PF ----------------
        initial_pf = safe_powerflow(system, timeout=3.0)
        if initial_pf is None:
            return {"status": "error", "message": "Initial power flow failed."}

        initial_voltages = {
            n.topology_node.name: (
                abs(n.voltage_pu),
                math.degrees(cmath.phase(n.voltage_pu))
            )
            for n in initial_pf.nodes
        }

        # ---------------- Log voltages ----------------
        logging.info(f"{'Node':<5} {'Case I (Initial)':<20} {'Case I (Final)':<20}")
        for node_name in sorted(initial_voltages.keys(), key=lambda x: int(x[1:]) if x[1:].isdigit() else x):
            mag_i, ang_i = initial_voltages[node_name]

            logging.info(f"{node_name:<5} "
                         f"{mag_i:.3f}∠ {ang_i:6.2f}°   ")

        # ---------------- Compute impedance distances ----------------
        G = nx.Graph()
        for branch in system.branches:
            if branch.start_node and branch.end_node:
                z_mag = abs(branch.z_pu)
                G.add_edge(branch.start_node.uuid, branch.end_node.uuid, weight=z_mag)

        slack_nodes = [n for n in system.nodes if n.type.name.lower() == "slack"]
        if not slack_nodes:
            logging.warning("⚠️ No slack bus found.")
            return
        slack_uuid = slack_nodes[0].uuid

        distances = nx.single_source_dijkstra_path_length(G, slack_uuid, weight="weight")

        # ---------------- Log priority scores ----------------
        logging.info("📊 Priority Score Summary:")
        for node in system.nodes:
            p = node.power.real
            q = node.power.imag
            impedance = distances.get(node.uuid, float("inf"))
            logging.info(f"🔌 Bus {node.name} (UUID: {node.uuid})")
            logging.info(f"   Load: P = {p:.3f}, Q = {q:.3f}")
            logging.info(f"   Impedance from Slack: {impedance:.4f} pu")

        # ---------------- Sensitivity-Based Remoteness ----------------
        logging.info("🔁 Starting sensitivity-based remoteness analysis...")

        # Initialize accumulator for ΔV across all injections
        delta_v_accumulator = {node.uuid: [] for node in system.nodes}

        for injection_node in system.nodes:
            try:
                # Reset all reactive injections
                for node in system.nodes:
                    node.reactive_power = 0.0

                # Inject reactive power at current node
                injected_q_mvar = 0.5
                injection_node.reactive_power = injected_q_mvar / base_apparent_power
                logging.info(f"⚡ Injected {injected_q_mvar} MVAR at Bus {injection_node.name}")

                # Run power flow
                updated_pf = safe_powerflow(system, timeout=5.0)
                if updated_pf is None:
                    logging.warning(f"❌ PF failed after injection at Bus {injection_node.name}")
                    continue

                # Record ΔV for all buses
                for solved_node in updated_pf.nodes:
                    node = system.get_node_by_uuid(solved_node.topology_node.uuid)
                    if node:
                        final_voltage = abs(solved_node.voltage)
                        initial_voltage = initial_voltages.get(node.uuid, None)
                        if initial_voltage is not None:
                            delta_v = abs(final_voltage - initial_voltage)
                            delta_v_accumulator[node.uuid].append(delta_v)

            except Exception as e:
                logging.error(f"⚠️ Injection at Bus {injection_node.name} failed: {e}")

        # ---------------- Compute Hybrid Remoteness Score ----------------
        logging.info("🧮 Computing hybrid electrical remoteness scores...")

        # Normalize impedance scores
        all_impedances = [distances.get(node.uuid, float("inf")) for node in system.nodes]
        z_min = min(all_impedances)
        z_max = max(all_impedances)
        impedance_scores = {
            node.uuid: (distances.get(node.uuid, float("inf")) - z_min) / (z_max - z_min)
            for node in system.nodes
        }

        # Compute average ΔV across all injections
        avg_delta_v_scores = {
            uuid: np.mean(deltas) if deltas else 0.0
            for uuid, deltas in delta_v_accumulator.items()
        }

        max_avg_delta_v = max(avg_delta_v_scores.values())

        sensitivity_scores = {
            uuid: 1.0 - (avg_delta_v / max_avg_delta_v) if max_avg_delta_v > 0 else 1.0
            for uuid, avg_delta_v in avg_delta_v_scores.items()
        }

        # Combine both scores
        for node in system.nodes:
            r_imp = impedance_scores.get(node.uuid, 1.0)
            r_sens = sensitivity_scores.get(node.uuid, 1.0)
            r_elec = 0.6 * r_imp + 0.4 * r_sens

            logging.info(f"📌 Bus {node.name} | R_imp = {r_imp:.4f} | R_sens = {r_sens:.4f} | R_elec = {r_elec:.4f}")


    except Exception as e:
        logging.error(f"❌ Optimization failed: {e}")
        return {"status": "error", "message": str(e)}

@app.route("/health", methods=["GET"])
def health_check():
    return "🟢 Volt/VAR control module is live on port 4002", 200

if __name__ == "__main__":
    optimize_powerflow()
    app.run(host="0.0.0.0", port=4003)