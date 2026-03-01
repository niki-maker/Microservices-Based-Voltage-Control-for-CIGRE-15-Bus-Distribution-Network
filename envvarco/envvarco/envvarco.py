import logging
from pathlib import Path
import cimpy
import multiprocessing
import cmath
import math
import json
import requests
from openpyxl import Workbook
from openpyxl.styles import Font
import numpy as np
from pyvolt import network
from flask import Flask, request
from src.oma_algorithm import fungal_growth_optimizer
#from src.grid_exporter import export_grid_to_excel

# Flask setup
app = Flask(__name__)

# Logging configuration
logging.basicConfig(filename='envvarco.log', level=logging.INFO)
logging.info("🚀 Volt/VAR Control Module Started...")

# ---------------- Timeout-enforced PF ---------------- #
def _solve_worker(system, return_dict):
    try:
        from pyvolt.nv_powerflow import solve
        results_pf, _ = solve(system)
        return_dict["results"] = results_pf
    except Exception as e:
        return_dict["error"] = str(e)

def safe_powerflow(system, timeout=10.0):
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

@app.route("/optimize", methods=["POST"])
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
        base_apparent_power = 25
        system.load_cim_data(res["topology"], base_apparent_power)
        logging.info("✅ System loaded successfully.")

        #try:
        #    node_n8 = system.get_node_by_uuid("N8")
        #    injected_q_mvar = 3.8
        #    node_n8.reactive_power = injected_q_mvar / base_apparent_power
        #    logging.info(f"🔌 Injected {injected_q_mvar} MVAR capacitor at N8")
        #except Exception as e:
        #    logging.error(f"⚠️ Failed to inject capacitor at N8: {e}")

        

        # ---------------- Initial PF ----------------
        initial_pf = safe_powerflow(system, timeout=10.0)
        if initial_pf is None:
            return {"status": "error", "message": "Initial power flow failed."}

        initial_voltages = {
            n.topology_node.name: (
                abs(n.voltage_pu),
                math.degrees(cmath.phase(n.voltage_pu))
            )
            for n in initial_pf.nodes
        }

        # ---------------- Device sets ----------------
        # ---------------- Load device config from JSON ----------------
        try:
            json_path = Path("/app/data/compensator_device.json")
            with open(json_path, "r") as f:
                device_config = json.load(f)

            capacitor_reactive_power = device_config.get("capacitor_reactive_power", {})
            shunt_reactor_reactive_power = device_config.get("shunt_reactor_reactive_power", {})

            logging.info(f"📄 Loaded device config from {json_path}")
        except Exception as e:
            logging.error(f"❌ Failed to load device config: {e}")
            return {"status": "error", "message": f"Device config load failed: {e}"}

        n_caps = len(capacitor_reactive_power)
        n_reacs = len(shunt_reactor_reactive_power)

        activated_capacitors = set()
        activated_reactors = set()

        # ---------------- Objective function ----------------
        def combined_objective(solution):
            binary_solution = [1 if s >= 0.5 else 0 for s in solution]
            cap_bits = binary_solution[:n_caps]
            reac_bits = binary_solution[n_caps:]

            # Reset all devices
            for node_name in capacitor_reactive_power.keys():
                node = system.get_node_by_uuid(node_name)
                if node:
                    node.reactive_power = 0.0
            for node_name in shunt_reactor_reactive_power.keys():
                node = system.get_node_by_uuid(node_name)
                if node:
                    node.reactive_power = 0.0

            # Apply capacitors
            for idx, (node_name, q_mvar) in enumerate(capacitor_reactive_power.items()):
                if cap_bits[idx] == 1:
                    node = system.get_node_by_uuid(node_name)
                    if node:
                        node.reactive_power = q_mvar / base_apparent_power

            # Apply reactors
            for idx, (node_name, q_mvar) in enumerate(shunt_reactor_reactive_power.items()):
                if reac_bits[idx] == 1:
                    node = system.get_node_by_uuid(node_name)
                    if node:
                        node.reactive_power = -(q_mvar / base_apparent_power)

            results_pf_inner = safe_powerflow(system, timeout=3.0)
            if results_pf_inner is None:
                return [float("inf"), float("inf")]

            voltages = {n.topology_node.name: abs(n.voltage_pu) for n in results_pf_inner.nodes}
            voltage_deviation = sum((max(v - 1.05, 0) + max(0.95 - v, 0)) ** 2 for v in voltages.values())
            wear_and_tear = sum(binary_solution)
            return [voltage_deviation, wear_and_tear]

        # ---------------- Run optimizer ----------------
        _, best_sol = fungal_growth_optimizer(
            50, 10,
            [1] * (n_caps + n_reacs),
            [0] * (n_caps + n_reacs),
            n_caps + n_reacs,
            combined_objective
        )

        binary_solution = [1 if s >= 0.5 else 0 for s in best_sol[:-2]]
        cap_bits = binary_solution[:n_caps]
        reac_bits = binary_solution[n_caps:]

        for idx, (node_name, q_mvar) in enumerate(capacitor_reactive_power.items()):
            if cap_bits[idx] == 1:
                activated_capacitors.add((node_name, q_mvar))

        for idx, (node_name, q_mvar) in enumerate(shunt_reactor_reactive_power.items()):
            if reac_bits[idx] == 1:
                activated_reactors.add((node_name, q_mvar))

        # ---------------- Apply final chosen devices ----------------
        for node in system.nodes:
            node.reactive_power = 0.0

        for node_name, q_mvar in activated_capacitors:
            node = system.get_node_by_uuid(node_name)
            if node:
                node.reactive_power = q_mvar / base_apparent_power

        for node_name, q_mvar in activated_reactors:
            node = system.get_node_by_uuid(node_name)
            if node:
                node.reactive_power = -(q_mvar / base_apparent_power)

        # ---------------- Final PF ----------------
        
        try:
            node_n8 = system.get_node_by_uuid("N8")
            injected_q_mvar = 3.8
            node_n8.reactive_power = injected_q_mvar / base_apparent_power
            logging.info(f"🔌 Injected {injected_q_mvar} MVAR capacitor at N8")
        except Exception as e:
            logging.error(f"⚠️ Failed to inject capacitor at N8: {e}")

        final_pf = safe_powerflow(system, timeout=3.0)
        if final_pf is None:
            return {"status": "error", "message": "Final power flow failed."}

        # ✅ Update system with final PF results so export has final values
        for solved_node in final_pf.nodes:
            node = system.get_node_by_uuid(solved_node.topology_node.uuid)
            if node:
                node.voltage = solved_node.voltage
                node.voltage_pu = solved_node.voltage / node.baseVoltage
                node.power = solved_node.power
                node.power_pu = solved_node.power / node.base_apparent_power

        final_voltages = {
            n.topology_node.name: (
                abs(n.voltage_pu),
                math.degrees(cmath.phase(n.voltage_pu))
            )
            for n in final_pf.nodes
        }

        # ---------------- Log voltages ----------------
        logging.info(f"{'Node':<5} {'Case I (Initial)':<20} {'Case I (Final)':<20}")
        for node_name in sorted(initial_voltages.keys(), key=lambda x: int(x[1:]) if x[1:].isdigit() else x):
            mag_i, ang_i = initial_voltages[node_name]
            mag_f, ang_f = final_voltages[node_name]
            logging.info(f"{node_name:<5} "
                         f"{mag_i:.3f}∠ {ang_i:6.2f}°   "
                         f"{mag_f:.3f}∠ {ang_f:6.2f}°")
            
        # ---------------- Tap-Changing Transformer Control ---------------- #
        lower_limit = 0.95
        upper_limit = 1.05
        max_iterations = 10
        converged = False
        iteration = 0

        # Create a voltage dictionary from final_pf
        voltages_second_pf = {
            n.topology_node.name: abs(n.voltage_pu)
            for n in final_pf.nodes
        }

        while not converged and iteration < max_iterations:
            print(f"\nIteration {iteration + 1}: Adjusting Transformer Tap Ratios")
            converged = True

            tap_adjustment_map = {}

            for branch in system.branches:
                branch.initial_tap_ratio = getattr(branch, "tap_ratio", 1.0)  # store initial
                branch.tap_adjust_count = 0  # initialize counter
                if (branch.start_node.name, branch.end_node.name) in [("N11", "N12"), ("N12", "N13")]:
                    downstream_node = branch.end_node.name
                    voltage = voltages_second_pf.get(downstream_node, 1.0)
                    print(f"Voltage at {downstream_node}: {voltage:.3f} p.u.")

                    if voltage < lower_limit:
                        branch.tap_ratio += 0.01
                        branch.tap_updated = True
                        voltages_second_pf[downstream_node] *= branch.tap_ratio
                        print(f"Increased tap ratio of {branch.start_node.name}-{downstream_node} to {branch.tap_ratio:.3f}")
                        converged = False

                    elif voltage > upper_limit:
                        branch.tap_ratio -= 0.01
                        branch.tap_updated = True
                        voltages_second_pf[downstream_node] *= branch.tap_ratio
                        print(f"Decreased tap ratio of {branch.start_node.name}-{downstream_node} to {branch.tap_ratio:.3f}")
                        converged = False

            # Sync updated voltages back to system
            for node in system.nodes:
                if node.name in voltages_second_pf:
                    node.voltage_pu = voltages_second_pf[node.name]

            print("\nUpdated Voltages After Tap Adjustment:")
            for node_name, voltage_value in voltages_second_pf.items():
                print(f"{node_name} = {voltage_value:.3f} p.u.")

            iteration += 1
                

        # ---------------- Display Final Voltages ---------------- #
        print(f"\n{'Node':<5} {'Final Voltage (Post-Tap)':<25}")
        for node_name in sorted(voltages_second_pf.keys(), key=lambda x: int(x[1:]) if x[1:].isdigit() else x):
            voltage = voltages_second_pf[node_name]
            print(f"{node_name:<5} {voltage:.3f} p.u.")
        
        

        # ---------------- Export final grid state to Excel ----------------
        #export_success = export_grid_to_excel(system)
        #if export_success:
        #    logging.info("✅ Grid state exported successfully after final PF")
        #else:
        #    logging.warning("⚠️ Grid export failed after final PF")






        # ---------------- Prepare unified export ----------------
        wb = Workbook()
        ws = wb.active
        ws.title = "Voltage Optimization Summary"

        # Header
        ws.append([
            "Node", "Device Type", "Reactive Power (MVAR)",
            "Voltage Status", "Tap Changed", "No. of Tap Adjustments"
        ])
        bold_font = Font(bold=True)
        for cell in ws[1]:
            cell.font = bold_font

        # ---------------- Build node-wise export ----------------
        all_nodes = {n.name for n in system.nodes}
        device_map = {}  # node_name → list of device types and reactive power

        for node_name, q_mvar in activated_capacitors:
            device_map.setdefault(node_name, []).append(("Capacitor", q_mvar))

        for node_name, q_mvar in activated_reactors:
            device_map.setdefault(node_name, []).append(("Reactor", q_mvar))

        for node_name in sorted(all_nodes, key=lambda x: int(x[1:]) if x[1:].isdigit() else x):
            # Voltage status
            node_obj = next((n for n in initial_pf.nodes if n.topology_node.name == node_name), None)
            v_pu = abs(node_obj.voltage_pu) if node_obj else 1.0
            if v_pu > 1.05:
                v_status = "Overvoltage"
            elif v_pu < 0.95:
                v_status = "Undervoltage"
            else:
                v_status = "Normal"

            # Tap info
            tap_count = tap_adjustment_map.get(node_name, 0)
            tap_changed = "Yes" if tap_count > 0 else "No"

            # Devices
            if node_name in device_map:
                for dev_type, q_mvar in device_map[node_name]:
                    ws.append([
                        node_name, dev_type, q_mvar,
                        v_status, tap_changed, tap_count
                    ])
            else:
                ws.append([
                    node_name, "", "",  # No device
                    v_status, tap_changed, tap_count
                ])

        # ---------------- Save Workbook ----------------
        try:
            wb.save("/shared_volume/voltage_optimization_summary.xlsx")
            logging.info("✅ Voltage optimization summary exported successfully.")
        except Exception as e:
            logging.error(f"❌ Failed to export voltage summary: {e}")


        try:
            response = requests.post(
                "http://trigger_var_control:4004/optimize",
                json={"base_apparent_power": base_apparent_power}
            )
            if response.status_code == 200:
                print("✅ Volt/VAR control completed.")
            else:
                print(f"⚠️ Volt/VAR module returned: {response.status_code} - {response.text}")
        except Exception as e:
            logging.error(f"❌ Failed to trigger Volt/VAR module: {e}")
            print(f"❌ Failed to trigger Volt/VAR module: {e}")



        # ---------------- Return results ----------------
        return {
            "status": "success",
            "activated_capacitors": sorted(activated_capacitors),
            "activated_reactors": sorted(activated_reactors)
        }

    except Exception as e:
        logging.error(f"❌ Optimization failed: {e}")
        return {"status": "error", "message": str(e)}

@app.route("/health", methods=["GET"])
def health_check():
    return "🟢 Volt/VAR control module is live on port 4002", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=4002)