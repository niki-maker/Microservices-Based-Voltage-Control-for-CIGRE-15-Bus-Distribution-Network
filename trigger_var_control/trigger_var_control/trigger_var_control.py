import logging
from pathlib import Path
import cimpy
import multiprocessing
import cmath
import math
import json
from openpyxl import Workbook
from openpyxl.styles import Font
import numpy as np
from pyvolt import network
from flask import Flask, request
import time
import pandas as pd
#from src.grid_exporter import export_grid_to_excel


# Flask setup
app = Flask(__name__)

# Logging configuration
logging.basicConfig(filename='trigger_var_control.log', level=logging.INFO)
logging.info("🚀 Volt/VAR Control Module Started...")

def update_voltage_pu_in_excel(system, path="/shared_volume/grid_data.xlsx"):
    try:
        # Read existing Excel
        df = pd.read_excel(path, sheet_name="nodes")

        # Build latest voltages from system (handles Node or ResultsNode)
        latest_voltages = {
            getattr(getattr(node, "topology_node", node), "name", "Unknown"): abs(node.voltage_pu)
            for node in system.nodes
        }

        # Update voltage_pu column where name matches
        df["voltage_pu"] = df["name"].map(lambda n: latest_voltages.get(n, df.loc[df["name"] == n, "voltage_pu"].values[0]))

        # Save back to same Excel, preserving other sheets
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name="nodes", index=False)

        logging.info(f"✅ Updated voltage_pu in {path}")
        return True

    except Exception as e:
        logging.error(f"❌ Failed to update voltage_pu: {e}")
        return False



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
        data = request.get_json()
        base_apparent_power = data.get("base_apparent_power", 25)
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

        #try:
        #    node_n8 = system.get_node_by_uuid("N8")
        #    injected_q_mvar = 3.8
        #    node_n8.reactive_power = injected_q_mvar / base_apparent_power
        #    logging.info(f"🔌 Injected {injected_q_mvar} MVAR capacitor at N8")
        #except Exception as e:
        #    logging.error(f"⚠️ Failed to inject capacitor at N8: {e}")
        
        # Load Excel files
        summary_df = pd.read_excel("/shared_volume/voltage_optimization_summary.xlsx")
        priority_df = pd.read_excel("/shared_volume/priority_score.xlsx")

        # Merge and sort devices by priority
        def merge_with_priority(df):
            return df.merge(priority_df, left_on="Node", right_on="bus_name").sort_values(by="priority_score", ascending=False)

        capacitor_df = merge_with_priority(summary_df[summary_df["Device Type"] == "Capacitor"])
        reactor_df = merge_with_priority(summary_df[summary_df["Device Type"] == "Reactor"])

        # Convert to list of dicts for easy popping
        capacitor_queue = capacitor_df.to_dict("records")
        reactor_queue = reactor_df.to_dict("records")

        # Track activated nodes
        activated_nodes = set()

        # Initial delay before first activation
        logging.info("⏱️ Waiting 45 seconds before first activation...")
        time.sleep(45)

        # Stepwise activation loop
        activation_count = 0
        while capacitor_queue or reactor_queue:
            # Run power flow
            pf_result = safe_powerflow(system, timeout=10.0)

            if pf_result is None:
                logging.warning("⚠️ Power flow failed during loop.")
                break
            
            # Log all node voltages
            logging.info("📊 Node Voltages After Power Flow:")
            for node in pf_result.nodes:
                name = node.topology_node.name
                mag = abs(node.voltage_pu)
                ang = math.degrees(cmath.phase(node.voltage_pu))
                logging.info(f"{name:<5} = {mag:.3f} ∠ {ang:6.2f}° p.u.")
                export_success = update_voltage_pu_in_excel(pf_result, path="/shared_volume/grid_data.xlsx")
                if export_success:
                    logging.info("✅ voltage_pu column updated successfully after final PF")
                else:
                    logging.warning("⚠️ Failed to update voltage_pu column after final PF")




            # Extract voltages
            voltages = {n.topology_node.name: abs(n.voltage_pu) for n in pf_result.nodes}
            over_count = sum(1 for v in voltages.values() if v > 1.05)
            under_count = sum(1 for v in voltages.values() if v < 0.95)

            # ✅ Exit if all voltages are within range
            if all(0.95 <= v <= 1.05 for v in voltages.values()):
                logging.info("✅ All node voltages within 0.95–1.05 p.u. Exiting control loop.")
                break

            # Determine voltage condition
            if over_count > under_count:
                condition = "Overvoltage"
            elif under_count > over_count:
                condition = "Undervoltage"
            else:
                condition = "Normal"

            logging.info(f"🔁 Voltage condition: {condition}")

            # Select next device based on condition and priority
            next_device = None
            if condition == "Undervoltage":
                # Prefer capacitor first
                while capacitor_queue:
                    candidate = capacitor_queue.pop(0)
                    if candidate["Node"] not in activated_nodes:
                        next_device = candidate
                        break
                # Fallback to reactor
                if not next_device:
                    while reactor_queue:
                        candidate = reactor_queue.pop(0)
                        if candidate["Node"] not in activated_nodes:
                            next_device = candidate
                            break

            elif condition == "Overvoltage":
                # Prefer reactor first
                while reactor_queue:
                    candidate = reactor_queue.pop(0)
                    if candidate["Node"] not in activated_nodes:
                        next_device = candidate
                        break
                # Fallback to capacitor
                if not next_device:
                    while capacitor_queue:
                        candidate = capacitor_queue.pop(0)
                        if candidate["Node"] not in activated_nodes:
                            next_device = candidate
                            break

            else:
                logging.info("✅ Voltage stabilized. Ending control loop.")
                break

            # Delay before activation
            if activation_count == 0:
                logging.info("⏱️ First activation already delayed. Proceeding...")
            else:
                logging.info("⏱️ Waiting 20 seconds before next activation...")
                time.sleep(20)


            # Activate selected device
            if next_device:
                node_name = next_device["Node"]
                q_mvar = next_device["Reactive Power (MVAR)"]
                device_type = next_device["Device Type"]
                try:
                    node = system.get_node_by_uuid(node_name)
                    if node:
                        node.reactive_power = q_mvar / base_apparent_power if device_type == "Capacitor" else -(q_mvar / base_apparent_power)
                        activated_nodes.add(node_name)
                        activation_count += 1
                        logging.info(f"🔌 Activated {device_type} at {node_name} with {q_mvar} MVAR")

                        # 🔁 Re-run power flow after activation
                        pf_result = safe_powerflow(system, timeout=10.0)
                        if pf_result:
                            logging.info("📊 Node Voltages After Activation:")
                            for node in pf_result.nodes:
                                name = node.topology_node.name
                                mag = abs(node.voltage_pu)
                                ang = math.degrees(cmath.phase(node.voltage_pu))
                                logging.info(f"{name:<5} = {mag:.3f} ∠ {ang:6.2f}° p.u.")

                            # 📤 Export grid state to Excel
                            export_success = update_voltage_pu_in_excel(pf_result, path="/shared_volume/grid_data.xlsx")
                            if export_success:
                                logging.info("✅ voltage_pu column updated successfully after final PF")
                            else:
                                logging.warning("⚠️ Failed to update voltage_pu column after final PF")

                        else:
                            logging.warning("⚠️ Power flow failed after activation.")
                except Exception as e:
                    logging.error(f"⚠️ Failed to activate {device_type} at {node_name}: {e}")


        return {
            "status": "success",
            "message": "Optimization completed.",
            "activated_nodes": sorted(list(activated_nodes)),
            "activations": activation_count
        }, 200

    except Exception as e:
        logging.error(f"❌ Optimization failed: {e}")

@app.route("/health", methods=["GET"])
def health_check():
    return "🟢 Volt/VAR control module is live on port 4004", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=4004)