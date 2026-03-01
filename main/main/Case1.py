import os
import logging
import pandas as pd
import numpy as np
import requests
from flask import Flask
from pathlib import Path
import cimpy
from pyvolt import network
from pyvolt import nv_powerflow
import time

# Flask app
app = Flask(__name__)

# Logging setup
logging.basicConfig(filename="main.log", level=logging.INFO)
logging.info("🔁 Starting grid parser in main.py...")

# 📁 Ensure shared volume folder exists
os.makedirs("/shared_volume", exist_ok=True)

def export_grid_to_excel(system, path="/shared_volume/grid_data.xlsx"):
    try:
        node_records = []
        for node in system.nodes:
            node_records.append({
                "name": node.name,
                "uuid": node.uuid,
                "voltage_pu": abs(node.voltage_pu),
                "voltage_angle_deg": np.angle(node.voltage_pu, deg=True),
                "power_pu": abs(node.power_pu),
                "power_angle_deg": np.degrees(np.angle(node.power_pu)),
                "base_voltage": node.baseVoltage,
                "base_apparent_power": node.base_apparent_power,
                "real_power": node.power.real,
                "imag_power": node.power.imag,
                "voltage_real": node.voltage.real,
                "voltage_imag": node.voltage.imag
            })
        node_df = pd.DataFrame(node_records)

        branch_records = []
        for branch in system.branches:
            branch_records.append({
                "uuid": branch.uuid,
                "from": branch.start_node.name if branch.start_node else "Unknown",
                "to": branch.end_node.name if branch.end_node else "Unknown",
                "r": branch.r,
                "x": branch.x,
                "bch": branch.bch,
                "bch_pu": branch.bch_pu,
                "length": branch.length,
                "base_voltage": branch.baseVoltage,
                "base_apparent_power": branch.base_apparent_power,
                "r_pu": branch.r_pu,
                "x_pu": branch.x_pu,
                "z_real": branch.z.real,
                "z_imag": branch.z.imag,
                "z_pu_real": branch.z_pu.real,
                "z_pu_imag": branch.z_pu.imag,
                "type": "transformer" if "TR" in branch.uuid else "line"
            })
        branch_df = pd.DataFrame(branch_records)

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            node_df.to_excel(writer, sheet_name="nodes", index=False)
            branch_df.to_excel(writer, sheet_name="branches", index=False)

        logging.info(f"📁 Excel exported to {path}")
        return True
    except Exception as e:
        logging.error(f"❌ Excel export failed: {e}")
        return False

def parse_and_export():
    try:
        this_file_folder = Path(__file__).resolve().parent
        xml_path = this_file_folder / "network"
        xml_files = [
            str(xml_path / "Rootnet_FULL_NE_06J16h_DI.xml"),
            str(xml_path / "Rootnet_FULL_NE_06J16h_EQ.xml"),
            str(xml_path / "Rootnet_FULL_NE_06J16h_SV.xml"),
            str(xml_path / "Rootnet_FULL_NE_06J16h_TP.xml")
        ]

        res = cimpy.cim_import(xml_files, "cgmes_v2_4_15")
        system = network.System()
        base_apparent_power = 25
        system.load_cim_data(res["topology"], base_apparent_power)
        logging.info("✅ System loaded successfully.")

        results_pf = nv_powerflow.solve(system)[0]

        # Inject solved voltages and powers into system.nodes
        uuid_to_node_map = {node.uuid: node for node in system.nodes}
        for solved_node in results_pf.nodes:
            uuid = solved_node.topology_node.uuid
            target_node = uuid_to_node_map.get(uuid)
            if target_node:
                target_node.voltage = solved_node.voltage
                target_node.voltage_pu = solved_node.voltage / target_node.baseVoltage
                target_node.power = solved_node.power
                target_node.power_pu = solved_node.power / target_node.base_apparent_power

        export_success = export_grid_to_excel(system)

        time.sleep(100)
        
        # Check if any voltage violates the limits
        voltage_violation = any(
            not 0.95 <= abs(node.voltage_pu) <= 1.05 for node in system.nodes
        )
        logging.info("⚠️ hi.")
        if voltage_violation:
            logging.info("⚠️ Voltage violation detected. Triggering Volt/VAR control.")
            print("⚠️ Voltage violation detected. Triggering Volt/VAR control...")

            try:
                response = requests.post(
                    "http://envvarco:4002/optimize",
                    json={"base_apparent_power": base_apparent_power}
                )
                if response.status_code == 200:
                    print("✅ Volt/VAR control completed.")
                else:
                    print(f"⚠️ Volt/VAR module returned: {response.status_code} - {response.text}")
            except Exception as e:
                logging.error(f"❌ Failed to trigger Volt/VAR module: {e}")
                print(f"❌ Failed to trigger Volt/VAR module: {e}")

        return export_success

    except Exception as e:
        logging.error(f"❌ Export failed: {e}")
        return False

@app.route('/health', methods=['GET'])
def health():
    return "🟢 main.py is live on port 4001", 200

if __name__ == "__main__":
    success = parse_and_export()
    if success:
        print("✅ Grid data exported to Excel.")
    else:
        print("❌ Grid export failed. See logs.")
    app.run(host="0.0.0.0", port=4001)