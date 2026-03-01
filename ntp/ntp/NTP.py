import os
import time
import logging
import json
from pathlib import Path
import pandas as pd
from flask import Flask, jsonify, request
import threading
from influxdb_client import InfluxDBClient, Point, WriteOptions
from datetime import datetime, timezone

# Flask app
app = Flask(__name__)

# Logging Setup
logging.basicConfig(filename="ntp.log", level=logging.INFO, filemode="w")

# InfluxDB Configuration
INFLUXDB_URL = os.getenv("INFLUXDB_URL", "http://influxdb:8086")
INFLUXDB_BUCKET = os.getenv("INFLUXDB_BUCKET", "Hello")
INFLUXDB_ORG = os.getenv("INFLUXDB_ORG", "my-org")
TOKEN_FILE = "/token_storage/token.txt"
EXCEL_PATH = "/shared_volume/grid_data.xlsx"

# Load token
if os.path.exists(TOKEN_FILE):
    with open(TOKEN_FILE, "r") as file:
        INFLUXDB_TOKEN = file.read().strip()
else:
    logging.error("Token file missing! Exiting.")
    exit(1)

# InfluxDB connection
def connect_to_influxdb():
    for _ in range(5):
        try:
            client = InfluxDBClient(url=INFLUXDB_URL, token=INFLUXDB_TOKEN, org=INFLUXDB_ORG)
            logging.info("Connected to InfluxDB.")
            return client
        except Exception as err:
            logging.error(f"InfluxDB connection failed: {err}")
            time.sleep(5)
    logging.error("Unable to connect to InfluxDB after retries. Exiting.")
    exit(1)

client = connect_to_influxdb()
write_api = client.write_api(write_options=WriteOptions(batch_size=1))

# In-memory data cache
node_data = []
branch_data = []

def load_excel_data():
    global node_data, branch_data
    if not os.path.exists(EXCEL_PATH):
        logging.error("Excel file not found.")
        node_data = []
        branch_data = []
        return False
    try:
        df_nodes = pd.read_excel(EXCEL_PATH, sheet_name="nodes")
        df_branches = pd.read_excel(EXCEL_PATH, sheet_name="branches")
        node_data = df_nodes.to_dict(orient="records")
        branch_data = df_branches.to_dict(orient="records")
        logging.info("✅ Excel data reloaded into memory.")
        return True
    except Exception as e:
        logging.error(f"❌ Failed to reload Excel: {e}")
        node_data = []
        branch_data = []
        return False

def ntp_powerflow():
    logging.info("🔁 Starting NTP telemetry push to InfluxDB...")
    if not load_excel_data():
        logging.error("Excel load failed, skipping InfluxDB push.")
        return

    for node in node_data:
        point = (
            Point("node")
            .tag("name", node["name"])
            .tag("uuid", node["uuid"])
            .field("voltage_pu_mag", node["voltage_pu"])
            .field("voltage_pu_phase_deg", node["voltage_angle_deg"])
            .field("load_pu_mag", node["power_pu"])
            .field("load_pu_angle_deg", node["power_angle_deg"])
            .field("base_voltage", node["base_voltage"])
            .field("base_apparent_power", node["base_apparent_power"])
            .field("load_mag", node["real_power"])
            .field("load_angle_deg", node["imag_power"])
            .field("voltage_real", node["voltage_real"])
            .field("voltage_imag", node["voltage_imag"])
        )
        write_api.write(bucket=INFLUXDB_BUCKET, org=INFLUXDB_ORG, record=point)

    for branch in branch_data:
        measurement = "transformer" if branch["type"] == "transformer" else "branch"
        point = (
            Point(measurement)
            .tag("uuid", branch["uuid"])
            .tag("from", branch["from"])
            .tag("to", branch["to"])
            .field("r", branch["r"])
            .field("x", branch["x"])
            .field("base_voltage", branch["base_voltage"])
            .field("base_apparent_power", branch["base_apparent_power"])
            .field("r_pu", branch["r_pu"])
            .field("x_pu", branch["x_pu"])
            .field("z_real", branch["z_real"])
            .field("z_imag", branch["z_imag"])
            .field("z_pu_real", branch["z_pu_real"])
            .field("z_pu_imag", branch["z_pu_imag"])
            .field("bch", branch["bch"])
            .field("bch_pu", branch["bch_pu"])
            .field("length", branch["length"])
            .field("short_circuit_temp", branch.get("short_circuit_temp", 0.0))
        )
        write_api.write(bucket=INFLUXDB_BUCKET, org=INFLUXDB_ORG, record=point)

    logging.info("✅ Grid telemetry posted to InfluxDB.")

def continuous_telemetry_loop():
    while True:
        try:
            ntp_powerflow()
        except Exception as e:
            logging.error(f"⚠️ Exception in continuous loop: {e}")
        time.sleep(0.5)

@app.route("/run", methods=["POST"])
def trigger():
    threading.Thread(target=ntp_powerflow).start()
    return jsonify({"status": "NTP module started"}), 200

@app.route("/reload", methods=["POST"])
def reload_data():
    success = load_excel_data()
    if success:
        return jsonify({"status": "Reload successful"}), 200
    else:
        return jsonify({"error": "Reload failed"}), 500

@app.route("/grafana_data", methods=["GET"])
@app.route("/grafana_data", methods=["GET"])
def grafana_data():
    if not node_data or not branch_data:
        return jsonify({"error": "Data not loaded"}), 500

    connections = []
    nodes = []

    # ---------------- Load capacitor/reactor config ----------------
    capacitor_reactive_power = {}
    shunt_reactor_reactive_power = {}
    try:
        json_path = Path("/app/data/compensator_device.json")  # volume mount path
        with open(json_path, "r") as f:
            device_config = json.load(f)
        capacitor_reactive_power = device_config.get("capacitor_reactive_power", {})
        shunt_reactor_reactive_power = device_config.get("shunt_reactor_reactive_power", {})
        logging.info(f"📄 Loaded device config from {json_path}")
    except Exception as e:
        logging.error(f"❌ Failed to load device config: {e}")

    # ---------------- Branch connections ----------------
    for branch in branch_data:
        present_voltage = f"{branch['z_pu_real']:.4f}∠{branch['z_pu_imag']:.2f}°"
        conn = {
            "id": branch["uuid"],
            "source": branch["from"],
            "target": branch["to"],
            "base_voltage": f"{branch['base_voltage']} kV",
            "reactance_pu": round(branch["x"], 5),
            "resistance_pu": round(branch["r"], 5),
            "shunt_susceptance_pu": round(branch["bch_pu"], 5),
            "present_voltage": present_voltage,
            "thickness": 7,
            "Message": (
                f"Base voltage: {branch['base_voltage']} kV; "
                f"Reactance_pu (X): {round(branch['x'], 5)}; "
                f"Resistance_pu (R): {round(branch['r'], 5)}; "
                f"Shunt susceptance_pu: {round(branch['bch_pu'], 5)}"
            )
        }
        connections.append(conn)

    # ---------------- Main grid nodes ----------------
    color_palette = [
        "red", "orange", "yellow", "green", "blue", "indigo", "violet", "pink",
        "brown", "cyan", "lime", "magenta", "gold", "silver", "teal"
    ]

    for i, node in enumerate(node_data):
        present_voltage = f"{node['voltage_pu']:.4f}∠{node['voltage_angle_deg']:.2f}°"
        timestamp = datetime.now(timezone.utc).isoformat()
        node_obj = {
            "base_voltage": f"{node['base_voltage']} kV",
            "id": node["name"],
            "label": f"Node {node['name']}",
            "present_voltage": present_voltage,
            "timestamp": timestamp,
            "color": color_palette[i % len(color_palette)],
            "size": 40  # normal size for main nodes
        }
        nodes.append(node_obj)

    # ---------------- Capacitor/Reactors as separate yellow nodes ----------------
    for cap_node, q_mvar in capacitor_reactive_power.items():
        cap_id = f"Cap_{cap_node}"
        nodes.append({
            "id": cap_id,
            "label": f"Capacitor {cap_node}",
            "color": "yellow",
            "size": 20,  # smaller size
            "icon": "square",
            "Message": f"Capacitor-{cap_node}-{q_mvar}MVAR"
        })
        connections.append({
            "id": f"conn_{cap_id}",
            "source": cap_id,
            "target": cap_node,
            "thickness": 3,
            "Message": f"Capacitor-{cap_node}-{q_mvar}MVAR"
        })

    for reac_node, q_mvar in shunt_reactor_reactive_power.items():
        reac_id = f"Reac_{reac_node}"
        nodes.append({
            "id": reac_id,
            "label": f"Reactor {reac_node}",
            "color": "yellow",
            "size": 20,  # smaller size
            "icon": "star",
            "Message": f"Reactor-{reac_node}-{q_mvar}MVAR"
        })
        connections.append({
            "id": f"conn_{reac_id}",
            "source": reac_id,
            "target": reac_node,
            "thickness": 3,
            "Message": f"Reactor-{reac_node}-{q_mvar}MVAR"
        })

    return jsonify({
        "connections": connections,
        "nodes": nodes
    })


@app.route("/health", methods=["GET"])
def health():
    return "🟢 NTP module is live", 200

# Start loop at startup
threading.Thread(target=continuous_telemetry_loop, daemon=True).start()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=4000)