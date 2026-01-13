import os
import subprocess
import threading
import openpyxl
from datetime import datetime
from flask import Flask, render_template, jsonify, request

app = Flask(__name__)

# List of states to display
STATES = sorted([
    "Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chhattisgarh", 
    "Goa", "Gujarat", "Haryana", "Himachal Pradesh", "Jharkhand", 
    "Karnataka", "Kerala", "Madhya Pradesh", "Maharashtra", "Manipur", 
    "Meghalaya", "Mizoram", "Nagaland", "Odisha", "Punjab", 
    "Rajasthan", "Sikkim", "Tamil Nadu", "Telangana", "Tripura", 
    "Uttar Pradesh", "Uttarakhand", "West Bengal", 
    "Andaman and Nicobar Islands", "Chandigarh", "Dadra and Nagar Haveli and Daman and Diu",
    "Delhi", "Jammu and Kashmir", "Ladakh", "Lakshadweep", "Puducherry"
])

# Global state
CURRENT_PROCESSING_STATE = None
AGENT_LOGS = []
IS_AGENT_RUNNING = False

def run_script(script_name, display_name):
    global CURRENT_PROCESSING_STATE, AGENT_LOGS
    CURRENT_PROCESSING_STATE = display_name
    timestamp = datetime.now().strftime("%H:%M:%S")
    AGENT_LOGS.append(f"[{timestamp}] Starting {script_name}...")
    
    # Set UTF-8 encoding environment variable to fix UnicodeEncodeError in scraper.py
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    
    try:
        # Using -u for unbuffered output to ensure real-time logs in the monitor
        process = subprocess.Popen(
            ["python", "-u", script_name],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True,
            env=env
        )
        
        for line in process.stdout:
            timestamp = datetime.now().strftime("%H:%M:%S")
            AGENT_LOGS.append(f"[{timestamp}] {line.strip()}")
            if len(AGENT_LOGS) > 1000:
                AGENT_LOGS.pop(0)
                
        process.wait()
        if process.returncode == 0:
            AGENT_LOGS.append(f"[{timestamp}] {script_name} finished successfully.")
        else:
            AGENT_LOGS.append(f"[{timestamp}] Error: {script_name} exited with code {process.returncode}")
            
    except Exception as e:
        AGENT_LOGS.append(f"[{timestamp}] Exception: {str(e)}")

def agent_worker():
    global IS_AGENT_RUNNING, CURRENT_PROCESSING_STATE, AGENT_LOGS
    IS_AGENT_RUNNING = True
    
    scripts = [
        ("scraper.py", "Scraping"),
        ("ists.py", "ISTS"),
        ("Meghalaya.py", "Meghalaya"),
        ("Rajastan.py", "Rajasthan"),
        ("Madyapradesh.py", "Madhya Pradesh"),
        ("bihar.py", "Bihar"),
        ("puducherry.py", "Puducherry"),
        ("Himachalpradesh.py", "Himachal Pradesh"),
        ("Assam.py", "Assam"),
        ("uttarpradesh.py", "Uttar Pradesh")
    ]
    
    for script, display in scripts:
        run_script(script, display)
        
    CURRENT_PROCESSING_STATE = None
    IS_AGENT_RUNNING = False
    AGENT_LOGS.append(f"[{datetime.now().strftime('%H:%M:%S')}] All tasks completed.")

@app.route('/')
def index():
    state_status = []
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for state in STATES:
        # Check for multiple variants (with/without spaces, lowercase/original)
        variants = [
            f"{state}.xlsx",
            f"{state.lower()}.xlsx",
            f"{state.replace(' ', '')}.xlsx",
            f"{state.replace(' ', '').lower()}.xlsx"
        ]
        has_file = any(os.path.exists(os.path.join(base_dir, v)) for v in variants)
        state_status.append({"name": state, "has_file": has_file})
        
    return render_template('index.html', states=state_status)

@app.route('/start-agent', methods=['POST'])
def start_agent():
    global IS_AGENT_RUNNING, AGENT_LOGS
    if IS_AGENT_RUNNING:
        return jsonify({"status": "error", "message": "Agent is already running."}), 400
    
    AGENT_LOGS = [f"[{datetime.now().strftime('%H:%M:%S')}] Agent started manually."]
    threading.Thread(target=agent_worker, daemon=True).start()
    return jsonify({"status": "success", "message": "Agent started."})

@app.route('/get-status', methods=['GET'])
def get_status():
    status = {}
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for state in STATES:
        variants = [
            f"{state}.xlsx",
            f"{state.lower()}.xlsx",
            f"{state.replace(' ', '')}.xlsx",
            f"{state.replace(' ', '').lower()}.xlsx"
        ]
        has_file = any(os.path.exists(os.path.join(base_dir, v)) for v in variants)
        status[state] = has_file
    return jsonify(status)

@app.route('/get-progress', methods=['GET'])
def get_progress():
    return jsonify({
        "current_task": CURRENT_PROCESSING_STATE,
        "is_running": IS_AGENT_RUNNING
    })

@app.route('/get-logs', methods=['GET'])
def get_logs():
    after = request.args.get('after', type=int, default=0)
    current_logs = AGENT_LOGS[after:]
    return jsonify({
        "logs": current_logs,
        "next_index": len(AGENT_LOGS),
        "is_running": IS_AGENT_RUNNING
    })

@app.route('/get-state-data/<state_name>', methods=['GET'])
def get_state_data(state_name):
    try:
        filename = None
        variants = [
            f"{state_name}.xlsx",
            f"{state_name.lower()}.xlsx",
            f"{state_name.replace(' ', '')}.xlsx",
            f"{state_name.replace(' ', '').lower()}.xlsx"
        ]
        for v in variants:
            if os.path.exists(v):
                filename = v
                break
                
        if not filename:
            return jsonify({"status": "error", "message": "File not found."}), 404
            
        wb = openpyxl.load_workbook(filename, data_only=True)
        sheet = wb.active
        data = []
        for row in sheet.iter_rows(values_only=True):
            if any(cell is not None for cell in row):
                data.append([str(c) if c is not None else "" for c in row])
        return jsonify({"status": "success", "data": data})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
