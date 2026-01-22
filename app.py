import os
import shutil
import stat
import subprocess
import threading
import openpyxl
from datetime import datetime
from flask import Flask, render_template, jsonify, request
from dotenv import load_dotenv

def delete_folder_contents(folder_path):
    def remove_readonly(func, path, _):
        os.chmod(path, stat.S_IWRITE)
        func(path)

    if os.path.exists(folder_path):
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.chmod(file_path, stat.S_IWRITE)
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path, onerror=remove_readonly)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

# Run cleanup at Flask startup
base_dir = os.path.dirname(os.path.abspath(__file__))
for folder in ["Extraction", "Download", "ists_pdf", "ists_charge_pdf", "ists_extracted"]:
    folder_path = os.path.join(base_dir, folder)
    if os.path.exists(folder_path):
        delete_folder_contents(folder_path)

load_dotenv()

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
        # Using the virtual environment's python if it exists, otherwise fallback to "python"
        python_exe = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".venv", "Scripts", "python.exe")
        if not os.path.exists(python_exe):
            python_exe = "python"

        # Using -u for unbuffered output to ensure real-time logs in the monitor
        process = subprocess.Popen(
            [python_exe, "-u", script_name],
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
        ("clear_excels.py", "Clearing Excels"),
        ("Automation.py", "Automation"),
        ("Auomation_ists.py", "ISTS Automation"),
        ("scraper.py", "Scraping"),
        ("ists.py", "ISTS"),
        ("chhattisgarh.py", "Chhattisgarh"),
        ("Meghalaya.py", "Meghalaya"),
        ("Rajasthan.py", "Rajasthan"),
        ("Madyapradesh.py", "Madhya Pradesh"),
        ("bihar.py", "Bihar"),
        ("puducherry.py", "Puducherry"),
        ("Himachalpradesh.py", "Himachal Pradesh"),
        ("Assam.py", "Assam"),
        ("uttarpradesh.py", "Uttar Pradesh")
    ]
    
    for script, display in scripts:
        # If we are about to start Scraping, clean previous extraction data
        if script == "scraper.py":
            AGENT_LOGS.append(f"[{datetime.now().strftime('%H:%M:%S')}] Cleaning previous extracted data...")
            for folder in ["Extraction", "ists_extracted"]:
                folder_path = os.path.join(base_dir, folder)
                if os.path.exists(folder_path):
                    delete_folder_contents(folder_path)

        run_script(script, display)
        # Sync to DB if it's a state script
        if script not in ["Automation.py", "scraper.py", "ists.py", "Auomation_ists.py"]:
            state_name = display
            excel_variants = [f"{state_name}.xlsx", f"{state_name.lower()}.xlsx", f"{state_name.replace(' ', '')}.xlsx"]
            for v in excel_variants:
                if os.path.exists(v):
                    try:
                        from database.database_utils import sync_excel_to_db
                        sync_excel_to_db(v)
                        AGENT_LOGS.append(f"[{datetime.now().strftime('%H:%M:%S')}] {v} synced to SQLite database.")
                        break
                    except Exception as e:
                        AGENT_LOGS.append(f"[{datetime.now().strftime('%H:%M:%S')}] DB Sync Error: {str(e)}")
        
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

@app.route('/get-db-data', methods=['GET'])
def get_db_data():
    try:
        import sqlite3
        db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tariff_orders.db")
        if not os.path.exists(db_path):
            return jsonify({"status": "error", "message": "Database not found."}), 404
            
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM tariff_data")
        rows = cursor.fetchall()
        
        data = [dict(row) for row in rows]
        conn.close()
        return jsonify({"status": "success", "data": data})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv("PORT", 5000))
    debug = os.getenv("FLASK_DEBUG", "True").lower() == "true"
    app.run(debug=debug, port=port)
