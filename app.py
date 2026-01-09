import os
import subprocess
import threading
import openpyxl
from datetime import datetime, date
from flask import Flask, render_template, jsonify

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

def get_today_update_count():
    count = 0
    today = date.today()
    for state in STATES:
        # Check specific state excel file
        # Note: mapping.py creates files like "Meghalaya.xlsx"
        filename = f"{state}.xlsx"
        if os.path.exists(filename):
            try:
                mtime = os.path.getmtime(filename)
                file_date = datetime.fromtimestamp(mtime).date()
                if file_date == today:
                    count += 1
            except: pass
    return count

@app.route('/')
def index():
    updates = 0
    state_status = []
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for state in STATES:
        excel_path = os.path.join(base_dir, f"{state}.xlsx")
        has_file = os.path.exists(excel_path)
        state_status.append({"name": state, "has_file": has_file})
    return render_template('index.html', states=state_status, updated_count=updates)

# Global variable to track current processing state
# Values: None, "Scraping", "ISTS", or State Name (e.g. "Meghalaya")
CURRENT_PROCESSING_STATE = None

@app.route('/start-agent', methods=['POST'])
def start_agent():
    global CURRENT_PROCESSING_STATE
    try:
        # Run scripts in sequence
        
        print("Running scraper.py...")
        CURRENT_PROCESSING_STATE = "Scraping"
        subprocess.run(["python", "scraper.py"], check=True)
        
        print("Running ists.py...")
        CURRENT_PROCESSING_STATE = "ISTS"
        subprocess.run(["python", "ists.py"], check=True)
        
        # Sequentially run state processors
        print("Running Meghalaya.py...")
        CURRENT_PROCESSING_STATE = "Meghalaya"
        try:
            subprocess.run(["python", "Meghalaya.py"], check=False)
        except Exception as e: print(f"Meghalaya error: {e}")

        print("Running Rajastan.py...")
        CURRENT_PROCESSING_STATE = "Rajasthan" # Clean name matching STATES list if possible
        try:
            subprocess.run(["python", "Rajastan.py"], check=False)
        except Exception as e: print(f"Rajastan error: {e}")

        print("Running Madyapradesh.py...")
        CURRENT_PROCESSING_STATE = "Madhya Pradesh"
        try:
            subprocess.run(["python", "Madyapradesh.py"], check=False)
        except Exception as e: print(f"Madyapradesh error: {e}")

        print("Running bihar.py...")
        CURRENT_PROCESSING_STATE = "Bihar"
        try:
            subprocess.run(["python", "bihar.py"], check=False)
        except Exception as e: print(f"Bihar error: {e}")

        print("Running puducherry.py...")
        CURRENT_PROCESSING_STATE = "Puducherry"
        try:
            subprocess.run(["python", "puducherry.py"], check=False)
        except Exception as e: print(f"Puducherry error: {e}")

        print("Running Assam.py...")
        CURRENT_PROCESSING_STATE = "Assam"
        try:
            subprocess.run(["python", "Assam.py"], check=False)
        except Exception as e: print(f"Assam error: {e}")
        
        CURRENT_PROCESSING_STATE = None # Done
        
        # Get fresh count after run
        new_count = get_today_update_count()
        
        return jsonify({"status": "success", "message": "All agents finished successfully.", "new_count": new_count})
    except subprocess.CalledProcessError as e:
        CURRENT_PROCESSING_STATE = None
        return jsonify({"status": "error", "message": str(e)}), 500
    except Exception as e:
        CURRENT_PROCESSING_STATE = None
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/get-status', methods=['GET'])
def get_all_status():
    # Return simple status for backward compatibility / cached clients
    status = {}
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for state in STATES:
        excel_path = os.path.join(base_dir, f"{state}.xlsx")
        status[state] = os.path.exists(excel_path)
    return jsonify(status)

@app.route('/get-progress', methods=['GET'])
def get_progress():
    global CURRENT_PROCESSING_STATE
    state_files = {}
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for state in STATES:
        excel_path = os.path.join(base_dir, f"{state}.xlsx")
        state_files[state] = {
            "has_file": os.path.exists(excel_path),
            "is_processing": (state == CURRENT_PROCESSING_STATE)
        }
    return jsonify({"states": state_files, "current_task": CURRENT_PROCESSING_STATE})

@app.route('/get-state-data/<state_name>', methods=['GET'])
def get_state_data(state_name):
    # ... (rest of function)
    try:
        # Assuming the excel file is named "{StateName}.xlsx"
        filename = f"{state_name}.xlsx"
        if not os.path.exists(filename):
            return jsonify({"status": "error", "message": f"File {filename} not found. Please run the agent first."}), 404
            
        wb = openpyxl.load_workbook(filename, data_only=True)
        sheet = wb.active
        
        data = []
        for row in sheet.iter_rows(values_only=True):
            # rudimentary check to skip empty rows
            if any(cell is not None for cell in row):
                # Convert None to "" for JSON
                clean_row = [str(cell) if cell is not None else "" for cell in row]
                data.append(clean_row)
                
        return jsonify({"status": "success", "data": data, "state": state_name})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
