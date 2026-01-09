# SolarTariff AI Sentinel ⚡

SolarTariff AI Sentinel is an agentic automation system designed to extract, process, and monitor Landed Tariff data across various Indian states. It automatically scrapes PDF tariff orders, extracts complex data into structured formats, and provides a real-time monitoring dashboard.

## 🚀 Getting Started

### 1. Prerequisites
Ensure you have **Python 3.10 or higher** installed on your system.

### 2. Installation
Open your terminal in the project root directory and install the required dependencies:

```bash
pip install flask openpyxl pdfplumber requests
```

## 🛠️ How to Run

### **Step 1: Start the Dashboard**
Open a terminal and run the Flask server:

```powershell
python app.py
```
*The dashboard will be available at `http://127.0.0.1:5000`*

### **Step 2: Start the Automation Agent**
1. Open your browser and navigate to `http://127.0.0.1:5000`.
2. Click the **"Start Agent"** button in the sidebar.
3. This will trigger the following automated sequence:
   - **Scraper**: Processes PDFs from the `Download/` folder.
   - **ISTS Logic**: Extracted All India transmission losses.
   - **State Processors**: Updates individual state Excel files (`Assam.xlsx`, `Rajasthan.xlsx`, etc.) with extracted charges, losses, and rebates.

### **Step 3: View Results**
Once a state card on the dashboard turns **Green**, click on it to view the live extracted data directly in your browser.

## 📁 Project Structure

- `app.py`: The central Flask application and dashboard orchestrator.
- `scraper.py`: Core logic for extracting tables from PDF files.
- `ists.py`: Utility for fetching transmission loss data.
- `{State}.py`: State-specific logic for mapping extracted data to the final Excel format.
- `templates/`: HTML/CSS for the web dashboard.
- `.gitignore`: Configured to keep the repository clean from caches and bulky data.

## 🔍 Data Source
The system expects PDF files to be placed in the `Download/{StateName}/` directories. The extracted data is stored in `Extraction/` before being mapped to the final `.xlsx` reports.
