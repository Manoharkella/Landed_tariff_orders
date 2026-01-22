
import os
import time
import re
import requests
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from urllib.parse import unquote

def parse_date_from_text(text):
    """
    Look for dates like 'January 2026' or 'January,2026' or '25.11.2025' or '20261901'
    Returns a sortable value.
    """
    # 1. Try YYYYDDMM-
    match_file = re.search(r'(\d{4})(\d{2})(\d{2})-', text)
    if match_file:
        year, day, month = match_file.groups()
        try:
            return datetime(int(year), int(month), int(day))
        except:
            pass

    # 2. Try DD.MM.YYYY
    match_ymd = re.search(r'(\d{2})[\.\-\/](\d{2})[\.\-\/](\d{4})', text)
    if match_ymd:
        day, month, year = match_ymd.groups()
        try:
            return datetime(int(year), int(month), int(day))
        except:
            pass
            
    # 3. Try Month YYYY
    months = ["january", "february", "march", "april", "may", "june", 
              "july", "august", "september", "october", "november", "december"]
    lower_text = text.lower()
    for i, m in enumerate(months):
        if m in lower_text:
            match_year = re.search(r'(\d{4})', lower_text)
            if match_year:
                year = int(match_year.group(1))
                return datetime(year, i + 1, 1)
    
    return None

def download_latest_pdf(driver, target_url, download_dir, exclude_keyword=None):
    """
    Navigates to URL, finds the newest PDF (non-corrigendum), and downloads it.
    """
    print(f"\n--- Processing: {target_url} ---")
    
    # Ensure directory exists and clear it
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    else:
        print(f"Clearing old files in {download_dir}...")
        for filename in os.listdir(download_dir):
            file_path = os.path.join(download_dir, filename)
            try:
                if os.path.isfile(file_path): os.unlink(file_path)
                elif os.path.isdir(file_path): shutil.rmtree(file_path)
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")

    try:
        driver.get(target_url)
        print("Waiting for page load...")
        time.sleep(10)
        
        rows = driver.find_elements(By.TAG_NAME, "tr")
        if not rows:
            # Fallback for pages without standard tables
            rows = driver.find_elements(By.TAG_NAME, "a")

        candidates = []
        
        for row in rows:
            try:
                # Get links and text
                if row.tag_name == 'a':
                    href = row.get_attribute("href")
                    text = row.text.strip()
                else:
                    link_elements = row.find_elements(By.TAG_NAME, "a")
                    href = None
                    for a in link_elements:
                        h = a.get_attribute("href")
                        if h and "pdf" in h.lower():
                            href = h
                            break
                    text = row.text.strip()
                
                if not href or "pdf" not in href.lower():
                    continue
                
                filename = href.split("/")[-1]
                
                # Filter out excluded keywords (like Corrigendum)
                if exclude_keyword and (exclude_keyword.lower() in text.lower() or exclude_keyword.lower() in filename.lower()):
                    continue
                
                # Extract date for sorting
                date_val = parse_date_from_text(text)
                if not date_val:
                    date_val = parse_date_from_text(filename)
                
                candidates.append({
                    'url': href,
                    'text': text,
                    'date': date_val,
                    'filename': filename
                })
            except:
                continue
                
        if not candidates:
            print("No valid PDF links found.")
            return

        # Sort by date (newest first)
        candidates.sort(key=lambda x: (x['date'] is not None, x['date']), reverse=True)
        
        selected = candidates[0]
        print(f"Selected newest: {selected['text'][:80]}...")
        
        # Download
        requests.packages.urllib3.disable_warnings()
        resp = requests.get(selected['url'], stream=True, verify=False)
        if resp.status_code == 200:
            clean_name = unquote(selected['filename']).split("?")[0]
            if not clean_name.lower().endswith(".pdf"):
                clean_name += ".pdf"
                
            out_path = os.path.join(download_dir, clean_name)
            with open(out_path, 'wb') as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"Successfully downloaded to: {out_path}")
        else:
            print(f"Failed to download. Status: {resp.status_code}")

    except Exception as e:
        print(f"Error during processing {target_url}: {e}")

def main():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new") 
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--log-level=3")

    driver = webdriver.Chrome(options=chrome_options)
    try:
        # 1. Download Transmission Losses -> ists_pdf
        download_latest_pdf(
            driver, 
            "https://grid-india.in/en/markets/transmission-losses", 
            r"c:\Users\hi\OneDrive\Desktop\demo\ists_pdf"
        )
        
        # 2. Download Transmission Charges -> ists_charge_pdf
        download_latest_pdf(
            driver, 
            "https://grid-india.in/en/markets/notification-of-transmission-charges-for-the-dics", 
            r"c:\Users\hi\OneDrive\Desktop\demo\ists_charge_pdf",
            exclude_keyword="corrigendum"
        )

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
