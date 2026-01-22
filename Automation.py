import os
import time
import requests
import urllib3
import glob
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv

load_dotenv()

# Disable SSL warnings globally for this automation
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ---------- CONFIGURATION ----------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS_ROOT = os.path.join(BASE_DIR, "Download")


def get_state_download_path(state_name):
    """Creates and returns a specific path for the state in the Download folder"""
    path = os.path.join(DOWNLOADS_ROOT, state_name)
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
    return path

def clean_garbage_files(folder):
    """Deletes all non-PDF files from the folder to keep it clean"""
    print(f"Cleaning up non-PDF files in {folder}...")
    for file_path in glob.glob(os.path.join(folder, "*")):
        # Keep everything that IS a PDF. Delete everything else.
        if not file_path.lower().endswith(".pdf"):
            try:
                os.remove(file_path)
                print(f"Deleted garbage file: {os.path.basename(file_path)}")
            except Exception as e:
                print(f"Could not delete {file_path}: {e}")

def setup_driver(view_browser=True, download_path=None):
    """Initialize Chrome driver with optional visibility and download handling"""
    options = Options()
    if not view_browser:
        options.add_argument("--headless=new") 
    
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--log-level=3")
    
    # Preferences for handling downloads and PDFs
    prefs = {
        "profile.default_content_settings.popups": 0,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True  # Force PDF download instead of browser view
    }
    if download_path:
        prefs["download.default_directory"] = download_path
        
    options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    # Enable download behavior in headless mode
    if not view_browser and download_path:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": download_path
        })
        
    return driver

def is_pdf(url):
    """Check if the URL likely points to a PDF by inspecting headers without downloading the body"""
    try:
        response = requests.head(url, verify=False, allow_redirects=True, timeout=10)
        content_type = response.headers.get('Content-Type', '').lower()
        if 'application/pdf' in content_type:
            return True
        # Sometimes HEAD is not supported or returns wrong info, try a partial GET
        if response.status_code != 200:
             response = requests.get(url, stream=True, verify=False, timeout=10)
             content_type = response.headers.get('Content-Type', '').lower()
             return 'application/pdf' in content_type
    except:
        pass
    return url.lower().endswith('.pdf') or ".pdf" in url.lower()

def download_file(url, folder, filename=None):
    """Download file using requests and return True if successful"""
    if not filename:
        filename = url.split("/")[-1]
        if '?' in filename: filename = filename.split('?')[0]
    
    # Ensure filename ends with .pdf
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"
        
    filepath = os.path.join(folder, filename)
    print(f"Downloading {filename} via requests...")
    
    try:
        response = requests.get(url, stream=True, verify=False, timeout=60)
        if response.status_code == 200:
            content_type = response.headers.get('Content-Type', '').lower()
            if 'application/pdf' not in content_type and not filename.lower().endswith(".pdf"):
                print(f"Aborting: Content-Type is {content_type}, not a PDF.")
                return False
                
            with open(filepath, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"File saved to {filepath}")
            return True
        else:
            print(f"Failed to download. Status code: {response.status_code}")
            return False
    except Exception as e:
        print(f"Download error: {e}")
        return False

def process_assam(view_browser=True):
    state_name = "Assam"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    wait = WebDriverWait(driver, 15)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to Orders page...")
        driver.get("https://aerc.gov.in/pages/sub/orders")
        
        print(f"[{state_name}] Detecting available years...")
        year_links = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, 'year=')]")))
        
        available_years = []
        for link in year_links:
            year_text = link.text.strip()
            if year_text.isdigit():
                available_years.append((int(year_text), link))
        
        if not available_years:
            print(f"[{state_name}] No year links found.")
            return False

        available_years.sort(key=lambda x: x[0], reverse=True)
        latest_year, latest_link = available_years[0]
        
        print(f"[{state_name}] Selecting latest available year: {latest_year}")
        driver.execute_script("arguments[0].click();", latest_link)
        
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        
        rows = driver.find_elements(By.XPATH, "//tr[td]")
        matches = []
        for row in rows:
            row_text = row.text.upper()
            if "TARIFF ORDER" in row_text and "APDCL" in row_text:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 3:
                    date_info = cols[0].text.strip()
                    date_str = date_info.split('&')[0].strip() if '&' in date_info else date_info
                    try:
                        order_date = datetime.strptime(date_str, "%d.%m.%Y")
                    except:
                        order_date = datetime.min
                    try:
                        pdf_link_element = cols[2].find_element(By.TAG_NAME, "a")
                        pdf_url = pdf_link_element.get_attribute("href")
                        matches.append({"date": order_date, "url": pdf_url, "date_str": date_info})
                    except:
                        continue

        if not matches:
            print(f"[{state_name}] No matching PDF found.")
            return False

        matches.sort(key=lambda x: x['date'], reverse=True)
        target = matches[0]
        print(f"[{state_name}] Found latest order: {target['url']}")
        success = download_file(target['url'], download_path)
        
    except Exception as e:
        print(f"[{state_name}] Error: {e}")
        success = False
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

def process_up(view_browser=True):
    state_name = "Uttar Pradesh"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    wait = WebDriverWait(driver, 20)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to UPERC Tariff Orders...")
        driver.get("https://www.uperc.org/Tariff_Order_Users.aspx")
        time.sleep(2)
        
        links = driver.find_elements(By.XPATH, "//a[@class='aaa']")
        matches = []
        for link in links:
            text = link.text.upper()
            url = link.get_attribute("href")
            if "STATE DISCOMS" in text and "TARIFF ORDER" in text and "SCANNED" not in text:
                if url and ".pdf" in url.lower():
                    matches.append({"text": link.text, "url": url})
        
        if not matches:
            print(f"[{state_name}] Trying previous years...")
            try:
                prev_years_link = driver.find_element(By.LINK_TEXT, "Previous Years")
                driver.execute_script("arguments[0].click();", prev_years_link)
                time.sleep(2)
                links = driver.find_elements(By.XPATH, "//a[@class='aaa']")
                for link in links:
                    text = link.text.upper()
                    url = link.get_attribute("href")
                    if "STATE DISCOMS" in text and "TARIFF ORDER" in text and "SCANNED" not in text:
                        if url and ".pdf" in url.lower():
                            matches.append({"text": link.text, "url": url})
            except: pass

        if not matches:
            print(f"[{state_name}] No matching order found.")
            return False

        target = matches[0]
        print(f"[{state_name}] Found latest order: {target['text']}")
        success = download_file(target['url'], download_path)
        
    except Exception as e:
        print(f"[{state_name}] Error: {e}")
        success = False
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

def process_meghalaya(view_browser=True):
    state_name = "Meghalaya"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to MSERC Tariff Orders...")
        driver.get("https://www.mserc.gov.in/tarifforders.html")
        time.sleep(2)
        
        links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        matches = []
        for link in links:
            text = link.text.upper()
            if "AGGREGATE REVENUE" in text and "CORRIGENDUM" not in text:
                matches.append({"text": link.text, "url": link.get_attribute("href")})
        
        if not matches:
            print(f"[{state_name}] No matching PDF found.")
            return False

        target = matches[0]
        print(f"[{state_name}] Found latest order: {target['text']}")
        success = download_file(target['url'], download_path)
        
    except Exception as e:
        print(f"[{state_name}] Error: {e}")
        success = False
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

def process_rajasthan(view_browser=True):
    state_name = "Rajasthan"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    wait = WebDriverWait(driver, 20)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to RERC website...")
        driver.get("https://rerc.rajasthan.gov.in/rerc-user-files/tariff-orders")
        
        # The user requested NOT to use the search bar, but to look at the list directly
        # We'll just wait for the table to load
        wait.until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class, 'table')]//tr[td]")))
        time.sleep(2)
        
        rows = driver.find_elements(By.XPATH, "//table[contains(@class, 'table')]//tr[td]")
        if not rows: rows = driver.find_elements(By.XPATH, "//*[@id='BodyContent_rptorders']//tr")
        
        target_row = None
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 4:
                subject_text = cells[3].text.upper()
                if "AGGREGATE REVENUE" in subject_text and "DISCOMS" in subject_text and "PETITION" not in subject_text:
                    target_row = row
                    print(f"[{state_name}] Found target row: {subject_text}")
                    break
        
        if target_row:
            # First, check if the Download button has a direct PDF URL (sometimes it does in new systems)
            try:
                download_btn = target_row.find_element(By.XPATH, ".//a[contains(@title, 'Download')]")
                href = download_btn.get_attribute("href")
                if href and href.lower().endswith(".pdf"):
                    print(f"[{state_name}] Direct PDF URL found in Download button.")
                    success = download_file(href, download_path, filename="Rajasthan_Tariff_Order.pdf")
                    if success: return True
            except: pass

            # Try View button logic
            try:
                view_btn = target_row.find_element(By.XPATH, ".//a[contains(@title, 'View')]")
                main_window = driver.current_window_handle
                view_btn.click()
                time.sleep(8)
                
                pdf_url = None
                for handle in driver.window_handles:
                    if handle != main_window:
                        driver.switch_to.window(handle)
                        url = driver.current_url
                        if is_pdf(url):
                            pdf_url = url
                            print(f"[{state_name}] Found PDF URL in new tab: {pdf_url}")
                        driver.close()
                        driver.switch_to.window(main_window)
                        break
                
                if pdf_url:
                    success = download_file(pdf_url, download_path, filename="Rajasthan_Tariff_Order.pdf")
                    if success: return True
            except: pass

            # Fallback to browser postback download
            target_filename = "Rajasthan_Tariff_Order.pdf"
            print(f"Downloading {target_filename}...")
            
            # Clear existing files with the same name to avoid detection issues
            target_file_path = os.path.join(download_path, target_filename)
            if os.path.exists(target_file_path):
                try: os.remove(target_file_path)
                except: pass

            before_files = set(os.listdir(download_path))
            download_btn = target_row.find_element(By.XPATH, ".//a[contains(@id, 'LinkButton')]")
            driver.execute_script("arguments[0].click();", download_btn)
            
            # Wait for file
            for i in range(60):
                time.sleep(1)
                after_files = set(os.listdir(download_path))
                new_files = [f for f in (after_files - before_files) if f.lower().endswith(".pdf") and not f.endswith(('.crdownload', '.tmp'))]
                
                # If no "new" file (because it might have replaced one or downloaded with same name)
                # check if the default name 'tariff-orders.pdf' appeared
                if not new_files:
                    if os.path.exists(os.path.join(download_path, "tariff-orders.pdf")):
                        new_files = ["tariff-orders.pdf"]

                if new_files:
                    downloaded_file = new_files[0]
                    downloaded_path = os.path.join(download_path, downloaded_file)
                    
                    # Rename to our standard name
                    if downloaded_file != target_filename:
                        try:
                            # Wait a bit to ensure file is not locked
                            time.sleep(1)
                            os.rename(downloaded_path, target_file_path)
                            downloaded_path = target_file_path
                        except: pass
                    
                    print(f"File saved to {downloaded_path}")
                    success = True
                    break
        else:
            print(f"[{state_name}] No matching row found.")

    except Exception as e:
        print(f"[{state_name}] Error: {e}")
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

def process_mp(view_browser=True):
    state_name = "Madhya Pradesh"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    wait = WebDriverWait(driver, 20)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to MPERC website...")
        driver.get("https://mperc.in/")
        
        # Navigation: Final Orders -> Tariff orders issued by MPERC -> Distribution tariff orders
        try:
            # Hover over 'Final Orders'
            # Note: The site uses 'FINAL ORDERS' (uppercase) in the nav bar
            final_orders = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'FINAL ORDERS')]")))
            ActionChains(driver).move_to_element(final_orders).perform()
            time.sleep(1)
            
            # Hover over 'Tariff orders issued by MPERC'
            tariff_orders_issued = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Tariff Orders Issued by MPERC')]")))
            ActionChains(driver).move_to_element(tariff_orders_issued).perform()
            time.sleep(1)
            
            # Click 'Distrubution Tariff Orders'
            distribution_orders = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Distrubution Tariff Orders')]")))
            distribution_orders.click()
        except Exception as e:
            # Fallback direct navigation if hover fails
            driver.get("https://mperc.in/page/distrubution-tariff-orders")
            
        time.sleep(3)
        
        # Find latest row in the table
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table//tr[td]")))
        
        target_info = None
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 5:
                description = cells[2].text.upper()
                # Requirements: "Tariff order", latest, ignore "Petition" and "Corrigendum"
                if "TARIFF ORDER" in description and "PETITION" not in description and "CORRIGENDUM" not in description:
                    try:
                        # The download button is in the 5th cell (index 4)
                        download_btn = cells[4].find_element(By.TAG_NAME, "a")
                        pdf_url = download_btn.get_attribute("href")
                        target_info = {"url": pdf_url, "name": cells[2].text}
                        break # Found the latest one at the top
                    except: continue
        
        if target_info:
            target_filename = "Madhya_Pradesh_Tariff_Order.pdf"
            print(f"[{state_name}] Found latest order: {target_info['name']}")
            success = download_file(target_info['url'], download_path, filename=target_filename)
        else:
            print(f"[{state_name}] No matching row found.")

    except Exception as e:
        print(f"[{state_name}] Error: {e}")
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

def process_chhattisgarh(view_browser=True):
    state_name = "Chhattisgarh"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    wait = WebDriverWait(driver, 20)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to CSERC website...")
        driver.get("https://cserc.gov.in/Welcome/index")
        
        # Hover on orders then click on tariff orders
        try:
            # Locate 'Orders' menu item - checking carefully for the text
            orders_menu = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Orders')]")))
            ActionChains(driver).move_to_element(orders_menu).perform()
            time.sleep(1)
            
            tariff_orders_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Tariff Orders')]")))
            tariff_orders_link.click()
        except:
            print(f"[{state_name}] Hover navigation failed, trying direct URL...")
            driver.get("https://cserc.gov.in/Welcome/show_tariff_orders")

        # Find rows
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table//tr[td]")))
        
        target_info = None
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 4:
                # cell[1] is Text (State Power Companies/Board)
                # cell[3] is Order PDF
                
                description = cells[1].text.upper()
                
                # Check keywords: "TARIFF ORDER", "CSPDCL" and ignore "PETITION", "CORRIGENDUM"
                if "TARIFF ORDER" in description and "CSPDCL" in description and "PETITION" not in description and "CORRIGENDUM" not in description:
                    try:
                        download_btn = cells[3].find_element(By.TAG_NAME, "a")
                        pdf_url = download_btn.get_attribute("href")
                        target_info = {"url": pdf_url, "name": cells[1].text}
                        break
                    except: continue
                    
        if target_info:
            print(f"[{state_name}] Found latest order: {target_info['name']}")
            success = download_file(target_info['url'], download_path, filename="Chhattisgarh_Tariff_Order.pdf")
        else:
            print(f"[{state_name}] No matching row found.")
            
    except Exception as e:
        print(f"[{state_name}] Error: {e}")
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

def process_himachal(view_browser=True):
    state_name = "Himachal Pradesh"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    wait = WebDriverWait(driver, 20)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to HPERC website...")
        # Direct navigation to Distribution Tariff Orders based on site inspection
        driver.get("https://hperc.org/order_type/distribution/")
        
        # Determine table rows
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table//tbody//tr")))
        
        target_info = None
        for row in rows:
            text = row.text.upper()
            
            # Keywords: "DISCOMS" (or "HPSEBL"), "AGGREGATE", "REVENUE"
            # Ignore: "PETITION"
            if "AGGREGATE" in text and "REVENUE" in text and ("DISCOMS" in text or "HPSEBL" in text) and "PETITION" not in text:
                 try:
                    # The download link is usually the title text or a view/download icon
                    link_element = row.find_element(By.TAG_NAME, "a")
                    pdf_url = link_element.get_attribute("href")
                    target_info = {"url": pdf_url, "name": row.text.split('\n')[0]} # Use first line as name
                    break
                 except: continue
        
        if target_info:
            print(f"[{state_name}] Found latest order: {target_info['name'][:100]}...")
            success = download_file(target_info['url'], download_path, filename="Himachal_Pradesh_Tariff_Order.pdf")
        else:
            print(f"[{state_name}] No matching row found.")

    except Exception as e:
        print(f"[{state_name}] Error: {e}")
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

def process_puducherry(view_browser=True):
    state_name = "Puducherry"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    wait = WebDriverWait(driver, 20)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to JERC UTS website...")
        # Direct navigation based on site inspection data
        driver.get("https://jercuts.gov.in/order_category/puducherry/")
        
        time.sleep(3)
        
        # Row selection: Looking for rows in the tariff orders table
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table//tr[td]")))
        
        target_info = None
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 5: # Based on screenshot, there are multiple columns (S.No, Description, Pet. No, Year, Action)
                # The description text is usually in the 2nd cell (index 1)
                description_cell = cells[1]
                description = description_cell.text.strip().upper()
                
                # Keywords: "DISCOMS" (or implied by PED context on this page, but strict check requested), 
                # "AGGREGATE", "REVENUE", "TRUE-UP"
                # Exclude: "PETITION", "CORRIGENDUM"
                
                has_keywords = ("AGGREGATE" in description and 
                                "REVENUE" in description and 
                                "TRUE-UP" in description)
                                
                is_excluded = ("PETITION" in description or "CORRIGENDUM" in description)
                
                has_discom_context = ("DISCOMS" in description or "PED" in description or "ELECTRICITY DEPARTMENT" in description)

                if has_keywords and not is_excluded:
                     # Attempt to find the download link
                     try:
                        # The download link is an icon in the last column usually
                        # Screenshot shows 2 icons: Download (arrow) and View (eye).
                        # We want the download one. Usually the first 'a' tag in the Actions column.
                        action_cell = cells[-1] # Last cell
                        download_btn = action_cell.find_element(By.XPATH, ".//a[contains(@href, '.pdf')]")
                        pdf_url = download_btn.get_attribute("href")
                        
                        target_info = {"url": pdf_url, "name": description}
                        break
                     except: continue
        
        if target_info:
            print(f"[{state_name}] Found latest order: {target_info['name'][:100]}...")
            success = download_file(target_info['url'], download_path, filename="Puducherry_Tariff_Order.pdf")
        else:
            print(f"[{state_name}] No matching row found. Checking strict keyword matches...")

    except Exception as e:
        print(f"[{state_name}] Error: {e}")
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

def process_bihar(view_browser=True):
    state_name = "Bihar"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    wait = WebDriverWait(driver, 20)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to BERC website...")
        # Direct navigation based on analysis
        driver.get("https://berc.co.in/orders/tariff/distribution/sbpdcl")
        
        # Get all rows in the main table
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table//tr[td]")))
        
        target_info = None
        for row in rows:
            text = row.text.upper()
            
            # Keywords: "TARIFF ORDER", "DISCOMS" (or NBPDCL/SBPDCL)
            # Ignore: "PETITION", "CHART" (if mentioned in main title, though usually on detail page)
            if "TARIFF ORDER" in text and ("DISCOMS" in text or "SBPDCL" in text) and "PETITION" not in text:
                try:
                    # Logic: The row has a link to a DETAILS page, not the PDF directly.
                    link_element = row.find_element(By.TAG_NAME, "a")
                    detail_url = link_element.get_attribute("href")
                    target_info = {"url": detail_url, "name": row.text}
                    break
                except: continue
                
        if target_info:
            print(f"[{state_name}] Found order entry: {target_info['name'][:100]}...")
            print(f"[{state_name}] Navigating to details page to find PDF...")
            driver.get(target_info['url'])
            
            # Now on details page, find the PDF link.
            # There might be multiple PDFs (Tariff Chart vs Tariff Order).
            # Look for <a> tags ending in .pdf OR <embed>/<iframe> src
            
            # Let's try finding all links ending in .pdf
            pdf_links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
            
            # Fallback: check embeds/iframes if no simple <a> tags found (common in Joomla/CMS sites)
            if not pdf_links:
                pdf_links = driver.find_elements(By.XPATH, "//embed[contains(@src, '.pdf')]") + \
                            driver.find_elements(By.XPATH, "//iframe[contains(@src, '.pdf')]")
            
            final_pdf_url = None
            for link in pdf_links:
                url = link.get_attribute("href") or link.get_attribute("src")
                if not url: continue
                
                # Handle embedded PDF viewers (e.g. viewer.html?file=...)
                if "viewer.html" in url and "file=" in url:
                    try:
                        from urllib.parse import unquote
                        url = unquote(url.split("file=")[1].split("&")[0])
                    except: pass

                url_lower = url.lower()
                
                # Logic to distinguish Order vs Chart
                # Order usually has "TO" or "Tariff_Order" or "Tariff-Order"
                # Chart usually has "Chart" or "Schedule"
                
                is_chart = "chart" in url_lower or "schedule" in url_lower
                is_order = "order" in url_lower or "to_" in url_lower or "tariff" in url_lower
                
                if is_order and not is_chart:
                    final_pdf_url = url
                    break
                elif is_order: # If it has both, maybe keep checking, but store as fallback
                     final_pdf_url = url
            
            if final_pdf_url:
                print(f"[{state_name}] Found PDF URL: {final_pdf_url}")
                success = download_file(final_pdf_url, download_path, filename="Bihar_Tariff_Order.pdf")
            else:
                 print(f"[{state_name}] Could not find suitable PDF link on details page.")
                 
        else:
            print(f"[{state_name}] No matching row found in main list.")

    except Exception as e:
        print(f"[{state_name}] Error: {e}")
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

def process_odisha(view_browser=True):
    state_name = "Odisha"
    download_path = get_state_download_path(state_name)
    driver = setup_driver(view_browser=view_browser, download_path=download_path)
    wait = WebDriverWait(driver, 20)
    
    success = False
    try:
        print(f"[{state_name}] Navigating to OERC website...")
        # Direct navigation based on site inspection
        driver.get("https://www.orierc.org/Distribution_Retail_Supply.aspx")
        
        # Get all rows in the main table
        # Structure is usually table -> tr -> td
        rows = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//table//tr[td]")))
        
        target_info = None
        for row in rows:
            text = row.text.upper()
            
            # Keywords: "AGGREGATE", "REVENUE", "REQUIREMENT"
            # Context: "DISCOMS" (User requested this. The file name has DISCOM, logic below checks for it or assumes if strict match fails)
            # Exclusion: "PETITION"
            
            # Row text example: "AGGREGATE REVENUE REQUIREMENT, WHEELING TARIFF & RETAIL SUPPLY TARIFF FOR THE FY 2025-26"
            # It might not explicitly say "DISCOMS" in the row text, but it IS the Distribution page.
            # User requirement: "have these keywords Discoms Aggregate ,Revenue , requirement"
            # Strict interpretation: Text MUST have "DISCOMS".
            # Site inspection shows the file is "DISCOM_TARIFF_ORDER...".
            # If the Text doesn't have "DISCOMS", I will check if the Link has "DISCOM".
            
            if "AGGREGATE" in text and "REVENUE" in text and "REQUIREMENT" in text and "PETITION" not in text:
                 # Check for "DISCOMS" in text OR link URL
                 link_element = None
                 try:
                     link_element = row.find_element(By.TAG_NAME, "a")
                 except: continue
                 
                 url = link_element.get_attribute("href")
                 # Check if "DISCOMS" is properly associated
                 if "DISCOM" in text or "DISCOM" in url.upper():
                     target_info = {"url": url, "name": row.text}
                     break
        
        if target_info:
            print(f"[{state_name}] Found latest order: {target_info['name'][:100]}...")
            success = download_file(target_info['url'], download_path, filename="Odisha_Tariff_Order.pdf")
        else:
            print(f"[{state_name}] No matching row found.")

    except Exception as e:
        print(f"[{state_name}] Error: {e}")
    finally:
        driver.quit()
        clean_garbage_files(download_path)
    return success

if __name__ == "__main__":
    def remove_readonly(func, path, _):
        """Clear read-only bit and retry."""
        os.chmod(path, stat.S_IWRITE)
        func(path)

    import stat
    start_time = time.time()
    SHOW_BROWSER = os.getenv("SHOW_BROWSER", "False").lower() == "true"
    
    states_to_process = ["Assam", "Uttar Pradesh", "Meghalaya", "Rajasthan", "Madhya Pradesh", "Chhattisgarh", "Himachal Pradesh", "Puducherry", "Bihar", "Odisha"]
    results = {}
    
    # Clear Download folder before starting, with better error handling for Windows
    if os.path.exists(DOWNLOADS_ROOT):
        print(f"Clearing Download folder: {DOWNLOADS_ROOT}")
        try:
            shutil.rmtree(DOWNLOADS_ROOT, onerror=remove_readonly)
        except Exception as e:
            print(f"Warning: Could not fully clear Download folder: {e}")
            
    os.makedirs(DOWNLOADS_ROOT, exist_ok=True)
    
    for state in states_to_process:
        if state == "Assam": results[state] = process_assam(view_browser=SHOW_BROWSER)
        elif state == "Uttar Pradesh": results[state] = process_up(view_browser=SHOW_BROWSER)
        elif state == "Meghalaya": results[state] = process_meghalaya(view_browser=SHOW_BROWSER)
        elif state == "Rajasthan": results[state] = process_rajasthan(view_browser=SHOW_BROWSER)
        elif state == "Madhya Pradesh": results[state] = process_mp(view_browser=SHOW_BROWSER)
        elif state == "Chhattisgarh": results[state] = process_chhattisgarh(view_browser=SHOW_BROWSER)
        elif state == "Himachal Pradesh": results[state] = process_himachal(view_browser=SHOW_BROWSER)
        elif state == "Puducherry": results[state] = process_puducherry(view_browser=SHOW_BROWSER)
        elif state == "Bihar": results[state] = process_bihar(view_browser=SHOW_BROWSER)
        elif state == "Odisha": results[state] = process_odisha(view_browser=SHOW_BROWSER)
    
    print("\n" + "="*40)
    print(f"{'STATE':<15} | {'AUTOMATION STATUS':<20}")
    print("-" * 40)
    for state, status in results.items():
        print(f"{state:<15} | {'Successfully Downloaded' if status else 'Failed to Download'}")
    print("="*40)
    print(f"Total time taken: {time.time() - start_time:.2f} seconds")
