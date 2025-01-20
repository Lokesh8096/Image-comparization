import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import requests
import cv2
from skimage.metrics import structural_similarity as ssim
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
import numpy as np
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os  # To create directories

def validate_url(url):
    parsed_url = urlparse(url)
    return f"https://{url}" if not parsed_url.scheme else url

def download_image_as_bytes(drive_link):
    try:
        file_id = drive_link.split("/d/")[1].split("/")[0] if "drive.google.com" in drive_link else None
        download_url = f"https://drive.google.com/uc?id={file_id}&export=download" if file_id else drive_link
        response = requests.get(download_url, stream=True, timeout=10)
        response.raise_for_status()
        return BytesIO(response.content)
    except Exception as e:
        print(f"Error downloading image: {e}")
        return None

def capture_screenshot_as_bytes(url, mobile_view=False):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=1920,1080")
    if mobile_view:
        chrome_options.add_experimental_option("mobileEmulation", {"deviceName": "iPhone X"})
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    try:
        driver.get(url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        screenshot = driver.get_screenshot_as_png()
        return BytesIO(screenshot)
    except Exception as e:
        print(f"Error capturing screenshot: {e}")
        return None
    finally:
        driver.quit()

def compare_images_bytes(img1_bytes, img2_bytes):
    img1 = cv2.imdecode(np.frombuffer(img1_bytes.getbuffer(), np.uint8), cv2.IMREAD_GRAYSCALE)
    img2 = cv2.imdecode(np.frombuffer(img2_bytes.getbuffer(), np.uint8), cv2.IMREAD_GRAYSCALE)
    if img1 is None or img2 is None:
        return 0.0
    img2_resized = cv2.resize(img2, (img1.shape[1], img1.shape[0]))
    similarity_index, _ = ssim(img1, img2_resized, full=True)
    return round(similarity_index * 100, 2)

def process_google_sheet(sheet_url, credentials_path, local_excel_path):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(creds)

    print("Connecting to Google Sheet...")
    sheet = client.open_by_url(sheet_url)
    worksheet = sheet.get_worksheet(0)
    data = worksheet.get_all_records()
    df = pd.DataFrame(data)

    def process_row(index, row):
        try:
            mobile_link = row['Mobile Response Link']
            desktop_link = row['Desktop Response Link']
            website_url = validate_url(row['Website URL'])

            # Download and capture images in parallel
            with ThreadPoolExecutor() as executor:
                future_mobile_ref = executor.submit(download_image_as_bytes, mobile_link)
                future_desktop_ref = executor.submit(download_image_as_bytes, desktop_link)
                future_mobile_site = executor.submit(capture_screenshot_as_bytes, website_url, True)
                future_desktop_site = executor.submit(capture_screenshot_as_bytes, website_url, False)

            mobile_ref = future_mobile_ref.result()
            desktop_ref = future_desktop_ref.result()
            mobile_site = future_mobile_site.result()
            desktop_site = future_desktop_site.result()

            # Save website screenshots locally
            output_dir = "website_screenshots"
            os.makedirs(output_dir, exist_ok=True)  # Create the directory if it doesn't exist

            mobile_site_path = os.path.join(output_dir, f"mobile_{index}.png")
            desktop_site_path = os.path.join(output_dir, f"desktop_{index}.png")

            if mobile_site:
                with open(mobile_site_path, "wb") as f:
                    f.write(mobile_site.getbuffer())

            if desktop_site:
                with open(desktop_site_path, "wb") as f:
                    f.write(desktop_site.getbuffer())

            # Compare images
            mobile_similarity = compare_images_bytes(mobile_ref, mobile_site) if mobile_ref and mobile_site else 0.0
            desktop_similarity = compare_images_bytes(desktop_ref, desktop_site) if desktop_ref and desktop_site else 0.0

            return mobile_similarity, desktop_similarity
        except Exception as e:
            print(f"Error processing row {index}: {e}")
            return 0.0, 0.0

    # Process rows in parallel
    results = []
    with ThreadPoolExecutor(max_workers=10) as executor:  # Control concurrency
        futures = {executor.submit(process_row, index, row): index for index, row in df.iterrows()}
        for future in as_completed(futures):
            index = futures[future]
            try:
                results.append((index, *future.result()))
            except Exception as e:
                print(f"Error in row {index}: {e}")

    # Update DataFrame
    for index, mobile_similarity, desktop_similarity in results:
        df.at[index, 'Mobile Match %'] = mobile_similarity
        df.at[index, 'Desktop Match %'] = desktop_similarity

    # Batch update Google Sheet
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())
    print("Google Sheet updated successfully!")

    # Save results locally
    df.to_excel(local_excel_path, index=False)
    print(f"Results saved to {local_excel_path}")

sheet_url = input("Enter Google Sheet URL: ").strip()
credentials_path = r"C:\Users\umama\Desktop\port\Assignment\inspiring-folio-447308-r3-37521cb4f054.json"
local_excel_path = "local_google_sheet.xlsx"

process_google_sheet(sheet_url, credentials_path, local_excel_path)
