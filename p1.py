import time
import pyautogui
import pyperclip
import pandas as pd
from google import genai
import os

# --- Configuration ---
# 1. Put your Gemini API Key here



client = genai.Client()

# 2. File paths
EXCEL_FILE = "GST_Notices_Master.xlsx"

# 3. Screen Coordinates (You MUST adjust these for your specific monitor)
# You can find these by running pyautogui.mouseInfo() in your terminal
CAPTCHA_REGION = (800, 500, 200, 80) # (Left, Top, Width, Height) of the CAPTCHA image
TEXT_BOX_CLICK = (850, 600)          # Where to click to type the CAPTCHA
TABLE_CLICK_AREA = (500, 500)        # Somewhere safe to click inside the notices page to select the table

def solve_captcha_with_gemini():
    print("📸 Taking screenshot of the CAPTCHA...")
    captcha_img = pyautogui.screenshot(region=CAPTCHA_REGION)
    
    # Save temporarily to send to Gemini
    captcha_path = "temp_captcha.png"
    captcha_img.save(captcha_path)

    print("🧠 Sending to Gemini API for text extraction...")
    try:
        # Upload the image to Gemini and ask for just the text
        img_file = client.files.upload()
        prompt = "Read the text in this CAPTCHA image. Respond ONLY with the exact letters and numbers. No spaces, no punctuation, no extra words."


        response = client.models.generate_content(
                model="gemini-3-flash-preview", contents=prompt
            )
        
        captcha_text = response.text.strip().replace(" ", "")
        print(f"✅ Gemini read the CAPTCHA as: {captcha_text}")
        
        # Cleanup
        genai.delete_file(img_file.name)
        os.remove(captcha_path)
        
        return captcha_text
    except Exception as e:
        print(f"❌ Failed to read CAPTCHA: {e}")
        return None

def login_and_wait(captcha_text):
    print("⌨️ Typing CAPTCHA and logging in...")
    # Click the input box
    pyautogui.click(TEXT_BOX_CLICK)
    time.sleep(0.5)
    
    # Type the text and press Enter
    pyautogui.write(captcha_text, interval=0.1)
    pyautogui.press('enter')
    
    print("⏳ Waiting 40 seconds for Winman to route to the Notices page...")
    time.sleep(40)

def extract_and_update_excel():
    print("📋 Extracting table data from the screen...")
    
    # Click somewhere inside the notices page to ensure the window is active
    pyautogui.click(TABLE_CLICK_AREA)
    time.sleep(1)
    
    # Use Ctrl+A to select all, and Ctrl+C to copy
    # (Since it's built on HTML tags, the clipboard will retain the table structure)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)
    
    print("📊 Processing clipboard data...")
    try:
        # pandas has a magical function that reads tabular data straight from your clipboard!
        df_new = pd.read_clipboard()
        
        # Basic check to ensure we actually copied a table
        if df_new.empty or len(df_new.columns) < 2:
            print("❌ No valid table data found on clipboard. Did the page load correctly?")
            return
            
    except Exception as e:
        print(f"❌ Failed to parse clipboard table: {e}")
        return

    print("📁 Updating Excel Workbook...")
    # Check if our master Excel file already exists
    if os.path.exists(EXCEL_FILE):
        # Load the old data
        df_existing = pd.read_excel(EXCEL_FILE)
        
        # Combine the old data and the newly scraped data
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        
        # Drop duplicates! 
        # (Assuming there is a 'Notice Number' or 'Reference Number' column to identify unique rows)
        # You will need to change 'Notice Number' to the exact column header the GST portal uses
        try:
            # We keep the first instance and drop the rest to avoid duplicates
            df_combined.drop_duplicates(subset=['Notice Number'], keep='first', inplace=True)
            new_records_count = len(df_combined) - len(df_existing)
            print(f"✅ Found and added {new_records_count} new notice(s).")
        except KeyError:
            print("⚠️ Could not find 'Notice Number' column to remove duplicates. Saving all rows.")
            
    else:
        # If the file doesn't exist yet, this is our first time running it
        df_combined = df_new
        print("✅ Created new Excel file with the extracted data.")

    # Save it back to the Excel file
    df_combined.to_excel(EXCEL_FILE, index=False)
    print("🎉 Done! Excel file updated.")

# --- Main Execution ---
if __name__ == "__main__":
    print("⚠️ Bring the Winman login screen to the front! Starting in 5 seconds...")
    time.sleep(5)
    
    captcha = solve_captcha_with_gemini()
    if captcha:
        login_and_wait(captcha)
        extract_and_update_excel()