import time
import pandas as pd
import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load Excel file
df = pd.read_excel("hostel_room_booking.xlsx")

# Setup Chrome WebDriver
service = Service("chromedriver.exe")
driver = webdriver.Chrome(service=service)

# Open WhatsApp Web
driver.get("https://web.whatsapp.com")
print("🔑 Please scan the QR code in Chrome to log in WhatsApp Web")
time.sleep(40)  # wait for QR scan

# Counters
success_count = 0
failed_count = 0
failed_list = []   # store failed rows

# Loop through students
for i, row in df.iterrows():
    if pd.isna(row['Phone']) or pd.isna(row['Name']) or pd.isna(row['S_Date']) or pd.isna(row['E_Date']):
        print(f"⏭️ Skipped row {i+1} (missing data)")
        continue

    phone = f"+91{int(row['Phone'])}"
    s_date = pd.to_datetime(row['S_Date']).strftime("%d.%m.%Y")
    e_date = pd.to_datetime(row['E_Date']).strftime("%d.%m.%Y")

    # Message content (dynamic dates + signature)
    message = (
        f"Good Morning, {row['Name']}.\n\n"
        f"The PMSS Scholarship renewal process started on *{s_date}*.\n\n"
        "Students from the *AD, AM, and BM* departments must submit their scholarship "
        f"renewal forms on or before *{e_date}*.\n\n"
        "If the forms are not submitted, the renewal date will be announced later.\n\n"
        "The following documents must be submitted along with the scholarship form:\n\n"
        "📌 *Aadhaar Card*\n"
        "📌 *Bank Passbook*\n"
        "📌 *Declaration Form*\n\n"
        "Regards,\n"
        "*KPRIET Scholarship Section*"
    )

    encoded_message = urllib.parse.quote(message)
    driver.get(f"https://web.whatsapp.com/send?phone={phone}&text={encoded_message}&app_absent=0")

    try:
        # Wait for input box
        input_box = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )

        # Just hit ENTER (don’t retype the message)
        input_box.send_keys(Keys.ENTER)
        print(f"✅ Sent to {row['Name']} ({phone})")
        success_count += 1

    except Exception as e:
        print(f"⚠️ Failed to send to {row['Name']} ({phone}) → {e}")
        failed_count += 1
        failed_list.append({
            "Name": row['Name'],
            "Phone": phone
        })

    time.sleep(3)

# Print summary
print("\n📊 ----- SUMMARY -----")
print(f"✅ Successfully sent: {success_count}")
print(f"❌ Failed to send: {failed_count}")

# Print failed list if any
if failed_list:
    print("\n❌ Failed Message Report:")
    for f in failed_list:
        print(f"   - {f['Name']} | {f['Phone']}")

driver.quit()
