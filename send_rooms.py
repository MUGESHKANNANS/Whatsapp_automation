import pandas as pd
import pywhatkit as kit
import time

# Load Excel file
df = pd.read_excel("hostel_room_booking.xlsx")

for i, row in df.iterrows():
    phone = f"+91{row['Phone']}"   # add country code
    message = f"Hello {row['Name']}, your hostel room number is {row['Room']}."

    print(f"Sending to {phone}: {message}")

    # Send instantly (requires WhatsApp Web open & logged in)
    kit.sendwhatmsg_instantly(phone, message, wait_time=15, tab_close=True)

    time.sleep(15)  # delay between messages

