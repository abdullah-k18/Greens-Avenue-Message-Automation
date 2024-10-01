import pywhatkit as kit
import pandas as pd

data = pd.read_excel("data.xlsx")

for index, row in data.iterrows():
    name = row['NAME']
    phone_number = str(int(row['NUMBER'])).strip()
    remaining_amount = row['BALANCE']

    if not phone_number.startswith("+"):
        phone_number = "+" + phone_number

    message = (f"Assalam o Alaikum! Sir/Madam {name}, your remaining amount for Greens Avenue monthly fee is {remaining_amount}RS. kindly pay it as soon aas possible.")

    hour = 16

    minute = (0 + index)

    print(f"Sending message to {phone_number}")

    try:
        kit.sendwhatmsg(phone_number, message, hour, minute)
        print(f"Message sent to {name} at {phone_number}")
    except Exception as e:
        print(f"Failed to send message to {phone_number}: {e}")

