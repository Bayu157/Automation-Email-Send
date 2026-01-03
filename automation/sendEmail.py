import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv, set_key
from pathlib import Path
import pandas as pd
import time

env_path = Path('your_env_file_path') 
excel_path = Path(__file__).parent / 'your_file.xlsx' 

# Functions
def account ():
    if env_path.exists():
        with open(env_path) as v:
            for i in v:
                if i.strip() and not i.startswith('#'):
                    key, value = i.strip().split('=', 1)
                    os.environ[key] = value
                    continue
    
    elif not env_path.exists():
        print("Please input your email and apppassword")
        email = input("Your Email:")
        password = input("Your App Password:")

        set_key(env_path,"Email", email)
        set_key(env_path,"Password", password)

def message (from_your_email, password, send_to_email, subject, body):
    msg = MIMEMultipart()

    msg['Subject'] = subject
    msg['From'] = from_your_email
    msg['To'] = send_to_email

    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(from_your_email, password)
            server.send_message(msg)
            # print("Email sent successfully!")

    except Exception as e:
        print(f"Failed to send email: {e}")

def start_SendEmail ():
    start = time.perf_counter()

    load_dotenv()

    a = os.getenv("Email")
    b = os.getenv("Password")

    df = pd.read_excel(excel_path)

    for i in range(len(df)):
        message(
            from_your_email= a,
            password       = b,
            send_to_email  = df['Email'].iloc[i],
            subject        = df['Subject'].iloc[i],
            body           = df['Message'].iloc[i]
        )
        print(f"Sending email to: {df['Email'].iloc[i]}")

    end = time.perf_counter()
    print(f"Finished in {end - start:.2f} seconds")

# Main Program 

print("===================================")
print("  v0.0.1 - Automation Send Email   ")
print("===================================")
print("")

start_SendEmail() # Call the function to start sending emails