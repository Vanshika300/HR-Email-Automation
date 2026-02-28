import pandas as pd
import smtplib
from email.message import EmailMessage
import time

# ---------------- CONFIG ---------------- #

YOUR_EMAIL = os.getenv("EMAIL_ID")
APP_PASSWORD = os.getenv("EMAIL_PASS")    # <-- Gmail App Password

EXCEL_FILE = "hr_contacts.xlsx"
RESUME_FILE = "Resume.pdf"

DAILY_LIMIT = 30
sent_today = 0

# -------------------------------------- #

# Load Excel
df = pd.read_excel(EXCEL_FILE)
df["Status"] = df["Status"].astype(str)
# Connect Gmail SMTP
server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server.login(YOUR_EMAIL, APP_PASSWORD)

sent_count = 0
df = df.drop_duplicates(subset=["Email"])

for index, row in df.iterrows():

    if str(row["Status"]).strip().lower() != "sent":

        name = row["Name"]
        company = row["Company"]
        receiver = row["Email"]

        msg = EmailMessage()
        msg["From"] = YOUR_EMAIL
        msg["To"] = receiver
        msg["Subject"] = f"Application for Fresher / Intern Role at {company}"

        body = f"""
Dear {name},

I hope this email finds you well.

My name is Vanshika Shukla, and I am currently pursuing B.Tech in Computer Science (AI/ML).
I am writing to express my interest in any fresher or internship opportunities at {company}.

I have attached my resume for your review and would welcome the opportunity to discuss how my skills align with your team’s requirements.

Thank you for your time and consideration.

Warm regards,
Vanshika Shukla
Email: vanshikashukla065@gmail.com
"""

        msg.set_content(body)

        with open(RESUME_FILE, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename="Vanshika_Shukla_Resume.pdf"
            )

        try:
            server.send_message(msg)
            sent_count += 1
            sent_today += 1

            if sent_today >= DAILY_LIMIT:
                print("Daily limit reached ✅")
                break
            
            df.at[index, "Status"] = "Sent"
            df.to_excel(EXCEL_FILE, index=False)   # SAVE AFTER EACH SEND

            print(f"Sent to {receiver}")

            time.sleep(10)

        except Exception as e:
            print("Stopped due to:", e)
            break


print(f"\nDONE ✅ Total emails sent: {sent_count}")
