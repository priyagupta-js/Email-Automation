import pandas as pd
import smtplib
import time
import logging
from email.message import EmailMessage
from datetime import datetime
from jinja2 import Template
from config import SMTP_SERVER, SMTP_PORT, SENDER_EMAIL, APP_PASSWORD, DAILY_LIMIT, DELAY_SECONDS
import os


# Logging
logging.basicConfig(
    filename="logs/email.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)




# load data
df = pd.read_excel("data/HR_lists.xlsx", dtype=str)

# Normalize column names
df.columns = [c.strip().lower() for c in df.columns]

# Ensure required columns exist
for col in ["status", "date", "error"]:
    if col not in df.columns:
        df[col] = ""

# Normalize values
df["status"] = (
    df["status"]
    .fillna("")
    .str.strip()
    .str.lower()
    .replace("", "pending")
)
df["date"] = df["date"].fillna("")
df["error"] = df["error"].fillna("")


# Load template
with open("email_template.txt", "r", encoding="utf-8") as f:
    template = Template(f.read())

def get_first_name(full_name: str) -> str:
    if not full_name:
        return ""
    return full_name.strip().split()[0]

# Send function
def send_email(row):
    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = row["email"]
    msg["Subject"] = "Software Engineer – Immediate Availability"

    first_name = get_first_name(row["name"])

    body = template.render(
        name=first_name,
        company=row["company"]
    )
    msg.set_content(body)

    with open("resume/Priya_Gupta_Resume_SDE.pdf", "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename="Priya_Gupta_Resume_SDE.pdf"
        )

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)

# Select rows
pending = df[df["status"] == "pending"].head(DAILY_LIMIT)

logging.info(f"Found {len(pending)} pending emails to process")


for idx, row in pending.iterrows():
    try:
        send_email(row)
        df.loc[idx, "status"] = "sent"
        df.loc[idx, "date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df.loc[idx, "error"] = ""
        logging.info(f"Sent to {row['email']}")
    except Exception as e:
        df.loc[idx, "status"] = "failed"
        df.loc[idx, "error"] = str(e)
        logging.error(f"Failed for {row['email']} - {e}")

    temp_file = "data/HR_lists_temp.xlsx"
    df.to_excel(temp_file, index=False)
    os.replace(temp_file, "data/HR_lists.xlsx")

    logging.info(f"Excel updated after processing {row['email']}")


    time.sleep(DELAY_SECONDS)

logging.info("=== Email automation script finished ===")