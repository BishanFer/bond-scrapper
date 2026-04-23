#!/usr/bin/env python3
import base64
import os
import sys
from datetime import datetime

import requests

latest_file = "treasury_bond_yields.xlsx"

api_key = os.environ.get("RESEND_API_KEY", "")
if not api_key:
    print("RESEND_API_KEY is not configured, skipping email")
    sys.exit(0)

to_email = os.environ.get("TO_EMAIL", "")
if not to_email:
    print("TO_EMAIL is not configured, skipping email")
    sys.exit(0)

if not os.path.exists(latest_file):
    print(f"No Excel file found at {latest_file}")
    sys.exit(0)

with open(latest_file, "rb") as f:
    file_b64 = base64.b64encode(f.read()).decode()

today = datetime.now().strftime("%Y-%m-%d")
from_email = os.environ.get("FROM_EMAIL") or "onboarding@resend.dev"

payload = {
    "from": from_email,
    "to": [to_email],
    "subject": f"Treasury Bond Yields - {today}",
    "html": f"<p>Please find attached the Treasury Bond Yields report for {today}.</p>",
    "attachments": [{"filename": latest_file, "content": file_b64}],
}

print(f"Sending to: {to_email}")
print(f"From: {from_email}")

response = requests.post(
    "https://api.resend.com/emails",
    headers={
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    },
    json=payload,
    timeout=30,
)

print(f"Response: {response.status_code}")
print(f"Body: {response.text}")

if response.status_code not in (200, 201, 202):
    sys.exit(1)
