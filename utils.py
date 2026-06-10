#!/usr/bin/env python3
"""
Shared utilities for DGCA automation
"""

import os
import time
import smtplib
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait

# ==========================================================
# LOAD CONFIG
# ==========================================================
CONFIG_FILE = "config.ini"
config = configparser.ConfigParser()
config.read(CONFIG_FILE)

# Email config (env vars override config)
SMTP_USER = os.getenv("SMTP_USER", config.get("EMAIL_SENDER", "EMAIL", fallback=""))
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", config.get("EMAIL_SENDER", "APP_PASSWORD", fallback=""))
SMTP_SERVER = os.getenv("SMTP_SERVER", config.get("SMTP", "SERVER", fallback="smtp.gmail.com"))
SMTP_PORT = int(os.getenv("SMTP_PORT", config.get("SMTP", "PORT", fallback=587)))
RECIPIENTS = os.getenv("RECIPIENTS", "").split(",")
if not RECIPIENTS and "EMAIL_RECIPIENTS" in config:
    RECIPIENTS = [value.strip() for key, value in config["EMAIL_RECIPIENTS"].items() if key.startswith("recipient_")]

# Chatbot config (used by sanity module)
CHATBOT_URL = os.getenv("CHATBOT_URL", config.get("CHATBOT", "URL", fallback="https://www.dgca.gov.in/digigov-portal/"))
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")

# Superset config (used by dashboard module)
SUPERSET_BASE_URL = "http://20.244.27.216:8088"
SUPERSET_HOME_URL = f"{SUPERSET_BASE_URL}/superset/welcome/"
DASHBOARD_NAME = "DGCA Chatbot Dashboard"
USERNAME = "viewer"
PASSWORD = "Dashboard@2026"

OUTPUT_DIR = os.path.abspath("./reports")
os.makedirs(OUTPUT_DIR, exist_ok=True)
SCREENSHOT_DIR = os.path.join(OUTPUT_DIR, "screenshots")
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# ==========================================================
# DRIVER CREATION
# ==========================================================
def create_driver():
    options = Options()
    headless = os.getenv("HEADLESS", "true").lower() == "true"
    if headless:
        options.add_argument("--headless=new")
    else:
        options.add_argument("--start-maximized")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    if os.path.exists("/usr/bin/chromium-browser"):
        options.binary_location = "/usr/bin/chromium-browser"
    elif os.path.exists("/usr/bin/chromium"):
        options.binary_location = "/usr/bin/chromium"
    
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(120)
    try:
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    except:
        pass
    return driver

def wait(driver, seconds=30):
    return WebDriverWait(driver, seconds)

def safe_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    time.sleep(0.5)
    try:
        element.click()
    except Exception:
        driver.execute_script("arguments[0].click();", element)

def save_error_screenshot(driver, filename="error.png"):
    path = os.path.join(OUTPUT_DIR, filename)
    try:
        driver.save_screenshot(path)
        print(f"   Error screenshot saved: {path}")
    except:
        pass

def send_email(attachments, subject="DGCA Automation Report"):
    if not SMTP_USER or not SMTP_PASSWORD or not RECIPIENTS:
        print("   Email not configured, skipping.")
        return
    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(RECIPIENTS)
    msg["Subject"] = subject
    body = "Please find attached the daily DGCA automation reports."
    msg.attach(MIMEText(body, "plain"))
    for file_path in attachments:
        with open(file_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
            msg.attach(part)
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)
    print("   Email sent successfully.")
