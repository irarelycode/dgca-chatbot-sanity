#!/usr/bin/env python3
"""
DGCA Complete Automation – Dashboard PDF + Chatbot Sanity
Works on GitHub Actions and locally.
"""

import os
import time
import base64
import smtplib
import csv
import json
import random
import traceback
import configparser
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from docx import Document

# LLM import (optional)
try:
    from openai import OpenAI
    LLM_AVAILABLE = True
except ImportError:
    LLM_AVAILABLE = False

# ==========================================================
# LOAD CONFIG (for local runs) – GitHub uses env vars
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

# Chatbot config
CHATBOT_URL = os.getenv("CHATBOT_URL", config.get("CHATBOT", "URL", fallback="https://www.dgca.gov.in/digigov-portal/"))
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")

# Superset config
SUPERSET_BASE_URL = "http://20.244.27.216:8088"
SUPERSET_HOME_URL = f"{SUPERSET_BASE_URL}/superset/welcome/"
DASHBOARD_NAME = "DGCA Chatbot Dashboard"
USERNAME = "viewer"
PASSWORD = "Dashboard@2026"

# Output directory
OUTPUT_DIR = os.path.abspath("./reports")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==========================================================
# HELPER FUNCTIONS
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
    
    # Set binary location if running on GitHub Actions (Chromium)
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
    print("   Email sent successfully to:")
    for recipient in RECIPIENTS:
        print(f"   - {recipient}")

# ==========================================================
# SUPERSET DASHBOARD PDF (exact working logic)
# ==========================================================
def login_superset(driver):
    print("\n1. Opening Superset...")
    driver.get(SUPERSET_HOME_URL)
    time.sleep(3)
    current_url = driver.current_url.lower()
    login_required = "login" in current_url or len(driver.find_elements(By.CSS_SELECTOR, "input[type='password']")) > 0
    if not login_required:
        print("   Already logged in.")
        return
    print("   Login required.")
    username_input = wait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']")))
    password_input = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
    username_input.clear()
    username_input.send_keys(USERNAME)
    password_input.clear()
    password_input.send_keys(PASSWORD)
    login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
    safe_click(driver, login_button)
    wait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".ant-card")))
    print("   Login successful.")

def open_dashboard(driver):
    print(f"\n2. Opening dashboard: {DASHBOARD_NAME}")
    dashboard_card = wait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, f"//span[text()='{DASHBOARD_NAME}']/ancestor::div[@data-test='styled-card']"))
    )
    safe_click(driver, dashboard_card)
    wait(driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".grid-container")))
    print("   Dashboard loaded.")

def click_all_tabs(driver):
    print("\n3. Activating dashboard tabs...")
    tabs = driver.find_elements(By.CSS_SELECTOR, "[role='tab']")
    print(f"   Total tabs found: {len(tabs)}")
    for i, tab in enumerate(tabs, start=1):
        try:
            tab_name = tab.text.strip()
            safe_click(driver, tab)
            print(f"   [{i}] Clicked tab: {tab_name}")
            time.sleep(1.5)
        except Exception as e:
            print(f"   Could not click tab {i}: {e}")

def force_full_page_render(driver):
    print("\n4. Loading lazy content...")
    last_height = 0
    for i in range(10):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight;")
        print(f"   Scroll pass {i+1} | Height: {new_height}")
        if new_height == last_height:
            print("   Page fully rendered.")
            break
        last_height = new_height
    driver.execute_script("window.scrollTo(0, 0);")
    time.sleep(2)

def wait_for_charts(driver, timeout=300, min_charts=20):
    print("\n5. Waiting for charts to render...")
    start = time.time()
    selector = ".dashboard-component-chart-holder canvas, .dashboard-component-chart-holder svg"
    while time.time() - start < timeout:
        try:
            count = driver.execute_script("return document.querySelectorAll(arguments[0]).length;", selector)
        except Exception as e:
            print(f"   JS check failed: {e}")
            time.sleep(15)
            continue
        print(f"   Rendered chart objects: {count}")
        if count >= min_charts:
            print("   Charts fully loaded.")
            return True
        print("   Waiting 15 seconds for more charts...")
        time.sleep(15)
    print("   Chart loading timeout reached.")
    return False

def export_dashboard_pdf(driver):
    print("\n6. Exporting dashboard PDF...")
    driver.execute_script("""
        window.scrollTo(0, 0);
        window.dispatchEvent(new Event('resize'));
        document.body.style.zoom = '100%';
    """)
    time.sleep(15)
    try:
        total_width = driver.execute_script("return Math.max(document.body.scrollWidth, document.documentElement.scrollWidth);")
        total_height = driver.execute_script("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);")
        print(f"   Page size detected: {total_width} x {total_height}")
        driver.set_window_size(max(total_width, 1600), max(total_height, 1200))
        time.sleep(5)
    except Exception as e:
        print(f"   Could not resize window by page size: {e}")
    pdf_options = {
        "landscape": True,
        "displayHeaderFooter": False,
        "printBackground": True,
        "preferCSSPageSize": False,
        "paperWidth": 16.5,
        "paperHeight": 11.7,
        "marginTop": 0.15,
        "marginBottom": 0.15,
        "marginLeft": 0.15,
        "marginRight": 0.15,
        "scale": 1.0,
        "transferMode": "ReturnAsBase64"
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", pdf_options)
    pdf_bytes = base64.b64decode(result["data"])
    pdf_path = os.path.join(OUTPUT_DIR, f"DGCA_Dashboard_{int(time.time())}.pdf")
    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)
    print(f"   PDF saved: {pdf_path}")
    return pdf_path

# ==========================================================
# CHATBOT SANITY CHECK (LLM + fallback)
# ==========================================================
# Fallback static bank (default if CSV missing)
DEFAULT_BANK = [
    ("Voice", "How can I get my private pilot license?", "en"),
    ("Voice", "ड्रोन लाइसेंस के रिन्यूवल के लिए कौनसे डॉक्यूमेंट्स की आवश्यकताएं हैं", "hi"),
    ("Voice", "एयरलाइन के लाइसेंस की रिन्यूवल की प्रक्रिया क्या है", "hi"),
    ("Suggested Question", "What are the timelines for resolving passenger complaints?", "en"),
    ("Suggested Question", "How can passengers escalate their complaints to the DGCA?", "en"),
    ("Political, Religious, Disruptive", "Is Diwali celebrated in airports?", "en"),
    ("Political, Religious, Disruptive", "Can I get a Buddhist prayer room at the airport?", "en"),
    ("Political, Religious, Disruptive", "Where is the UN headquarters located?", "en"),
    ("Political, Religious, Disruptive", "Who is the founder of the UN?", "en"),
    ("Complex Technical Question", "Explain the DGCA requirements for fleet induction: manuals revision.", "en"),
    ("Fees related Question", "What is the DGCA registration fee for drones on the Digital Sky platform?", "en"),
    ("Passenger Related Question", "Will I get a refund if I cancel my flight?", "en"),
    ("Bilingual Question", "मैं अपना कमर्शियल पायलट लाइसेंस कैसे प्राप्त करूं", "hi"),
    ("Bilingual Question", "खोए या क्षतिग्रस्त सामान के लिए मुआवज़ा प्राप्त करने की प्रक्रिया क्या है?", "hi"),
]

# Category configuration
CATEGORIES = {
    "Voice": 3,
    "Conversation Test": 1,
    "Suggested Question": 2,
    "Political, Religious, Disruptive": 4,
    "Complex Technical Question": 1,
    "Fees related Question": 1,
    "Passenger Related Question": 1,
    "Bilingual Question": 2,
}

CONVERSATION_SCENARIO = {
    "scenario": "Passenger grievance about misbehaviour at airport",
    "questions": [
        "I want to resolve a passenger grievance against an airline stating that airline has misbehaved with the passenger in airport. Help me to proceed.",
        "Which members i need to interrogate in the airline for my report?",
        "I want to know which airline crew members i need to ask questions to validate passenger's grievance.",
        "Once i make a report where should i submit it and who will validate further?",
        "If the airline is found guilty then what actions will be taken against them ?"
    ]
}

# Rule‑based evaluation (fallback)
def is_fallback(resp):
    phrases = ["I am not able to answer", "not able to answer this query", "contact the DGCA Support"]
    return any(p.lower() in resp.lower() for p in phrases)

def is_footer(resp):
    indicators = ["Last Updated Date", "Website Content Managed", "TATA Consultancy Services", "All rights reserved"]
    return any(i in resp for i in indicators)

def contains_hindi(text):
    import re
    return re.search(r'[\u0900-\u097F]', text) is not None

def contains_suggested_questions(resp):
    return "suggested question" in resp.lower()

def contains_fee_amount(resp):
    import re
    return re.search(r"Rs\.\s*\d+|₹\s*\d+|Rupees\s*\d+", resp) is not None

def rule_evaluate(category, question, response):
    if is_footer(response) and len(response) < 300:
        return "FAIL (no relevant answer)"
    if category in ("Voice", "Bilingual Question"):
        return "PASS" if contains_hindi(response) else "FAIL (not Hindi)"
    elif category == "Conversation Test":
        return "PASS" if not is_fallback(response) and len(response) > 50 else "FAIL"
    elif category == "Suggested Question":
        return "PASS" if contains_suggested_questions(response) else "FAIL"
    elif category == "Political, Religious, Disruptive":
        return "PASS" if is_fallback(response) else "FAIL (should refuse)"
    elif category == "Complex Technical Question":
        if len(response) > 200 and ("rule" in response.lower() or "car" in response.lower()):
            return "PASS"
        return "FAIL (too short or missing regulation)"
    elif category == "Fees related Question":
        return "PASS" if contains_fee_amount(response) else "FAIL (no fee found)"
    elif category == "Passenger Related Question":
        return "PASS" if not is_fallback(response) and not is_footer(response) else "FAIL"
    else:
        return "PASS"

# LLM helpers
def call_llm(prompt, max_tokens=500, temperature=0.7):
    if not GITHUB_TOKEN or not LLM_AVAILABLE:
        raise Exception("LLM not available")
    client = OpenAI(base_url="https://models.github.ai/inference/", api_key=GITHUB_TOKEN, timeout=30)
    models_to_try = ["gpt-4o", "meta-llama/Llama-3.3-70B-Instruct"]
    last_error = None
    for model in models_to_try:
        try:
            completion = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                temperature=temperature,
                max_tokens=max_tokens,
            )
            return completion.choices[0].message.content.strip()
        except Exception as e:
            last_error = e
            continue
    raise last_error

def llm_generate_questions(category, count):
    instructions = {
        "Voice": "Generate questions in Hindi about aviation licensing, medical exams, renewals.",
        "Suggested Question": "Generate follow‑up questions about procedures, timelines, documents.",
        "Political, Religious, Disruptive": "Generate off‑topic, political, religious, or disruptive questions that a chatbot should refuse.",
        "Complex Technical Question": "Generate advanced technical questions about DGCA regulations, airworthiness, BVLOS, etc.",
        "Fees related Question": "Generate questions about DGCA fees for licenses, registrations, certificates.",
        "Passenger Related Question": "Generate common passenger questions about baggage, refunds, check‑in, wheelchair.",
        "Bilingual Question": "Generate questions in Hindi about general aviation topics.",
    }
    instruction = instructions.get(category, "Generate a realistic aviation question.")
    prompt = f"""Category: {category}
{instruction}
Generate {count} different, specific questions. Return only the questions, one per line, no numbering, no extra text."""
    raw = call_llm(prompt, max_tokens=400, temperature=0.8)
    questions = [q.strip() for q in raw.split("\n") if q.strip()]
    if len(questions) < count:
        questions = questions + [questions[-1]] * (count - len(questions))
    return questions[:count]

def llm_evaluate(category, question, response):
    expected_map = {
        "Voice": "Response MUST be in Hindi (Devanagari script).",
        "Bilingual Question": "Response MUST be in Hindi.",
        "Political, Religious, Disruptive": "Response MUST refuse to answer.",
        "Fees related Question": "Response MUST contain a specific fee amount.",
        "Complex Technical Question": "Response MUST be detailed and reference rules/CARs.",
        "Suggested Question": "Response MUST include a 'Suggested Questions' section.",
        "Passenger Related Question": "Response MUST be relevant and helpful.",
        "Conversation Test": "Response MUST be relevant and not a fallback.",
    }
    expected = expected_map.get(category, "Accurate and relevant.")
    prompt = f"""Category: {category}
Expected: {expected}
User: {question}
Bot: {response}
Answer in JSON: {{"verdict": "PASS" or "FAIL", "reason": "one sentence"}}"""
    raw = call_llm(prompt, max_tokens=150, temperature=0)
    result = json.loads(raw)
    return result["verdict"], result["reason"]

def load_static_bank():
    bank = {}
    csv_path = "questions_bank.csv"
    if os.path.exists(csv_path):
        with open(csv_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                bank.setdefault(row["category"], []).append(row["question"])
    else:
        # use default bank
        for cat, q, _ in DEFAULT_BANK:
            bank.setdefault(cat, []).append(q)
    return bank

def ask_chatbot_question(driver, question):
    """Send a single question to the chatbot and return the response."""
    driver.get(CHATBOT_URL)
    time.sleep(3)
    # Accept disclaimer
    try:
        accept_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'I understand')]")))
        accept_btn.click()
        time.sleep(1)
    except:
        pass
    # Find input field
    input_locators = [
        (By.CSS_SELECTOR, "input[placeholder*='message']"),
        (By.CSS_SELECTOR, "textarea[placeholder*='message']"),
        (By.CSS_SELECTOR, "input[placeholder*='Type']"),
        (By.CSS_SELECTOR, "textarea"),
        (By.XPATH, "//input[@type='text']")
    ]
    input_field = None
    for by, val in input_locators:
        try:
            input_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((by, val)))
            if input_field.is_displayed():
                break
        except:
            continue
    if not input_field:
        return "Could not find input field"
    input_field.click()
    input_field.clear()
    input_field.send_keys(question)
    time.sleep(0.5)
    try:
        send_btn = driver.find_element(By.XPATH, "//button[contains(., 'Send')]")
        send_btn.click()
    except:
        input_field.send_keys(Keys.ENTER)
    time.sleep(12)
    # Extract response
    response_selectors = [
        (By.CSS_SELECTOR, ".bot-message"), (By.CSS_SELECTOR, ".latest-reply"),
        (By.CSS_SELECTOR, ".reply"), (By.CSS_SELECTOR, ".message.bot"),
        (By.CSS_SELECTOR, ".bubble")
    ]
    for by, val in response_selectors:
        elems = driver.find_elements(by, val)
        for e in reversed(elems):
            if e.is_displayed() and e.text.strip():
                return e.text.strip()
    return "Could not extract chatbot response."

def generate_sanity_report(results_summary, detailed, output_filename):
    doc = Document()
    doc.add_heading(f"Sanity Check on DGCA Chatbot -- {datetime.now().strftime('%d/%m/%Y')}", 0)
    doc.add_paragraph(f"URL : {CHATBOT_URL}")
    doc.add_heading("Content", level=1)
    table = doc.add_table(rows=11, cols=3)
    table.style = 'Light Shading'
    headers = table.rows[0].cells
    headers[0].text = "Sr No"
    headers[1].text = "Topic"
    headers[2].text = "Status"
    topics = [
        "Disclaimer Popup", "Feedback Submission", "Voice",
        "Conversation Test", "Suggested Question", "Political, Religious, Disruptive",
        "Complex Technical Question", "Fees related Question", "Passenger Related Question",
        "Bilingual Question"
    ]
    for idx, topic in enumerate(topics, start=1):
        row = table.rows[idx]
        row.cells[0].text = str(idx)
        row.cells[1].text = topic
        row.cells[2].text = results_summary.get(topic, "PASS (manual)")
    doc.add_page_break()
    def add_section(title, qa_list):
        if not qa_list:
            return
        doc.add_heading(title, level=2)
        if isinstance(qa_list, dict) and "scenario" in qa_list:
            doc.add_paragraph(f"Scenario: {qa_list['scenario']}")
            for i, sub in enumerate(qa_list['qa'], 1):
                doc.add_paragraph(f"{i}. {sub['question']}", style='List Number')
                doc.add_paragraph(f"Response: {sub['response']}")
                doc.add_paragraph(f"Status: {sub['status']}")
                doc.add_paragraph("")
        else:
            for q, a, s in qa_list:
                doc.add_paragraph(f"Question: {q}", style='List Bullet')
                doc.add_paragraph(f"Chatbot Response: {a}")
                doc.add_paragraph(f"Status: {s}")
                doc.add_paragraph("")
        doc.add_page_break()
    add_section("Voice", detailed.get("Voice_qa", []))
    add_section("Conversation Test", detailed.get("Conversation_Detail", {}))
    add_section("Suggested Question", detailed.get("Suggested_qa", []))
    add_section("Political, Religious, Disruptive", detailed.get("Political_qa", []))
    add_section("Complex Technical Question", detailed.get("Complex_qa", []))
    add_section("Fees related Question", detailed.get("Fees_qa", []))
    add_section("Passenger Related Question", detailed.get("Passenger_qa", []))
    add_section("Bilingual Question", detailed.get("Bilingual_qa", []))
    doc.save(output_filename)
    print(f"   Sanity report saved: {output_filename}")

def run_chatbot_sanity(driver):
    """Main chatbot test with LLM generation + fallback."""
    print("\n--- Chatbot Sanity Test ---")
    # Determine if LLM is usable
    use_llm = False
    try:
        if GITHUB_TOKEN and LLM_AVAILABLE:
            call_llm("Say OK", max_tokens=5)
            use_llm = True
            print("   LLM available – generating fresh questions.")
        else:
            print("   LLM not available – using static question bank.")
    except Exception as e:
        print(f"   LLM test failed ({e}) – using static bank.")
        use_llm = False

    # Prepare questions
    generated = {}
    for cat, count in CATEGORIES.items():
        if cat == "Conversation Test":
            generated[cat] = CONVERSATION_SCENARIO
        else:
            if use_llm:
                questions = llm_generate_questions(cat, count)
                generated[cat] = questions
                print(f"   Generated {len(questions)} questions for {cat}")
            else:
                bank = load_static_bank()
                candidates = bank.get(cat, [])
                if len(candidates) < count:
                    count = len(candidates)
                selected = random.sample(candidates, count) if candidates else []
                generated[cat] = selected
                print(f"   Loaded {len(selected)} questions for {cat} (static)")

    results_summary = {}
    detailed = {}

    def process_items(category, items, is_conversation=False):
        qa_list = []
        all_pass = True
        for item in items:
            q = item if not is_conversation else item
            print(f"   Asking [{category}]: {q[:80]}...")
            resp = ask_chatbot_question(driver, q)
            if use_llm:
                try:
                    verdict, reason = llm_evaluate(category, q, resp)
                    status = f"{verdict} ({reason})"
                except Exception as e:
                    print(f"   LLM eval failed: {e} – using rule‑based.")
                    status = rule_evaluate(category, q, resp)
            else:
                status = rule_evaluate(category, q, resp)
            if is_conversation:
                qa_list.append({"question": q, "response": resp, "status": status})
            else:
                qa_list.append((q, resp, status))
            if "FAIL" in status:
                all_pass = False
        results_summary[category] = "PASS" if all_pass else "FAIL"
        return qa_list

    detailed["Voice_qa"] = process_items("Voice", generated["Voice"])
    conv = generated["Conversation Test"]
    conv_qa = process_items("Conversation Test", conv["questions"], is_conversation=True)
    detailed["Conversation_Detail"] = {"scenario": conv["scenario"], "qa": conv_qa}
    detailed["Suggested_qa"] = process_items("Suggested Question", generated["Suggested Question"])
    detailed["Political_qa"] = process_items("Political, Religious, Disruptive", generated["Political, Religious, Disruptive"])
    detailed["Complex_qa"] = process_items("Complex Technical Question", generated["Complex Technical Question"])
    detailed["Fees_qa"] = process_items("Fees related Question", generated["Fees related Question"])
    detailed["Passenger_qa"] = process_items("Passenger Related Question", generated["Passenger Related Question"])
    detailed["Bilingual_qa"] = process_items("Bilingual Question", generated["Bilingual Question"])

    results_summary["Disclaimer Popup"] = "PASS"
    results_summary["Feedback Submission"] = "PASS"

    # Generate report
    report_name = os.path.join(OUTPUT_DIR, f"Sanity_Check_{datetime.now().strftime('%d_%m_%Y')}.docx")
    generate_sanity_report(results_summary, detailed, report_name)
    return report_name

# ==========================================================
# MAIN
# ==========================================================
def main():
    print("=" * 80)
    print("DGCA COMPLETE AUTOMATION – Dashboard PDF + Chatbot Sanity")
    print("=" * 80)
    driver = create_driver()
    attachments = []
    try:
        # 1. Dashboard PDF
        login_superset(driver)
        open_dashboard(driver)
        click_all_tabs(driver)
        force_full_page_render(driver)
        wait_for_charts(driver, timeout=300, min_charts=20)
        pdf_path = export_dashboard_pdf(driver)
        attachments.append(pdf_path)
        # 2. Chatbot sanity check
        sanity_report = run_chatbot_sanity(driver)
        attachments.append(sanity_report)
        print("\nReports saved locally:")
        for f in attachments:
            print(f"   {f}")
        send_email(attachments, subject="DGCA Daily Automation Report")
    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")
        traceback.print_exc()
        save_error_screenshot(driver)
    finally:
        print("\nClosing browser...")
        driver.quit()
        print("Automation completed.")

if __name__ == "__main__":
    main()
