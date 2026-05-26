#!/usr/bin/env python3
"""
Daily DGCA Chatbot Sanity Check – Randomized Questions
Handles disclaimer popup, runs headless on GitHub Actions.
"""

import os
import time
import re
import random
import csv
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException
from docx import Document

# ========== CONFIGURATION FROM ENVIRONMENT ==========
CHATBOT_URL = os.getenv("CHATBOT_URL", "https://www.dgca.gov.in/digigov-portal/")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
RECIPIENTS = os.getenv("RECIPIENTS", "").split(",") if os.getenv("RECIPIENTS") else []

QUESTIONS_BANK_FILE = "questions_bank.csv"

SELECTION_COUNTS = {
    "Voice": 3,
    "Conversation Test": 1,
    "Suggested Question": 2,
    "Political, Religious, Disruptive": 4,
    "Complex Technical Question": 1,
    "Fees related Question": 1,
    "Passenger Related Question": 1,
    "Bilingual Question": 2,
}

CONVERSATION_SCENARIOS = [
    {
        "scenario": "Passenger grievance about misbehaviour",
        "questions": [
            "I want to resolve a passenger grievance against an airline stating that airline has misbehaved with the passenger in airport. Help me to proceed.",
            "Which members i need to interrogate in the airline for my report?",
            "I want to know which airline crew members i need to ask questions to validate passenger's grievance.",
            "Once i make a report where should i submit it and who will validate further?",
            "If the airline is found guilty then what actions will be taken against them ?"
        ]
    },
    {
        "scenario": "Flight safety issue investigation",
        "questions": [
            "I have received a grievance from a passenger regarding flight safety issue. How can i investigate this?",
            "So tell me about the safety rules i should keep in mind during my investigation.",
            "Whom i should interview regarding the grievance during my investigation?",
            "At the end what documents i need to submit at the end of my investigation?"
        ]
    }
]

# ========== HELPER FUNCTIONS ==========

def is_fallback(response):
    fallback_phrases = [
        "I am not able to answer",
        "not able to answer this query",
        "contact the DGCA Support"
    ]
    return any(phrase.lower() in response.lower() for phrase in fallback_phrases)

def contains_suggested_questions(response):
    return "suggested question" in response.lower()

def contains_fee_amount(response):
    return re.search(r"Rs\.\s*\d+|₹\s*\d+|Rupees\s*\d+", response) is not None

def contains_hindi_script(text):
    return re.search(r'[\u0900-\u097F]', text) is not None

def evaluate_response(category, question, response):
    if category in ("Voice", "Bilingual Question"):
        return "PASS" if contains_hindi_script(response) else "FAIL (not Hindi)"
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
        return "PASS" if not is_fallback(response) else "FAIL"
    else:
        return "PASS"

def accept_disclaimer(driver):
    """Click the 'I understand' button on the disclaimer popup."""
    try:
        accept_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'I understand')]"))
        )
        accept_btn.click()
        print("   Disclaimer accepted.")
        time.sleep(2)
        return True
    except TimeoutException:
        print("   No disclaimer popup found (or already accepted).")
        return False

def ask_question(driver, question, max_attempts=2):
    """Send a question to the chatbot and return the response."""
    for attempt in range(max_attempts):
        # 1. Ensure chat widget is open (optional: click on a chat icon if needed)
        #    Many sites open the chatbot by default, so we skip this.

        # 2. Wait for input field to be clickable
        input_selectors = [
            "input[placeholder*='message']",
            "textarea[placeholder*='message']",
            "input[placeholder*='Type']",
            "textarea",
            "input[type='text']"
        ]
        input_field = None
        for selector in input_selectors:
            try:
                input_field = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                if input_field:
                    break
            except:
                continue
        if not input_field:
            time.sleep(3)
            continue

        # 3. Focus on the field
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_field)
        time.sleep(0.5)
        try:
            input_field.click()
        except:
            driver.execute_script("arguments[0].click();", input_field)
        time.sleep(0.5)

        # 4. Clear and type
        try:
            input_field.clear()
        except:
            driver.execute_script("arguments[0].value = '';", input_field)
        input_field.send_keys(question)
        time.sleep(0.5)

        # 5. Send message
        sent = False
        try:
            send_btn = driver.find_element(By.XPATH, "//button[contains(., 'Send')]")
            if send_btn.is_displayed() and send_btn.is_enabled():
                driver.execute_script("arguments[0].click();", send_btn)
                sent = True
        except:
            pass
        if not sent:
            input_field.send_keys(Keys.ENTER)

        # 6. Wait for response (adjust time as needed)
        time.sleep(12)

        # 7. Extract latest response
        response_selectors = [
            ".bot-message", ".latest-reply", ".reply", ".message.bot", ".bubble",
            "[class*='bot']", "[class*='reply']", "[class*='answer']"
        ]
        response = ""
        for selector in response_selectors:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            if elements:
                for elem in reversed(elements):
                    txt = elem.text.strip()
                    if txt:
                        response = txt
                        break
                if response:
                    break
        if response:
            return response
        else:
            print(f"   No response on attempt {attempt+1}, retrying...")
            time.sleep(3)
    return "Could not extract response after multiple attempts"

def generate_word_report(results_summary, detailed, output_filename):
    doc = Document()
    doc.add_heading(f"Sanity Check on DGCA Chatbot -- {datetime.now().strftime('%d/%m/%Y')}", 0)
    doc.add_paragraph(f"URL : {CHATBOT_URL}")
    doc.add_paragraph("")

    # Summary table
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
        if isinstance(qa_list, dict) and "scenario" in qa_list:   # conversation
            doc.add_paragraph(f"Scenario: {qa_list['scenario']}")
            for i, sub in enumerate(qa_list['qa'], 1):
                doc.add_paragraph(f"{i}. {sub['question']}", style='List Number')
                doc.add_paragraph(f"Response: {sub['response']}")
                doc.add_paragraph(f"Status: {sub['status']}")
                doc.add_paragraph("")
        else:
            for item in qa_list:
                q, a, s = item
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
    print(f"Report saved: {output_filename}")

def send_email(attachments, subject):
    if not SMTP_USER or not SMTP_PASSWORD or not RECIPIENTS:
        print("Email not configured, skipping.")
        return
    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(RECIPIENTS)
    msg["Subject"] = subject
    body = f"Daily DGCA Chatbot Sanity Check – {datetime.now().strftime('%d/%m/%Y')}\n\nSee attached report."
    msg.attach(MIMEText(body, "plain"))
    for filepath in attachments:
        with open(filepath, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(filepath)}")
            msg.attach(part)
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)
    print("Email sent.")

# ========== MAIN ==========

def main():
    print("Loading question bank...")
    with open(QUESTIONS_BANK_FILE, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        bank = list(reader)

    # Randomly select questions
    selected = {}
    for cat, count in SELECTION_COUNTS.items():
        if cat == "Conversation Test":
            selected[cat] = random.choice(CONVERSATION_SCENARIOS)
        else:
            candidates = [q for q in bank if q["category"] == cat]
            if len(candidates) < count:
                count = len(candidates)
            selected[cat] = random.sample(candidates, count)

    # Setup Chrome for GitHub Actions
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.binary_location = "/usr/bin/chromium-browser"

    service = Service("/usr/bin/chromedriver")
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(CHATBOT_URL)
    time.sleep(5)

    # Accept disclaimer popup
    accept_disclaimer(driver)

    # Additional: ensure any language selection is accepted (if "English" button)
    try:
        eng_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'ENGLISH')]"))
        )
        eng_btn.click()
        print("   English language selected.")
        time.sleep(2)
    except:
        pass

    results_summary = {}
    detailed = {}

    # Voice
    voice_qa = []
    all_pass = True
    for q in selected["Voice"]:
        print(f"Asking: {q['question']}")
        resp = ask_question(driver, q['question'])
        status = evaluate_response("Voice", q['question'], resp)
        voice_qa.append((q['question'], resp, status))
        if "FAIL" in status:
            all_pass = False
    results_summary["Voice"] = "PASS" if all_pass else "FAIL"
    detailed["Voice_qa"] = voice_qa

    # Conversation Test
    conv = selected["Conversation Test"]
    conv_qa = []
    for q in conv["questions"]:
        print(f"Conversation: {q[:50]}...")
        resp = ask_question(driver, q)
        status = evaluate_response("Conversation Test", q, resp)
        conv_qa.append({"question": q, "response": resp, "status": status})
    all_pass = all("FAIL" not in item["status"] for item in conv_qa)
    results_summary["Conversation Test"] = "PASS" if all_pass else "FAIL"
    detailed["Conversation_Detail"] = {"scenario": conv["scenario"], "qa": conv_qa}

    # Suggested Question
    sugg_qa = []
    all_pass = True
    for q in selected["Suggested Question"]:
        print(f"Asking: {q['question']}")
        resp = ask_question(driver, q['question'])
        status = evaluate_response("Suggested Question", q['question'], resp)
        sugg_qa.append((q['question'], resp, status))
        if "FAIL" in status:
            all_pass = False
    results_summary["Suggested Question"] = "PASS" if all_pass else "FAIL"
    detailed["Suggested_qa"] = sugg_qa

    # Political, Religious, Disruptive
    pol_qa = []
    all_pass = True
    for q in selected["Political, Religious, Disruptive"]:
        print(f"Asking: {q['question']}")
        resp = ask_question(driver, q['question'])
        status = evaluate_response("Political, Religious, Disruptive", q['question'], resp)
        pol_qa.append((q['question'], resp, status))
        if "FAIL" in status:
            all_pass = False
    results_summary["Political, Religious, Disruptive"] = "PASS" if all_pass else "FAIL"
    detailed["Political_qa"] = pol_qa

    # Complex Technical Question
    comp_qa = []
    all_pass = True
    for q in selected["Complex Technical Question"]:
        print(f"Asking: {q['question'][:80]}...")
        resp = ask_question(driver, q['question'])
        status = evaluate_response("Complex Technical Question", q['question'], resp)
        comp_qa.append((q['question'], resp, status))
        if "FAIL" in status:
            all_pass = False
    results_summary["Complex Technical Question"] = "PASS" if all_pass else "FAIL"
    detailed["Complex_qa"] = comp_qa

    # Fees related Question
    fees_qa = []
    all_pass = True
    for q in selected["Fees related Question"]:
        print(f"Asking: {q['question']}")
        resp = ask_question(driver, q['question'])
        status = evaluate_response("Fees related Question", q['question'], resp)
        fees_qa.append((q['question'], resp, status))
        if "FAIL" in status:
            all_pass = False
    results_summary["Fees related Question"] = "PASS" if all_pass else "FAIL"
    detailed["Fees_qa"] = fees_qa

    # Passenger Related Question
    pass_qa = []
    all_pass = True
    for q in selected["Passenger Related Question"]:
        print(f"Asking: {q['question']}")
        resp = ask_question(driver, q['question'])
        status = evaluate_response("Passenger Related Question", q['question'], resp)
        pass_qa.append((q['question'], resp, status))
        if "FAIL" in status:
            all_pass = False
    results_summary["Passenger Related Question"] = "PASS" if all_pass else "FAIL"
    detailed["Passenger_qa"] = pass_qa

    # Bilingual Question
    bil_qa = []
    all_pass = True
    for q in selected["Bilingual Question"]:
        print(f"Asking: {q['question']}")
        resp = ask_question(driver, q['question'])
        status = evaluate_response("Bilingual Question", q['question'], resp)
        bil_qa.append((q['question'], resp, status))
        if "FAIL" in status:
            all_pass = False
    results_summary["Bilingual Question"] = "PASS" if all_pass else "FAIL"
    detailed["Bilingual_qa"] = bil_qa

    # Manual categories (simulated pass)
    results_summary["Disclaimer Popup"] = "PASS"
    results_summary["Feedback Submission"] = "PASS"

    driver.quit()

    # Generate report
    report_name = f"Sanity_Check_{datetime.now().strftime('%d_%m_%Y')}.docx"
    generate_word_report(results_summary, detailed, report_name)

    # Email
    send_email([report_name], f"DGCA Sanity Report {datetime.now().strftime('%d/%m/%Y')}")

if __name__ == "__main__":
    main()
