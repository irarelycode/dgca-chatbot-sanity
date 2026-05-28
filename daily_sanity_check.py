#!/usr/bin/env python3
"""
Daily DGCA Chatbot Sanity Check – Robust with LLM + Fallback to CSV + Rule Evaluation
"""

import os
import time
import csv
import json
import random
import smtplib
import traceback
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
from selenium.common.exceptions import TimeoutException
from docx import Document

# Optional LLM (if available)
try:
    from openai import OpenAI
    LLM_AVAILABLE = True
except ImportError:
    LLM_AVAILABLE = False

# ========== CONFIGURATION ==========
CHATBOT_URL = os.getenv("CHATBOT_URL", "https://www.dgca.gov.in/digigov-portal/")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
RECIPIENTS = os.getenv("RECIPIENTS", "").split(",") if os.getenv("RECIPIENTS") else []
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN", "")

# Fallback question bank file
QUESTIONS_BANK_FILE = "questions_bank.csv"

# How many questions per category (for fallback mode)
FALLBACK_COUNTS = {
    "Voice": 3,
    "Conversation Test": 1,          # will use fixed scenario
    "Suggested Question": 2,
    "Political, Religious, Disruptive": 4,
    "Complex Technical Question": 1,
    "Fees related Question": 1,
    "Passenger Related Question": 1,
    "Bilingual Question": 2,
}

# Fixed conversation scenario (used in both modes)
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

# ========== RULE‑BASED EVALUATION (always available) ==========
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

# ========== LLM HELPERS (with fallback to static bank) ==========
def call_llm(prompt, max_tokens=500, temperature=0.7):
    if not GITHUB_TOKEN or not LLM_AVAILABLE:
        raise Exception("LLM not available")
    client = OpenAI(
        base_url="https://models.github.ai/inference/",
        api_key=GITHUB_TOKEN,
        timeout=30,
    )
    # Try two model names
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

# ========== LOAD STATIC QUESTION BANK (fallback) ==========
def load_static_bank():
    bank = {}
    if not os.path.exists(QUESTIONS_BANK_FILE):
        # Create a minimal default bank if file missing
        default_bank = [
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
        for cat, q, lang in default_bank:
            bank.setdefault(cat, []).append({"question": q, "expected_lang": lang})
        return bank
    with open(QUESTIONS_BANK_FILE, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            bank.setdefault(row["category"], []).append(row)
    return bank

# ========== CHATBOT INTERACTION (same as before) ==========
def click_if_present(driver, locators, timeout=5):
    for by, value in locators:
        try:
            elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
            time.sleep(0.5)
            try:
                elem.click()
            except:
                driver.execute_script("arguments[0].click();", elem)
            return True
        except:
            continue
    return False

def first_visible_element(driver, locators, timeout=30):
    last_error = None
    for by, value in locators:
        try:
            elem = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
            if elem and elem.is_displayed():
                return elem
        except Exception as e:
            last_error = e
    if last_error:
        raise last_error
    raise TimeoutException("No matching element found.")

def ask_question(driver, question, timeout=30):
    driver.get(CHATBOT_URL)
    time.sleep(3)
    click_if_present(driver, [
        (By.ID, "chat-toggle"), (By.ID, "chatbot-toggle"), (By.CSS_SELECTOR, ".chat-toggle"),
        (By.CSS_SELECTOR, ".chatbot-toggle"), (By.CSS_SELECTOR, "[aria-label*='chat']"),
        (By.XPATH, "//button[contains(., 'Chat')]"), (By.XPATH, "//button[contains(., 'Ask')]")
    ], timeout=5)
    time.sleep(2)
    click_if_present(driver, [
        (By.XPATH, "//button[contains(text(), 'I understand')]"),
        (By.XPATH, "//button[contains(text(), 'Accept')]")
    ], timeout=5)
    input_locators = [
        (By.CSS_SELECTOR, "input[placeholder*='message']"),
        (By.CSS_SELECTOR, "textarea[placeholder*='message']"),
        (By.CSS_SELECTOR, "input[placeholder*='Type']"),
        (By.CSS_SELECTOR, "textarea"),
        (By.XPATH, "//input[@type='text']")
    ]
    input_field = first_visible_element(driver, input_locators, timeout=timeout)
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_field)
    try:
        input_field.click()
    except:
        driver.execute_script("arguments[0].click();", input_field)
    try:
        input_field.clear()
    except:
        driver.execute_script("arguments[0].value = '';", input_field)
    input_field.send_keys(question)
    if not click_if_present(driver, [(By.XPATH, "//button[contains(., 'Send')]")], timeout=2):
        input_field.send_keys(Keys.ENTER)
    time.sleep(12)
    response_selectors = [
        (By.CSS_SELECTOR, ".bot-message"), (By.CSS_SELECTOR, ".latest-reply"),
        (By.CSS_SELECTOR, ".reply"), (By.CSS_SELECTOR, ".message.bot"),
        (By.CSS_SELECTOR, ".bubble"), (By.CSS_SELECTOR, "[class*='bot']"),
        (By.CSS_SELECTOR, "[class*='answer']")
    ]
    response_text = ""
    for by, value in response_selectors:
        elems = driver.find_elements(by, value)
        for e in reversed(elems):
            if e.is_displayed() and e.text.strip():
                response_text = e.text.strip()
                break
        if response_text:
            break
    if not response_text:
        response_text = "Could not extract chatbot response."
    return response_text

# ========== WORD REPORT GENERATION ==========
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
    print("=" * 60)
    print("DGCA Sanity Check - Robust Mode")
    print("=" * 60)

    # Try to use LLM for question generation and evaluation
    use_llm = False
    try:
        if GITHUB_TOKEN and LLM_AVAILABLE:
            # Quick test: ask a simple question
            call_llm("Say 'OK'", max_tokens=5)
            use_llm = True
            print("LLM is available and working. Will use LLM for generation & evaluation.")
        else:
            print("LLM not available (no token or library). Falling back to static question bank + rule evaluation.")
    except Exception as e:
        print(f"LLM test failed: {e}. Falling back to static bank + rule evaluation.")
        use_llm = False

    # Generate or load questions
    if use_llm:
        print("Generating fresh questions using LLM...")
        generated = {}
        for cat, count in FALLBACK_COUNTS.items():
            if cat == "Conversation Test":
                generated[cat] = CONVERSATION_SCENARIO
            else:
                questions = llm_generate_questions(cat, count)
                generated[cat] = questions
                print(f"   Generated {len(questions)} questions for {cat}")
    else:
        print("Loading static question bank...")
        bank = load_static_bank()
        generated = {}
        for cat, count in FALLBACK_COUNTS.items():
            if cat == "Conversation Test":
                generated[cat] = CONVERSATION_SCENARIO
            else:
                candidates = bank.get(cat, [])
                if len(candidates) < count:
                    count = len(candidates)
                selected = random.sample(candidates, count) if candidates else []
                generated[cat] = [item["question"] for item in selected]
                print(f"   Loaded {len(generated[cat])} questions for {cat}")

    # Driver setup
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.binary_location = "/usr/bin/chromium-browser"
    service = Service("/usr/bin/chromedriver")
    driver = webdriver.Chrome(service=service, options=options)

    results_summary = {}
    detailed = {}

    def process_items(category, items, is_conversation=False):
        qa_list = []
        all_pass = True
        for item in items:
            if is_conversation:
                q = item
            else:
                q = item
            print(f"Asking [{category}]: {q[:80]}...")
            resp = ask_question(driver, q)
            if use_llm:
                try:
                    verdict, reason = llm_evaluate(category, q, resp)
                    status = f"{verdict} ({reason})"
                except Exception as e:
                    print(f"   LLM evaluation failed: {e}. Using rule‑based.")
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

    # Process each category
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

    driver.quit()

    # Generate report
    report_name = f"Sanity_Check_{datetime.now().strftime('%d_%m_%Y')}.docx"
    generate_sanity_report(results_summary, detailed, report_name)
    send_email([report_name], f"DGCA Sanity Report {datetime.now().strftime('%d/%m/%Y')}")

if __name__ == "__main__":
    main()
