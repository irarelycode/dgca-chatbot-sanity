"""
Superset Dashboard – Full page screenshot (reliable fallback)
"""

import time
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image

from utils import (
    create_driver, wait, safe_click, OUTPUT_DIR,
    SUPERSET_HOME_URL, DASHBOARD_NAME, USERNAME, PASSWORD
)

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
        except:
            pass

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
        except:
            time.sleep(15)
            continue
        print(f"   Rendered chart objects: {count}")
        if count >= min_charts:
            print("   Charts fully loaded.")
            return True
        time.sleep(15)
    print("   Chart loading timeout reached.")
    return False

def dashboard_full_page_screenshot(driver):
    """Take a full-page screenshot and convert to PDF."""
    print("\n6. Taking full-page screenshot of dashboard...")
    # Get full page dimensions
    total_width = driver.execute_script("return Math.max(document.body.scrollWidth, document.documentElement.scrollWidth);")
    total_height = driver.execute_script("return Math.max(document.body.scrollHeight, document.documentElement.scrollHeight);")
    print(f"   Page size: {total_width} x {total_height}")
    
    # Temporarily set window size to full page
    original_size = driver.get_window_size()
    driver.set_window_size(max(total_width, 1600), max(total_height, 1200))
    time.sleep(3)
    
    # Take screenshot
    screenshot_path = os.path.join(OUTPUT_DIR, f"dashboard_screenshot_{int(time.time())}.png")
    driver.save_screenshot(screenshot_path)
    print(f"   Screenshot saved: {screenshot_path}")
    
    # Convert to PDF
    pdf_path = screenshot_path.replace(".png", ".pdf")
    image = Image.open(screenshot_path)
    image.convert("RGB").save(pdf_path)
    print(f"   PDF saved: {pdf_path}")
    
    # Restore window size
    driver.set_window_size(original_size['width'], original_size['height'])
    return pdf_path

def main():
    driver = create_driver()
    try:
        login_superset(driver)
        open_dashboard(driver)
        click_all_tabs(driver)
        force_full_page_render(driver)
        wait_for_charts(driver, timeout=300, min_charts=20)
        pdf_path = dashboard_full_page_screenshot(driver)
        print(f"\nDashboard PDF saved: {pdf_path}")
        return pdf_path
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
