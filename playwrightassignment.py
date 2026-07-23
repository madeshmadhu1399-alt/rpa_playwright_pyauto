from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from datetime import datetime
import json
import os

from pyautogui import sleep

today = datetime.now().strftime("%Y-%m-%d")

json_report = f"whatsapp_report_{today}.json"
excel_report = f"whatsapp_report_{today}.xlsx"

results = []

with sync_playwright() as p:
    browser = p.chromium.launch_persistent_context(
        user_data_dir="whatsapp_session",
        headless=False
    )

    page = browser.new_page()

    print("Opening WhatsApp Web...")
    page.goto("https://web.whatsapp.com")

    print("Please scan the QR code if prompted...")
    page.wait_for_selector("div[role='grid']", timeout=120000)
    print("WhatsApp Web loaded successfully.")

    screenshot = f"whatsapp_home_{today}.png"
    page.screenshot(path=screenshot)

    results.append({
        "date": today,
        "status": "Logged In",
        "title": page.title(),
        "url": page.url,
        "screenshot": screenshot
    })
    sleep(5)
    page.keyboard.press("Tab", delay=100)
    page.keyboard.press("Tab", delay=100)
    page.keyboard.press("Tab", delay=100)
    page.keyboard.press("Tab", delay=100)
    sleep(5)
    # fill the phone number and message
    page.keyboard.type("9600811370", delay=100)  # Replace with the actual phone number
    #page.wait_for_selector("selector_for_whatever_appears_next")
    sleep(5)
    # select the chat and send a message
    page.keyboard.press("Tab", delay=100)
    page.keyboard.press("Tab", delay=100)
    page.keyboard.press("Enter", delay=100)
    sleep(5)
    #select the message box and type a message
    page.keyboard.type("Hello, this is an automated message from Playwright!", delay=100)
    page.keyboard.press("Enter", delay=100)
    sleep(5)



with open(json_report, "w") as f:
    json.dump(results, f, indent=4)

wb = Workbook()
ws = wb.active
ws.append(["Date", "Status", "Title", "URL", "Screenshot"])

for item in results:
    ws.append([item["date"], item["status"], item["title"], item["url"], item["screenshot"]])

wb.save(excel_report)

print("JSON Report :", os.path.abspath(json_report))
print("Excel Report:", os.path.abspath(excel_report))
print("Completed Successfully.")

