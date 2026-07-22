#!/usr/bin/env python3
"""Safe WhatsApp Web helper using Playwright.

This script demonstrates how to:
- launch WhatsApp Web in a Playwright browser
- wait for the WhatsApp UI to become ready
- read contact names from an Excel workbook
- search for contacts and open their chat
- take screenshots of the WhatsApp interface
- save a JSON and XLSX report of the browser session

This implementation intentionally avoids sending automated WhatsApp messages to arbitrary contacts.
"""

from __future__ import annotations

import argparse
import json
import random
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from openpyxl import Workbook, load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, sync_playwright

DEFAULT_CONTACTS_FILE = "contacts.xlsx"
DEFAULT_OUTPUT_DIR = "reports"
DEFAULT_MAX_CONTACTS = 5

SEARCH_INPUT_SELECTORS = [
    'div[contenteditable="true"][data-tab="3"]',
    'div[title="Search input textbox"]',
    'div[role="textbox"][contenteditable="true"]',
]

CONTACT_RESULT_SELECTOR = 'span[title]'


def random_delay(min_seconds: float = 0.8, max_seconds: float = 2.0) -> float:
    return random.uniform(min_seconds, max_seconds)


def normalize_filename(value: str, max_length: int = 80) -> str:
    cleaned = "".join(
        ch if ch.isalnum() or ch in "-_ ." else "_" for ch in value
    ).strip()
    if len(cleaned) > max_length:
        cleaned = cleaned[:max_length].rstrip("_.-")
    return cleaned or "item"


def load_contacts(workbook_path: Path) -> List[Dict[str, str]]:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Contacts workbook not found: {workbook_path}")

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    worksheet = workbook.active

    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Contacts workbook is empty")

    headers = [str(value).strip().lower() if value is not None else "" for value in rows[0]]
    if not any(headers):
        raise ValueError("Contacts workbook header row is empty")

    def normalize_key(value: str) -> str:
        value = value.strip().lower()
        if value in {"name", "contact", "contact name", "person"}:
            return "name"
        if value in {"number", "phone", "phone number"}:
            return "number"
        if value in {"message", "text", "note"}:
            return "message"
        return value

    normalized_headers = [normalize_key(header) for header in headers]
    contacts: List[Dict[str, str]] = []

    for row in rows[1:]:
        if row is None:
            continue
        record: Dict[str, str] = {}
        for key, value in zip(normalized_headers, row):
            if not key:
                continue
            record[key] = str(value).strip() if value is not None else ""
        if record.get("name") or record.get("number"):
            contacts.append(record)

    return contacts


def find_search_box(page) -> None:
    for selector in SEARCH_INPUT_SELECTORS:
        try:
            page.wait_for_selector(selector, timeout=30000)
            return selector
        except PlaywrightTimeoutError:
            continue
    raise RuntimeError(
        "Unable to find WhatsApp search box. Ensure you are logged into WhatsApp Web and the page has loaded."
    )


def wait_for_whatsapp_ready(page) -> None:
    try:
        selector = find_search_box(page)
        page.wait_for_selector(selector, timeout=60000)
    except PlaywrightTimeoutError as exc:
        raise RuntimeError(
            "WhatsApp Web did not become ready in time. Scan the QR code and wait for the chat list to appear."
        ) from exc


def search_contact_and_open_chat(page, contact_query: str) -> bool:
    selector = find_search_box(page)
    search_box = page.locator(selector)
    search_box.click()

    modifier = "Meta" if sys.platform == "darwin" else "Control"
    page.keyboard.press(f"{modifier}+A")
    page.keyboard.press("Backspace")
    search_box.type(contact_query, delay=random.randint(50, 120))
    time.sleep(random_delay(0.3, 0.6))

    page.wait_for_timeout(1000)

    results = page.locator(f'{CONTACT_RESULT_SELECTOR}:has-text("{contact_query}")')
    if results.count() == 0:
        results = page.locator(f'text="{contact_query}"')

    if results.count() == 0:
        return False

    try:
        results.first.click()
    except PlaywrightTimeoutError:
        return False

    page.wait_for_timeout(random_delay(1.2, 2.5))
    return True


def capture_screenshot(page, destination: Path) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(destination), full_page=True)
    return destination


def save_json_report(report: Dict, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, ensure_ascii=False)


def save_excel_report(report: Dict, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "WhatsApp Report"
    headers = ["Contact Name", "Search Query", "Status", "Screenshot", "Notes"]
    worksheet.append(headers)

    for entry in report.get("entries", []):
        worksheet.append([
            entry.get("contact_name", ""),
            entry.get("search_query", ""),
            entry.get("status", ""),
            entry.get("screenshot", ""),
            entry.get("notes", ""),
        ])

    workbook.save(destination)


def build_report(contacts: List[Dict[str, str]], output_dir: Path, max_contacts: int, headless: bool) -> Dict:
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    report_dir = output_dir.resolve()
    report_file_base = f"whatsapp_report_{timestamp}"

    report: Dict = {
        "generated_at": datetime.now().isoformat(),
        "contacts_file": str(DEFAULT_CONTACTS_FILE),
        "headless": headless,
        "max_contacts": max_contacts,
        "entries": [],
    }

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=headless, args=["--start-maximized"])
        try:
            context = browser.new_context(viewport={"width": 1280, "height": 900})
            page = context.new_page()
            page.goto("https://web.whatsapp.com")
            print("WhatsApp Web opened. Please scan the QR code if needed.")
            wait_for_whatsapp_ready(page)
            print("WhatsApp Web is ready.")
            capture_screenshot(page, report_dir / f"{report_file_base}_whatsapp_home.png")

            for contact in contacts[:max_contacts]:
                search_query = contact.get("name") or contact.get("number") or ""
                if not search_query:
                    continue

                contact_name = contact.get("name") or contact.get("number")
                entry: Dict[str, str] = {
                    "contact_name": contact_name,
                    "search_query": search_query,
                    "status": "pending",
                    "screenshot": "",
                    "notes": "",
                }
                try:
                    print(f"Searching for contact: {search_query}")
                    opened = search_contact_and_open_chat(page, search_query)
                    if not opened:
                        entry["status"] = "not_found"
                        entry["notes"] = "Contact did not appear in search results."
                        print(f"Contact not found: {search_query}")
                    else:
                        entry["status"] = "opened"
                        screenshot_name = f"{report_file_base}_{normalize_filename(search_query)}.png"
                        screenshot_path = report_dir / screenshot_name
                        capture_screenshot(page, screenshot_path)
                        entry["screenshot"] = str(screenshot_path)
                        entry["notes"] = "Chat opened and screenshot captured."
                        print(f"Opened chat for: {search_query}")
                except Exception as exc:
                    entry["status"] = "error"
                    entry["notes"] = str(exc)
                    print(f"Error while processing {search_query}: {exc}")
                finally:
                    report["entries"].append(entry)
                    page.goto("https://web.whatsapp.com")
                    wait_for_whatsapp_ready(page)
                    time.sleep(random_delay(0.8, 1.5))

        finally:
            browser.close()

    report_path_json = report_dir / f"{report_file_base}.json"
    report_path_xlsx = report_dir / f"{report_file_base}.xlsx"
    save_json_report(report, report_path_json)
    save_excel_report(report, report_path_xlsx)
    print(f"Saved report: {report_path_json}")
    print(f"Saved report: {report_path_xlsx}")
    return report


def create_sample_contacts_file(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Contacts"
    worksheet.append(["Name", "Number", "Message"])
    worksheet.append(["Friend Name", "+1234567890", "Hello! This is a sample message."])
    worksheet.append(["Family Member", "+1987654321", "Automated message template."])
    workbook.save(path)
    print(f"Created sample contacts file at: {path}")


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build a safe WhatsApp Web helper that opens contact chats and captures reports."
    )
    parser.add_argument(
        "--contacts",
        default=DEFAULT_CONTACTS_FILE,
        help="Path to the contacts Excel file. The workbook should include Name or Number columns.",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_DIR,
        help="Directory where JSON/XLSX reports and screenshots are saved.",
    )
    parser.add_argument(
        "--max-contacts",
        type=int,
        default=DEFAULT_MAX_CONTACTS,
        help="Maximum number of contacts to process in a single run.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run the browser in headless mode. Note: WhatsApp Web login may be difficult without a visible browser.",
    )
    parser.add_argument(
        "--create-sample-contacts",
        action="store_true",
        help="Create a sample contacts.xlsx file and exit.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_arguments()
    output_dir = Path(args.output_dir)
    contacts_path = Path(args.contacts)

    if args.create_sample_contacts:
        create_sample_contacts_file(contacts_path)
        return 0

    if not contacts_path.exists():
        print(f"Contacts file not found: {contacts_path}")
        print("Use --create-sample-contacts to generate a sample contacts.xlsx file.")
        return 1

    contacts = load_contacts(contacts_path)
    if not contacts:
        print("No contacts found in the contacts workbook.")
        return 1

    try:
        build_report(contacts, output_dir, args.max_contacts, args.headless)
    except Exception as exc:
        print(f"Fatal error: {exc}")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
