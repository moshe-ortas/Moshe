#!/usr/bin/env python3
"""
Pazoom Invoice Checker
מוצא חשבוניות פזומט בGmail, מוריד PDF ושומר לCSV
"""

import csv
import email
import imaplib
import os
import re
import sys
from datetime import datetime
from email.header import decode_header

from dotenv import load_dotenv

load_dotenv()

GMAIL_USER = os.getenv("GMAIL_USER")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", "./invoices")
CSV_OUTPUT = os.getenv("CSV_OUTPUT", "./pazoom_invoices.csv")

IMAP_HOST = "imap.gmail.com"
IMAP_PORT = 993

# Search terms for Pazoom invoices
PAZOOM_SENDERS = ["pazoom", "pelephone", "partner"]
PAZOOM_SUBJECTS = ["חשבון", "חשבונית", "פקטורה", "invoice", "pazoom", "פזומט"]


def decode_str(value):
    if not value:
        return ""
    parts = decode_header(value)
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(part)
    return "".join(decoded)


def extract_amount(text):
    """Extract invoice amount from email body (ILS)."""
    patterns = [
        r"(?:סה\"כ|סהכ|total|לתשלום)[^\d]*(\d[\d,\.]+)\s*(?:₪|ש\"ח|NIS|ILS)?",
        r"(\d[\d,\.]+)\s*(?:₪|ש\"ח)",
        r"(?:₪|ש\"ח)\s*(\d[\d,\.]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            amount = match.group(1).replace(",", "")
            try:
                return float(amount)
            except ValueError:
                pass
    return None


def extract_invoice_number(text):
    """Extract invoice number from email body."""
    patterns = [
        r"(?:מספר חשבון(?:ית)?|חשבונית מס(?:פר)?\.?|invoice\s*#?|invoice no\.?)[^\d]*(\d[\d\-]+)",
        r"(?:מס'?|no\.?)[^\d]*(\d{5,})",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1)
    return None


def is_pazoom_email(sender, subject):
    sender_lower = sender.lower()
    subject_lower = subject.lower()
    sender_match = any(term in sender_lower for term in PAZOOM_SENDERS)
    subject_match = any(term in subject_lower for term in PAZOOM_SUBJECTS)
    return sender_match or (subject_match and "pazoom" in sender_lower)


def connect_gmail():
    if not GMAIL_USER or not GMAIL_APP_PASSWORD:
        print("Error: Missing GMAIL_USER or GMAIL_APP_PASSWORD in .env file")
        print("Copy .env.example to .env and fill in your credentials.")
        sys.exit(1)

    print(f"Connecting to Gmail as {GMAIL_USER}...")
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(GMAIL_USER, GMAIL_APP_PASSWORD)
    print("Connected successfully.")
    return mail


def search_pazoom_emails(mail):
    mail.select("INBOX")

    # Search by sender domains
    all_ids = set()
    search_queries = [
        '(FROM "pazoom")',
        '(FROM "pelephone")',
        '(SUBJECT "חשבון")',
        '(SUBJECT "חשבונית")',
        '(SUBJECT "pazoom")',
        '(SUBJECT "פזומט")',
    ]

    for query in search_queries:
        try:
            _, data = mail.search(None, query)
            if data and data[0]:
                ids = data[0].split()
                all_ids.update(ids)
        except Exception:
            # Some IMAP servers don't support Hebrew search - skip silently
            pass

    # Also search ALL and filter manually if the above returns nothing
    if not all_ids:
        print("Falling back to full inbox scan...")
        _, data = mail.search(None, "ALL")
        if data and data[0]:
            all_ids = set(data[0].split())

    return list(all_ids)


def process_email(mail, msg_id, download_dir):
    _, data = mail.fetch(msg_id, "(RFC822)")
    raw = data[0][1]
    msg = email.message_from_bytes(raw)

    sender = decode_str(msg.get("From", ""))
    subject = decode_str(msg.get("Subject", ""))
    date_str = msg.get("Date", "")

    # Parse date
    try:
        date = email.utils.parsedate_to_datetime(date_str)
        date_formatted = date.strftime("%Y-%m-%d")
    except Exception:
        date_formatted = date_str[:10] if date_str else ""

    if not is_pazoom_email(sender, subject):
        return None

    body_text = ""
    attachments = []

    for part in msg.walk():
        content_type = part.get_content_type()
        disposition = str(part.get("Content-Disposition", ""))

        if content_type == "text/plain" and "attachment" not in disposition:
            payload = part.get_payload(decode=True)
            if payload:
                body_text += payload.decode("utf-8", errors="replace")

        elif content_type == "text/html" and "attachment" not in disposition and not body_text:
            payload = part.get_payload(decode=True)
            if payload:
                # Strip basic HTML tags for text extraction
                html = payload.decode("utf-8", errors="replace")
                body_text += re.sub(r"<[^>]+>", " ", html)

        elif "attachment" in disposition or content_type == "application/pdf":
            filename = part.get_filename()
            if filename:
                filename = decode_str(filename)
                if filename.lower().endswith(".pdf") or content_type == "application/pdf":
                    safe_name = re.sub(r"[^\w\-_\. ]", "_", filename)
                    filepath = os.path.join(download_dir, f"{date_formatted}_{safe_name}")
                    os.makedirs(download_dir, exist_ok=True)
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    attachments.append(filepath)
                    print(f"  Downloaded: {filepath}")

    amount = extract_amount(body_text)
    invoice_number = extract_invoice_number(body_text)

    return {
        "date": date_formatted,
        "sender": sender,
        "subject": subject,
        "invoice_number": invoice_number or "",
        "amount_ils": amount or "",
        "attachments": "; ".join(attachments),
    }


def save_csv(invoices, path):
    os.makedirs(os.path.dirname(path) if os.path.dirname(path) else ".", exist_ok=True)
    fieldnames = ["date", "sender", "subject", "invoice_number", "amount_ils", "attachments"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(invoices)
    print(f"\nSaved {len(invoices)} invoices to {path}")


def print_table(invoices):
    if not invoices:
        print("No Pazoom invoices found.")
        return

    print(f"\n{'=' * 80}")
    print(f"{'DATE':<12} {'AMOUNT':>10}  {'INVOICE #':<15} {'SUBJECT'}")
    print(f"{'=' * 80}")
    for inv in sorted(invoices, key=lambda x: x["date"], reverse=True):
        amount = f"₪{inv['amount_ils']}" if inv["amount_ils"] else "N/A"
        print(
            f"{inv['date']:<12} {amount:>10}  {inv['invoice_number']:<15} {inv['subject'][:45]}"
        )
    print(f"{'=' * 80}")
    total = sum(float(i["amount_ils"]) for i in invoices if i["amount_ils"])
    print(f"Total invoices found: {len(invoices)}")
    if total:
        print(f"Total amount: ₪{total:,.2f}")


def main():
    mail = connect_gmail()

    print("Searching for Pazoom invoices...")
    msg_ids = search_pazoom_emails(mail)
    print(f"Scanning {len(msg_ids)} emails...")

    invoices = []
    for i, msg_id in enumerate(msg_ids, 1):
        if i % 50 == 0:
            print(f"  Processed {i}/{len(msg_ids)}...")
        try:
            result = process_email(mail, msg_id, DOWNLOAD_DIR)
            if result:
                invoices.append(result)
        except Exception as e:
            print(f"  Warning: Could not process message {msg_id}: {e}")

    mail.logout()

    print_table(invoices)

    if invoices:
        save_csv(invoices, CSV_OUTPUT)
    else:
        print(
            "\nNo Pazoom invoices found. Make sure GMAIL_USER and GMAIL_APP_PASSWORD are set correctly."
        )


if __name__ == "__main__":
    main()
