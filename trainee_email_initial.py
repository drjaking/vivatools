"""
trainee_email_initial.py (v3)
==============================

Sends initial notification emails to each student (trainee), listing assigned
examiners. For students marked deferred, adds an extra note about deferred status.

This script supports a dry-run mode (default) and an explicit send mode.

Requirements
------------
• Python with pywin32 (`pip install pywin32`)
• psycopg2 for PostgreSQL access
• Windows with Outlook configured

Usage
-----
# Dry run (default)
python trainee_email_initial.py --dsn "dbname=vivas user=… host=…"

# Actually send emails
python trainee_email_initial.py --dsn "…" --send

Options
-------
--send       Send emails via Outlook instead of printing them (default: dry-run)
"""
import argparse
import logging
from pathlib import Path

import psycopg2
from docxtpl import DocxTemplate
import win32com.client as win32

# ---------------------------------------------------------------------------
# Fetch assignments, contact info, and deferred status
# ---------------------------------------------------------------------------
def fetch_student_assignments(cur):
    cur.execute("""
        SELECT
            s.student_id,
            s.full_name   AS student_name,
            s.email       AS student_email,
            COALESCE(s.deferred, FALSE) AS is_deferred,
            i.full_name   AS internal_name,
            i.email       AS internal_email,
            e.full_name   AS external_name,
            e.email       AS external_email
        FROM examiner_assignments ea
        JOIN students s ON s.student_id = ea.student_id
        LEFT JOIN examiners i ON i.examiner_id = (
            SELECT examiner_id FROM examiner_assignments
             WHERE student_id = s.student_id AND role = 'internal'
        )
        LEFT JOIN examiners e ON e.examiner_id = (
            SELECT examiner_id FROM examiner_assignments
             WHERE student_id = s.student_id AND role = 'external'
        )
        WHERE ea.role IN ('internal', 'external')
        GROUP BY s.student_id, s.full_name, s.email, s.deferred,
                 i.full_name, i.email, e.full_name, e.email
    """)
    return cur.fetchall()

# ---------------------------------------------------------------------------
# Compose and send or print emails, including deferred note
# ---------------------------------------------------------------------------
def send_notifications(records, template_path=None, send=False):
    outlook = win32.Dispatch('Outlook.Application')
    for (_sid, student_name, student_email, is_deferred,
         int_name, _int_email, ext_name, _ext_email) in records:
        subject = "Your Viva Examiners (provisional)"
        
        if template_path and template_path.exists():
            try:
                doc = DocxTemplate(str(template_path))
                context = {
                    'student_name': student_name,
                    'internal_name': int_name,
                    'external_name': ext_name,
                    'is_deferred': is_deferred
                }
                doc.render(context)
                body = "\n\n".join(p.text for p in doc.docx.paragraphs if p.text.strip())
            except Exception as e:
                logging.error(f"Error rendering template, falling back to text: {e}")
                template_path = None # Force fallback
                
        if not template_path or not template_path.exists():
            lines = [
                f"Dear {student_name},",
                "",
                "You have been provisionally assigned the following viva examiners:",
                f"  • Internal Examiner: {int_name}",
                f"  • External Examiner: {ext_name}",
                "",
                "If you believe there is any conflict of interest (for example, an examiner who \
 has supervised your project), please contact john.king@ucl.ac.uk as soon as possible.",
                "",
                "We don't need to hear about tutors, proposal reviewers, or stats demonstrators.",
                ""
            ]
            if is_deferred:
                lines.extend([
                    "Note: your deferred status is not affected by this allocation; \
these are the examiners who will conduct your viva when the time comes.",
                    ""
                ])
            lines.extend(["Kind regards,", "John"])
            body = "\n".join(lines)

        if not send:
            print("---")
            print(f"To: {student_email}")
            print(f"Subject: {subject}\n")
            print(body)
            continue

        mail = outlook.CreateItem(0)
        mail.To = student_email
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        logging.info(f"Sent email to {student_name} <{student_email}>")

# ---------------------------------------------------------------------------
# Main entrypoint
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Send initial viva assignment emails to trainees."
    )
    parser.add_argument('--dsn', required=True, help='PostgreSQL DSN')
    parser.add_argument('--template', type=Path, default=Path('templates/Trainee_Initial_Email.docx'), help='Path to DOCX template')
    parser.add_argument('--send', action='store_true', help='Actually send emails (default: dry-run)')
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO)
    with psycopg2.connect(args.dsn) as conn:
        cur = conn.cursor()
        records = fetch_student_assignments(cur)

    send_notifications(records, template_path=args.template, send=args.send)

if __name__ == '__main__':
    main()
