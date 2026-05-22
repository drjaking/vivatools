"""
examiner_mailmerge.py (v2: include peer examiner in email)
==========================================================

Performs a mail-merge for examiners (internal and external) using a
`.docx` template. The filled template is inserted into the email body
instead of as an attachment.

Placeholders in the template must match the keys below:
  - {{examiner_name}}
  - {{examiner_email}}
  - {{student_name}}
  - {{student_email}}
  - {{role}}                   # 'Internal' or 'External'
  - {{peer_examiner_name}}     # the other examiner
  - {{peer_examiner_email}}

Dependencies
------------
• psycopg2
• docxtpl (pip install docxtpl)
• pywin32

Usage
-----
python examiner_mailmerge.py \
  --dsn "dbname=vivas user=… host=…" \
  --template path/to/template.docx \
  [--send]

By default runs as dry-run, printing To/CC/Subject/Body. Use `--send` to deliver.
"""
import argparse
import logging
from pathlib import Path

import psycopg2
from docxtpl import DocxTemplate
import win32com.client as win32

# ---------------------------------------------------------------------------
# Fetch assignments along with peer examiner
# ---------------------------------------------------------------------------
def fetch_examiners(cur):
    cur.execute("""
        SELECT
          ea1.student_id,
          s.full_name    AS student_name,
          s.email        AS student_email,
          ea1.examiner_id,
          e1.full_name   AS examiner_name,
          e1.email       AS examiner_email,
          ea1.role       AS role,
          ea2.examiner_id    AS peer_id,
          e2.full_name       AS peer_examiner_name,
          e2.email           AS peer_examiner_email
        FROM examiner_assignments ea1
        JOIN students s ON s.student_id = ea1.student_id
        JOIN examiners e1 ON e1.examiner_id = ea1.examiner_id
        LEFT JOIN examiner_assignments ea2
          ON ea2.student_id = ea1.student_id
         AND ea2.role <> ea1.role
        LEFT JOIN examiners e2 ON e2.examiner_id = ea2.examiner_id
        WHERE ea1.role IN ('internal','external')
    """)
    return cur.fetchall()

# ---------------------------------------------------------------------------
# Merge and send
# ---------------------------------------------------------------------------
def process_records(records, template_path: Path, send=False):
    outlook = win32.Dispatch('Outlook.Application')
    for (_sid, student_name, student_email,
         examiner_id, examiner_name, examiner_email, role,
         peer_id, peer_name, peer_email) in records:
        # Load and render the DOCX template
        doc = DocxTemplate(str(template_path))
        context = {
            'examiner_name': examiner_name,
            'examiner_email': examiner_email,
            'student_name': student_name,
            'student_email': student_email,
            'role': role.title(),
            'peer_examiner_name': peer_name or '',
            'peer_examiner_email': peer_email or '',
        }
        doc.render(context)
        # Extract rendered text for email body
        paragraphs = [p.text for p in doc.docx.paragraphs]
        body = "\n".join(paragraphs)

        # Email metadata
        subject = f"Viva Assignment: {student_name} ({role.title()} Examiner)"
        cc = student_email

        if not send:
            print("---")
            print(f"To: {examiner_email}")
            print(f"CC: {cc}")
            print(f"Subject: {subject}\n")
            print(body)
            continue

        # Send via Outlook
        mail = outlook.CreateItem(0)
        mail.To = examiner_email
        mail.CC = cc
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        logging.info(f"Sent email to {examiner_name} <{examiner_email}>")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Mail-merge Viva examiner notifications.")
    parser.add_argument('--dsn', required=True, help='PostgreSQL DSN')
    parser.add_argument('--template', required=True, type=Path, help='Path to .docx template')
    parser.add_argument('--send', action='store_true', help='Actually send emails')
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO)
    with psycopg2.connect(args.dsn) as conn:
        cur = conn.cursor()
        records = fetch_examiners(cur)

    process_records(records, args.template, send=args.send)

if __name__ == '__main__':
    main()
