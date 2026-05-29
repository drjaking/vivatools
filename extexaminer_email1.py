"""
extexaminer_email1.py
======================

Mail-merge for external examiners with thesis title, supervisors, and deferred flag.
Inserts a plain-text formatting table into the email body and groups by examiner.
"""
import argparse
import logging
from pathlib import Path

import psycopg2
from docxtpl import DocxTemplate
import win32com.client as win32

# ---------------------------------------------------------------------------
# Fetch external assignments with thesis title, deferred flag, and supervisors
# ---------------------------------------------------------------------------
def fetch_external_assignments(cur):
    cur.execute("""
        SELECT
          ia.student_id,
          s.full_name       AS student_name,
          s.email           AS student_email,
          s.project_title   AS thesis_title,
          COALESCE(s.deferred, FALSE) AS is_deferred,
          ia.examiner_id    AS examiner_id,
          e.full_name       AS examiner_name,
          e.email           AS examiner_email,
          sup.supervisors   AS supervisors
        FROM examiner_assignments ia
        JOIN students s ON s.student_id = ia.student_id
        JOIN examiners e ON e.examiner_id = ia.examiner_id
        LEFT JOIN student_supervisors sup ON sup.student_id = ia.student_id
        WHERE ia.role = 'external'
        ORDER BY e.full_name, s.full_name
    """)
    return cur.fetchall()

# ---------------------------------------------------------------------------
# Merge and send
# ---------------------------------------------------------------------------
def process_records(records, template_path: Path, send=False):
    from collections import defaultdict
    grouped = defaultdict(list)
    for row in records:
        # row: (student_id, student_name, student_email, thesis_title, is_deferred, examiner_id, examiner_name, examiner_email, supervisors)
        eid = (row[5], row[6], row[7]) # (examiner_id, examiner_name, examiner_email)
        grouped[eid].append(row)
        
    outlook = win32.Dispatch('Outlook.Application')
    for (examiner_id, examiner_name, examiner_email), assignments in grouped.items():
        # Render template
        doc = DocxTemplate(str(template_path))
        
        tbl_rows = []
        for row in assignments:
            is_def = row[4]
            def_status = "Deferred" if is_def else "submitting"
            tbl_rows.append({
                'student_name': row[1],
                'thesis_title': row[3] or '[No Thesis Title]',
                'supervisors': row[8] or '[No Supervisors Listed]',
                'deferred_status': def_status
            })
            
        context = {
            'examiner_name': examiner_name,
            'tbl_rows': tbl_rows
        }
        doc.render(context)
        
        # Build plain text body from paragraphs
        body_paras = [p.text for p in doc.docx.paragraphs if p.text.strip()]
        body = "\n\n".join(body_paras)
        
        # Format clean plain text table to append to the email body
        table_text = "\n\n" + f"{'Student Name':<25} | {'Thesis Title':<45} | {'Supervisors':<30} | {'Status':<12}\n"
        table_text += "=" * 118 + "\n"
        for r in tbl_rows:
            s_name = r['student_name'][:25]
            t_title = r['thesis_title'][:45]
            sups = r['supervisors'][:30]
            status = r['deferred_status']
            table_text += f"{s_name:<25} | {t_title:<45} | {sups:<30} | {status:<12}\n"
            
        body += table_text
        
        subject = "Provisional DClinPsy Viva Examiner Allocations"

        if not send:
            print("---")
            print(f"To: {examiner_email}")
            print(f"Subject: {subject}\n")
            print(body)
            continue

        mail = outlook.CreateItem(0)
        mail.To = examiner_email
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        logging.info(f"Sent email to {examiner_name} <{examiner_email}>")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="Mail-merge for external examiners with consolidated allocations."
    )
    parser.add_argument('--dsn', required=True, help='PostgreSQL DSN')
    parser.add_argument('--template', required=True, type=Path, help='Path to DOCX template')
    parser.add_argument('--send', action='store_true', help='Actually send emails')
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO)
    with psycopg2.connect(args.dsn) as conn:
        cur = conn.cursor()
        records = fetch_external_assignments(cur)

    process_records(records, args.template, send=args.send)

if __name__ == '__main__':
    main()
