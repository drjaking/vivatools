"""
intexaminer_email1.py (v6: flag deferred in subject)
============================================

Mail-merge for internal examiners with thesis title, deferred flag, and top-5 categories.
Inserts rendered DOCX template into email body and annotates subject for deferred students.
"""
import argparse
import logging
from pathlib import Path

import psycopg2
from docxtpl import DocxTemplate
import win32com.client as win32

# ---------------------------------------------------------------------------
# Fetch internal assignments with peer examiner, thesis title, deferred flag, and top-5 categories
# ---------------------------------------------------------------------------
def fetch_internal_assignments(cur):
    cur.execute("""
        SELECT
          ia.student_id,
          s.full_name       AS student_name,
          s.email           AS student_email,
          s.project_title   AS thesis_title,
          COALESCE(s.deferred, FALSE) AS is_deferred,
          ia.examiner_id    AS examiner_id,
          i.full_name       AS examiner_name,
          i.email           AS examiner_email,
          ia.role           AS role,
          ea.examiner_id    AS peer_id,
          e.full_name       AS peer_examiner_name,
          e.email           AS peer_examiner_email,
          t5.top_5_categories
        FROM examiner_assignments ia
        JOIN students s ON s.student_id = ia.student_id
        JOIN examiners i ON i.examiner_id = ia.examiner_id
        LEFT JOIN examiner_assignments ea
          ON ea.student_id = ia.student_id
         AND ea.role = 'external'
        LEFT JOIN examiners e ON e.examiner_id = ea.examiner_id
        LEFT JOIN LATERAL (
            SELECT string_agg(sub.name, ', ') AS top_5_categories
            FROM (
                SELECT c.name
                FROM thesis_categories tc
                JOIN categories c ON c.category_id = tc.category_id
                WHERE tc.student_id = s.student_id AND tc.in_scope
                ORDER BY c.default_weight DESC
                LIMIT 5
            ) AS sub
        ) AS t5 ON TRUE
        WHERE ia.role = 'internal'
    """)
    return cur.fetchall()

# ---------------------------------------------------------------------------
# Merge and send
# ---------------------------------------------------------------------------
def process_records(records, template_path: Path, send=False):
    outlook = win32.Dispatch('Outlook.Application')
    for (_sid, student_name, student_email, thesis_title, is_deferred,
         examiner_id, examiner_name, examiner_email, role,
         peer_id, peer_name, peer_email, top5) in records:
        # Render template
        doc = DocxTemplate(str(template_path))
        context = {
            'examiner_name': examiner_name,
            'examiner_email': examiner_email,
            'student_name': student_name,
            'student_email': student_email,
            'thesis_title': thesis_title or '',
            'role': role.title(),
            'peer_examiner_name': peer_name or '',
            'peer_examiner_email': peer_email or '',
            'top_5_categories': top5 or ''
        }
        doc.render(context)
        body = "\n".join(p.text for p in doc.docx.paragraphs)

        # Subject with deferred flag
        subject = f"Viva Assignment: {student_name}"
        if is_deferred:
            subject += " (deferred)"

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
        description="Mail-merge for internal examiners with deferred annotation."
    )
    parser.add_argument('--dsn', required=True, help='PostgreSQL DSN')
    parser.add_argument('--template', required=True, type=Path, help='Path to DOCX template')
    parser.add_argument('--send', action='store_true', help='Actually send emails')
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO)
    with psycopg2.connect(args.dsn) as conn:
        cur = conn.cursor()
        records = fetch_internal_assignments(cur)

    process_records(records, args.template, send=args.send)

if __name__ == '__main__':
    main()
