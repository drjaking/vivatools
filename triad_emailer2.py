"""
triad_emailer.py (v3: fix tuple unpack + deferred template)
==========================================================

* Replaces the invalid `*rest, is_deferred, *_tail` unpack (which caused the
  `SyntaxError`) with a simpler index‑based approach.
* Builds a `processed` list via list‑comprehension: `(rec, pick_template(rec[4]))`.
* Ensures newline handling in `body` and preview print.
"""
import argparse
import logging
from pathlib import Path

import psycopg2
from docxtpl import DocxTemplate
import win32com.client as win32

ADMIN_ADDR = "cehp.dclinpsyresearchsupport@ucl.ac.uk"

# ---------------------------------------------------------------------------
# Data fetch
# ---------------------------------------------------------------------------

def fetch_triads(cur):
    cur.execute(
        """
        SELECT
          s.student_id,
          s.full_name                    AS student_name,
          s.email                        AS student_email,
          s.project_title                AS thesis_title,
          COALESCE(s.deferred,FALSE)     AS is_deferred,
          i.full_name                    AS internal_name,
          i.email                        AS internal_email,
          e.full_name                    AS external_name,
          e.email                        AS external_email,
          t5.top_5_categories
        FROM students s
        JOIN examiner_assignments ia ON ia.student_id = s.student_id AND ia.role='internal'
        JOIN examiners            i ON i.examiner_id  = ia.examiner_id
        JOIN examiner_assignments ea ON ea.student_id = s.student_id AND ea.role='external'
        JOIN examiners            e ON e.examiner_id  = ea.examiner_id
        LEFT JOIN LATERAL (
            SELECT string_agg(c.name, ', ') AS top_5_categories
            FROM (
                SELECT c.name
                FROM thesis_categories tc
                JOIN categories c ON c.category_id = tc.category_id
                WHERE tc.student_id = s.student_id AND tc.in_scope
                ORDER BY c.default_weight DESC
                LIMIT 5
            ) c
        ) t5 ON TRUE
        ORDER BY s.full_name;
        """
    )
    return cur.fetchall()

# ---------------------------------------------------------------------------
# Merge + send/preview helper
# ---------------------------------------------------------------------------

def render_email(record: tuple, template_path: Path) -> tuple[str, str, str]:
    """Return (to_addr, subject, body) for the given record."""
    (sid, student_name, student_email, thesis_title, is_def,
     internal_name, internal_email, external_name, external_email, top5) = record

    doc = DocxTemplate(str(template_path))
    ctx = dict(student_name=student_name,
               student_email=student_email,
               internal_name=internal_name,
               internal_email=internal_email,
               external_name=external_name,
               external_email=external_email,
               thesis_title=thesis_title or '',
               top_5_categories=top5 or '',
               deferred="Yes" if is_def else "No")
    doc.render(ctx)
    body = "\n\n".join(p.text for p in doc.docx.paragraphs if p.text.strip())

    subj = "IMPORTANT - UCL DClinPsy Thesis allocations" + (" (deferred)" if is_def else "")
    to_addr = "; ".join(e for e in [internal_email, external_email, student_email] if e)
    return to_addr, subj, body

# ---------------------------------------------------------------------------
# Main driver
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Triad mail‑merge emailer")
    parser.add_argument('--dsn', required=True)
    parser.add_argument('--template', required=True, type=Path, help='Standard DOCX template')
    parser.add_argument('--template-deferred', type=Path, help='Deferred DOCX template')
    parser.add_argument('--send', action='store_true')
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO)

    with psycopg2.connect(args.dsn) as conn:
        recs = fetch_triads(conn.cursor())

    # Choose template path per record
    processed = [(rec, (args.template_deferred if (rec[4] and args.template_deferred) else args.template))
                 for rec in recs]

    outlook = win32.Dispatch('Outlook.Application')
    for rec, tmpl in processed:
        to_addr, subject, body = render_email(rec, tmpl)
        if not args.send:
            print("---")
            print(f"To : {to_addr}\nCC : {ADMIN_ADDR}\nSubj: {subject}\n")
            print(body[:500] + ("…" if len(body) > 500 else ""))
            continue

        msg = outlook.CreateItem(0)
        msg.To = to_addr
        msg.CC = ADMIN_ADDR
        msg.Subject = subject
        msg.Body = body
        msg.Send()
        logging.info("Sent triad email for student %s", rec[1])

if __name__ == '__main__':
    main()
