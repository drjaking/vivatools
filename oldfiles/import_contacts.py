"""
import_contacts.py
==================

Imports populated CSVs from `./contact_csvs/` into the database:
  • `contact_csvs/students.csv`           (student_id, full_name, email, project_title)
  • `contact_csvs/internal_examiners.csv` (examiner_id, full_name, email)
  • `contact_csvs/external_examiners.csv` (examiner_id, full_name, email)

Usage:
  python import_contacts.py --dsn "dbname=vivas user=… host=…"

This script will:
 1. Update `students.email` and, if present, `students.project_title`.
 2. Update `examiners.email` for all examiners.

Requires:
  • psycopg2
"""
import csv
import logging
from pathlib import Path

import psycopg2

# Hard-coded CSV paths
CONTACT_DIR = Path('contact_csvs')
STUDENTS_CSV = CONTACT_DIR / 'students.csv'
INTERNAL_CSV = CONTACT_DIR / 'internal_examiners.csv'
EXTERNAL_CSV = CONTACT_DIR / 'external_examiners.csv'


def update_students(conn):
    cur = conn.cursor()
    # Ensure email and project_title columns exist
    cur.execute("ALTER TABLE students ADD COLUMN IF NOT EXISTS email TEXT")
    cur.execute("ALTER TABLE students ADD COLUMN IF NOT EXISTS project_title TEXT")
    if not STUDENTS_CSV.exists():
        logging.error(f"Students CSV not found at {STUDENTS_CSV}")
        return
    cur = conn.cursor()
    if not STUDENTS_CSV.exists():
        logging.error(f"Students CSV not found at {STUDENTS_CSV}")
        return
    with STUDENTS_CSV.open(newline='', encoding='utf-8', errors='replace') as f:
        reader = csv.DictReader(f)
        # Ensure project_title column exists
        cur.execute("ALTER TABLE students ADD COLUMN IF NOT EXISTS project_title TEXT")
        for row in reader:
            sid = row['student_id']
            email = row['email'].strip() or None
            title = row.get('project_title', '').strip() or None
            cur.execute(
                "UPDATE students SET email = %s, project_title = %s WHERE student_id = %s",
                (email, title, sid)
            )
    conn.commit()
    cur.close()
    logging.info("Updated students from %s", STUDENTS_CSV)


def update_examiners(conn, csv_path: Path):
    cur = conn.cursor()
    # Ensure email column exists
    cur.execute("ALTER TABLE examiners ADD COLUMN IF NOT EXISTS email TEXT")
    if not csv_path.exists():
        logging.error(f"Examiners CSV not found at {csv_path}")
        return
    with csv_path.open(newline='', encoding='utf-8', errors='replace') as f:
        reader = csv.DictReader(f)
        for row in reader:
            eid = row['examiner_id']
            email = row['email'].strip() or None
            cur.execute(
                "UPDATE examiners SET email = %s WHERE examiner_id = %s",
                (email, eid)
            )
    conn.commit()
    cur.close()
    logging.info("Updated examiners from %s", csv_path)


def main():
    logging.basicConfig(level=logging.INFO)
    import argparse
    parser = argparse.ArgumentParser(
        description='Import contact emails from hard-coded CSV paths.'
    )
    parser.add_argument('--dsn', required=True, help='PostgreSQL DSN')
    args = parser.parse_args()

    with psycopg2.connect(args.dsn) as conn:
        update_students(conn)
        update_examiners(conn, INTERNAL_CSV)
        update_examiners(conn, EXTERNAL_CSV)

if __name__ == '__main__':
    main()
