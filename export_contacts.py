"""
export_contacts.py (v5)
======================

Exports CSV files for students and examiners. If the pywin32 Outlook API
is available, it auto-populates emails via Contacts and GAL; otherwise, it
defaults to empty emails with a warning.

Outputs:
  • students.csv           (student_id, full_name, email, project_title)
  • internal_examiners.csv (examiner_id, full_name, email)
  • external_examiners.csv (examiner_id, full_name, email)

Requirements:
  • Python with psycopg2
  • Optional: pywin32 for Outlook lookup (`pip install pywin32`)

Usage:
  python export_contacts.py --dsn "dbname=vivas user=… host=…" --outdir ./contact_csvs
"""
import argparse
import csv
import logging
from pathlib import Path

import psycopg2

# Try optional Outlook integration
try:
    import win32com.client
    _use_outlook = True
    _outlook = win32com.client.Dispatch('Outlook.Application')
    _namespace = _outlook.GetNamespace('MAPI')
    logging.info("Outlook: COM interface loaded for email lookup.")
except ImportError:
    _use_outlook = False
    logging.warning("pywin32 not installed; emails will be blank. To enable Outlook lookup, install pywin32.")

# Load Contacts into map
_contact_map = {}
if _use_outlook:
    try:
        contacts_folder = _namespace.GetDefaultFolder(10)  # olFolderContacts
        for item in contacts_folder.Items:
            try:
                name = item.FullName.strip().lower()
                email = getattr(item, 'Email1Address', '') or ''
                if name and email:
                    _contact_map[name] = email
            except Exception:
                continue
        logging.info(f"Loaded {_contact_map.__len__()} contacts from Outlook.")
    except Exception as e:
        logging.warning(f"Failed to load Outlook contacts: {e}")

# Email resolver
def get_email_for(name: str) -> str:
    if not _use_outlook:
        return ''
    key = name.strip().lower()
    # Contacts
    if key in _contact_map:
        return _contact_map[key]
    # Global Address List
    try:
        recip = _namespace.CreateRecipient(name)
        recip.Resolve()
        if recip.Resolved:
            entry = recip.AddressEntry
            exch = entry.GetExchangeUser()
            if exch:
                return exch.PrimarySmtpAddress or ''
    except Exception:
        pass
    return ''

# Export students
# Export students
def export_students(cur, filepath: Path):
    """
    Dumps student template CSV with blank email and project_title columns.
    """
    # Only select existing columns; project_title unsupported, so leave blank
    cur.execute("""
        SELECT student_id, full_name
          FROM students
        ORDER BY student_id
    """
    )
    rows = cur.fetchall()
    with filepath.open('w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['student_id', 'full_name', 'email', 'project_title'])
        for sid, full_name in rows:
            email = get_email_for(full_name)
            writer.writerow([sid, full_name, email, ''])

# Export examiners
def export_examiners(cur, internal_path: Path, external_path: Path):
    cur.execute("""
        SELECT examiner_id, full_name, examiner_type
          FROM examiners
        ORDER BY examiner_id
    """
    )
    internal = []
    external = []
    for eid, full_name, etype in cur.fetchall():
        email = get_email_for(full_name)
        record = (eid, full_name, email)
        if etype.lower() == 'internal':
            internal.append(record)
        else:
            external.append(record)
    # Write files
    for path, rows in [(internal_path, internal), (external_path, external)]:
        with path.open('w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['examiner_id', 'full_name', 'email'])
            writer.writerows(rows)

# Main entry

def main():
    logging.basicConfig(level=logging.INFO)
    parser = argparse.ArgumentParser(
        description='Export CSVs with optional Outlook email lookup.'
    )
    parser.add_argument('--dsn', required=True, help='PostgreSQL DSN')
    parser.add_argument('--outdir', default='.', help='Directory to write CSV files')
    args = parser.parse_args()

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    st_csv = outdir / 'students.csv'
    int_csv = outdir / 'internal_examiners.csv'
    ext_csv = outdir / 'external_examiners.csv'

    with psycopg2.connect(args.dsn) as conn:
        cur = conn.cursor()
        export_students(cur, st_csv)
        export_examiners(cur, int_csv, ext_csv)

    print(f"Exported templates to: {outdir.resolve()}")
    print(f" - {st_csv.name}")
    print(f" - {int_csv.name}")
    print(f" - {ext_csv.name}")

if __name__ == '__main__':
    main()
