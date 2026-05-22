"""
import_student_supervisors.py
============================

Import a two‑column Excel (.xlsx) file containing **student** and
**supervisors** into a `student_supervisors` table.

Excel layout (first sheet):
| student          | supervisors               |
|------------------|---------------------------|
| Jane Bloggs      | Dr A. Smith               |
| John Doe         | Prof Z. Jones; Dr Y. Li   |

*The **supervisors** cell is taken verbatim – no splitting is attempted.*

Assumed database schema
-----------------------
```sql
CREATE TABLE IF NOT EXISTS student_supervisors (
    student_id      INT PRIMARY KEY,
    supervisors     TEXT
);
```
*If the table does not exist the script will create it.*

Usage
-----
```bash
python import_student_supervisors.py \
  --dsn "dbname=vivas user=postgres password=treefrog host=localhost" \
  --xlsx ./student_supervisors.xlsx [--dry-run]
```

* If **student** values match `students.full_name`, the script looks up
  `student_id`.  Unknown names are reported and skipped.
* Existing rows are **upserted** (updated if present).
* `--dry-run` parses and reports without touching the database.

Dependencies
------------
```
pip install pandas openpyxl psycopg2-binary
```
"""
import argparse
import logging
from pathlib import Path

import pandas as pd
import psycopg2

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

def ensure_table(cur):
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS student_supervisors (
            student_id   INT PRIMARY KEY REFERENCES students(student_id),
            supervisors  TEXT
        );
        """
    )


def load_xlsx(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    if df.columns[0].lower() != "student" or df.columns[1].lower() != "supervisors":
        raise ValueError("First two columns must be 'student' and 'supervisors'.")
    return df[[df.columns[0], df.columns[1]]]


def insert_rows(conn, df: pd.DataFrame, dry: bool = False):
    cur = conn.cursor()
    ensure_table(cur)

    # map names → ids
    cur.execute("SELECT student_id, full_name FROM students")
    name_to_id = {name: sid for sid, name in cur.fetchall()}

    skipped = []
    for _, row in df.iterrows():
        name = str(row[0]).strip()
        sup  = str(row[1]).strip() if not pd.isna(row[1]) else ""
        sid  = name_to_id.get(name)
        if not sid:
            skipped.append(name)
            logging.warning("Student not found: %s", name)
            continue
        logging.info("Upserting supervisors for %s → %s", name, sup)
        if dry:
            continue
        cur.execute(
            """
            INSERT INTO student_supervisors (student_id, supervisors)
                 VALUES (%s, %s)
            ON CONFLICT (student_id)
            DO UPDATE SET supervisors = EXCLUDED.supervisors;
            """,
            (sid, sup)
        )
    if not dry:
        conn.commit()
    if skipped:
        logging.warning("Skipped %d students (not found in students table).", len(skipped))


def main():
    p = argparse.ArgumentParser(description="Import student-supervisor Excel mapping")
    p.add_argument("--dsn", required=True, help="PostgreSQL DSN string")
    p.add_argument("--xlsx", required=True, type=Path, help="Path to Excel file (.xlsx)")
    p.add_argument("--dry-run", action="store_true", help="Parse only; no DB writes")
    args = p.parse_args()

    df = load_xlsx(args.xlsx)
    logging.info("Loaded %d rows from %s", len(df), args.xlsx)

    if args.dry_run:
        for _, r in df.iterrows():
            logging.info("%s → %s", r[0], r[1])
        logging.info("Dry run complete – no database changes.")
        return

    with psycopg2.connect(args.dsn) as conn:
        insert_rows(conn, df, dry=False)
    logging.info("Import finished.")

if __name__ == "__main__":
    main()
