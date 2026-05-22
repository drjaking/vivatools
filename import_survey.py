'''
import_survey.py
────────────────
Idempotently load examiner competences and thesis characteristics
from three Excel workbooks into the PostgreSQL schema printed earlier.

Prerequisites
• PostgreSQL running locally, database  vivas
• Tables created as in the “full revised schema”
• Packages: pandas, openpyxl, psycopg2-binary, pandera[pandas], tqdm
'''

from pathlib import Path
from typing import Dict, List

import pandas as pd
import pandera.pandas as pa
from pandera import Column, Check
import psycopg2
from psycopg2.extras import execute_values
from tqdm import tqdm
import re
import unicodedata as ud


# ────────────────────────────────────────────────────────────────
#  CONFIGURATION – edit these literals only
# ────────────────────────────────────────────────────────────────
DB_DSN = (
    "dbname=vivas "
    "user=vivas_admin "
    "password=treefrog "
    "host=localhost "
    "port=5432"
)

BASE_DIR = Path(r"C:\Users\john\vivatools")
PATH_INT  = BASE_DIR / "InternalCompetences2025.xlsx"
PATH_EXT  = BASE_DIR / "ExtrernalCompetences2025.xlsx"
PATH_PROJ = BASE_DIR / "ThesisCharacteristics2025.xlsx"


# ────────────────────────────────────────────────────────────────
#  UTILITY FUNCTIONS
# ────────────────────────────────────────────────────────────────
def tidy_columns(df: pd.DataFrame, first_col_name: str) -> pd.DataFrame:
    '''Normalise headers to snake_case; set first column name explicitly.'''
    df.columns = [
        (first_col_name if i == 0 else str(c)
            .replace('\u00A0', ' ')
            .replace('\n', ' ')
            .strip()
            .lower()
            .replace(' ', '_'))
        for i, c in enumerate(df.columns)
    ]
    return df

# ────────────────────────────────────────────────────────────────
#  Validation helpers – use DataFrameSchema (no annotations needed)
# ────────────────────────────────────────────────────────────────
def validate_examiner_df(df: pd.DataFrame, comps: List[str]) -> None:
    """Ensure every competence cell is one of 'can', 'could', 'can't'."""
    valid = {'can', 'could', "can't"}

    def _ok(series: pd.Series) -> pd.Series:
        return series.astype(str).str.strip().str.lower().isin(valid)

    # build the schema dict programmatically
    columns = {
        'examiner': pa.Column(str, checks=pa.Check.str_length(1, 200)),
        **{
            comp: pa.Column(str, checks=pa.Check(_ok))
            for comp in comps
        }
    }
    pa.DataFrameSchema(columns, strict=True).validate(df)


def validate_project_df(df: pd.DataFrame, comps: List[str]) -> None:
    """Accept Yes/No/Y/N/1/0 (case-insensitive, blanks ⇒ No)."""
    yn = {"yes", "no", "y", "n", "1", "0", ""}

    def _ok(series: pd.Series) -> pd.Series:
        return series.fillna("").astype(str).str.strip().str.lower().isin(yn)

    columns = {
        "student": pa.Column(str, checks=pa.Check.str_length(1, 200)),
        **{comp: pa.Column(object, checks=pa.Check(_ok)) for comp in comps},
    }
    pa.DataFrameSchema(columns, strict=True).validate(df)

def slug(s: str) -> str:
    """snake_case helper: ASCII-fold, lowercase, replace non-alnum with '_'."""
    s = ud.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = re.sub(r"[^\w]+", "_", s)
    return s.strip("_").lower()

def read_sheet(path: Path, key_col_name: str) -> pd.DataFrame:
    """
    • Read the worksheet using a single-row header.
    • Convert all column names to snake_case.
    • Replace the first column’s name with *key_col_name*.
    """
    # Read only the first row as the header
    df = pd.read_excel(path, header=0)

    # Build the new column list
    cols = df.columns.to_list()
    cols[0] = key_col_name                       # explicit rename of the key column
    cols = [slug(col) if idx else col            # normalise all other column names
            for idx, col in enumerate(cols)]

    df.columns = cols
    return df


# ────────────────────────────────────────────────────────────────
#  MAIN IMPORT WORKFLOW
# ────────────────────────────────────────────────────────────────
def main() -> None:
    # 1 ▸ Read the three workbooks
    df_int  = read_sheet(PATH_INT,  key_col_name="examiner")
    df_ext  = read_sheet(PATH_EXT,  key_col_name="examiner")
    df_proj = read_sheet(PATH_PROJ, key_col_name="student")

    # 2 ▸ Tidy column headers
    df_int  = tidy_columns(df_int,  'examiner')
    df_ext  = tidy_columns(df_ext,  'examiner')
    df_proj = tidy_columns(df_proj, 'student')

    # 3 ▸ Verify competence columns match
    set_int = set(df_int.columns) - {'examiner'}
    set_ext = set(df_ext.columns) - {'examiner'}
    if set_int != set_ext:
        print("↯ Column mismatch between examiner sheets:")
        print("  only in INTERNAL :", sorted(set_int - set_ext))
        print("  only in EXTERNAL :", sorted(set_ext - set_int))
        raise SystemExit("Fix the column list and rerun.")

    competences = sorted(set_int)  # canonical order

    # confirm thesis sheet matches
    set_proj = set(df_proj.columns) - {'student'}
    if set_proj != set_int:
        print("↯ Column mismatch between examiner and thesis sheets:")
        print("  only in THESIS :", sorted(set_proj - set_int))
        print("  missing in THESIS :", sorted(set_int - set_proj))
        raise SystemExit("Fix the column list and rerun.")

    # 4 ▸ Validate dataframes
    validate_examiner_df(df_int, competences)
    validate_examiner_df(df_ext, competences)
    validate_project_df(df_proj, competences)

    # 5 ▸ Database transaction
    conn = psycopg2.connect(DB_DSN)
    conn.autocommit = False

    try:
        with conn, conn.cursor() as cur:
            # 5a ▸ Upsert categories
            execute_values(
                cur,
                "INSERT INTO categories (name) VALUES %s "
                "ON CONFLICT (name) DO NOTHING",
                [(c,) for c in competences]
            )
            cur.execute("SELECT category_id, name FROM categories")
            cat_id: Dict[str, int] = {n: i for i, n in cur.fetchall()}

            # 5b ▸ Helper to process an examiner workbook
            def upsert_examiners(df: pd.DataFrame, e_type: str) -> None:
                execute_values(
                    cur,
                    "INSERT INTO examiners (full_name, examiner_type) "
                    "VALUES %s "
                    "ON CONFLICT (full_name) DO UPDATE "
                    "SET examiner_type = EXCLUDED.examiner_type",
                    [(row.examiner, e_type) for row in df.itertuples()]
                )
                cur.execute(
                    "SELECT examiner_id, full_name "
                    "FROM examiners WHERE examiner_type = %s",
                    (e_type,)
                )
                ex_id = {n: i for i, n in cur.fetchall()}

                tuples = []
                map_level = {'can': 'can', 'could': 'could', "can't": 'cannot'}
                for row in tqdm(df.itertuples(),
                                total=len(df),
                                desc=f"{e_type.title()} examiners"):
                    eid = ex_id[row.examiner]
                    for comp in competences:
                        level = map_level[getattr(row, comp).strip().lower()]
                        tuples.append((eid, cat_id[comp], level))

                execute_values(
                    cur,
                    "INSERT INTO examiner_competences "
                    "      (examiner_id, category_id, competence) "
                    "VALUES %s "
                    "ON CONFLICT (examiner_id, category_id) DO UPDATE "
                    "SET competence = EXCLUDED.competence",
                    tuples
                )

            upsert_examiners(df_int, 'internal')
            upsert_examiners(df_ext, 'external')

            # 5c ▸ Upsert students
            execute_values(
                cur,
                "INSERT INTO students (full_name) VALUES %s "
                "ON CONFLICT (full_name) DO NOTHING",
                [(row.student,) for row in df_proj.itertuples()]
            )
            cur.execute("SELECT student_id, full_name FROM students")
            stu_id = {n: i for i, n in cur.fetchall()}

            # 5d ▸ Upsert thesis → competence flags
            tuples = []
            for row in tqdm(df_proj.itertuples(), total=len(df_proj), desc="Theses"):
                sid = stu_id[row.student]
                for comp in competences:
                    raw = str(getattr(row, comp)).strip().lower()
                    in_scope = raw in {"yes", "y", "1"}          # anything else ⇒ False
                    tuples.append((sid, cat_id[comp], in_scope))

            cur.execute(
                "DELETE FROM thesis_categories "
                "WHERE student_id = ANY(%s)",
                (list(stu_id.values()),)
            )
            execute_values(
                cur,
                "INSERT INTO thesis_categories "
                "      (student_id, category_id, in_scope) "
                "VALUES %s",
                tuples
            )

        conn.commit()
        print("\n✓ Import completed successfully.\n")

    except Exception as exc:
        conn.rollback()
        print("\n✗ Import aborted – transaction rolled back.")
        raise exc

    finally:
        conn.close()


# ────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    main()
