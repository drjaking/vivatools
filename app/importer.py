# app/importer.py
import os
import re
import unicodedata as ud
from pathlib import Path
from typing import Dict, List, Tuple
import pandas as pd
import pandera.pandas as pa
from pandera import Column, Check
import psycopg2
from psycopg2.extras import execute_values

def slug(s: str) -> str:
    """snake_case helper: ASCII-fold, lowercase, replace non-alnum with '_'."""
    s = ud.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = re.sub(r"[^\w]+", "_", s)
    return s.strip("_").lower()

def tidy_columns(df: pd.DataFrame, first_col_name: str) -> pd.DataFrame:
    """Normalise headers to snake_case; set first column name explicitly."""
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

def read_sheet(path: Path, key_col_name: str) -> pd.DataFrame:
    """Read worksheet and normalize column headers."""
    df = pd.read_excel(path, header=0)
    cols = df.columns.to_list()
    cols[0] = key_col_name
    cols = [slug(col) if idx else col for idx, col in enumerate(cols)]
    df.columns = cols
    return df

def validate_examiner_df(df: pd.DataFrame, comps: List[str]) -> Tuple[bool, str]:
    """Ensure every competence cell is one of 'can', 'could', 'can't'."""
    valid = {'can', 'could', "can't"}
    errors = []
    
    # 1. Check for empty examiner names
    if df['examiner'].isna().any():
        errors.append("Empty examiner name found in sheet.")
        
    # 2. Check cells
    for comp in comps:
        invalid_mask = ~df[comp].astype(str).str.strip().str.lower().isin(valid)
        if invalid_mask.any():
            invalid_vals = df[invalid_mask][['examiner', comp]]
            for _, row in invalid_vals.iterrows():
                errors.append(f"Examiner '{row['examiner']}' has invalid competence level '{row[comp]}' in category '{comp}'. Expected 'can', 'could', or \"can't\".")
                
    if errors:
        return False, "\n".join(errors[:15]) + (f"\n... and {len(errors)-15} more errors" if len(errors) > 15 else "")
    return True, "Validation successful."

def map_row_text_to_canonical(text: str) -> str | None:
    text = str(text).strip()
    if text.startswith("Sample characteristics - "):
        sub = text.replace("Sample characteristics - ", "")
        return f"Sample - {sub}"
    elif text.startswith("Methodology - "):
        return text
    elif text.startswith("Field of study (multiple responses are OK) - "):
        sub = text.replace("Field of study (multiple responses are OK) - ", "")
        return f"Field - {sub}"
    return None


NICKNAMES = {
    'liz': ['elizabeth', 'eliza'],
    'elizabeth': ['liz', 'eliza'],
    'josh': ['joshua'],
    'joshua': ['josh'],
    'will': ['william'],
    'william': ['will', 'bill'],
    'vicky': ['victoria'],
    'victoria': ['vicky'],
    'tom': ['thomas'],
    'thomas': ['tom'],
    'andy': ['andrew'],
    'andrew': ['andy'],
    'chris': ['christopher', 'christina', 'christine'],
    'christina': ['chris'],
    'alex': ['alexander', 'alexandra', 'alexina'],
    'kate': ['katherine', 'kathryn', 'katrina'],
    'katrina': ['kate'],
}

def clean_name(name):
    if not name:
        return ""
    name = re.sub(r'^(Dr|Prof|Professor|Mr|Mrs|Ms|Associate Professor|Assoc Prof)\.?\s+', '', name, flags=re.IGNORECASE)
    name = re.sub(r'\s*\(.*?\)', '', name)
    name = re.sub(r'\s+ucl.*$', '', name, flags=re.IGNORECASE)
    name = re.sub(r'[^a-zA-Z\s\-]', '', name)
    return re.sub(r'\s+', ' ', name).strip()

def get_tokens(name):
    cleaned = clean_name(name).lower()
    return [t for t in cleaned.split() if t]

def match_name(excel_name, db_candidates):
    excel_name_clean = clean_name(excel_name)
    if not excel_name_clean:
        return None
        
    excel_tokens = get_tokens(excel_name_clean)
    if not excel_tokens:
        return None
        
    # 1. Try exact match first
    for ex_name, eid in db_candidates.items():
        if clean_name(ex_name).lower() == excel_name_clean.lower():
            return eid, ex_name, "exact"
            
    # 2. Try token-based matching (handles reversed names like 'Krebs Georgina')
    for ex_name, eid in db_candidates.items():
        ex_tokens = get_tokens(ex_name)
        if not ex_tokens:
            continue
            
        match_count = 0
        for t1 in excel_tokens:
            for t2 in ex_tokens:
                if t1 == t2 or t1 in NICKNAMES.get(t2, []) or t2 in NICKNAMES.get(t1, []):
                    match_count += 1
                    
        if match_count == len(excel_tokens) and len(excel_tokens) == len(ex_tokens):
            return eid, ex_name, "fuzzy"
            
        # Check if we match surname and first name
        if len(excel_tokens) > 1 and len(ex_tokens) > 1:
            s1 = excel_tokens[-1]
            s2 = ex_tokens[-1]
            is_reversed = (excel_tokens[0] == ex_tokens[-1] and excel_tokens[-1] == ex_tokens[0])
            
            if s1 == s2 or is_reversed:
                f1 = excel_tokens[0] if not is_reversed else excel_tokens[1]
                f2 = ex_tokens[0]
                if f1 == f2 or f1 in NICKNAMES.get(f2, []) or f2 in NICKNAMES.get(f1, []):
                    return eid, ex_name, "fuzzy"
                    
                import difflib
                sim = difflib.SequenceMatcher(None, f1, f2).ratio()
                if sim > 0.8:
                    return eid, ex_name, "fuzzy"
                    
    # 3. Standard fuzzy string match if ratio is extremely high (> 0.85)
    import difflib
    db_names = list(db_candidates.keys())
    close_matches = difflib.get_close_matches(excel_name_clean, db_names, n=1, cutoff=0.85)
    if close_matches:
        ex_name = close_matches[0]
        ex_tokens = get_tokens(ex_name)
        overlap = set(excel_tokens) & set(ex_tokens)
        if overlap:
            return db_candidates[ex_name], ex_name, "fuzzy"
            
    return None


def validate_project_df(df: pd.DataFrame, comps: List[str]) -> Tuple[bool, str]:
    errors = []
    
    # Check that required columns exist
    required_cols = ['RecipientLastName', 'RecipientFirstName', 'RecipientEmail', 'QTitle', 'Supv', 'Pairs']
    for rc in required_cols:
        if rc not in df.columns:
            errors.append(f"Missing required column: '{rc}'")
            
    if errors:
        return False, "\n".join(errors)
        
    # Map row 0 descriptions to canonical category names to identify category columns
    col_to_category = {}
    for col in df.columns:
        if col in required_cols:
            continue
        val0 = df.iloc[0][col]
        canon = map_row_text_to_canonical(val0)
        if canon:
            slug_canon = slug(canon)
            if slug_canon in comps:
                col_to_category[col] = slug_canon
            
    # Check rows starting from Row 1
    yn = {"yes", "no", "y", "n", "1", "0", ""}
    for idx in range(1, len(df)):
        row = df.iloc[idx]
        first = str(row['RecipientFirstName']).strip() if not pd.isna(row['RecipientFirstName']) else ""
        last = str(row['RecipientLastName']).strip() if not pd.isna(row['RecipientLastName']) else ""
        student = f"{first} {last}".strip()
        
        if not first or not last or student.lower() == "nan nan" or student.lower() == "":
            errors.append(f"Row {idx+1}: Empty student name found.")
            
        for col, cat in col_to_category.items():
            val = str(row[col]).strip().lower() if not pd.isna(row[col]) else ""
            if val not in yn:
                errors.append(f"Row {idx+1}: Student '{student}' has invalid value '{row[col]}' for category '{cat}'.")
                
    if errors:
        return False, "\n".join(errors[:15]) + (f"\n... and {len(errors)-15} more errors" if len(errors) > 15 else "")
        
    return True, "Validation successful."


def validate_survey_files(path_int: Path, path_ext: Path, path_proj: Path) -> Tuple[bool, str, Dict]:
    """Runs a complete dry-run validation of the three main survey sheets."""
    try:
        df_int = read_sheet(path_int, key_col_name="examiner")
        df_ext = read_sheet(path_ext, key_col_name="examiner")
        df_proj = pd.read_excel(path_proj, header=0)
        
        df_int = tidy_columns(df_int, 'examiner')
        df_ext = tidy_columns(df_ext, 'examiner')
        
        # Verify examiner columns match
        set_int = set(df_int.columns) - {'examiner'}
        set_ext = set(df_ext.columns) - {'examiner'}
        
        if set_int != set_ext:
            mismatch_msg = []
            if set_int - set_ext:
                mismatch_msg.append(f"Only in Internal: {sorted(set_int - set_ext)}")
            if set_ext - set_int:
                mismatch_msg.append(f"Only in External: {sorted(set_ext - set_int)}")
            return False, "Column mismatch between internal and external sheets:\n" + "\n".join(mismatch_msg), {}
            
        examiner_comps = sorted(set_int)
        
        # Validate examiner sheets
        ok_int, msg_int = validate_examiner_df(df_int, examiner_comps)
        if not ok_int:
            return False, f"Internal Sheet Error:\n{msg_int}", {}
            
        ok_ext, msg_ext = validate_examiner_df(df_ext, examiner_comps)
        if not ok_ext:
            return False, f"External Sheet Error:\n{msg_ext}", {}
            
        # Collect categories from project sheet (row 0)
        project_categories = set()
        for col in df_proj.columns:
            if col in ['RecipientLastName', 'RecipientFirstName', 'RecipientEmail', 'QTitle', 'Supv', 'Pairs']:
                continue
            if len(df_proj) > 0:
                val0 = df_proj.iloc[0][col]
                canon = map_row_text_to_canonical(val0)
                if canon:
                    project_categories.add(slug(canon))
                    
        all_competences = sorted(list(set_int | project_categories))
        
        # Validate project sheet
        ok_proj, msg_proj = validate_project_df(df_proj, all_competences)
        if not ok_proj:
            return False, f"Thesis Sheet Error:\n{msg_proj}", {}
            
        stats = {
            "internal_count": len(df_int),
            "external_count": len(df_ext),
            "student_count": len(df_proj) - 1,
            "categories_count": len(all_competences)
        }
        return True, "All sheets validated successfully. Ready to import.", stats
        
    except Exception as e:
        return False, f"Error reading/parsing spreadsheets: {str(e)}", {}


def commit_survey_to_db(db_dsn: str, path_int: Path, path_ext: Path, path_proj: Path, overwrite: bool = False) -> Tuple[bool, str]:
    """Transactionally load survey data into PostgreSQL, detecting fuzzy mappings."""
    try:
        df_int = tidy_columns(read_sheet(path_int, key_col_name="examiner"), 'examiner')
        df_ext = tidy_columns(read_sheet(path_ext, key_col_name="examiner"), 'examiner')
        
        df_proj = pd.read_excel(path_proj, header=0)
        
        # Collect examiner categories
        set_int = set(df_int.columns) - {'examiner'}
        
        # Collect categories from project sheet (row 0)
        project_categories = set()
        for col in df_proj.columns:
            if col in ['RecipientLastName', 'RecipientFirstName', 'RecipientEmail', 'QTitle', 'Supv', 'Pairs']:
                continue
            if len(df_proj) > 0:
                val0 = df_proj.iloc[0][col]
                canon = map_row_text_to_canonical(val0)
                if canon:
                    project_categories.add(slug(canon))
                    
        competences = sorted(list(set_int | project_categories))
        
        conn = psycopg2.connect(db_dsn)
        conn.autocommit = False
        with conn, conn.cursor() as cur:
            if overwrite:
                cur.execute("DELETE FROM allocation_audit")
                cur.execute("DELETE FROM examiner_assignments")
                cur.execute("DELETE FROM examiner_student_bans")
                cur.execute("DELETE FROM student_supervisors")
                cur.execute("DELETE FROM project_groups")
                cur.execute("DELETE FROM thesis_categories")
                cur.execute("DELETE FROM examiner_competences")
                cur.execute("DELETE FROM examiner_limits")
                cur.execute("DELETE FROM students")
                cur.execute("DELETE FROM examiners")
                cur.execute("DELETE FROM categories")
                cur.execute("DELETE FROM pending_mappings")
            else:
                cur.execute("DELETE FROM pending_mappings")

            # 1. Upsert categories
            execute_values(
                cur,
                "INSERT INTO categories (name) VALUES %s ON CONFLICT (name) DO NOTHING",
                [(c,) for c in competences]
            )
            cur.execute("SELECT category_id, name FROM categories")
            cat_id = {n: i for i, n in cur.fetchall()}
            
            # Helper to upsert examiners
            def upsert_examiners(df: pd.DataFrame, e_type: str):
                execute_values(
                    cur,
                    "INSERT INTO examiners (full_name, examiner_type) VALUES %s "
                    "ON CONFLICT (full_name) DO UPDATE SET examiner_type = EXCLUDED.examiner_type",
                    [(row.examiner, e_type) for row in df.itertuples()]
                )
                cur.execute("SELECT examiner_id, full_name FROM examiners WHERE examiner_type = %s", (e_type,))
                ex_id = {n: i for i, n in cur.fetchall()}
                
                tuples = []
                map_level = {'can': 'can', 'could': 'could', "can't": 'cannot'}
                for row in df.itertuples():
                    eid = ex_id[row.examiner]
                    for comp in competences:
                        if comp in df.columns:
                            level = map_level[str(getattr(row, comp)).strip().lower()]
                        else:
                            level = 'cannot'
                        tuples.append((eid, cat_id[comp], level))
                        
                execute_values(
                    cur,
                    "INSERT INTO examiner_competences (examiner_id, category_id, competence) VALUES %s "
                    "ON CONFLICT (examiner_id, category_id) DO UPDATE SET competence = EXCLUDED.competence",
                    tuples
                )
                
            upsert_examiners(df_int, 'internal')
            upsert_examiners(df_ext, 'external')
            
            # Load examiners for fuzzy matching
            cur.execute("SELECT examiner_id, full_name FROM examiners")
            db_examiners = {r[1].strip(): r[0] for r in cur.fetchall()}

            # Identify category columns in project sheet
            col_to_category = {}
            for col in df_proj.columns:
                if col in ['RecipientLastName', 'RecipientFirstName', 'RecipientEmail', 'QTitle', 'Supv', 'Pairs']:
                    continue
                val0 = df_proj.iloc[0][col]
                canon = map_row_text_to_canonical(val0)
                if canon:
                    slug_canon = slug(canon)
                    if slug_canon in competences:
                        col_to_category[col] = slug_canon

            # 2. Ingest Students starting from Row 1
            student_records = []
            for idx in range(1, len(df_proj)):
                row = df_proj.iloc[idx]
                first = str(row['RecipientFirstName']).strip() if not pd.isna(row['RecipientFirstName']) else ""
                last = str(row['RecipientLastName']).strip() if not pd.isna(row['RecipientLastName']) else ""
                name = f"{first} {last}".strip()
                if not name:
                    continue
                
                email = str(row['RecipientEmail']).strip() if not pd.isna(row['RecipientEmail']) else None
                title = str(row['QTitle']).strip() if not pd.isna(row['QTitle']) else None
                
                # Freetext fields
                sample_other = str(row['Qsampleother']).strip() if not pd.isna(row['Qsampleother']) else None
                methods_other = str(row['Qothermethod']).strip() if not pd.isna(row['Qothermethod']) else None
                field_other = str(row['Q14']).strip() if not pd.isna(row['Q14']) else None
                additional_char = str(row['Q6']).strip() if not pd.isna(row['Q6']) else None

                student_records.append((
                    name, email, title, sample_other, methods_other, field_other, additional_char
                ))

            execute_values(
                cur,
                """
                INSERT INTO students (
                    full_name, email, project_title, 
                    sample_other_desc, methodology_other_desc, 
                    field_other_desc, additional_characteristics
                ) 
                VALUES %s
                ON CONFLICT (full_name) DO UPDATE SET
                    email = EXCLUDED.email,
                    project_title = EXCLUDED.project_title,
                    sample_other_desc = EXCLUDED.sample_other_desc,
                    methodology_other_desc = EXCLUDED.methodology_other_desc,
                    field_other_desc = EXCLUDED.field_other_desc,
                    additional_characteristics = EXCLUDED.additional_characteristics
                """,
                student_records
            )

            # Map student names to IDs
            cur.execute("SELECT student_id, full_name FROM students")
            db_students = {r[1].strip(): r[0] for r in cur.fetchall()}

            # 3. Upsert thesis categories & compile relationships to fuzzy match
            thesis_tuples = []
            pending_tuples = []
            
            for idx in range(1, len(df_proj)):
                row = df_proj.iloc[idx]
                first = str(row['RecipientFirstName']).strip() if not pd.isna(row['RecipientFirstName']) else ""
                last = str(row['RecipientLastName']).strip() if not pd.isna(row['RecipientLastName']) else ""
                name = f"{first} {last}".strip()
                if not name:
                    continue
                sid = db_students[name]

                # Thesis Categories
                for col, cat in col_to_category.items():
                    raw_val = str(row[col]).strip().lower() if not pd.isna(row[col]) else ""
                    in_scope = raw_val in {"yes", "y", "1"}
                    thesis_tuples.append((sid, cat_id[cat], in_scope))

                # Supervisors (Supv)
                sup_str = str(row['Supv']).strip() if not pd.isna(row['Supv']) else ""
                if sup_str and sup_str.lower() != "nan":
                    parts = re.split(r'[,;]', sup_str)
                    for part in parts:
                        raw_sup = part.strip()
                        if not raw_sup:
                            continue
                        
                        match_res = match_name(raw_sup, db_examiners)
                        if match_res:
                            suggested_eid, suggested_name, confidence = match_res
                            pending_tuples.append((sid, 'supervisor', raw_sup, suggested_eid, confidence))
                        else:
                            pending_tuples.append((sid, 'supervisor', raw_sup, None, 'none'))

                # Partners (Pairs)
                pair_str = str(row['Pairs']).strip() if not pd.isna(row['Pairs']) else ""
                if pair_str and pair_str.lower() not in {"nan", "no", "n/a", "none", "not applicable", "not applicable."}:
                    parts = re.split(r'[,;]', pair_str)
                    for part in parts:
                        raw_partner = part.strip()
                        if not raw_partner:
                            continue
                        cleaned_partner = re.sub(r'\s*\(.*?\)', '', raw_partner).strip()
                        if not cleaned_partner or cleaned_partner.lower() in {"no", "n/a", "none"}:
                            continue
                            
                        match_res = match_name(cleaned_partner, db_students)
                        if match_res:
                            suggested_sid, suggested_name, confidence = match_res
                            pending_tuples.append((sid, 'partner', raw_partner, suggested_sid, confidence))
                        else:
                            pending_tuples.append((sid, 'partner', raw_partner, None, 'none'))

            # Write thesis categories
            cur.execute("DELETE FROM thesis_categories WHERE student_id = ANY(%s)", (list(db_students.values()),))
            execute_values(
                cur,
                "INSERT INTO thesis_categories (student_id, category_id, in_scope) VALUES %s",
                thesis_tuples
            )

            # Write pending mappings
            if pending_tuples:
                execute_values(
                    cur,
                    """
                    INSERT INTO pending_mappings (
                        student_id, relationship_type, raw_name, suggested_id, confidence, status
                    ) 
                    VALUES %s
                    """,
                    [(t[0], t[1], t[2], t[3], t[4], 'confirmed' if t[4] == 'exact' else 'pending') for t in pending_tuples]
                )

                # Automatically link exact matches
                for sid, r_type, raw_name, sug_id, conf in pending_tuples:
                    if conf == 'exact' and sug_id:
                        if r_type == 'supervisor':
                            cur.execute("""
                                INSERT INTO examiner_student_bans (examiner_id, student_id, reason)
                                VALUES (%s, %s, %s)
                                ON CONFLICT (examiner_id, student_id) DO NOTHING
                            """, (sug_id, sid, f"Project Supervisor ({raw_name})"))
                        elif r_type == 'partner':
                            # Check if already in a group
                            cur.execute("SELECT group_id FROM project_groups WHERE student_id = %s", (sid,))
                            g1 = cur.fetchone()
                            cur.execute("SELECT group_id FROM project_groups WHERE student_id = %s", (sug_id,))
                            g2 = cur.fetchone()
                            
                            if g1 and g2:
                                g1_id = g1[0]
                                g2_id = g2[0]
                                if g1_id != g2_id:
                                    cur.execute("UPDATE project_groups SET group_id = %s WHERE group_id = %s", (g1_id, g2_id))
                            elif g1:
                                cur.execute("INSERT INTO project_groups (group_id, student_id) VALUES (%s, %s) ON CONFLICT DO NOTHING", (g1[0], sug_id))
                            elif g2:
                                cur.execute("INSERT INTO project_groups (group_id, student_id) VALUES (%s, %s) ON CONFLICT DO NOTHING", (g2[0], sid))
                            else:
                                cur.execute("SELECT COALESCE(MAX(group_id), 0) + 1 FROM project_groups")
                                new_gid = cur.fetchone()[0]
                                cur.execute("INSERT INTO project_groups (group_id, student_id) VALUES (%s, %s) ON CONFLICT DO NOTHING", (new_gid, sid))
                                cur.execute("INSERT INTO project_groups (group_id, student_id) VALUES (%s, %s) ON CONFLICT DO NOTHING", (new_gid, sug_id))

            # Write student supervisors list
            student_sups = []
            for idx in range(1, len(df_proj)):
                row = df_proj.iloc[idx]
                first = str(row['RecipientFirstName']).strip() if not pd.isna(row['RecipientFirstName']) else ""
                last = str(row['RecipientLastName']).strip() if not pd.isna(row['RecipientLastName']) else ""
                name = f"{first} {last}".strip()
                if not name:
                    continue
                sid = db_students[name]
                sup_str = str(row['Supv']).strip() if not pd.isna(row['Supv']) else ""
                if sup_str and sup_str.lower() != "nan":
                    student_sups.append((sid, sup_str))
            if student_sups:
                cur.execute("DELETE FROM student_supervisors WHERE student_id = ANY(%s)", (list(db_students.values()),))
                execute_values(
                    cur,
                    "INSERT INTO student_supervisors (student_id, supervisors) VALUES %s ON CONFLICT (student_id) DO UPDATE SET supervisors = EXCLUDED.supervisors",
                    student_sups
                )

        conn.commit()
        conn.close()
        return True, "Survey imported successfully."
    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, f"Database transaction failed: {str(e)}"


def import_supervisors(db_dsn: str, path: Path, overwrite: bool = False) -> Tuple[bool, str]:
    """Parse and load student supervisors excel."""
    try:
        df = pd.read_excel(path, engine="openpyxl")
        if df.columns[0].lower() != "student" or df.columns[1].lower() != "supervisors":
            return False, "Spreadsheet columns must be exactly 'student' and 'supervisors'."
            
        conn = psycopg2.connect(db_dsn)
        with conn, conn.cursor() as cur:
            # Ensure student_supervisors table exists
            cur.execute("""
                CREATE TABLE IF NOT EXISTS student_supervisors (
                    student_id   INT PRIMARY KEY REFERENCES students(student_id),
                    supervisors  TEXT
                );
            """)
            
            if overwrite:
                cur.execute("DELETE FROM student_supervisors")
            
            # Map names -> ids
            cur.execute("SELECT student_id, full_name FROM students")
            name_to_id = {name: sid for sid, name in cur.fetchall()}
            
            upserted = 0
            new_members = 0
            for _, row in df.iterrows():
                name = str(row.iloc[0]).strip()
                sup = str(row.iloc[1]).strip() if not pd.isna(row.iloc[1]) else ""
                if not name or name == "nan":
                    continue
                sid = name_to_id.get(name)
                if not sid:
                    # Insert the new student
                    cur.execute("INSERT INTO students (full_name) VALUES (%s) RETURNING student_id", (name,))
                    sid = cur.fetchone()[0]
                    name_to_id[name] = sid
                    new_members += 1
                    
                cur.execute("""
                    INSERT INTO student_supervisors (student_id, supervisors) VALUES (%s, %s)
                    ON CONFLICT (student_id) DO UPDATE SET supervisors = EXCLUDED.supervisors;
                """, (sid, sup))
                upserted += 1
                
        conn.commit()
        conn.close()
        msg = f"Imported supervisors for {upserted} students."
        if new_members > 0:
            msg += f" Added {new_members} new student members to the database."
        return True, msg
    except Exception as e:
        return False, f"Import supervisors failed: {str(e)}"


def import_contacts_csvs(db_dsn: str, path_students: Path, path_internal: Path, path_external: Path, overwrite: bool = False) -> Tuple[bool, str]:
    """Load emails and project titles from Contacts CSV templates."""
    try:
        import csv
        conn = psycopg2.connect(db_dsn)
        with conn, conn.cursor() as cur:
            # Ensure columns exist
            cur.execute("ALTER TABLE students ADD COLUMN IF NOT EXISTS email TEXT")
            cur.execute("ALTER TABLE students ADD COLUMN IF NOT EXISTS project_title TEXT")
            cur.execute("ALTER TABLE examiners ADD COLUMN IF NOT EXISTS email TEXT")
            
            if overwrite:
                cur.execute("UPDATE students SET email = NULL, project_title = NULL")
                cur.execute("UPDATE examiners SET email = NULL")
            
            # Map name to id helper
            cur.execute("SELECT student_id, full_name FROM students")
            student_name_to_id = {name: sid for sid, name in cur.fetchall()}
            
            cur.execute("SELECT examiner_id, full_name FROM examiners")
            examiner_name_to_id = {name: eid for eid, name in cur.fetchall()}
            
            new_students = 0
            new_examiners = 0
            
            # Load students
            if path_students and path_students.exists():
                with path_students.open(newline='', encoding='utf-8', errors='replace') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        sid = row.get('student_id')
                        name = row.get('full_name', '').strip()
                        email = row.get('email', '').strip() or None
                        title = row.get('project_title', '').strip() or None
                        
                        student_id = None
                        if sid and sid.strip():
                            student_id = int(sid)
                            
                        if not student_id and name:
                            student_id = student_name_to_id.get(name)
                            
                        if not student_id and name:
                            # Insert new student
                            cur.execute("INSERT INTO students (full_name) VALUES (%s) RETURNING student_id", (name,))
                            student_id = cur.fetchone()[0]
                            student_name_to_id[name] = student_id
                            new_students += 1
                            
                        if student_id:
                            cur.execute(
                                "UPDATE students SET email = %s, project_title = %s WHERE student_id = %s",
                                (email, title, student_id)
                            )
            
            def load_examiners(path: Path, etype: str):
                nonlocal new_examiners
                if path and path.exists():
                    with path.open(newline='', encoding='utf-8', errors='replace') as f:
                        reader = csv.DictReader(f)
                        for row in reader:
                            eid = row.get('examiner_id')
                            name = row.get('full_name', '').strip()
                            email = row.get('email', '').strip() or None
                            
                            examiner_id = None
                            if eid and eid.strip():
                                examiner_id = int(eid)
                                
                            if not examiner_id and name:
                                examiner_id = examiner_name_to_id.get(name)
                                
                            if not examiner_id and name:
                                # Insert new examiner
                                cur.execute("INSERT INTO examiners (full_name, examiner_type) VALUES (%s, %s) RETURNING examiner_id", (name, etype))
                                examiner_id = cur.fetchone()[0]
                                examiner_name_to_id[name] = examiner_id
                                new_examiners += 1
                                
                                # Replenish limit for new examiner
                                cur.execute("INSERT INTO examiner_limits (examiner_id, limit_n) VALUES (%s, 3)", (examiner_id,))
                                
                            if examiner_id:
                                cur.execute("UPDATE examiners SET email = %s WHERE examiner_id = %s", (email, examiner_id))
            
            load_examiners(path_internal, 'internal')
            load_examiners(path_external, 'external')
            
        conn.commit()
        conn.close()
        
        msg = "Contact CSV details imported successfully."
        if new_students > 0 or new_examiners > 0:
            msg += f" Added {new_students} new students and {new_examiners} new examiners to the database."
        return True, msg
    except Exception as e:
        return False, f"Import contacts failed: {str(e)}"
