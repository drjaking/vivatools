"""
streamlit_examiner_gui.py (v3: fix NaN capacity, remove duplicate widgets)
=======================================================================

* Handles NULL capacity or assigned counts by treating them as 0 so the
  remaining calculation doesn’t raise a NaN→int error.
* Removes the accidental duplicate selectbox definitions.
* Keeps caching hash‑safe.
"""
import argparse
from functools import lru_cache
from typing import List, Tuple

import pandas as pd
import psycopg2
import streamlit as st

# ---------------------------------------------------------------------------
# Cached helpers (hashable parameters only)
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_write_conn(dsn: str):
    return psycopg2.connect(dsn)

@st.cache_data(show_spinner=False)
def load_students(dsn: str) -> List[Tuple[int, str]]:
    with psycopg2.connect(dsn) as conn:
        cur = conn.cursor()
        cur.execute("SELECT student_id, full_name FROM students ORDER BY full_name")
        return cur.fetchall()

@st.cache_data(show_spinner=False)
def load_examiners_table(dsn: str) -> pd.DataFrame:
    with psycopg2.connect(dsn) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            SELECT ex.examiner_id,
                   ex.full_name,
                   ex.examiner_type,
                   COALESCE(el.limit_n,0) AS capacity,
                   COALESCE(a.assigned,0)  AS assigned
            FROM   examiners ex
            LEFT JOIN examiner_limits el USING (examiner_id)
            LEFT JOIN (
                SELECT examiner_id, COUNT(*) AS assigned
                FROM   examiner_assignments
                GROUP  BY examiner_id
            ) a USING (examiner_id)
            ORDER BY ex.full_name;
            """
        )
        rows = cur.fetchall()
    df = pd.DataFrame(rows, columns=['examiner_id','full_name','examiner_type','capacity','assigned'])
    df[['capacity','assigned']] = df[['capacity','assigned']].fillna(0).astype(int)
    return df

@lru_cache(maxsize=None)
def get_assignment(dsn: str, student_id: int):
    with psycopg2.connect(dsn) as conn:
        cur = conn.cursor()
        cur.execute(
            """SELECT role, examiner_id, full_name
                   FROM examiner_assignments ea
                   JOIN examiners e USING (examiner_id)
                  WHERE ea.student_id = %s""", (student_id,))
        return {r[0]: (r[1], r[2]) for r in cur.fetchall()}

# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

def label_with_remaining(row):
    remain = int(row['capacity']) - int(row['assigned'])
    return f"{row['full_name']} (rem {remain})"

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dsn', required=True)
    cli_args, _ = parser.parse_known_args()
    dsn = cli_args.dsn

    write_conn = get_write_conn(dsn)

    st.title("Viva Examiner Assignment Tool")

    students = load_students(dsn)
    exams_df = load_examiners_table(dsn)
    int_df = exams_df[exams_df['examiner_type']=='internal']
    ext_df = exams_df[exams_df['examiner_type']=='external']

    # Student selection
    student_select = st.selectbox("Select a student", [name for _id,name in students])
    if not student_select:
        st.stop()
    student_id = next(sid for sid,name in students if name==student_select)

    assignment = get_assignment(dsn, student_id)
    cur_int_id, cur_int_name = assignment.get('internal',(None,'—'))
    cur_ext_id, cur_ext_name = assignment.get('external',(None,'—'))

    st.write(f"**Current Internal:** {cur_int_name}")
    st.write(f"**Current External:** {cur_ext_name}")
    st.divider()

    # Option lists with human‑readable labels
    int_labels = int_df.apply(label_with_remaining, axis=1).tolist()
    ext_labels = ext_df.apply(label_with_remaining, axis=1).tolist()

    cur_int_label = label_with_remaining(int_df[int_df['examiner_id']==cur_int_id].iloc[0]) if cur_int_id in int_df['examiner_id'].values else int_labels[0]
    cur_ext_label = label_with_remaining(ext_df[ext_df['examiner_id']==cur_ext_id].iloc[0]) if cur_ext_id in ext_df['examiner_id'].values else ext_labels[0]

    int_choice = st.selectbox("New Internal Examiner", int_labels, index=int_labels.index(cur_int_label))
    new_int_id = int_df.iloc[int_labels.index(int_choice)]['examiner_id']

    ext_choice = st.selectbox("New External Examiner", ext_labels, index=ext_labels.index(cur_ext_label))
    new_ext_id = ext_df.iloc[ext_labels.index(ext_choice)]['examiner_id']

    if st.button("Update Assignment"):
        try:
            with write_conn:
                with write_conn.cursor() as cur:
                    cur.execute("UPDATE examiner_assignments SET examiner_id=%s WHERE student_id=%s AND role='internal'", (int(new_int_id), student_id))
                    cur.execute("UPDATE examiner_assignments SET examiner_id=%s WHERE student_id=%s AND role='external'", (int(new_ext_id), student_id))
            st.success("Assignment updated.")
            get_assignment.cache_clear()
            load_examiners_table.clear()
        except Exception as e:
            st.error(f"Update failed: {e}")

if __name__ == '__main__':
    main()
