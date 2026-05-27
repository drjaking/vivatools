"""
diagnostics.py (v7)
===================

Provides two diagnostic tables after allocation:

1. **STUDENT-EXAMINER DIAGNOSTICS** – a comprehensive per-student report, including:
   - Assigned examiner names
   - Positivity scores and penalties
   - Project category details (top 5, YES categories, can/could, cannot)

2. **EXAMINER WORKLOAD VS LIMITS** – for each examiner:
   - Name, type, capacity, assigned load, remaining slots

Usage:

    python diagnostics.py --dsn "..." --weight-could 3 --weight-cannot 10 \
                         [--csv-out ./diagnostics]

If `--csv-out` is set, CSVs are written under that directory.
"""
from __future__ import annotations

import argparse
import csv
import logging
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Tuple

import psycopg2

# ---------------------------------------------------------------------------
# Domain models
# ---------------------------------------------------------------------------
class Category:
    def __init__(self, cid: int, name: str, default_weight: float):
        self.id = cid
        self.name = name
        self.default_weight = default_weight

class Student:
    def __init__(self, sid: int, name: str, categories: Dict[int, float | None]):
        self.id = sid
        self.name = name
        self.categories = categories  # weight==0/None => use default

class Examiner:
    def __init__(self, eid: int, name: str, is_internal: bool, limit: int, competences: Dict[int, str], has_dclinpsy: bool = True):
        self.id = eid
        self.name = name
        self.is_internal = is_internal
        self.limit = limit
        self.competences = competences
        self.has_dclinpsy = has_dclinpsy

# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------
def fetch_categories(cur) -> Dict[int, Category]:
    cur.execute("SELECT category_id, name, default_weight FROM categories")
    return {cid: Category(cid, name, float(w) if w is not None else 1.0) for cid, name, w in cur.fetchall()}


def fetch_students(cur) -> Dict[int, Student]:
    cur.execute(
        """
        SELECT student_id, category_id, in_scope, weight
          FROM thesis_categories
        """
    )
    tmp: Dict[int, Dict[int, float | None]] = defaultdict(dict)
    for sid, cid, in_scope, w in cur.fetchall():
        if in_scope:
            tmp[sid][cid] = None if w is None else float(w)
    cur.execute("SELECT student_id, full_name FROM students")
    names = {sid: n for sid, n in cur.fetchall()}
    return {sid: Student(sid, names.get(sid, f"student {sid}"), tmp[sid]) for sid in tmp}


def fetch_examiners(cur) -> Dict[int, Examiner]:
    cur.execute("SELECT examiner_id, limit_n FROM examiner_limits")
    limits = {eid: lim for eid, lim in cur.fetchall()}
    cur.execute("SELECT examiner_id, examiner_type, has_dclinpsy FROM examiners")
    ex_data = {row[0]: (row[1] == 'internal', bool(row[2])) for row in cur.fetchall()}
    cur.execute("SELECT examiner_id, category_id, competence FROM examiner_competences")
    comp: Dict[int, Dict[int, str]] = defaultdict(dict)
    for eid, cid, lvl in cur.fetchall():
        comp[eid][cid] = lvl
    cur.execute("SELECT examiner_id, full_name FROM examiners")
    names = {eid: n for eid, n in cur.fetchall()}
    return {
        eid: Examiner(
            eid,
            names.get(eid, f"examiner {eid}"),
            ex_data.get(eid, (True, True))[0],
            limits.get(eid, 3),
            comp[eid],
            ex_data.get(eid, (True, True))[1]
        )
        for eid in names
    }


def fetch_assignments(cur) -> List[Tuple[int,int,int]]:
    cur.execute(
        """
        SELECT student_id,
               MAX(CASE WHEN role='internal' THEN examiner_id END),
               MAX(CASE WHEN role='external' THEN examiner_id END)
          FROM examiner_assignments
         GROUP BY student_id
        """
    )
    return cur.fetchall()

# ---------------------------------------------------------------------------
# Scoring utilities
# ---------------------------------------------------------------------------
def _effective_weight(raw: float | None, cat: Category) -> float:
    return raw if raw and raw > 0 else cat.default_weight


def _penalty(level: str, w_could: int, w_cannot: int) -> int:
    if level == 'can': return 0
    if level == 'could': return w_could
    return w_cannot


def _score(level: str) -> float:
    return 1.0 if level=='can' else 0.5 if level=='could' else 0.0


def examiner_metrics(st: Student, ex: Examiner, cats: Dict[int, Category], w_could: int, w_cannot: int) -> Tuple[float, float]:
    tot_w = tot_pen = tot_sc = 0.0
    for cid, raw in st.categories.items():
        w = _effective_weight(raw, cats[cid])
        lvl = ex.competences.get(cid, 'cannot')
        tot_w += w
        tot_pen += w * _penalty(lvl, w_could, w_cannot)
        tot_sc += w * _score(lvl)
    positivity = 100 * tot_sc / tot_w if tot_w else 0.0
    return round(positivity, 1), round(tot_pen, 2)

# ---------------------------------------------------------------------------
# Diagnostics builders
# ---------------------------------------------------------------------------
def build_full_diag(assignments, students, examiners, cats, w_could, w_cannot) -> List[Dict[str,str]]:
    rows = []
    for sid, ieid, eeid in assignments:
        st = students[sid]
        iex = examiners[ieid]; eex = examiners[eeid]
        i_score, i_pen = examiner_metrics(st, iex, cats, w_could, w_cannot)
        e_score, e_pen = examiner_metrics(st, eex, cats, w_could, w_cannot)
        total_pen = round(i_pen + e_pen, 2)
        overall_score = round((i_score + e_score) / 2, 1)
        # categories
        top5 = ", ".join(
            cats[cid].name for cid, _ in sorted(
                st.categories.items(),
                key=lambda kv: -_effective_weight(kv[1], cats[kv[0]])
            )[:5]
        )
        yes_cats = ", ".join(sorted(cats[cid].name for cid in st.categories))
        can_int = ", ".join(sorted(
            cats[cid].name for cid, lvl in iex.competences.items()
            if lvl in ('can', 'could') and cid in st.categories
        ))
        can_ext = ", ".join(sorted(
            cats[cid].name for cid, lvl in eex.competences.items()
            if lvl in ('can', 'could') and cid in st.categories
        ))
        cannot_i = ", ".join(sorted(
            cats[cid].name for cid, lvl in iex.competences.items()
            if lvl == 'cannot' and cid in st.categories
        ))
        cannot_e = ", ".join(sorted(
            cats[cid].name for cid, lvl in eex.competences.items()
            if lvl == 'cannot' and cid in st.categories
        ))
        rows.append({
            'student': st.name,
            'internal_examiner': iex.name, 'external_examiner': eex.name,
            'internal_score': i_score, 'external_score': e_score, 'overall_score': overall_score,
            'internal_penalty': i_pen, 'external_penalty': e_pen, 'total_penalty': total_pen,
            'top_5_categories': top5, 'student_yes_categories': yes_cats,
            'can_could_categories_internal': can_int, 'can_could_categories_external': can_ext,
            'cant_categories_internal': cannot_i, 'cant_categories_external': cannot_e
        })
    return rows


def build_examiner_workload(examinations: List[Tuple[int,int,int]], examiners: Dict[int, Examiner]) -> List[Dict[str,str]]:
    load = defaultdict(int)
    for _, ieid, eeid in examinations:
        load[ieid] += 1; load[eeid] += 1
    rows = []
    for eid, ex in examiners.items():
        rows.append({
            'examiner': ex.name,
            'type': 'internal' if ex.is_internal else 'external',
            'capacity': str(ex.limit),
            'assigned': str(load[eid]),
            'remaining': str(ex.limit - load[eid])
        })
    return rows

# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------
def print_table(rows: List[Dict[str,str]], title: str):
    if not rows:
        print(f"\n{title}: <no data>\n")
        return
    cols = list(rows[0].keys())
    widths = {c: max(len(c), *[len(str(r[c])) for r in rows]) for c in cols}
    sep = "| " + " | ".join("-" * widths[c] for c in cols) + " |"
    print(f"\n{title}\n{sep}")
    print("| " + " | ".join(f"{c:{widths[c]}}" for c in cols) + " |")
    print(sep)
    for r in rows:
        print("| " + " | ".join(f"{str(r[c]):{widths[c]}}" for c in cols) + " |")
    print(sep + "\n")


def write_csv(rows: List[Dict[str,str]], path: Path):
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Diagnostics: full and workload tables")
    parser.add_argument('--dsn', required=True)
    parser.add_argument('--weight-could', type=int, default=3)
    parser.add_argument('--weight-cannot', type=int, default=10)
    parser.add_argument('--csv-out', type=Path, help="Directory for CSV output")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO)

    with psycopg2.connect(args.dsn) as conn:
        cur = conn.cursor()
        cats = fetch_categories(cur)
        students = fetch_students(cur)
        examiners = fetch_examiners(cur)
        assignments = fetch_assignments(cur)

    full_rows = build_full_diag(assignments, students, examiners, cats, args.weight_could, args.weight_cannot)
    workload_rows = build_examiner_workload(assignments, examiners)

    print_table(full_rows, 'STUDENT-EXAMINER DIAGNOSTICS')
    print_table(workload_rows, 'EXAMINER WORKLOAD VS LIMITS')

    if args.csv_out:
        write_csv(full_rows, args.csv_out / 'student_examiner_diagnostics.csv')
        write_csv(workload_rows, args.csv_out / 'examiner_workload.csv')
        logging.info('CSV files written to %s', args.csv_out.resolve())

if __name__ == '__main__':
    main()
