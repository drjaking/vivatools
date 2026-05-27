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
    cur.execute("SELECT category_id, name FROM categories")
    cat_names = {cid: name for cid, name in cur.fetchall()}
    
    cur.execute(
        """
        SELECT student_id, category_id, in_scope, weight
          FROM thesis_categories
        """
    )
    tmp: Dict[int, Dict[int, float | None]] = defaultdict(dict)
    for sid, cid, in_scope, w in cur.fetchall():
        if in_scope:
            cname = cat_names.get(cid, "")
            if cname.endswith("_other") or cname == "other":
                continue
            tmp[sid][cid] = None if w is None else float(w)
    cur.execute("SELECT student_id, full_name FROM students")
    names = {sid: n for sid, n in cur.fetchall()}
    return {sid: Student(sid, names.get(sid, f"student {sid}"), tmp[sid]) for sid in tmp}


def patch_mixed_competence(examiners: Dict[int, Examiner], categories: Dict[int, Category]):
    categories_by_name = {cat.name: cid for cid, cat in categories.items()}
    qual_id = categories_by_name.get('methodology_qualitative')
    quant_id = categories_by_name.get('methodology_quantitative')
    mixed_id = categories_by_name.get('methodology_mixed_qual_quant')
    
    if qual_id is None or quant_id is None or mixed_id is None:
        return
        
    for ex in examiners.values():
        qual_lvl = ex.competences.get(qual_id, "cannot")
        quant_lvl = ex.competences.get(quant_id, "cannot")
        
        if qual_lvl in ('can', 'could') and quant_lvl in ('can', 'could'):
            if qual_lvl == 'can' and quant_lvl == 'can':
                inferred_lvl = 'can'
            else:
                inferred_lvl = 'could'
                
            current_lvl = ex.competences.get(mixed_id, "cannot")
            lvl_rank = {'can': 2, 'could': 1, 'cannot': 0}
            if lvl_rank[inferred_lvl] > lvl_rank[current_lvl]:
                ex.competences[mixed_id] = inferred_lvl


def fetch_examiners(cur) -> Dict[int, Examiner]:
    # Fetch categories to identify qualitative/quantitative IDs
    cur.execute("SELECT category_id, name, default_weight FROM categories")
    categories = {cid: Category(cid, name, float(w) if w is not None else 1.0) for cid, name, w in cur.fetchall()}

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
    
    examiners = {
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
    
    patch_mixed_competence(examiners, categories)
    return examiners


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


def get_joint_mixed_competence(
    iex: Examiner | None, eex: Examiner | None,
    qual_id: int | None, quant_id: int | None
) -> str:
    if qual_id is None or quant_id is None:
        return 'cannot'
    i_qual = iex.competences.get(qual_id, 'cannot') if iex else 'cannot'
    i_quant = iex.competences.get(quant_id, 'cannot') if iex else 'cannot'
    e_qual = eex.competences.get(qual_id, 'cannot') if eex else 'cannot'
    e_quant = eex.competences.get(quant_id, 'cannot') if eex else 'cannot'
    
    # Check qual coverage
    has_qual_can = (i_qual == 'can' or e_qual == 'can')
    has_qual_could = (i_qual == 'could' or e_qual == 'could')
    qual_covered = has_qual_can or has_qual_could
    
    # Check quant coverage
    has_quant_can = (i_quant == 'can' or e_quant == 'can')
    has_quant_could = (i_quant == 'could' or e_quant == 'could')
    quant_covered = has_quant_can or has_quant_could
    
    if not (qual_covered and quant_covered):
        return 'cannot'
        
    if has_qual_can and has_quant_can:
        return 'can'
        
    return 'could'


def examiner_metrics(
    st: Student, ex: Examiner, cats: Dict[int, Category],
    w_could: int, w_cannot: int,
    skip_mixed_id: int | None = None,
    override_competences: Dict[int, str] | None = None
) -> Tuple[float, float]:
    tot_w = tot_pen = tot_sc = 0.0
    for cid, raw in st.categories.items():
        if cid == skip_mixed_id:
            continue
        w = _effective_weight(raw, cats[cid])
        lvl = ex.competences.get(cid, 'cannot')
        if override_competences and cid in override_competences:
            lvl = override_competences[cid]
        tot_w += w
        tot_pen += w * _penalty(lvl, w_could, w_cannot)
        tot_sc += w * _score(lvl)
    positivity = 100 * tot_sc / tot_w if tot_w else 0.0
    return round(positivity, 1), round(tot_pen, 2)


# ---------------------------------------------------------------------------
# Diagnostics builders
# ---------------------------------------------------------------------------
def build_full_diag(assignments, students, examiners, cats, w_could, w_cannot) -> List[Dict[str,str]]:
    categories_by_name = {cat.name: cid for cid, cat in cats.items()}
    qual_id = categories_by_name.get('methodology_qualitative')
    quant_id = categories_by_name.get('methodology_quantitative')
    mixed_id = categories_by_name.get('methodology_mixed_qual_quant')

    rows = []
    for sid, ieid, eeid in assignments:
        st = students[sid]
        iex = examiners.get(ieid) if ieid else None
        eex = examiners.get(eeid) if eeid else None
        
        # Determine joint mixed level
        if mixed_id and mixed_id in st.categories:
            p_level = get_joint_mixed_competence(iex, eex, qual_id, quant_id)
            overrides = {mixed_id: p_level}
            skip_id = mixed_id
            w_mixed = _effective_weight(st.categories[mixed_id], cats[mixed_id])
            mixed_pen = w_mixed * _penalty(p_level, w_could, w_cannot)
        else:
            overrides = None
            skip_id = None
            mixed_pen = 0.0
            
        i_score, i_pen = (0.0, 0.0)
        i_pen_non_mixed = 0.0
        if iex:
            i_score, i_pen = examiner_metrics(st, iex, cats, w_could, w_cannot, override_competences=overrides)
            i_pen_non_mixed = examiner_metrics(st, iex, cats, w_could, w_cannot, skip_mixed_id=skip_id)[1]
            
        e_score, e_pen = (0.0, 0.0)
        e_pen_non_mixed = 0.0
        if eex:
            e_score, e_pen = examiner_metrics(st, eex, cats, w_could, w_cannot, override_competences=overrides)
            e_pen_non_mixed = examiner_metrics(st, eex, cats, w_could, w_cannot, skip_mixed_id=skip_id)[1]
            
        total_pen = round(i_pen_non_mixed + e_pen_non_mixed + mixed_pen, 2)
        overall_score = round((i_score + e_score) / 2, 1) if (iex and eex) else (i_score or e_score)
        
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
        )) if iex else ""
        can_ext = ", ".join(sorted(
            cats[cid].name for cid, lvl in eex.competences.items()
            if lvl in ('can', 'could') and cid in st.categories
        )) if eex else ""
        cannot_i = ", ".join(sorted(
            cats[cid].name for cid, lvl in iex.competences.items()
            if lvl == 'cannot' and cid in st.categories
        )) if iex else ""
        cannot_e = ", ".join(sorted(
            cats[cid].name for cid, lvl in eex.competences.items()
            if lvl == 'cannot' and cid in st.categories
        )) if eex else ""
        rows.append({
            'student': st.name,
            'internal_examiner': iex.name if iex else "None", 'external_examiner': eex.name if eex else "None",
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
