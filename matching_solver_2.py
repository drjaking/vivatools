"""
matching_solver.py (v8)
=========================

Corrects a bug where `internal_ids` and `external_ids` were out of scope in
`solve_and_write`, and fixes the DELETE statement's stray quote.  Now:

• Defines `internal_ids`/`external_ids` within `solve_and_write` to be in scope
• Ensures full overwrite by deleting all prior assignments

Usage and dependencies are unchanged.
"""
from __future__ import annotations

import argparse
import logging
from collections import defaultdict
from dataclasses import dataclass
from typing import Dict, List, Set, Tuple

import psycopg2
from ortools.sat.python import cp_model

# ---------------------------------------------------------------------------
# Domain models
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Category:
    id: int
    name: str
    default_weight: float

@dataclass(frozen=True)
class Student:
    id: int
    categories: Dict[int, float]
    group_id: int | None

@dataclass(frozen=True)
class Examiner:
    id: int
    is_internal: bool
    limit: int
    competences: Dict[int, str]

# ---------------------------------------------------------------------------
# Data ingestion
# ---------------------------------------------------------------------------
def load_categories(cur) -> Dict[int, Category]:
    cur.execute("SELECT category_id, name, default_weight FROM categories")
    return {cid: Category(cid, name, float(w) if w is not None else 1.0) for cid, name, w in cur.fetchall()}


def load_students(cur) -> Dict[int, Student]:
    # Exclude deferred students from the active solver pool
    cur.execute("SELECT student_id FROM students WHERE COALESCE(deferred, FALSE) = TRUE")
    deferred_ids = {row[0] for row in cur.fetchall()}

    cur.execute(
        """
        SELECT student_id, category_id, in_scope, COALESCE(weight, 0)
          FROM thesis_categories
        """
    )
    tmp: Dict[int, Dict[int, float]] = defaultdict(dict)
    for sid, cid, in_scope, w in cur.fetchall():
        if in_scope and (sid not in deferred_ids):
            tmp[sid][cid] = float(w)

    cur.execute("SELECT group_id, student_id FROM project_groups")
    group_of: Dict[int, int] = {sid: gid for gid, sid in cur.fetchall() if sid not in deferred_ids}

    return {
        sid: Student(sid, cats, group_of.get(sid))
        for sid, cats in tmp.items()
    }


def load_examiners(cur, active_sids: Set[int]) -> Dict[int, Examiner]:
    # Workload allocations for students who are NOT active in this solver run
    # (e.g. deferred students) still consume capacity.
    pre_matched_counts = defaultdict(int)
    cur.execute("SELECT student_id, examiner_id FROM examiner_assignments")
    for sid, eid in cur.fetchall():
        if sid not in active_sids:
            pre_matched_counts[eid] += 1

    cur.execute("""
        SELECT ex.examiner_id, ex.examiner_type, COALESCE(el.limit_n, 3) AS limit_n
        FROM examiners ex
        LEFT JOIN examiner_limits el USING (examiner_id)
    """)
    limits = {}
    types = {}
    for eid, etype, lim in cur.fetchall():
        limits[eid] = max(0, lim - pre_matched_counts[eid])
        types[eid] = (etype == "internal")

    cur.execute("SELECT examiner_id, category_id, competence FROM examiner_competences")
    comp: Dict[int, Dict[int, str]] = defaultdict(dict)
    for eid, cid, level in cur.fetchall():
        comp[eid][cid] = level

    return {
        eid: Examiner(eid, types[eid], limits[eid], comp[eid])
        for eid in limits
    }


def load_prematches(cur) -> Dict[int, Dict[str, int]]:
    cur.execute("""
        SELECT student_id, examiner_id, role
        FROM examiner_assignments
        JOIN students USING (student_id)
        WHERE pre_matched = TRUE
    """)
    prematches = defaultdict(dict)
    for sid, eid, role in cur.fetchall():
        prematches[sid][role] = eid
    return prematches


def load_bans(cur) -> Set[Tuple[int, int]]:
    cur.execute("SELECT examiner_id, student_id FROM examiner_student_bans")
    return {(eid, sid) for eid, sid in cur.fetchall()}

# ---------------------------------------------------------------------------
# Cost / penalty
# ---------------------------------------------------------------------------
def competence_penalty(level: str, w_could: int, w_cannot: int) -> int:
    if level == "can": return 0
    if level == "could": return w_could
    return w_cannot


def student_examiner_cost(
    st: Student, ex: Examiner, categories: Dict[int, Category],
    w_could: int, w_cannot: int
) -> int:
    total = 0.0
    for cid, weight in st.categories.items():
        lvl = ex.competences.get(cid, "cannot")
        pen = competence_penalty(lvl, w_could, w_cannot)
        eff_w = weight if weight > 0 else categories[cid].default_weight
        total += eff_w * pen
    return int(round(total * 100))

# ---------------------------------------------------------------------------
# Model building
# ---------------------------------------------------------------------------
def build_model(
    students: Dict[int, Student],
    examiners: Dict[int, Examiner],
    bans: Set[Tuple[int, int]],
    categories: Dict[int, Category],
    prematches: Dict[int, Dict[str, int]],
    w_could: int,
    w_cannot: int,
    w_group_miss: int
) -> Tuple[cp_model.CpModel, Dict[Tuple[int, int], cp_model.IntVar]]:
    model = cp_model.CpModel()
    x: Dict[Tuple[int, int], cp_model.IntVar] = {}
    internal_ids = {eid for eid, ex in examiners.items() if ex.is_internal}
    external_ids = set(examiners) - internal_ids

    terms: List[cp_model.LinearExpr] = []
    for sid, st in students.items():
        st_pre = prematches.get(sid, {})
        for eid, ex in examiners.items():
            # Ban check is overridden if this examiner is explicitly pre-matched to this student
            is_prematched = (st_pre.get('internal') == eid or st_pre.get('external') == eid)
            if (eid, sid) in bans and not is_prematched:
                continue
            var = model.NewBoolVar(f"x_{sid}_{eid}")
            x[(sid, eid)] = var
            cost = student_examiner_cost(st, ex, categories, w_could, w_cannot)
            terms.append(cost * var)

    # one internal & one external per student
    for sid in students:
        model.Add(sum(x[(sid, eid)] for eid in internal_ids if (sid, eid) in x) == 1)
        model.Add(sum(x[(sid, eid)] for eid in external_ids if (sid, eid) in x) == 1)

    # force pre-matched assignments
    for sid, st_pre in prematches.items():
        if sid in students:
            for role, eid in st_pre.items():
                if (sid, eid) in x:
                    model.Add(x[(sid, eid)] == 1)

    # capacity
    for eid, ex in examiners.items():
        model.Add(sum(x[(sid, eid)] for sid in students if (sid, eid) in x) <= ex.limit)

    # soft group sharing
    groups = defaultdict(list)
    for st in students.values():
        if st.group_id is not None: groups[st.group_id].append(st.id)
    for gid, members in groups.items():
        if len(members) < 2: continue
        y = model.NewBoolVar(f"group_miss_{gid}")
        share_vars = []
        for eid in examiners:
            z = model.NewBoolVar(f"z_{gid}_{eid}")
            share_vars.append(z)
            for sid in members:
                if (sid, eid) in x:
                    model.Add(x[(sid, eid)] >= z)
                else:
                    model.Add(z == 0)
            model.Add(sum(x[(sid, eid)] for sid in members if (sid, eid) in x) - len(members)*z >= 0)
        model.Add(sum(share_vars) + y >= 1)
        terms.append(w_group_miss * 100 * y)

    model.Minimize(sum(terms))
    return model, x

# ---------------------------------------------------------------------------
# Solver & persistence
# ---------------------------------------------------------------------------
def solve_and_write(conn, args):
    cur = conn.cursor()
    categories = load_categories(cur)
    students = load_students(cur)
    active_sids = set(students.keys())
    examiners = load_examiners(cur, active_sids)
    bans = load_bans(cur)
    prematches = load_prematches(cur)

    # re-compute internal/external sets here for solve_and_write
    internal_ids = {eid for eid, ex in examiners.items() if ex.is_internal}
    external_ids = set(examiners) - internal_ids

    model, x = build_model(
        students, examiners, bans, categories, prematches,
        args.weight_could, args.weight_cannot, args.weight_group_miss
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = args.max_seconds
    solver.parameters.num_search_workers = args.threads
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("No feasible allocation found.")

    logging.info("Objective: %.2f", solver.ObjectiveValue() / 100)
    cur.execute("BEGIN")
    # remove existing assignments only for active students that are about to be written
    processed_sids = list(students.keys())
    if processed_sids:
        cur.execute("DELETE FROM examiner_assignments WHERE student_id = ANY(%s)", (processed_sids,))

    ins_alloc = (
        "INSERT INTO examiner_assignments (student_id, examiner_id, role)"
        " VALUES (%s,%s,%s)"
    )
    ins_audit = (
        "INSERT INTO allocation_audit (student_id, internal_examiner_id, external_examiner_id, allocation_note)"
        " VALUES (%s,%s,%s,%s)"
    )
    for sid in students:
        interns = [eid for eid in internal_ids if (sid, eid) in x and solver.Value(x[(sid, eid)])]
        externs = [eid for eid in external_ids if (sid, eid) in x and solver.Value(x[(sid, eid)])]
        if interns and externs:
            cur.execute(ins_alloc, (sid, interns[0], 'internal'))
            cur.execute(ins_alloc, (sid, externs[0], 'external'))
            cur.execute(ins_audit, (sid, interns[0], externs[0], f"auto solver obj={solver.ObjectiveValue()/100:.2f}"))
    conn.commit()
    logging.info("Written assignments + audit.")

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--dsn", required=True)
    p.add_argument("--weight-could", type=int, default=3)
    p.add_argument("--weight-cannot", type=int, default=10)
    p.add_argument("--weight-group-miss", type=int, default=5)
    p.add_argument("--max-seconds", type=int, default=300)
    p.add_argument("--threads", type=int, default=8)
    return p.parse_args()

if __name__ == "__main__":
    args = parse_args()
    logging.basicConfig(level=logging.INFO)
    with psycopg2.connect(args.dsn) as conn:
        solve_and_write(conn, args)
