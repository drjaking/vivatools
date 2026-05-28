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
    has_dclinpsy: bool

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

    # Fetch category names to identify "other" categories
    cur.execute("SELECT category_id, name FROM categories")
    cat_names = {cid: name for cid, name in cur.fetchall()}

    cur.execute(
        """
        SELECT student_id, category_id, in_scope, COALESCE(weight, 0)
          FROM thesis_categories
        """
    )
    tmp: Dict[int, Dict[int, float]] = defaultdict(dict)
    for sid, cid, in_scope, w in cur.fetchall():
        if in_scope and (sid not in deferred_ids):
            cname = cat_names.get(cid, "")
            if cname.endswith("_other") or cname == "other":
                continue
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
        SELECT ex.examiner_id, ex.examiner_type, ex.has_dclinpsy, COALESCE(el.limit_n, 3) AS limit_n
        FROM examiners ex
        LEFT JOIN examiner_limits el USING (examiner_id)
    """)
    limits = {}
    types = {}
    has_dclinpsy = {}
    for eid, etype, dclinpsy, lim in cur.fetchall():
        limits[eid] = max(0, lim - pre_matched_counts[eid])
        types[eid] = (etype == "internal")
        has_dclinpsy[eid] = bool(dclinpsy)

    cur.execute("SELECT examiner_id, category_id, competence FROM examiner_competences")
    comp: Dict[int, Dict[int, str]] = defaultdict(dict)
    for eid, cid, level in cur.fetchall():
        comp[eid][cid] = level

    return {
        eid: Examiner(eid, types[eid], limits[eid], comp[eid], has_dclinpsy[eid])
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

# ---------------------------------------------------------------------------
# Cost / penalty
# ---------------------------------------------------------------------------
def competence_penalty(level: str, w_could: int, w_cannot: int) -> int:
    if level == "can": return 0
    if level == "could": return w_could
    return w_cannot


def get_min_desirable(limit: int) -> int:
    if limit >= 4: return 2
    if limit == 3: return 2
    if limit == 2: return 1
    if limit == 1: return 1
    return 0


def student_examiner_cost(
    st: Student, ex: Examiner, categories: Dict[int, Category],
    w_could: int, w_cannot: int,
    skip_mixed_id: int | None = None
) -> int:
    total = 0.0
    for cid, weight in st.categories.items():
        if cid == skip_mixed_id:
            continue
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
    w_group_miss: int,
    w_unoccupied: int
) -> Tuple[cp_model.CpModel, Dict[Tuple[int, int], cp_model.IntVar]]:
    model = cp_model.CpModel()
    x: Dict[Tuple[int, int], cp_model.IntVar] = {}
    internal_ids = {eid for eid, ex in examiners.items() if ex.is_internal}
    external_ids = set(examiners) - internal_ids

    categories_by_name = {cat.name: cid for cid, cat in categories.items()}
    qual_id = categories_by_name.get('methodology_qualitative')
    quant_id = categories_by_name.get('methodology_quantitative')
    mixed_id = categories_by_name.get('methodology_mixed_qual_quant')

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
            cost = student_examiner_cost(st, ex, categories, w_could, w_cannot, skip_mixed_id=mixed_id)
            terms.append(cost * var)

    # one internal & one external per student
    for sid in students:
        model.Add(sum(x[(sid, eid)] for eid in internal_ids if (sid, eid) in x) == 1)
        model.Add(sum(x[(sid, eid)] for eid in external_ids if (sid, eid) in x) == 1)
        
        # Mandatory constraint: at least one assigned examiner must have DClinPsy
        dclinpsy_vars = [x[(sid, eid)] for eid in examiners if (sid, eid) in x and examiners[eid].has_dclinpsy]
        if dclinpsy_vars:
            model.Add(sum(dclinpsy_vars) >= 1)

        # Dynamic qualification constraint for mixed projects:
        # A mixed project requires coverage of both qualitative and quantitative expertise.
        if mixed_id and mixed_id in students[sid].categories:
            if qual_id is not None and quant_id is not None:
                # Find examiner ids for each level
                qual_can_eids = [eid for eid, ex in examiners.items() if ex.competences.get(qual_id, 'cannot') == 'can']
                qual_could_eids = [eid for eid, ex in examiners.items() if ex.competences.get(qual_id, 'cannot') == 'could']
                quant_can_eids = [eid for eid, ex in examiners.items() if ex.competences.get(quant_id, 'cannot') == 'can']
                quant_could_eids = [eid for eid, ex in examiners.items() if ex.competences.get(quant_id, 'cannot') == 'could']
                
                # Variables indicating if any examiner of that type is assigned to student sid
                qual_can_vars = [x[(sid, eid)] for eid in qual_can_eids if (sid, eid) in x]
                qual_could_vars = [x[(sid, eid)] for eid in qual_could_eids if (sid, eid) in x]
                quant_can_vars = [x[(sid, eid)] for eid in quant_can_eids if (sid, eid) in x]
                quant_could_vars = [x[(sid, eid)] for eid in quant_could_eids if (sid, eid) in x]
                
                # Boolean variables for coverage
                has_qual_can = model.NewBoolVar(f"has_qual_can_{sid}")
                has_qual_could = model.NewBoolVar(f"has_qual_could_{sid}")
                has_quant_can = model.NewBoolVar(f"has_quant_can_{sid}")
                has_quant_could = model.NewBoolVar(f"has_quant_could_{sid}")
                
                if qual_can_vars:
                    model.Add(sum(qual_can_vars) >= 1).OnlyEnforceIf(has_qual_can)
                    model.Add(sum(qual_can_vars) == 0).OnlyEnforceIf(has_qual_can.Not())
                else:
                    model.Add(has_qual_can == 0)
                    
                if qual_could_vars:
                    model.Add(sum(qual_could_vars) >= 1).OnlyEnforceIf(has_qual_could)
                    model.Add(sum(qual_could_vars) == 0).OnlyEnforceIf(has_qual_could.Not())
                else:
                    model.Add(has_qual_could == 0)
                    
                if quant_can_vars:
                    model.Add(sum(quant_can_vars) >= 1).OnlyEnforceIf(has_quant_can)
                    model.Add(sum(quant_can_vars) == 0).OnlyEnforceIf(has_quant_can.Not())
                else:
                    model.Add(has_quant_can == 0)
                    
                if quant_could_vars:
                    model.Add(sum(quant_could_vars) >= 1).OnlyEnforceIf(has_quant_could)
                    model.Add(sum(quant_could_vars) == 0).OnlyEnforceIf(has_quant_could.Not())
                else:
                    model.Add(has_quant_could == 0)
                    
                # qual is covered if has_qual_can or has_qual_could
                qual_covered = model.NewBoolVar(f"qual_covered_{sid}")
                model.AddBoolOr([has_qual_can, has_qual_could]).OnlyEnforceIf(qual_covered)
                model.AddBoolAnd([has_qual_can.Not(), has_qual_could.Not()]).OnlyEnforceIf(qual_covered.Not())
                
                # quant is covered if has_quant_can or has_quant_could
                quant_covered = model.NewBoolVar(f"quant_covered_{sid}")
                model.AddBoolOr([has_quant_can, has_quant_could]).OnlyEnforceIf(quant_covered)
                model.AddBoolAnd([has_quant_can.Not(), has_quant_could.Not()]).OnlyEnforceIf(quant_covered.Not())
                
                # mixed_covered = qual_covered AND quant_covered
                mixed_covered = model.NewBoolVar(f"mixed_covered_{sid}")
                model.AddBoolAnd([qual_covered, quant_covered]).OnlyEnforceIf(mixed_covered)
                model.AddBoolOr([qual_covered.Not(), quant_covered.Not()]).OnlyEnforceIf(mixed_covered.Not())
                
                mixed_cannot = mixed_covered.Not()
                
                # mixed_can = has_qual_can AND has_quant_can
                mixed_can = model.NewBoolVar(f"mixed_can_{sid}")
                model.AddBoolAnd([has_qual_can, has_quant_can]).OnlyEnforceIf(mixed_can)
                model.AddBoolOr([has_qual_can.Not(), has_quant_can.Not()]).OnlyEnforceIf(mixed_can.Not())
                
                # mixed_could = 1 - mixed_cannot - mixed_can
                mixed_could = model.NewBoolVar(f"mixed_could_{sid}")
                model.Add(mixed_cannot + mixed_could + mixed_can == 1)
                
                # Calculate penalties
                eff_w = students[sid].categories[mixed_id]
                if eff_w <= 0:
                    eff_w = categories[mixed_id].default_weight
                cannot_cost = int(round(eff_w * w_cannot * 100))
                could_cost = int(round(eff_w * w_could * 100))
                
                # Add dynamic penalties to objective terms
                terms.append(cannot_cost * mixed_cannot)
                terms.append(could_cost * mixed_could)

    # force pre-matched assignments
    for sid, st_pre in prematches.items():
        if sid in students:
            for role, eid in st_pre.items():
                if (sid, eid) in x:
                    model.Add(x[(sid, eid)] == 1)

    # capacity
    for eid, ex in examiners.items():
        model.Add(sum(x[(sid, eid)] for sid in students if (sid, eid) in x) <= ex.limit)
        
        # soft minimum workload constraint
        min_desirable = get_min_desirable(ex.limit)
        if min_desirable > 0:
            under_alloc = model.NewIntVar(0, min_desirable, f"under_alloc_{eid}")
            assigned_vars = [x[(sid, eid)] for sid in students if (sid, eid) in x]
            model.Add(sum(assigned_vars) + under_alloc >= min_desirable)
            terms.append(w_unoccupied * 100 * under_alloc)

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
    patch_mixed_competence(examiners, categories)
    bans = load_bans(cur)
    prematches = load_prematches(cur)

    # re-compute internal/external sets here for solve_and_write
    internal_ids = {eid for eid, ex in examiners.items() if ex.is_internal}
    external_ids = set(examiners) - internal_ids

    model, x = build_model(
        students, examiners, bans, categories, prematches,
        args.weight_could, args.weight_cannot, args.weight_group_miss,
        args.weight_unoccupied
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
    p.add_argument("--weight-unoccupied", type=int, default=4)
    p.add_argument("--max-seconds", type=int, default=300)
    p.add_argument("--threads", type=int, default=8)
    return p.parse_args()

if __name__ == "__main__":
    args = parse_args()
    logging.basicConfig(level=logging.INFO)
    with psycopg2.connect(args.dsn) as conn:
        solve_and_write(conn, args)
