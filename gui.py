# gui.py  ─────────────────────────────────────────────────────────────
# Streamlit admin console for DClinPsy examiner-allocation database
#
# • Category weights
# • Examiner workload limits  (separate tabs for internal / external)
# • Examiner ↔ student bans
# • Project groups  (1-to-3 linked theses that must share examiners)
#
# Requires:
#   pip install streamlit pandas SQLAlchemy psycopg2-binary
#
# Run:
#   streamlit run gui.py
# --------------------------------------------------------------------

import streamlit as st
import pandas as pd
import sqlalchemy as sa
from sqlalchemy import text

# ────────────────────────────────────────────────────────────────────
#  Database engine (SQLAlchemy, connection pool)
# ────────────────────────────────────────────────────────────────────
ENGINE = sa.create_engine(
    "postgresql+psycopg2://vivas_admin:treefrog@localhost:5432/vivas",
    pool_pre_ping=True,
    echo=False,        # flip to True for SQL debugging
)

# ────────────────────────────────────────────────────────────────────
#  Cached reference data  (students, examiners, categories)
# ────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False, ttl=300)
def load_reference():
    cats = pd.read_sql("SELECT * FROM categories ORDER BY name", ENGINE)

    ex = pd.read_sql("""
        SELECT examiner_id, full_name, examiner_type
        FROM examiners
        ORDER BY full_name
    """, ENGINE)

    stu = pd.read_sql("""
        SELECT student_id, full_name
        FROM students
        ORDER BY full_name
    """, ENGINE)

    return cats, ex, stu


cats_df, ex_df, stu_df = load_reference()

# ────────────────────────────────────────────────────────────────────
#  Page layout: five tabs
# ────────────────────────────────────────────────────────────────────
tab_weights, tab_int, tab_ext, tab_bans, tab_grp = st.tabs(
    ["Category weights",
     "Internal limits", "External limits",
     "Banned links", "Linked projects"]
)

# ╔══════════════════════════════════════════════════════════════════╗
# ║  TAB 1 – CATEGORY WEIGHTS                                       ║
# ╚══════════════════════════════════════════════════════════════════╝
from sqlalchemy import text

with tab_weights:
    st.header("Default weight per competence")

    edited = st.data_editor(
        cats_df[["name", "default_weight"]],
        column_config={
            "name": st.column_config.Column("Competence"),
            "default_weight": st.column_config.NumberColumn("Weight"),
        },
        num_rows="static",
        use_container_width=True,
    )

    if st.button("Save weights"):
        # Prepare values as list of tuples
        val_tuples = [
            (int(cid), float(w) if w is not None else None)
            for cid, w in zip(cats_df.category_id, edited.default_weight)
        ]

        values_clause = ", ".join(["(%s, %s)"] * len(val_tuples))
        flat_params = [v for tup in val_tuples for v in tup]

        raw_sql = f"""
            UPDATE categories
            SET default_weight = data.default_weight
            FROM (VALUES {values_clause}) AS data(category_id, default_weight)
            WHERE categories.category_id = data.category_id
        """

        # Use the raw DBAPI connection to allow native `%s` parameters
        with ENGINE.begin() as conn:
            with conn.connection.cursor() as cur:
                cur.execute(raw_sql, flat_params)

        st.success("Weights updated")
        load_reference.clear()
        st.rerun()


# ╔══════════════════════════════════════════════════════════════════╗
# ║  TAB 2 & 3 – EXAMINER LIMITS                                    ║
# ╚══════════════════════════════════════════════════════════════════╝
def limits_editor(examiner_type: str, container):
    """
    Reusable pane for internal / external workload limits.

    • Shows every examiner of the given type with an editable integer column.
    • Values ≤ 0 are treated as “no explicit limit” and are NOT stored,
      which avoids violating the CHECK (limit_n > 0) constraint.
    """
    # ------------------------------------------------------------
    # 1 ▸ assemble the grid data
    # ------------------------------------------------------------
    sub = ex_df[ex_df.examiner_type == examiner_type].copy()

    existing = pd.read_sql("SELECT * FROM examiner_limits", ENGINE)
    sub["limit_n"] = (
        sub.examiner_id.map(existing.set_index("examiner_id")["limit_n"])
        .fillna(0)
        .astype(int)
    )

    edited = container.data_editor(
        sub[["full_name", "limit_n"]],
        column_config={
            "full_name": st.column_config.Column("Examiner"),
            "limit_n":  st.column_config.NumberColumn("Max projects", step=1),
        },
        num_rows="static",
        use_container_width=True,
        key=f"limits_{examiner_type}",
    )

    # ------------------------------------------------------------
    # 2 ▸ commit on click
    # ------------------------------------------------------------
    if container.button(f"Save {examiner_type} limits"):
        with ENGINE.begin() as conn:
            # a) remove only the limits for this examiner type
            conn.execute(
                text("""
                    DELETE FROM examiner_limits
                    USING examiners e
                    WHERE examiner_limits.examiner_id = e.examiner_id
                      AND e.examiner_type = :tp
                """),
                {"tp": examiner_type},
            )

            # b) insert rows whose limit_n > 0
            rows = [
                {"id": i, "n": int(n)}
                for i, n in zip(sub.examiner_id, edited.limit_n)
                if int(n) > 0
            ]
            if rows:
                conn.execute(
                    text("""
                        INSERT INTO examiner_limits (examiner_id, limit_n)
                        VALUES (:id, :n)
                    """),
                    rows,
                )

        container.success(f"{examiner_type.capitalize()} limits saved")

with tab_int:
    st.header("Internal examiner workload limits")
    limits_editor("internal", st)

with tab_ext:
    st.header("External examiner workload limits")
    limits_editor("external", st)

# ╔══════════════════════════════════════════════════════════════════╗
# ║  TAB 4 – BANNED LINKS                                            ║
# ╚══════════════════════════════════════════════════════════════════╝
with tab_bans:
    st.header("Examiner ↔ Student bans")
    st.caption(f"Loaded {len(ex_df)} examiners.")

    # 1 ▸ pick student
    stu_name = st.selectbox("Student", stu_df.full_name, key="ban_stu")
    sid = int(stu_df.set_index("full_name").loc[stu_name, "student_id"])

    # 2 ▸ current bans for this student
    banned_ids = pd.read_sql(
        text("""
            SELECT examiner_id
            FROM examiner_student_bans
            WHERE student_id = :sid
        """),
        ENGINE,
        params={"sid": sid},
    )["examiner_id"].tolist()

    # 3 ▸ tick-box list of examiners (default = already banned)
    chosen = st.multiselect(
        "Select examiners to BAN for this student",
        ex_df.full_name,
        default=ex_df[ex_df.examiner_id.isin(banned_ids)].full_name,
        key="ban_multi",
    )
    chosen_ids = ex_df.set_index("full_name").loc[chosen, "examiner_id"].tolist()

    if st.button("Save bans"):
        to_add    = set(chosen_ids) - set(banned_ids)
        to_remove = set(banned_ids) - set(chosen_ids)

        with ENGINE.begin() as conn:
            if to_add:
                conn.execute(
                    text("""
                        INSERT INTO examiner_student_bans (examiner_id, student_id)
                        VALUES (:e, :s)
                    """),
                    [{"e": eid, "s": sid} for eid in to_add],
                )
            if to_remove:
                conn.execute(
                    text("""
                        DELETE FROM examiner_student_bans
                        WHERE student_id = :s AND examiner_id = ANY(:arr)
                    """),
                    {"s": sid, "arr": list(to_remove)},
                )
        st.success("Bans updated")

    st.subheader("Current bans for this student")
    st.write(", ".join(
        ex_df.set_index("examiner_id").loc[chosen_ids, "full_name"]
    ) or "— none —")

    st.divider()
    st.subheader("All bans (read-only overview)")
    all_bans = pd.read_sql("""
        SELECT s.full_name AS student,
               e.full_name AS examiner
        FROM examiner_student_bans b
        JOIN students  s USING (student_id)
        JOIN examiners e USING (examiner_id)
        ORDER BY student, examiner
    """, ENGINE)
    st.dataframe(all_bans, use_container_width=True, height=250)

# ╔══════════════════════════════════════════════════════════════════╗
# ║  TAB 5 – LINKED PROJECT GROUPS (1-3 STUDENTS)                   ║
# ╚══════════════════════════════════════════════════════════════════╝


with tab_grp:
    st.header("Linked project groups (max three theses)")

    # ── state: increments after every successful save --------------
    st.session_state.setdefault("grp_refresh", 0)
    rv = st.session_state["grp_refresh"]      # short alias

    # ── current groups table ---------------------------------------
    current = pd.read_sql(
        """
        SELECT g.group_id,
               string_agg(s.full_name, ', ' ORDER BY s.full_name) AS members
        FROM project_groups g
        JOIN students s USING (student_id)
        GROUP BY g.group_id
        ORDER BY g.group_id
        """,
        ENGINE,
    )
    st.dataframe(current, use_container_width=True, height=250)

    # ── widgets ----------------------------------------------------
    st.subheader("Create or extend a group")

    sel_key = f"grp_select_{rv}"          # NEW — unique per refresh
    selected = st.multiselect(
        "Select 1–3 students",
        stu_df.full_name,
        max_selections=3,
        key=sel_key,
    )

    ids = current.group_id.tolist()
    new_grp = st.checkbox("Create new group", value=not ids, key=f"grp_new_{rv}")
    grp_id = None
    if not new_grp and ids:
        grp_id = st.selectbox("Or choose existing group", ids, key=f"grp_id_{rv}")

    # ── save operation --------------------------------------------
    if st.button("Save group", disabled=len(selected) == 0, key=f"grp_save_{rv}"):
        with ENGINE.begin() as conn:

            # 1 ▸ create new group if asked
            if new_grp or grp_id is None:
                first_sid = int(
                    stu_df.set_index("full_name").loc[selected[0], "student_id"]
                )
                grp_id = conn.execute(
                    text("""
                        INSERT INTO project_groups (student_id)
                        VALUES (:sid) RETURNING group_id
                    """),
                    {"sid": first_sid},
                ).scalar_one()
                remaining = selected[1:]
            else:
                remaining = selected

            # 2 ▸ add remaining students to the same group
            if remaining:
                rows = [
                    {
                        "gid": grp_id,
                        "sid": int(stu_df.set_index("full_name")
                                   .loc[name, "student_id"])
                    }
                    for name in remaining
                ]
                conn.execute(
                    text("""
                        INSERT INTO project_groups (group_id, student_id)
                        VALUES (:gid, :sid)
                        ON CONFLICT DO NOTHING
                    """),
                    rows,
                )

        st.success("Group updated")

        # 3 ▸ bump refresh counter → widgets get brand-new keys
        st.session_state["grp_refresh"] += 1
        st.rerun()                 # Streamlit ≥ 1.41


