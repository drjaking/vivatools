/* ================================================================
   PostgreSQL schema for DClinPsy thesis-examiner allocation
   ---------------------------------------------------------------
   • Normalised to 3 NF
   • Uses PostgreSQL-native features (ENUM, partial indexes, cascades)
   • Ready for import scripts and OR-Tools matching code
   ================================================================*/

/* ---------- 0.  ENUM type used in several tables ---------------- */

CREATE TYPE competence_level_enum AS ENUM ('can', 'could', 'cannot');

/* ---------- 1.  Core reference tables --------------------------- */

CREATE TABLE students (
    student_id                  SERIAL      PRIMARY KEY,
    full_name                   TEXT        UNIQUE NOT NULL,
    email                       TEXT,
    project_title               TEXT,
    deferred                    BOOLEAN     NOT NULL DEFAULT FALSE,
    pre_matched                 BOOLEAN     NOT NULL DEFAULT FALSE,
    sample_other_desc           TEXT,
    methodology_other_desc      TEXT,
    field_other_desc            TEXT,
    additional_characteristics   TEXT,
    created_at                  TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE categories (
    category_id      SERIAL      PRIMARY KEY,
    name             TEXT        UNIQUE NOT NULL,
    default_weight   NUMERIC(4,2)             -- optional weighting
);

CREATE TABLE examiners (
    examiner_id   SERIAL      PRIMARY KEY,
    full_name     TEXT        UNIQUE NOT NULL,
    examiner_type TEXT        NOT NULL
        CHECK (examiner_type IN ('internal', 'external')),
    email         TEXT,
    created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

/* ---------- 2.  Thesis ↔ competence mapping --------------------- */

CREATE TABLE thesis_categories (
    student_id    INT REFERENCES students   ON DELETE CASCADE,
    category_id   INT REFERENCES categories ON DELETE CASCADE,
    in_scope      BOOLEAN      NOT NULL,     -- TRUE = competence is required
    weight        NUMERIC(4,2),              -- NULL ⇒ use default_weight
    PRIMARY KEY (student_id, category_id)
);

/* ---------- 3.  Examiner competences ---------------------------- */

CREATE TABLE examiner_competences (
    examiner_id   INT  REFERENCES examiners  ON DELETE CASCADE,
    category_id   INT  REFERENCES categories ON DELETE CASCADE,
    competence    competence_level_enum      NOT NULL,
    PRIMARY KEY (examiner_id, category_id)
);

/* ---------- 4.  Examiner ↔ thesis assignments ------------------- */

CREATE TABLE examiner_assignments (
    student_id    INT REFERENCES students   ON DELETE CASCADE,
    examiner_id   INT REFERENCES examiners  ON DELETE CASCADE,
    role          TEXT NOT NULL              -- 'internal' | 'external'
        CHECK (role IN ('internal','external')),
    assigned_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (student_id, examiner_id)
);

/* ---------- 5.  Capacity limits -------------------------------- */

CREATE TABLE examiner_limits (
    examiner_id   INT  PRIMARY KEY REFERENCES examiners ON DELETE CASCADE,
    limit_n       INT  NOT NULL CHECK (limit_n >= 0)
);

/* ---------- 6.  Conflict-of-interest exclusions ----------------- */

CREATE TABLE examiner_student_bans (
    examiner_id   INT REFERENCES examiners ON DELETE CASCADE,
    student_id    INT REFERENCES students  ON DELETE CASCADE,
    reason        TEXT,
    PRIMARY KEY (examiner_id, student_id)
);

/* ---------- 7.  Coupled projects (must share examiners) --------- */

CREATE TABLE project_pairs (
    pair_id       SERIAL PRIMARY KEY,
    student_id_1  INT REFERENCES students ON DELETE CASCADE,
    student_id_2  INT REFERENCES students ON DELETE CASCADE,
    CONSTRAINT chk_distinct_students CHECK (student_id_1 <> student_id_2)
);

CREATE TABLE IF NOT EXISTS project_groups (
    group_id   INT NOT NULL,
    student_id INT UNIQUE REFERENCES students(student_id) ON DELETE CASCADE,
    PRIMARY KEY (group_id, student_id)
);

/* ---------- 8.  Audit log of completed allocations -------------- */

CREATE TABLE allocation_audit (
    audit_id              SERIAL PRIMARY KEY,
    timestamp_utc         TIMESTAMPTZ NOT NULL DEFAULT now(),
    student_id            INT REFERENCES students,
    internal_examiner_id  INT REFERENCES examiners,
    external_examiner_id  INT REFERENCES examiners,
    allocation_note       TEXT
);

/* ---------- 9.  Recommended indexes for matching performance ---- */

/* Who can/could examine a competence? */
CREATE INDEX idx_examiner_competence_can_could
    ON examiner_competences(category_id, competence)
    WHERE competence IN ('can','could');

/* All competences required by a given thesis */
CREATE INDEX idx_thesis_required_competence
    ON thesis_categories(student_id)
    WHERE in_scope;

/* Fast lookup of remaining capacity */
CREATE INDEX idx_examiner_limits_remaining
    ON examiner_limits(limit_n)
    WHERE limit_n > 0;

/* Optional: index to fetch internal/external examiner assignments separately */
CREATE INDEX idx_assignments_by_role
    ON examiner_assignments(role);

/* ---------- 10. Verification table for relationships ----------- */
CREATE TABLE IF NOT EXISTS pending_mappings (
    mapping_id          SERIAL PRIMARY KEY,
    student_id          INT NOT NULL REFERENCES students(student_id) ON DELETE CASCADE,
    relationship_type   TEXT NOT NULL CHECK (relationship_type IN ('supervisor', 'partner')),
    raw_name            TEXT NOT NULL,
    suggested_id        INT,
    confidence          TEXT NOT NULL CHECK (confidence IN ('exact', 'fuzzy', 'none')),
    status              TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'confirmed', 'rejected'))
);
