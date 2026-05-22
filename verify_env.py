# verify_env.py
import sys
import os

print("Python executable:", sys.executable)
print("Python version:", sys.version)

packages = [
    "fastapi",
    "uvicorn",
    "jinja2",
    "pandas",
    "pandera",
    "openpyxl",
    "psycopg2",
    "sqlalchemy",
    "ortools",
    "docxtpl",
    "win32com"
]

print("\n--- Package Import Verification ---")
for pkg in packages:
    try:
        __import__(pkg)
        print(f"  [OK] {pkg}")
    except ImportError as e:
        print(f"  [FAIL] {pkg}: {e}")

print("\n--- PostgreSQL Connectivity Verification ---")
DB_DSN = "dbname=vivas user=vivas_admin password=treefrog host=localhost port=5432"
try:
    import psycopg2
    conn = psycopg2.connect(DB_DSN)
    cur = conn.cursor()
    
    # 1. Fetch tables
    cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='public'")
    tables = [row[0] for row in cur.fetchall()]
    print("Tables found:", tables)
    
    # 2. Inspect columns of key tables
    for table in ["students", "examiners", "project_groups", "project_pairs"]:
        if table in tables:
            cur.execute(f"SELECT column_name, data_type FROM information_schema.columns WHERE table_name='{table}'")
            cols = cur.fetchall()
            print(f"\nColumns for '{table}':")
            for cname, dtype in cols:
                print(f"  - {cname} ({dtype})")
        else:
            print(f"\nTable '{table}' does NOT exist in the database.")
            
    conn.close()
    print("\n[OK] Database verification complete.")
except Exception as e:
    print(f"\n[FAIL] Database verification failed: {e}")
