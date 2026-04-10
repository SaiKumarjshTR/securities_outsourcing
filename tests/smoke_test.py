"""Local smoke test — run from sgml-pipeline-deployment/ directory."""
import sys
import os

# Ensure project root is on path
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

os.environ.setdefault("USE_LLM", "false")
os.environ.setdefault("RAG_ENABLED", "false")
os.environ.setdefault("TEMP_DIR", "C:\\Temp\\sgml_test")

from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app)
passed = 0
failed = 0


def check(label, got, expected):
    global passed, failed
    if got == expected:
        print(f"  PASS  {label}  (got {got})")
        passed += 1
    else:
        print(f"  FAIL  {label}  (expected {expected}, got {got})")
        failed += 1


print("\n=== SGML Pipeline API — Smoke Tests ===\n")

# --- GET / ---
r = client.get("/")
check("GET /  → 200", r.status_code, 200)
check("GET /  name field", r.json().get("name"), "SGML Pipeline API")

# --- GET /health ---
r = client.get("/health")
check("GET /health  → 200", r.status_code, 200)
assert r.json().get("status") in ("healthy", "degraded"), "status field missing"
print(f"  INFO  /health body: {r.json()}")
passed += 1

# --- POST /convert rejects .txt ---
r = client.post("/convert", files={"file": ("test.txt", b"hello world", "text/plain")})
check("POST /convert (.txt)  → 422", r.status_code, 422)

# --- POST /convert rejects empty docx ---
r = client.post(
    "/convert",
    files={"file": ("test.docx", b"", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
)
check("POST /convert (empty docx)  → 422", r.status_code, 422)

print(f"\n=== Results: {passed} passed, {failed} failed ===\n")
sys.exit(0 if failed == 0 else 1)
