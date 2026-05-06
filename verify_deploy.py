"""
verify_deploy.py - comprehensive production safety verification
"""
import re, sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent

DEPLOY     = str(_HERE / 'pipeline' / 'batch_runner_deploy.py')
STANDALONE = str(_HERE / 'pipeline' / 'batch_runner_standalone.py')
EXCEL      = str(_HERE / 'pipeline' / 'excel_batch_converter.py')
# SRC_STANDALONE is external — skip gracefully if not present
_src_path  = _HERE.parent.parent / 'securities-outsourcing-samples' / 'final_scripts' / 'batch_runner_standalone.py'
SRC_STANDALONE = str(_src_path) if _src_path.exists() else None

def _read(path):
    try:
        return open(path, encoding='utf-8', errors='ignore').read() if path else ''
    except FileNotFoundError:
        print(f'[SKIP] Not found: {path}')
        return ''

deploy     = _read(DEPLOY)
standalone = _read(STANDALONE)
excel      = _read(EXCEL)
src_sa     = _read(SRC_STANDALONE)

print('=== FILE SIZES ===')
print(f'pipeline/batch_runner_deploy.py    : {len(deploy):>10,} chars | {deploy.count(chr(10))+1:>6,} lines')
print(f'pipeline/batch_runner_standalone.py: {len(standalone):>10,} chars | {standalone.count(chr(10))+1:>6,} lines')
print(f'pipeline/excel_batch_converter.py  : {len(excel):>10,} chars | {excel.count(chr(10))+1:>6,} lines')
print(f'src final_scripts/batch_runner_standalone.py: {len(src_sa):>10,} chars | {src_sa.count(chr(10))+1:>6,} lines')

print()
print('=== PRODUCTION SAFETY — batch_runner_deploy.py ===')

env_checks = [
    ("ABBYY_CUSTOMER_ID via os.getenv",     "os.getenv('ABBYY_CUSTOMER_ID'"),
    ("ABBYY_LICENSE_PATH via os.getenv",    "os.getenv('ABBYY_LICENSE_PATH'"),
    ("WORKSPACE_ID via os.getenv",          "os.getenv('WORKSPACE_ID'"),
    ("ANTHROPIC_MODEL via os.getenv",       "os.getenv('ANTHROPIC_MODEL'"),
    ("OPUS_MODEL via os.getenv",            "os.getenv('OPUS_MODEL'"),
    ("KEYING_RULES_PATH via os.getenv",     "os.getenv('KEYING_RULES_PATH'"),
    ("TEMP_DIR via os.getenv",              "os.getenv('TEMP_DIR'"),
    ("RAG_ENABLED via os.getenv",           "os.getenv('RAG_ENABLED'"),
    ("VENDOR_SGMS_DIR via os.getenv",       "os.getenv('VENDOR_SGMS_DIR'"),
]

bool_checks = [
    ("No hardcoded C:\\Users\\C303180 path",   'C:\\Users\\C303180' not in deploy),
    ("No hardcoded chroma_db_v6",              'chroma_db_v6' not in deploy),
    ("No hardcoded sec-out-samples path",      'sec-out-samples' not in deploy),
    ("win32com guarded try/except",            'try:\n    import win32com' in deploy),
    ("_WIN32_AVAILABLE flag present",           '_WIN32_AVAILABLE' in deploy),
    ("vendor_sgms uses _discover_vendor_sgms", '_discover_vendor_sgms' in deploy),
    ("No notebook CODE CELL markers",          '# ====== CODE CELL ' not in deploy),
    ("Notebook TEST_BASE stripped",            'TEST_BASE' not in deploy),
    ("class SGMLGenerator present",            'class SGMLGenerator' in deploy),
    ("class CompletePipeline present",         'class CompletePipeline' in deploy),
    ("class PatternBasedTagger present",       'class PatternBasedTagger' in deploy),
    ("class RAGManager present",               'class RAGManager' in deploy),
    ("class ABBYYConverter present",           'class ABBYYConverter' in deploy),
    ("class AgenticLLMLayer present",          'AgenticLLMLayer' in deploy),
]

all_ok = True

for desc, pattern in env_checks:
    ok = pattern in deploy
    if not ok: all_ok = False
    print(f"  [{'OK  ' if ok else 'FAIL'}] {desc}")

for desc, result in bool_checks:
    if not result: all_ok = False
    print(f"  [{'OK  ' if result else 'FAIL'}] {desc}")

# _post_fix methods
print()
print('=== _post_fix METHODS in batch_runner_deploy.py ===')
post_fix_in_src    = re.findall(r'def (_post_fix_\w+)', src_sa)
post_fix_in_deploy = re.findall(r'def (_post_fix_\w+)', deploy)
seen = set()
unique_src    = [m for m in post_fix_in_src    if not (m in seen or seen.add(m))]
seen = set()
unique_deploy = [m for m in post_fix_in_deploy if not (m in seen or seen.add(m))]

print(f"  Source standalone: {len(unique_src)} unique _post_fix methods")
print(f"  Deploy file:       {len(unique_deploy)} unique _post_fix methods")

missing_in_deploy = set(unique_src) - set(unique_deploy)
if missing_in_deploy:
    all_ok = False
    for m in sorted(missing_in_deploy):
        print(f"  [FAIL] MISSING in deploy: {m}")
else:
    print("  [OK  ] All _post_fix methods present in deploy")

# Call sites
calls_in_src    = re.findall(r'self\.(_post_fix_\w+)\(', src_sa)
calls_in_deploy = re.findall(r'self\.(_post_fix_\w+)\(', deploy)
print(f"  Source call sites: {len(calls_in_src)}")
print(f"  Deploy call sites: {len(calls_in_deploy)}")
if len(calls_in_deploy) < len(calls_in_src):
    all_ok = False
    missing_calls = set(calls_in_src) - set(calls_in_deploy)
    for m in sorted(missing_calls):
        print(f"  [FAIL] MISSING call site: {m}")
else:
    print("  [OK  ] All call sites present in deploy")

# Excel checks
print()
print('=== PRODUCTION SAFETY — excel_batch_converter.py ===')
excel_checks = [
    ("No hardcoded Windows paths",   'C:\\Users\\C303180' not in excel),
    ("No win32com import",           'import win32com' not in excel),
    ("class EnhancedExcelToSGMLConverter present", 'class EnhancedExcelToSGMLConverter' in excel),
]
for desc, result in excel_checks:
    if not result: all_ok = False
    print(f"  [{'OK  ' if result else 'FAIL'}] {desc}")

# Reference standalone sync check
print()
print('=== REFERENCE STANDALONE IN DEPLOYMENT FOLDER ===')
src_lines    = len(src_sa.splitlines())
deploy_standalone_lines = len(standalone.splitlines())
in_sync = abs(src_lines - deploy_standalone_lines) <= 5  # allow minor CRLF diff
print(f"  Source lines: {src_lines}, Deploy copy lines: {deploy_standalone_lines}")
print(f"  [{'OK  ' if in_sync else 'FAIL'}] Reference standalone in sync with source")
if not in_sync: all_ok = False

print()
print('=' * 55)
print(f'  OVERALL: {"ALL CHECKS PASSED" if all_ok else "FAILURES FOUND — see above"}')
print('=' * 55)
sys.exit(0 if all_ok else 1)
