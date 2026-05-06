"""Page 2 — PDF → SGML HITL Review (delegates to hitl_review.py)."""
import runpy, sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_APP  = str(_ROOT / "app")

# ── Fix sys.path: root must be first; app/ must NOT shadow root/config.py ──
# Remove app/ so 'import config' finds root/config.py, not app/config.py
sys.path = [p for p in sys.path if p != _APP]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
else:
    sys.path.remove(str(_ROOT))
    sys.path.insert(0, str(_ROOT))

# ── Evict any stale 'config' module cached from app/config.py ─────────────
for _key in [k for k in sys.modules if k == "config" or k.startswith("config.")]:
    del sys.modules[_key]

runpy.run_path(str(_ROOT / "hitl_review.py"), run_name="__main__")
