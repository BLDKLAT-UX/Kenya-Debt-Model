# -*- coding: utf-8 -*-
"""
run_all.py — Kenya Debt Model Master Entry Point
=================================================
Runs the complete DSA suite in sequence:

  Step 1: Validate assumptions (config/assumptions.yaml)
  Step 2: Compute model (src/model.py)
  Step 3: Write Excel workbook (src/excel_writer.py)
  Step 4: Write PDF report (src/pdf_writer.py)
  Step 5: Write sensitivity tables (sensitivity.py)
  Step 6: Launch interactive dashboard (dashboard_app.py)

Usage:
    python run_all.py              # Full suite + dashboard
    python run_all.py --no-dash    # Generate files only
    python run_all.py --only-dash  # Dashboard only
    python run_all.py --help       # Show help

Outputs (in outputs/):
    Kenya_Debt_Model.xlsx
    Kenya_Debt_Report.pdf
    Kenya_Sensitivity.xlsx
"""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

# ── Ensure project root is on the path
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from utils.logger import get_logger, log_run_summary, log_section  # noqa: E402
from utils.config_loader import load_config  # noqa: E402
from utils.validators import validate_assumptions  # noqa: E402

log = get_logger("run_all")

OUTPUTS = ROOT / "outputs"
OUTPUTS.mkdir(exist_ok=True)


# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────

def _banner() -> None:
    log.info("=" * 57)
    log.info("  KENYA DEBT SUSTAINABILITY MODEL — FULL SUITE")
    log.info("  Public Debt Management Office  |  IMF LIC-DSF")
    log.info("=" * 57)


def _step(n: int, total: int, desc: str) -> None:
    log.info("[%d/%d]  %s", n, total, desc)


def _check_deps() -> None:
    log_section(log, "Checking Dependencies")
    required = {
        "openpyxl":  "pip install openpyxl",
        "reportlab": "pip install reportlab",
        "dash":      "pip install dash",
        "plotly":    "pip install plotly",
        "pandas":    "pip install pandas",
        "yaml":      "pip install pyyaml",
    }
    missing = []
    for pkg, cmd in required.items():
        try:
            __import__(pkg)
            log.info("  ✓  %s", pkg)
        except ImportError:
            log.warning("  ✗  %s  →  %s", pkg, cmd)
            missing.append(pkg)

    if missing:
        log.info("Installing missing packages ...")
        for pkg in missing:
            log.info("  Installing %s ...", pkg)
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", pkg, "-q"],
                capture_output=True, text=True,
            )
            if result.returncode != 0:
                log.error("Failed to install %s: %s", pkg, result.stderr.strip())
                sys.exit(1)
            log.info("  ✓  %s installed", pkg)


# ─────────────────────────────────────────────────────────────────
# GENERATION STEPS
# ─────────────────────────────────────────────────────────────────

def step_validate() -> None:
    _step(1, 5, "Validating assumptions (config/assumptions.yaml)")
    cfg = load_config()
    validate_assumptions(cfg)
    log.info("  ✓  Validation passed")


def step_excel() -> Path:
    _step(2, 5, "Writing Excel workbook (9 sheets)")
    from src.model import DebtModel
    from src.excel_writer import ExcelWriter
    t0 = time.time()
    model = DebtModel().compute()
    path = ExcelWriter(model).write(OUTPUTS / "Kenya_Debt_Model.xlsx")
    log.info("  ✓  %.1fs  →  %s  (%.0f KB)",
             time.time()-t0, path.name, path.stat().st_size/1024)
    return path


def step_pdf() -> Path:
    _step(3, 5, "Writing PDF executive report (3 pages)")
    from src.model import DebtModel
    from src.pdf_writer import PDFWriter
    t0 = time.time()
    model = DebtModel().compute()
    path = PDFWriter(model).write(OUTPUTS / "Kenya_Debt_Report.pdf")
    log.info("  ✓  %.1fs  →  %s  (%.0f KB)",
             time.time()-t0, path.name, path.stat().st_size/1024)
    return path


def step_sensitivity() -> Path:
    _step(4, 5, "Writing sensitivity analysis tables (3 grids)")
    import importlib.util
    t0 = time.time()
    spec = importlib.util.spec_from_file_location(
        "sensitivity", ROOT / "sensitivity.py")
    mod = importlib.util.module_from_spec(spec)
    # Patch output path to outputs/
    import sensitivity as _  # noqa: F401  (just to check it exists)
    result = subprocess.run(
        [sys.executable, str(ROOT / "sensitivity.py")],
        capture_output=True, text=True, cwd=str(OUTPUTS.parent)
    )
    # Move to outputs if generated at root
    src_file = ROOT / "Kenya_Sensitivity.xlsx"
    dst_file = OUTPUTS / "Kenya_Sensitivity.xlsx"
    if src_file.exists():
        src_file.rename(dst_file)
    elapsed = time.time() - t0
    if dst_file.exists():
        log.info("  ✓  %.1fs  →  Kenya_Sensitivity.xlsx  (%.0f KB)",
                 elapsed, dst_file.stat().st_size/1024)
    else:
        log.warning("  sensitivity.py ran but output file not found")
    return dst_file


def step_dashboard() -> None:
    _step(5, 5, "Launching interactive web dashboard")
    log.info("")
    log.info("  ┌──────────────────────────────────────────────────┐")
    log.info("  │  Dashboard running at:  http://127.0.0.1:8050   │")
    log.info("  │  Open the link above in your browser.            │")
    log.info("  │  Press Ctrl+C to stop the server.                │")
    log.info("  └──────────────────────────────────────────────────┘")
    log.info("")
    subprocess.run([sys.executable, str(ROOT / "dashboard_app.py")])


# ─────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────

def _parse_args() -> tuple[bool, bool]:
    args = sys.argv[1:]
    if "--help" in args or "-h" in args:
        print(__doc__)
        sys.exit(0)
    no_dash   = "--no-dash"   in args
    only_dash = "--only-dash" in args
    return no_dash, only_dash


def main() -> None:
    no_dash, only_dash = _parse_args()
    _banner()

    if only_dash:
        step_dashboard()
        return

    _check_deps()

    outputs: list[tuple[str, bool]] = []
    t_total = time.time()

    for fn, label, fname in [
        (step_validate,    "Validation",          "—"),
        (step_excel,       "Excel Workbook",       "Kenya_Debt_Model.xlsx"),
        (step_pdf,         "PDF Report",           "Kenya_Debt_Report.pdf"),
        (step_sensitivity, "Sensitivity Tables",   "Kenya_Sensitivity.xlsx"),
    ]:
        try:
            fn()
            outputs.append((fname, True))
        except Exception as exc:
            log.error("FAILED: %s — %s", label, exc)
            import traceback
            traceback.print_exc()
            outputs.append((fname, False))

    log.info("")
    log_run_summary(log, [(f, ok) for f, ok in outputs if f != "—"])
    log.info("Total time: %.1fs", time.time() - t_total)

    if not no_dash:
        step_dashboard()
    else:
        log.info("Dashboard skipped (--no-dash).  "
                 "Run 'python run_all.py --only-dash' to launch it later.")


if __name__ == "__main__":
    main()
