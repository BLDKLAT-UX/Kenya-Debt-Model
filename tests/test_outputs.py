# -*- coding: utf-8 -*-
"""
tests/test_outputs.py
=====================
Integration tests — verify Excel, PDF, and Sensitivity outputs
are generated correctly and contain expected content.

Run:
    pytest tests/ -v
"""

import os
import pytest
from pathlib import Path

OUTPUTS = Path(__file__).resolve().parent.parent / "outputs"


@pytest.fixture(scope="module", autouse=True)
def generate_outputs():
    """Generate all output files once before running output tests."""
    from src.model import DebtModel
    from src.excel_writer import ExcelWriter
    from src.pdf_writer import PDFWriter
    from utils.config_loader import reload_config
    from utils.validators import validate_assumptions

    reload_config()
    cfg = reload_config()
    validate_assumptions(cfg)

    model = DebtModel().compute()
    ExcelWriter(model).write(OUTPUTS / "Kenya_Debt_Model.xlsx")
    PDFWriter(model).write(OUTPUTS / "Kenya_Debt_Report.pdf")
    yield


# ── Excel Tests
class TestExcelOutput:
    def test_excel_file_exists(self):
        assert (OUTPUTS / "Kenya_Debt_Model.xlsx").exists()

    def test_excel_file_not_empty(self):
        size = (OUTPUTS / "Kenya_Debt_Model.xlsx").stat().st_size
        assert size > 10_000, f"Excel file too small: {size} bytes"

    def test_excel_has_all_sheets(self):
        from openpyxl import load_workbook
        wb = load_workbook(OUTPUTS / "Kenya_Debt_Model.xlsx")
        expected = [
            "COVER", "ASSUMPTIONS", "MACRO_FRAMEWORK", "DEBT_STOCK",
            "DEBT_SERVICE", "DSA_INDICATORS", "STRESS_TESTS",
            "CHARTS", "DASHBOARD",
        ]
        for sheet in expected:
            assert sheet in wb.sheetnames, f"Missing sheet: {sheet}"

    def test_cover_sheet_has_title(self):
        from openpyxl import load_workbook
        wb = load_workbook(OUTPUTS / "Kenya_Debt_Model.xlsx")
        ws = wb["COVER"]
        # Row 2, col 2 should contain the country name
        cell_val = ws.cell(2, 2).value
        assert cell_val is not None
        assert "KENYA" in str(cell_val).upper()

    def test_macro_framework_has_revenue_data(self):
        from openpyxl import load_workbook
        wb = load_workbook(OUTPUTS / "Kenya_Debt_Model.xlsx")
        ws = wb["MACRO_FRAMEWORK"]
        # Check at least one numeric cell exists in the sheet
        numeric_cells = [
            ws.cell(r, c).value
            for r in range(1, 50)
            for c in range(3, 12)
            if isinstance(ws.cell(r, c).value, (int, float))
        ]
        assert len(numeric_cells) > 20, "MACRO_FRAMEWORK has too few numeric cells"

    def test_debt_stock_has_data(self):
        from openpyxl import load_workbook
        wb = load_workbook(OUTPUTS / "Kenya_Debt_Model.xlsx")
        ws = wb["DEBT_STOCK"]
        numeric_cells = [
            ws.cell(r, c).value
            for r in range(1, 60)
            for c in range(2, 11)
            if isinstance(ws.cell(r, c).value, (int, float))
        ]
        assert len(numeric_cells) > 30

    def test_stress_tests_sheet_not_empty(self):
        from openpyxl import load_workbook
        wb = load_workbook(OUTPUTS / "Kenya_Debt_Model.xlsx")
        ws = wb["STRESS_TESTS"]
        cells_with_data = [
            ws.cell(r, c).value
            for r in range(1, 100)
            for c in range(2, 9)
            if ws.cell(r, c).value is not None
        ]
        assert len(cells_with_data) > 50

    def test_dashboard_sheet_has_kpi_data(self):
        from openpyxl import load_workbook
        wb = load_workbook(OUTPUTS / "Kenya_Debt_Model.xlsx")
        ws = wb["DASHBOARD"]
        assert ws.cell(2, 2).value is not None


# ── PDF Tests
class TestPDFOutput:
    def test_pdf_file_exists(self):
        assert (OUTPUTS / "Kenya_Debt_Report.pdf").exists()

    def test_pdf_file_not_empty(self):
        size = (OUTPUTS / "Kenya_Debt_Report.pdf").stat().st_size
        assert size > 50_000, f"PDF too small: {size} bytes"

    def test_pdf_is_valid(self):
        """Check PDF has valid header."""
        with open(OUTPUTS / "Kenya_Debt_Report.pdf", "rb") as f:
            header = f.read(5)
        assert header == b"%PDF-", "File is not a valid PDF"


# ── Validators
class TestValidators:
    def test_validate_passes_on_good_config(self):
        from utils.config_loader import load_config
        from utils.validators import validate_assumptions
        cfg = load_config()
        validate_assumptions(cfg)  # should not raise

    def test_validate_fails_on_wrong_length(self):
        from utils.config_loader import load_config
        from utils.validators import validate_assumptions, ValidationError
        import copy
        cfg = copy.deepcopy(load_config())
        # Break a series length
        cfg["real_sector"]["nominal_gdp_kes_bn"] = [1, 2, 3]
        with pytest.raises(ValidationError, match="nominal_gdp_kes_bn"):
            validate_assumptions(cfg)

    def test_validate_fails_on_negative_debt(self):
        from utils.config_loader import load_config
        from utils.validators import validate_assumptions, ValidationError
        import copy
        cfg = copy.deepcopy(load_config())
        cfg["debt_stock"]["domestic"]["tbills_91d"][0] = -100
        with pytest.raises(ValidationError):
            validate_assumptions(cfg)

    def test_validate_fails_on_bad_holder_shares(self):
        from utils.config_loader import load_config
        from utils.validators import validate_assumptions, ValidationError
        import copy
        cfg = copy.deepcopy(load_config())
        cfg["debt_stock"]["domestic"]["holder_shares"]["commercial_banks"] = 0.9
        with pytest.raises(ValidationError, match="holder_shares"):
            validate_assumptions(cfg)

    def test_validate_fails_on_missing_threshold(self):
        from utils.config_loader import load_config
        from utils.validators import validate_assumptions, ValidationError
        import copy
        cfg = copy.deepcopy(load_config())
        del cfg["thresholds"]["ds_revenue_pct"]
        with pytest.raises(ValidationError, match="ds_revenue_pct"):
            validate_assumptions(cfg)


# ── Config Loader
class TestConfigLoader:
    def test_loads_successfully(self):
        from utils.config_loader import load_config
        cfg = load_config()
        assert isinstance(cfg, dict)

    def test_has_required_top_level_keys(self):
        from utils.config_loader import load_config
        cfg = load_config()
        for key in ["meta", "years", "real_sector", "external_sector",
                    "rates", "debt_stock", "fiscal", "thresholds",
                    "stress_tests"]:
            assert key in cfg, f"Missing top-level key: {key}"

    def test_caching_returns_same_object(self):
        from utils.config_loader import load_config
        cfg1 = load_config()
        cfg2 = load_config()
        assert cfg1 is cfg2

    def test_reload_clears_cache(self):
        from utils.config_loader import load_config, reload_config
        cfg1 = load_config()
        cfg2 = reload_config()
        # Values should be equal even if different objects
        assert cfg1["meta"]["country"] == cfg2["meta"]["country"]

    def test_raises_on_missing_file(self):
        from utils.config_loader import reload_config
        with pytest.raises(FileNotFoundError):
            reload_config("/nonexistent/path/assumptions.yaml")
