# -*- coding: utf-8 -*-
"""
tests/test_model.py
===================
Unit tests for src/model.py — core DSA calculations.

Run:
    pytest tests/ -v
"""

import pytest
from src.model import DebtModel


@pytest.fixture(scope="module")
def model():
    """Shared computed model for all tests."""
    m = DebtModel()
    m.compute()
    return m


# ── Initialisation
class TestModelInit:
    def test_loads_config(self, model):
        assert model.cfg is not None

    def test_correct_year_count(self, model):
        assert model.n == 8

    def test_fiscal_years_labels(self, model):
        assert model.years[0] == "FY22/23"
        assert model.years[-1] == "FY29/30"

    def test_projection_index(self, model):
        assert model.proj_idx == 3

    def test_requires_compute(self):
        m = DebtModel()
        with pytest.raises(RuntimeError, match="compute"):
            m.check_thresholds()


# ── Real Sector
class TestRealSector:
    def test_gdp_length(self, model):
        assert len(model.gdp) == 8

    def test_gdp_positive(self, model):
        assert all(g > 0 for g in model.gdp)

    def test_gdp_increasing(self, model):
        assert all(model.gdp[i] <= model.gdp[i+1] for i in range(7))

    def test_gdp_base_year_value(self, model):
        # FY24/25 is index 2
        assert model.gdp[2] == 14_400

    def test_fx_rate_positive(self, model):
        assert all(r > 0 for r in model.fx_rate)


# ── Debt Stock
class TestDebtStock:
    def test_dom_debt_length(self, model):
        assert len(model.dom_debt) == 8

    def test_ext_debt_length(self, model):
        assert len(model.ext_debt) == 8

    def test_total_debt_is_sum(self, model):
        for i in range(8):
            expected = model.dom_debt[i] + model.ext_debt[i] + model.guar_debt[i]
            assert abs(model.total_debt[i] - expected) < 0.01

    def test_dom_debt_positive(self, model):
        assert all(d > 0 for d in model.dom_debt)

    def test_tbills_are_sum_of_components(self, model):
        cfg = model.cfg["debt_stock"]["domestic"]
        for i in range(8):
            expected = (cfg["tbills_91d"][i]
                        + cfg["tbills_182d"][i]
                        + cfg["tbills_364d"][i])
            assert abs(model.tbills[i] - expected) < 0.01

    def test_holder_shares_sum_to_total(self, model):
        for i in range(8):
            total = (model.hold_banks[i] + model.hold_pension[i]
                     + model.hold_cbk[i] + model.hold_other[i])
            assert abs(total - model.dom_debt[i]) < 1.0  # allow rounding


# ── Fiscal
class TestFiscal:
    def test_overall_balance_is_revenue_minus_expenditure(self, model):
        for i in range(8):
            expected = model.revenue[i] - model.expenditure[i]
            assert abs(model.overall_balance[i] - expected) < 0.01

    def test_primary_balance_adds_interest(self, model):
        for i in range(8):
            expected = model.overall_balance[i] + model.interest_total[i]
            assert abs(model.primary_balance[i] - expected) < 0.01

    def test_kenya_runs_fiscal_deficit(self, model):
        # All years should show overall deficit (expenditure > revenue)
        assert all(b < 0 for b in model.overall_balance)

    def test_financing_gap_is_positive(self, model):
        assert all(g > 0 for g in model.financing_gap)


# ── Debt Service
class TestDebtService:
    def test_dom_ds_is_principal_plus_interest(self, model):
        for i in range(8):
            expected = model.dom_principal[i] + model.interest_dom[i]
            assert abs(model.dom_ds[i] - expected) < 0.01

    def test_total_ds_is_sum(self, model):
        for i in range(8):
            expected = model.dom_ds[i] + model.ext_ds[i]
            assert abs(model.total_ds[i] - expected) < 0.01

    def test_ds_positive(self, model):
        assert all(d > 0 for d in model.total_ds)


# ── Ratios
class TestRatios:
    def test_debt_gdp_length(self, model):
        assert len(model.debt_gdp) == 8

    def test_debt_gdp_in_plausible_range(self, model):
        assert all(0 < r < 200 for r in model.debt_gdp)

    def test_ds_revenue_critical_year(self, model):
        # FY24/25 (index 2) should be ~82.2% — the headline finding
        assert 75 < model.ds_revenue[2] < 90

    def test_ds_revenue_above_threshold(self, model):
        # FY24/25 should breach the 30% threshold
        thr = model.thresholds["ds_revenue_pct"]
        assert model.ds_revenue[2] > thr

    def test_interest_revenue_positive(self, model):
        assert all(r > 0 for r in model.interest_rev)


# ── DSA Indicators
class TestDSAIndicators:
    def test_pv_ext_debt_less_than_face(self, model):
        # PV should be less than face value (concessionality discount)
        for i in range(8):
            assert model.pv_ext_debt[i] < model.ext_debt[i]

    def test_pv_total_gdp_length(self, model):
        assert len(model.pv_total_gdp) == 8

    def test_pv_total_gdp_below_anchor(self, model):
        # FY24/25 is currently below the 55% composite anchor (~46.9%)
        assert 40 < model.pv_total_gdp[2] < 55.0

    def test_ext_ds_revenue_positive(self, model):
        assert all(r > 0 for r in model.ext_ds_revenue)


# ── Threshold Checker
class TestThresholdChecker:
    def test_returns_all_indicators(self, model):
        checks = model.check_thresholds(2)
        expected_keys = [
            "pv_ext_debt_gdp", "pv_ext_debt_exports", "pv_ext_debt_revenue",
            "ext_ds_exports", "ext_ds_revenue", "pv_total_debt_gdp",
            "ds_revenue", "ds_gdp", "interest_revenue",
        ]
        for key in expected_keys:
            assert key in checks

    def test_ds_revenue_breach_flagged(self, model):
        checks = model.check_thresholds(2)
        assert checks["ds_revenue"]["breach"] is True

    def test_pv_total_debt_gdp_not_breached(self, model):
        checks = model.check_thresholds(2)
        assert checks["pv_total_debt_gdp"]["breach"] is False

    def test_ds_revenue_is_severe(self, model):
        checks = model.check_thresholds(2)
        assert checks["ds_revenue"]["severe"] is True  # >1.5x threshold

    def test_check_structure(self, model):
        checks = model.check_thresholds(2)
        for key, result in checks.items():
            assert "value" in result
            assert "threshold" in result
            assert "breach" in result
            assert "severe" in result
            assert isinstance(result["breach"], bool)
            assert isinstance(result["severe"], bool)

    def test_valid_year_indices(self, model):
        for i in range(model.n):
            checks = model.check_thresholds(i)
            assert len(checks) == 9
