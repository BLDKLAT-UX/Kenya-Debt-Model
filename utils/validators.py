# -*- coding: utf-8 -*-
"""
utils/validators.py
===================
Input validation layer for the Kenya Debt Model.

Validates all data loaded from config/assumptions.yaml before any
output is generated. Raises descriptive errors on bad data so
problems are caught early — not buried in a corrupt Excel file.

Usage::

    from utils.validators import validate_assumptions
    cfg = load_config()
    validate_assumptions(cfg)   # raises ValidationError if anything is wrong
"""

from __future__ import annotations

from typing import Any

from utils.logger import get_logger

log = get_logger(__name__)


class ValidationError(Exception):
    """Raised when model assumptions fail validation checks."""


# ── Helpers
def _check_length(name: str, series: list, expected: int) -> None:
    if len(series) != expected:
        raise ValidationError(
            f"{name}: expected {expected} values, got {len(series)}"
        )


def _check_positive(name: str, series: list) -> None:
    for i, v in enumerate(series):
        if v is not None and v < 0:
            raise ValidationError(
                f"{name}[{i}] = {v}: all values must be non-negative"
            )


def _check_range(name: str, series: list, lo: float, hi: float) -> None:
    for i, v in enumerate(series):
        if v is not None and not (lo <= v <= hi):
            raise ValidationError(
                f"{name}[{i}] = {v}: expected value in [{lo}, {hi}]"
            )


def _check_monotonic_increasing(name: str, series: list) -> None:
    for i in range(1, len(series)):
        if series[i] < series[i - 1]:
            log.warning(
                "%s is not monotonically increasing at index %d "
                "(%s → %s) — verify assumptions",
                name, i, series[i - 1], series[i],
            )


def _check_shares_sum_to_one(name: str, shares: dict) -> None:
    total = sum(shares.values())
    if abs(total - 1.0) > 0.001:
        raise ValidationError(
            f"{name}: holder shares sum to {total:.4f}, expected 1.0"
        )


def _get(cfg: dict, *keys: str) -> Any:
    """Safely navigate nested dict keys."""
    node = cfg
    for k in keys:
        if not isinstance(node, dict) or k not in node:
            raise ValidationError(
                f"Missing required config key: {' → '.join(keys)}"
            )
        node = node[k]
    return node


# ── Main validation entry point
def validate_assumptions(cfg: dict) -> None:
    """
    Run all validation checks on loaded config data.

    Parameters
    ----------
    cfg : dict
        Loaded YAML config (from config/assumptions.yaml).

    Raises
    ------
    ValidationError
        If any check fails. The message describes exactly what is wrong.
    """
    log.info("Validating model assumptions ...")

    n_years = len(_get(cfg, "years", "labels"))
    if n_years == 0:
        raise ValidationError("years.labels is empty")

    _validate_years(cfg, n_years)
    _validate_real_sector(cfg, n_years)
    _validate_external(cfg, n_years)
    _validate_rates(cfg, n_years)
    _validate_debt_stock(cfg, n_years)
    _validate_fiscal(cfg, n_years)
    _validate_thresholds(cfg)
    _validate_stress_tests(cfg)

    log.info("✓ All validation checks passed (%d fiscal years)", n_years)


def _validate_years(cfg: dict, n: int) -> None:
    integers = _get(cfg, "years", "integers")
    _check_length("years.integers", integers, n)
    proj_idx = _get(cfg, "years", "projection_from_index")
    if not (0 <= proj_idx < n):
        raise ValidationError(
            f"years.projection_from_index={proj_idx} is out of range [0, {n-1}]"
        )
    log.debug("years: OK (%d labels, projection from index %d)", n, proj_idx)


def _validate_real_sector(cfg: dict, n: int) -> None:
    rs = _get(cfg, "real_sector")
    fields = [
        ("nominal_gdp_kes_bn",    0,     100_000),
        ("real_gdp_growth_pct",  -10,    20),
        ("gdp_deflator_pct",     -5,     30),
        ("cpi_inflation_avg_pct",-5,     30),
        ("usd_kes_exchange_rate", 50,    500),
    ]
    for field, lo, hi in fields:
        series = _get(cfg, "real_sector", field)
        _check_length(f"real_sector.{field}", series, n)
        _check_range(f"real_sector.{field}", series, lo, hi)

    _check_monotonic_increasing(
        "real_sector.nominal_gdp_kes_bn",
        _get(cfg, "real_sector", "nominal_gdp_kes_bn")
    )
    log.debug("real_sector: OK")


def _validate_external(cfg: dict, n: int) -> None:
    fields = [
        ("exports_kes_bn",          0, 50_000),
        ("imports_kes_bn",          0, 50_000),
        ("current_account_pct_gdp",-50, 50),
        ("fdi_inflows_usd_mn",      0, 100_000),
        ("reserves_months_imports", 0, 36),
        ("remittances_usd_mn",      0, 100_000),
    ]
    for field, lo, hi in fields:
        series = _get(cfg, "external_sector", field)
        _check_length(f"external_sector.{field}", series, n)
        _check_range(f"external_sector.{field}", series, lo, hi)
    log.debug("external_sector: OK")


def _validate_rates(cfg: dict, n: int) -> None:
    rate_fields = [
        "cbr", "tbill_91d", "tbill_182d", "tbill_364d",
        "bond_2yr", "bond_5yr", "bond_10yr",
        "concessional_external", "china_exim", "eurobond_new_issuance",
    ]
    for field in rate_fields:
        series = _get(cfg, "rates", field)
        _check_length(f"rates.{field}", series, n)
        _check_range(f"rates.{field}", series, 0, 50)
    log.debug("rates: OK")


def _validate_debt_stock(cfg: dict, n: int) -> None:
    dom_fields = [
        "tbills_91d", "tbills_182d", "tbills_364d",
        "bonds_2yr", "bonds_5yr", "bonds_10yr", "bonds_15yr",
        "bonds_25yr", "bonds_infra", "cbk_overdraft",
    ]
    for field in dom_fields:
        series = _get(cfg, "debt_stock", "domestic", field)
        _check_length(f"debt_stock.domestic.{field}", series, n)
        _check_positive(f"debt_stock.domestic.{field}", series)

    ext_fields = [
        "world_bank_ida", "afdb", "imf", "china_exim",
        "paris_club", "eurobonds", "other_bilateral",
    ]
    for field in ext_fields:
        series = _get(cfg, "debt_stock", "external", field)
        _check_length(f"debt_stock.external.{field}", series, n)
        _check_positive(f"debt_stock.external.{field}", series)

    guar_fields = ["kengen", "kq", "kpa"]
    for field in guar_fields:
        series = _get(cfg, "debt_stock", "guaranteed", field)
        _check_length(f"debt_stock.guaranteed.{field}", series, n)
        _check_positive(f"debt_stock.guaranteed.{field}", series)

    shares = _get(cfg, "debt_stock", "domestic", "holder_shares")
    _check_shares_sum_to_one("debt_stock.domestic.holder_shares", shares)

    log.debug("debt_stock: OK")


def _validate_fiscal(cfg: dict, n: int) -> None:
    rev_fields = [
        "total", "tax", "kra_collections",
        "other_tax", "non_tax", "grants",
    ]
    for field in rev_fields:
        series = _get(cfg, "fiscal", "revenue", field)
        _check_length(f"fiscal.revenue.{field}", series, n)
        _check_positive(f"fiscal.revenue.{field}", series)

    exp_fields = [
        "total", "wages", "operations",
        "interest_dom", "interest_ext",
        "development", "net_lending",
    ]
    for field in exp_fields:
        series = _get(cfg, "fiscal", "expenditure", field)
        _check_length(f"fiscal.expenditure.{field}", series, n)
        _check_positive(f"fiscal.expenditure.{field}", series)

    ds_fields = ["domestic_principal", "external_principal"]
    for field in ds_fields:
        series = _get(cfg, "fiscal", "debt_service", field)
        _check_length(f"fiscal.debt_service.{field}", series, n)
        _check_positive(f"fiscal.debt_service.{field}", series)

    # Revenue should be less than expenditure (Kenya runs deficits)
    rev = _get(cfg, "fiscal", "revenue", "total")
    exp = _get(cfg, "fiscal", "expenditure", "total")
    for i, (r, e) in enumerate(zip(rev, exp)):
        if r > e * 1.5:
            log.warning(
                "fiscal: revenue[%d]=%s exceeds 150%% of expenditure — verify data",
                i, r
            )
    log.debug("fiscal: OK")


def _validate_thresholds(cfg: dict) -> None:
    thr = _get(cfg, "thresholds")
    required = [
        "pv_ext_debt_gdp_pct", "pv_ext_debt_exports_pct",
        "pv_ext_debt_revenue_pct", "ext_ds_exports_pct",
        "ext_ds_revenue_pct", "pv_total_debt_gdp_pct",
        "ds_revenue_pct", "ds_gdp_pct", "interest_revenue_pct",
    ]
    for key in required:
        if key not in thr:
            raise ValidationError(f"thresholds.{key} is missing")
        if thr[key] <= 0:
            raise ValidationError(f"thresholds.{key} must be positive, got {thr[key]}")
    log.debug("thresholds: OK")


def _validate_stress_tests(cfg: dict) -> None:
    scenarios = _get(cfg, "stress_tests", "scenarios")
    proj_years = _get(cfg, "stress_tests", "projection_years")
    n_proj = len(proj_years)

    required_keys = ["description", "debt_gdp", "ds_rev", "pv_ext_gdp", "primary_gdp"]
    for scen_name, scen_data in scenarios.items():
        for key in required_keys:
            if key not in scen_data:
                raise ValidationError(
                    f"stress_tests.scenarios.{scen_name}.{key} is missing"
                )
            if key != "description":
                _check_length(
                    f"stress_tests.{scen_name}.{key}",
                    scen_data[key],
                    n_proj
                )
        # Debt/GDP should be in plausible range
        _check_range(
            f"stress_tests.{scen_name}.debt_gdp",
            scen_data["debt_gdp"], 0, 200
        )
    log.debug("stress_tests: OK (%d scenarios, %d projection years)",
              len(scenarios), n_proj)
