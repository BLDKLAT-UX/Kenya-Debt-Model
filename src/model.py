# -*- coding: utf-8 -*-
"""
src/model.py
============
Core calculations for the Kenya DSA model.

All data is read from config/assumptions.yaml via utils/config_loader.
No numbers are hardcoded here.

Public API
----------
    from src.model import DebtModel
    model = DebtModel()
    model.compute()           # run all calculations
    print(model.debt_gdp)     # [52.3, 57.1, 63.7, ...]
"""

from __future__ import annotations

from utils.config_loader import load_config
from utils.logger import get_logger

log = get_logger(__name__)


class DebtModel:
    """
    Kenya Public Debt Sustainability Analysis — core model.

    Loads assumptions from config/assumptions.yaml, computes all
    derived series, and exposes them as attributes for use by
    Excel, PDF, and dashboard output modules.
    """

    def __init__(self, config_path: str | None = None) -> None:
        self.cfg = load_config(config_path)
        self._computed = False

        # shortcuts
        self.years = self.cfg["years"]["labels"]
        self.n = len(self.years)
        self.proj_idx = self.cfg["years"]["projection_from_index"]
        self.proj_years = self.years[self.proj_idx:]

        log.debug("DebtModel initialised: %d fiscal years", self.n)

    # ──────────────────────────────────────────────────────────────
    # PUBLIC
    # ──────────────────────────────────────────────────────────────

    def compute(self) -> "DebtModel":
        """Run all derived calculations. Returns self for chaining."""
        log.info("Computing model ...")
        self._compute_real_sector()
        self._compute_debt_stock()
        self._compute_fiscal()
        self._compute_debt_service()
        self._compute_ratios()
        self._compute_dsa_indicators()
        self._computed = True
        log.info("Model computation complete")
        return self

    def _require_computed(self) -> None:
        if not self._computed:
            raise RuntimeError("Call model.compute() before accessing results.")

    # ──────────────────────────────────────────────────────────────
    # REAL SECTOR
    # ──────────────────────────────────────────────────────────────

    def _compute_real_sector(self) -> None:
        rs = self.cfg["real_sector"]
        self.gdp = rs["nominal_gdp_kes_bn"]
        self.real_growth = rs["real_gdp_growth_pct"]
        self.deflator = rs["gdp_deflator_pct"]
        self.cpi = rs["cpi_inflation_avg_pct"]
        self.fx_rate = rs["usd_kes_exchange_rate"]
        log.debug("Real sector: OK")

    # ──────────────────────────────────────────────────────────────
    # DEBT STOCK
    # ──────────────────────────────────────────────────────────────

    def _compute_debt_stock(self) -> None:
        dom = self.cfg["debt_stock"]["domestic"]
        ext = self.cfg["debt_stock"]["external"]
        guar = self.cfg["debt_stock"]["guaranteed"]

        # Domestic
        self.tbills = [
            dom["tbills_91d"][i] + dom["tbills_182d"][i] + dom["tbills_364d"][i]
            for i in range(self.n)
        ]
        self.bonds = [
            dom["bonds_2yr"][i] + dom["bonds_5yr"][i] + dom["bonds_10yr"][i]
            + dom["bonds_15yr"][i] + dom["bonds_25yr"][i] + dom["bonds_infra"][i]
            for i in range(self.n)
        ]
        self.cbk_od = dom["cbk_overdraft"]
        self.dom_debt = [
            self.tbills[i] + self.bonds[i] + self.cbk_od[i]
            for i in range(self.n)
        ]

        # Holder breakdown
        hs = dom["holder_shares"]
        self.hold_banks   = [round(self.dom_debt[i] * hs["commercial_banks"], 1) for i in range(self.n)]
        self.hold_pension = [round(self.dom_debt[i] * hs["pension_funds"], 1)    for i in range(self.n)]
        self.hold_cbk     = [round(self.dom_debt[i] * hs["cbk"], 1)              for i in range(self.n)]
        self.hold_other   = [round(self.dom_debt[i] * hs["other"], 1)            for i in range(self.n)]

        # External
        self.ext_multilateral = [
            ext["world_bank_ida"][i] + ext["afdb"][i] + ext["imf"][i]
            for i in range(self.n)
        ]
        self.ext_bilateral = [
            ext["china_exim"][i] + ext["paris_club"][i] + ext["other_bilateral"][i]
            for i in range(self.n)
        ]
        self.ext_commercial = ext["eurobonds"]
        self.ext_debt = [
            self.ext_multilateral[i] + self.ext_bilateral[i] + self.ext_commercial[i]
            for i in range(self.n)
        ]

        # Guaranteed
        self.guar_debt = [
            guar["kengen"][i] + guar["kq"][i] + guar["kpa"][i]
            for i in range(self.n)
        ]

        # Total
        self.total_debt = [
            self.dom_debt[i] + self.ext_debt[i] + self.guar_debt[i]
            for i in range(self.n)
        ]

        log.debug(
            "Debt stock FY24/25: DOM=%s EXT=%s GUAR=%s TOTAL=%s",
            self.dom_debt[2], self.ext_debt[2],
            self.guar_debt[2], self.total_debt[2],
        )

    # ──────────────────────────────────────────────────────────────
    # FISCAL
    # ──────────────────────────────────────────────────────────────

    def _compute_fiscal(self) -> None:
        rev = self.cfg["fiscal"]["revenue"]
        exp = self.cfg["fiscal"]["expenditure"]

        self.revenue = rev["total"]
        self.tax_rev = rev["tax"]
        self.non_tax = rev["non_tax"]
        self.grants  = rev["grants"]

        self.expenditure  = exp["total"]
        self.wages        = exp["wages"]
        self.interest_dom = exp["interest_dom"]
        self.interest_ext = exp["interest_ext"]
        self.interest_total = [
            self.interest_dom[i] + self.interest_ext[i]
            for i in range(self.n)
        ]
        self.development = exp["development"]

        self.overall_balance = [
            self.revenue[i] - self.expenditure[i]
            for i in range(self.n)
        ]
        self.primary_balance = [
            self.overall_balance[i] + self.interest_total[i]
            for i in range(self.n)
        ]
        self.financing_gap = [-b for b in self.overall_balance]

        log.debug("Fiscal: OK")

    # ──────────────────────────────────────────────────────────────
    # DEBT SERVICE
    # ──────────────────────────────────────────────────────────────

    def _compute_debt_service(self) -> None:
        ds = self.cfg["fiscal"]["debt_service"]

        self.dom_principal = ds["domestic_principal"]
        self.ext_principal = ds["external_principal"]
        self.dom_ds = [
            self.dom_principal[i] + self.interest_dom[i]
            for i in range(self.n)
        ]
        self.ext_ds = [
            self.ext_principal[i] + self.interest_ext[i]
            for i in range(self.n)
        ]
        self.total_ds = [
            self.dom_ds[i] + self.ext_ds[i]
            for i in range(self.n)
        ]
        self.total_principal = [
            self.dom_principal[i] + self.ext_principal[i]
            for i in range(self.n)
        ]

        log.debug(
            "Debt service FY24/25: DOM=%s EXT=%s TOTAL=%s",
            self.dom_ds[2], self.ext_ds[2], self.total_ds[2],
        )

    # ──────────────────────────────────────────────────────────────
    # RATIOS (all as %)
    # ──────────────────────────────────────────────────────────────

    def _compute_ratios(self) -> None:
        exports = self.cfg["external_sector"]["exports_kes_bn"]

        def pct(num, den):
            return [round(n / d * 100, 1) for n, d in zip(num, den)]

        self.debt_gdp        = pct(self.total_debt,     self.gdp)
        self.dom_debt_gdp    = pct(self.dom_debt,       self.gdp)
        self.ext_debt_gdp    = pct(self.ext_debt,       self.gdp)
        self.revenue_gdp     = pct(self.revenue,        self.gdp)
        self.expenditure_gdp = pct(self.expenditure,    self.gdp)
        self.primary_gdp     = pct(self.primary_balance,self.gdp)
        self.overall_gdp     = pct(self.overall_balance,self.gdp)
        self.ds_revenue      = pct(self.total_ds,       self.revenue)
        self.ds_gdp          = pct(self.total_ds,       self.gdp)
        self.ds_exports      = pct(self.total_ds,       exports)
        self.interest_rev    = pct(self.interest_total, self.revenue)
        self.interest_gdp    = pct(self.interest_total, self.gdp)

        log.debug(
            "Ratios FY24/25: Debt/GDP=%.1f%%  DS/Rev=%.1f%%  Int/Rev=%.1f%%",
            self.debt_gdp[2], self.ds_revenue[2], self.interest_rev[2],
        )

    # ──────────────────────────────────────────────────────────────
    # DSA INDICATORS (PV-based)
    # ──────────────────────────────────────────────────────────────

    def _compute_dsa_indicators(self) -> None:
        """
        Compute IMF LIC-DSF Present Value indicators.
        PV discount: 30% concessionality applied to eligible external portion (60%).
        """
        exports = self.cfg["external_sector"]["exports_kes_bn"]
        pv_discount = 0.30 * 0.60  # net PV haircut

        self.pv_ext_debt = [
            round(self.ext_debt[i] * (1 - pv_discount), 1)
            for i in range(self.n)
        ]
        self.pv_total_debt = [
            self.dom_debt[i] + self.pv_ext_debt[i] + self.guar_debt[i]
            for i in range(self.n)
        ]

        def pct(num, den):
            return [round(n / d * 100, 1) for n, d in zip(num, den)]

        self.pv_ext_gdp     = pct(self.pv_ext_debt,   self.gdp)
        self.pv_ext_exports = pct(self.pv_ext_debt,   exports)
        self.pv_ext_revenue = pct(self.pv_ext_debt,   self.revenue)
        self.pv_total_gdp   = pct(self.pv_total_debt, self.gdp)
        self.ext_ds_exports = pct(self.ext_ds,        exports)
        self.ext_ds_revenue = pct(self.ext_ds,        self.revenue)

        log.debug(
            "DSA FY24/25: PV Total/GDP=%.1f%%  DS/Rev=%.1f%%",
            self.pv_total_gdp[2], self.ds_revenue[2],
        )

    # ──────────────────────────────────────────────────────────────
    # THRESHOLD CHECKER
    # ──────────────────────────────────────────────────────────────

    def check_thresholds(self, year_index: int = 2) -> dict[str, dict]:
        """
        Compare key indicators against IMF LIC-DSF thresholds
        for a given year index (default 2 = FY24/25).

        Returns a dict of {indicator: {value, threshold, breach, severe}}.
        """
        self._require_computed()
        thr = self.cfg["thresholds"]
        i = year_index

        checks = {
            "pv_ext_debt_gdp":      (self.pv_ext_gdp[i],     thr["pv_ext_debt_gdp_pct"]),
            "pv_ext_debt_exports":  (self.pv_ext_exports[i],  thr["pv_ext_debt_exports_pct"]),
            "pv_ext_debt_revenue":  (self.pv_ext_revenue[i],  thr["pv_ext_debt_revenue_pct"]),
            "ext_ds_exports":       (self.ext_ds_exports[i],  thr["ext_ds_exports_pct"]),
            "ext_ds_revenue":       (self.ext_ds_revenue[i],  thr["ext_ds_revenue_pct"]),
            "pv_total_debt_gdp":    (self.pv_total_gdp[i],    thr["pv_total_debt_gdp_pct"]),
            "ds_revenue":           (self.ds_revenue[i],      thr["ds_revenue_pct"]),
            "ds_gdp":               (self.ds_gdp[i],          thr["ds_gdp_pct"]),
            "interest_revenue":     (self.interest_rev[i],    thr["interest_revenue_pct"]),
        }

        results = {}
        for key, (value, threshold) in checks.items():
            results[key] = {
                "value":     value,
                "threshold": threshold,
                "breach":    value >= threshold,
                "severe":    value >= threshold * 1.5,
            }
        return results

    # ──────────────────────────────────────────────────────────────
    # CONVENIENCE PROPERTIES
    # ──────────────────────────────────────────────────────────────

    @property
    def scenarios(self) -> dict:
        """Raw stress test scenario data from config."""
        return self.cfg["stress_tests"]["scenarios"]

    @property
    def thresholds(self) -> dict:
        """IMF LIC-DSF thresholds from config."""
        return self.cfg["thresholds"]

    @property
    def meta(self) -> dict:
        """Model metadata from config."""
        return self.cfg["meta"]
