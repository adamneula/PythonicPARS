class TaxableBond:
    """
    Data container for a single taxable fixed-income holding
    (Corporate, U.S. Treasury, Agency, etc.).

    Mirrors the structure of Bond (muni) but replaces muni-specific fields
    with corp-relevant equivalents:

        muni_issue_type  →  security_type   (Corp / US GOVERNMENT / AGENCY)
        muni_source      →  industry_sector / industry_subgroup (sector breakdown)
        muni_purpose     →  industry_subgroup
        tax_provision    →  hardcoded "Taxable"
        credit_enhancmts →  payment_rank (seniority / capital structure)

    Adds spread analytics that are central to corporate credit analysis:
        oas_spread, z_spread, asw_spread

    All computed fields (portfolio_weight) are initialised to 0.0 and
    should be set by the calling pipeline after the full portfolio is loaded
    (so each weight reflects its share of total market value).
    """

    def __init__(self, cusip: str, face_value: float, bbg_data: dict = None):
        self.cusip = cusip
        self.face_value = float(face_value)

        # Default to empty dict if Bloomberg failed to return data
        if bbg_data is None:
            bbg_data = {}

        # ── Helpers ────────────────────────────────────────────────────────────

        def get_valid_float(fields, default=None):
            for f in fields:
                val = bbg_data.get(f)
                if val is not None and not (isinstance(val, str) and val.startswith("#N/A")):
                    try:
                        return float(val)
                    except (ValueError, TypeError):
                        pass
            return default

        def get_valid_str(fields, default="Unknown"):
            for f in fields:
                val = bbg_data.get(f)
                if val and isinstance(val, str) and not val.startswith("#N/A") and val not in ("N.S.", "N/A"):
                    return val.strip()
            return default

        def get_valid_rating(fields, default="NR"):
            for f in fields:
                val = bbg_data.get(f)
                if val and isinstance(val, str) and not val.startswith("#N/A") and val != "N.S.":
                    return val.strip()
            return default

        # ── Identifiers & Name ─────────────────────────────────────────────────
        self.security_name = get_valid_str(
            ["SHORT_NAME", "DSPLY_NAME", "ISSUER", "SECURITY_DES", "SECURITY_NAME"]
        )
        self.issuer = bbg_data.get("ISSUER", "Unknown")

        # ── Pricing ────────────────────────────────────────────────────────────
        self.current_price = get_valid_float(["PX_LAST"], default=100.0)
        self.coupon = get_valid_float(["CPN"])

        # ── Yields ─────────────────────────────────────────────────────────────
        self.yield_to_worst = get_valid_float([
            "YLD_CNV_MID", "YLD_YTW", "YAS_BOND_YLD",
            "YLD_YTW_MID", "YLD_YTW_BID", "YLD_TO_WORST",
        ])
        self.yield_to_maturity = get_valid_float([
            "YLD_YTM_MID", "YLD_YTM_BID", "YLD_TO_MATURITY",
        ])

        # ── Spreads (Corp-specific analytics) ──────────────────────────────────
        self.oas_spread = get_valid_float(["OAS_SPREAD_MID"])   # Option-Adjusted Spread (bps)
        self.z_spread   = get_valid_float(["YAS_ZSPREAD_MID"])  # Z-spread (bps)
        self.asw_spread = get_valid_float(["YAS_ASW_SPREAD"])   # Asset-swap spread (bps)

        # ── Dates ──────────────────────────────────────────────────────────────
        self.maturity_date = bbg_data.get("MATURITY", "N/A")
        self.next_call_date = bbg_data.get("NXT_CALL_DT", "NC")  # "NC" = Non-Callable
        self.is_callable = self.next_call_date != "NC"

        # Effective / workout date: OAM, YAS workout, or stated maturity fallback
        self.effective_maturity_date = (
            bbg_data.get("YAS_WORKOUT_DT")
            or bbg_data.get("WORKOUT_DT")
            or self.maturity_date
        )

        # ── Duration & Convexity ───────────────────────────────────────────────
        # OAS_DUR is preferred for callable corporates; fall back to DUR_ADJ_MID
        self.effective_duration = get_valid_float(["OAS_DUR", "DUR_ADJ_MID", "MID_MOD_DUR_MTY"])
        self.modified_duration  = get_valid_float(["MID_MOD_DUR_MTY", "DUR_ADJ_MID"])
        self.convexity          = get_valid_float(["CNVX_MID"])

        # ── Credit Ratings ─────────────────────────────────────────────────────
        # Corporates use RTG_SP / RTG_MOODY / RTG_FITCH directly;
        # no insured / enhanced layers, but we still try _UNDERLYING as a fallback.
        self.composite_rating = get_valid_rating(["BB_COMPOSITE"])
        self.sp_rating     = get_valid_rating(["RTG_SP_UNDERLYING", "RTG_SP"])
        self.moodys_rating = get_valid_rating(["RTG_MDY_UNDERLYING", "RTG_MOODY"])
        self.fitch_rating  = get_valid_rating(["RTG_FITCH_UNDERLYING", "RTG_FITCH"])

        # ── Security Classification (replaces muni_issue_type / muni_source) ───
        self.security_type    = get_valid_str(["SECURITY_TYP"])        # Corp / US GOVERNMENT / AGENCY
        self.market_sector    = get_valid_str(["MARKET_SECTOR_DES"])   # Corp / Govt / Mtge
        self.industry_sector  = get_valid_str(["ISSUER_INDUSTRY", "INDUSTRY_SECTOR"])  # INDUSTRIAL / FINANCIAL
        self.industry_group   = get_valid_str(["INDUSTRY_GROUP"])
        self.industry_subgroup = get_valid_str(["INDUSTRY_SUBGROUP"])  # "Casino Hotels", "REITS-Apartments"

        # ── Seniority / Capital Structure (replaces credit_enhancements) ───────
        self.payment_rank = get_valid_str(["PAYMENT_RANK"])  # "Sr Unsecured", "Subordinated"

        # ── Tax Status ─────────────────────────────────────────────────────────
        # Taxable bonds are always fully taxable — hard-coded, no BBG lookup needed.
        self.tax_provision = "Taxable"

        # ── Geography ──────────────────────────────────────────────────────────
        self.state_code      = bbg_data.get("STATE_CODE", "Unknown")   # Issuer domicile state
        self.country_of_risk = bbg_data.get("CNTRY_OF_RISK", "Unknown")
        self.country_iso     = bbg_data.get("CNTRY_ISSUE_ISO", "Unknown")

        # ── Backward-Compatibility Aliases (so Muni templates/scripts don't break) ───
        self.muni_issue_type     = self.security_type
        self.muni_source         = self.industry_sector    # groups by sector for structure charts
        self.muni_purpose        = self.industry_subgroup  # granular industry classification
        self.credit_enhancements = self.payment_rank

        # ── Computed Metrics (set by pipeline post-load) ───────────────────────
        self.market_value             = (self.face_value / 100.0) * self.current_price
        self.annual_income            = self.face_value * (self.coupon / 100.0) if self.coupon is not None else 0.0
        self.portfolio_weight         = 0.0   # populated by mainTaxible.py after full portfolio is loaded
        self.taxable_equivalent_yield = self.yield_to_worst or 0.0  # Already taxable!

    def __repr__(self):
        return (
            f"<TaxableBond {self.cusip} | {self.security_name} "
            f"| {self.security_type} | YTW: {self.yield_to_worst}% "
            f"| OAS: {self.oas_spread}bps>"
        )
