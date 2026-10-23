import blpapi


class BloombergClient:
    """
    A standalone client to handle batch requests to the Bloomberg Terminal.
    Includes a 'use_mock' flag to test the app without hitting Bloomberg limits.

    Args:
        host / port: Bloomberg BLPAPI server settings (localhost:8194 default).
        use_mock: When True, returns realistic fake data without a real connection.
        security_type: Bloomberg yellow-key market sector string appended to each
            CUSIP when building the request (e.g. "Muni", "Corp", "Govt").
            Pass None or "" to let Bloomberg auto-detect the security class —
            useful for mixed taxable portfolios that include both Corporates and
            Treasuries under different CUSIP prefixes.
            Defaults to "Muni" for backward compatibility with existing muni flows.
        fields: Optional explicit list of Bloomberg field mnemonics to request.
            When None the client picks the type-appropriate built-in field list.
    """

    # ── Default field list for Municipal portfolios (original behaviour, unchanged) ──
    _MUNI_FIELDS = [
        "SECURITY_NAME", "SECURITY_DES", "SHORT_NAME", "DSPLY_NAME", "ISSUER",
        "CPN", "PX_LAST", "MATURITY", "NXT_CALL_DT", "WORKOUT_DT", "YAS_WORKOUT_DT",
        "YLD_CNV_MID", "YLD_YTW", "YAS_BOND_YLD", "MUNI_YLD_TO_WORST",
        "YLD_YTW_MID", "YLD_YTW_BID", "YLD_TO_WORST",
        "YLD_YTM_MID", "YLD_YTM_BID", "YLD_TO_MATURITY",
        "DUR_ADJ_MID", "MID_MOD_DUR_MTY", "CNVX_MID",
        "BB_COMPOSITE",
        "RTG_SP_UNDERLYING", "RTG_MDY_UNDERLYING", "RTG_MDY_LT_UNDL",
        "RTG_FITCH_UNDERLYING", "RTG_SP_MUNI_LONG_TERM", "RTG_MDY_MUNI_LONG_TERM",
        "RTG_SP_INSURED", "RTG_MOODY_INSURED", "RTG_FITCH_INSURED",
        "RTG_SP_ENHANCED", "RTG_MOODY_ENHANCED", "RTG_FITCH_ENHANCED",
        "RTG_SP", "RTG_MOODY", "RTG_FITCH",
        "INDUSTRY_SECTOR", "INDUSTRY_SUBGROUP", "ISSUER_INDUSTRY",
        "STATE_CODE", "MUNI_TAX_PROV",
        "MUNI_ISSUE_TYP", "MUNI_PURPOSE", "MUNI_SOURCE", "CREDIT_ENHANCEMENTS",
    ]

    # ── Field list for taxable Corporate / Treasury / Agency portfolios ────────
    _TAXABLE_FIELDS = [
        # Identifiers & Name
        "SECURITY_NAME", "SECURITY_DES", "SHORT_NAME", "DSPLY_NAME", "ISSUER",
        # Pricing
        "CPN", "PX_LAST",
        # Dates
        "MATURITY", "NXT_CALL_DT", "WORKOUT_DT", "YAS_WORKOUT_DT",
        # Yields
        "YLD_CNV_MID", "YLD_YTW", "YAS_BOND_YLD",
        "YLD_YTW_MID", "YLD_YTW_BID", "YLD_TO_WORST",
        "YLD_YTM_MID", "YLD_YTM_BID", "YLD_TO_MATURITY",
        # Spreads — key analytics for taxables
        "OAS_SPREAD_MID",    # Option-Adjusted Spread vs. benchmark curve
        "YAS_ZSPREAD_MID",   # Z-spread (constant spread over spot rate curve)
        "YAS_ASW_SPREAD",    # Asset-swap spread
        # Duration & Convexity
        "OAS_DUR",           # OAS-adjusted effective duration (best for callables)
        "DUR_ADJ_MID", "MID_MOD_DUR_MTY", "CNVX_MID",
        # Credit Ratings
        "BB_COMPOSITE",
        "RTG_SP", "RTG_MOODY", "RTG_FITCH",
        "RTG_SP_UNDERLYING", "RTG_MDY_UNDERLYING", "RTG_FITCH_UNDERLYING",
        # Security Classification (replaces muni_issue_type / muni_source)
        "SECURITY_TYP",        # e.g. "Corp", "US GOVERNMENT", "AGENCY"
        "MARKET_SECTOR_DES",   # Bloomberg market sector: "Corp", "Govt", "Mtge"
        "INDUSTRY_SECTOR",     # High-level: "INDUSTRIAL", "FINANCIAL", "UTILITY"
        "INDUSTRY_GROUP",      # Mid-level grouping
        "INDUSTRY_SUBGROUP",   # Detailed: "Applications Software", "Casino Hotels"
        "ISSUER_INDUSTRY",
        # Seniority / Capital Structure
        "PAYMENT_RANK",        # "Sr Unsecured", "Subordinated", "Sr Secured", etc.
        # Geography
        "STATE_CODE",          # Issuer domicile state (US-only; blank for foreign)
        "CNTRY_OF_RISK",       # Country of primary economic risk
        "CNTRY_ISSUE_ISO",     # ISO country code of issuance
    ]

    def __init__(
        self,
        host: str = "localhost",
        port: int = 8194,
        use_mock: bool = True,
        security_type: str | None = "Muni",
        fields: list[str] | None = None,
        is_taxable: bool | None = None,
    ):
        self.host = host
        self.port = port
        self.use_mock = use_mock
        self.security_type = security_type  # yellow-key suffix; None = auto-detect

        if is_taxable is None:
            is_taxable = bool(security_type and security_type.strip().lower() != "muni")
        self.is_taxable = is_taxable

        # Pick field list: caller-supplied > type-appropriate built-in
        if fields is not None:
            self.fields = fields
        elif self.is_taxable:
            self.fields = self._TAXABLE_FIELDS
        else:
            self.fields = self._MUNI_FIELDS

        self.session = None

    def __enter__(self):
        if self.use_mock:
            print("DEBUG: Using MOCKED Bloomberg connection.")
            return self

        options = blpapi.SessionOptions()
        options.setServerHost(self.host)
        options.setServerPort(self.port)
        self.session = blpapi.Session(options)

        if not self.session.start() or not self.session.openService("//blp/refdata"):
            raise RuntimeError("Failed to start BLPAPI session.")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.session is not None:
            self.session.stop()

    def fetch_data(self, cusips: list[str]) -> dict:
        """
        Takes a list of CUSIPs and returns a dict keyed by CUSIP.
        Example return: {"123456789": {"YLD_YTW_MID": 4.5, "CPN": 3.0}}
        """
        if not cusips:
            return {}
        return self._get_mock_data(cusips) if self.use_mock else self._get_live_data(cusips)

    # ── Mock data ──────────────────────────────────────────────────────────────

    def _get_mock_data(self, cusips: list[str]) -> dict:
        print(f"DEBUG: Generating mock Bloomberg data for {len(cusips)} CUSIPs.")
        results = {}

        is_taxable = self.is_taxable

        if is_taxable:
            # Realistic corporate bond mock templates — cycles through a variety of
            # sectors, ratings, seniority tiers, and spread levels.
            templates = [
                {
                    "SECURITY_NAME": "APPLE INC", "ISSUER": "Apple Inc",
                    "CPN": 3.850, "PX_LAST": 97.45,
                    "YLD_YTW_MID": 4.42, "YLD_YTM_MID": 4.42,
                    "MATURITY": "2031-08-04", "NXT_CALL_DT": "2031-05-04",
                    "OAS_DUR": 5.82, "DUR_ADJ_MID": 5.82, "MID_MOD_DUR_MTY": 5.90, "CNVX_MID": 0.41,
                    "OAS_SPREAD_MID": 62,  "YAS_ZSPREAD_MID": 67,  "YAS_ASW_SPREAD": 58,
                    "BB_COMPOSITE": "AA+", "RTG_SP": "AA+", "RTG_MOODY": "Aaa", "RTG_FITCH": "AA+",
                    "SECURITY_TYP": "Corp", "MARKET_SECTOR_DES": "Corp",
                    "INDUSTRY_SECTOR": "INDUSTRIAL", "INDUSTRY_GROUP": "Technology",
                    "INDUSTRY_SUBGROUP": "Computers", "ISSUER_INDUSTRY": "INDUSTRIAL",
                    "PAYMENT_RANK": "Sr Unsecured",
                    "STATE_CODE": "CA", "CNTRY_OF_RISK": "US", "CNTRY_ISSUE_ISO": "US",
                },
                {
                    "SECURITY_NAME": "JPMORGAN CHASE & CO", "ISSUER": "JPMorgan Chase & Co",
                    "CPN": 4.800, "PX_LAST": 99.12,
                    "YLD_YTW_MID": 5.15, "YLD_YTM_MID": 5.08,
                    "MATURITY": "2033-07-23", "NXT_CALL_DT": "2032-07-23",
                    "OAS_DUR": 5.41, "DUR_ADJ_MID": 5.41, "MID_MOD_DUR_MTY": 5.55, "CNVX_MID": 0.38,
                    "OAS_SPREAD_MID": 118, "YAS_ZSPREAD_MID": 124, "YAS_ASW_SPREAD": 112,
                    "BB_COMPOSITE": "A+", "RTG_SP": "A+", "RTG_MOODY": "Aa2", "RTG_FITCH": "AA-",
                    "SECURITY_TYP": "Corp", "MARKET_SECTOR_DES": "Corp",
                    "INDUSTRY_SECTOR": "FINANCIAL", "INDUSTRY_GROUP": "Banks",
                    "INDUSTRY_SUBGROUP": "Diversified Banking Inst", "ISSUER_INDUSTRY": "FINANCIAL",
                    "PAYMENT_RANK": "Sr Unsecured",
                    "STATE_CODE": "NY", "CNTRY_OF_RISK": "US", "CNTRY_ISSUE_ISO": "US",
                },
                {
                    "SECURITY_NAME": "AMAZON.COM INC", "ISSUER": "Amazon.com Inc",
                    "CPN": 3.100, "PX_LAST": 91.80,
                    "YLD_YTW_MID": 4.88, "YLD_YTM_MID": 4.88,
                    "MATURITY": "2031-05-12", "NXT_CALL_DT": "2031-02-12",
                    "OAS_DUR": 5.70, "DUR_ADJ_MID": 5.70, "MID_MOD_DUR_MTY": 5.80, "CNVX_MID": 0.40,
                    "OAS_SPREAD_MID": 88,  "YAS_ZSPREAD_MID": 93,  "YAS_ASW_SPREAD": 82,
                    "BB_COMPOSITE": "AA",  "RTG_SP": "AA",  "RTG_MOODY": "A1",  "RTG_FITCH": "AA",
                    "SECURITY_TYP": "Corp", "MARKET_SECTOR_DES": "Corp",
                    "INDUSTRY_SECTOR": "INDUSTRIAL", "INDUSTRY_GROUP": "Retail",
                    "INDUSTRY_SUBGROUP": "Internet Content-Entmnt", "ISSUER_INDUSTRY": "INDUSTRIAL",
                    "PAYMENT_RANK": "Sr Unsecured",
                    "STATE_CODE": "WA", "CNTRY_OF_RISK": "US", "CNTRY_ISSUE_ISO": "US",
                },
                {
                    "SECURITY_NAME": "BERKSHIRE HATHAWAY", "ISSUER": "Berkshire Hathaway Inc",
                    "CPN": 4.200, "PX_LAST": 98.55,
                    "YLD_YTW_MID": 4.62, "YLD_YTM_MID": 4.62,
                    "MATURITY": "2028-08-15", "NXT_CALL_DT": "NC",
                    "OAS_DUR": 3.45, "DUR_ADJ_MID": 3.45, "MID_MOD_DUR_MTY": 3.52, "CNVX_MID": 0.15,
                    "OAS_SPREAD_MID": 58,  "YAS_ZSPREAD_MID": 62,  "YAS_ASW_SPREAD": 54,
                    "BB_COMPOSITE": "AA",  "RTG_SP": "AA",  "RTG_MOODY": "Aa2", "RTG_FITCH": "A+",
                    "SECURITY_TYP": "Corp", "MARKET_SECTOR_DES": "Corp",
                    "INDUSTRY_SECTOR": "FINANCIAL", "INDUSTRY_GROUP": "Insurance",
                    "INDUSTRY_SUBGROUP": "Multi-line Insurance", "ISSUER_INDUSTRY": "FINANCIAL",
                    "PAYMENT_RANK": "Sr Unsecured",
                    "STATE_CODE": "NE", "CNTRY_OF_RISK": "US", "CNTRY_ISSUE_ISO": "US",
                },
                {
                    "SECURITY_NAME": "US TREASURY NOTE", "ISSUER": "United States Treasury",
                    "CPN": 4.625, "PX_LAST": 99.80,
                    "YLD_YTW_MID": 4.65, "YLD_YTM_MID": 4.65,
                    "MATURITY": "2029-02-15", "NXT_CALL_DT": "NC",
                    "OAS_DUR": 3.85, "DUR_ADJ_MID": 3.85, "MID_MOD_DUR_MTY": 3.90, "CNVX_MID": 0.18,
                    "OAS_SPREAD_MID": 0,   "YAS_ZSPREAD_MID": 0,   "YAS_ASW_SPREAD": -5,
                    "BB_COMPOSITE": "AAA", "RTG_SP": "AA+", "RTG_MOODY": "Aaa", "RTG_FITCH": "AAA",
                    "SECURITY_TYP": "US GOVERNMENT", "MARKET_SECTOR_DES": "Govt",
                    "INDUSTRY_SECTOR": "GOVERNMENT", "INDUSTRY_GROUP": "Sovereign",
                    "INDUSTRY_SUBGROUP": "US Treasuries", "ISSUER_INDUSTRY": "GOVERNMENT",
                    "PAYMENT_RANK": "Sr Unsecured",
                    "STATE_CODE": "DC", "CNTRY_OF_RISK": "US", "CNTRY_ISSUE_ISO": "US",
                },
            ]
        else:
            # Original muni-style templates (unchanged for backward compat)
            templates = [
                {
                    "SECURITY_NAME": "NEW YORK ST URBAN DEV", "CPN": 5.000,
                    "PX_LAST": 105.20, "YLD_YTW_MID": 3.45, "YLD_YTM_MID": 3.75,
                    "MATURITY": "2035-03-15", "NXT_CALL_DT": "2025-03-15",
                    "DUR_ADJ_MID": 7.2, "CNVX_MID": 0.65,
                    "RTG_SP": "AA+", "RTG_MOODY": "Aa1",
                    "INDUSTRY_SECTOR": "Government", "STATE_CODE": "NY",
                    "MUNI_TAX_PROV": "Exempt",
                },
                {
                    "SECURITY_NAME": "CALIFORNIA GOVT", "CPN": 4.500,
                    "PX_LAST": 102.10, "YLD_YTW_MID": 3.20, "YLD_YTM_MID": 3.40,
                    "MATURITY": "2030-10-01", "NXT_CALL_DT": "2026-10-01",
                    "DUR_ADJ_MID": 5.5, "CNVX_MID": 0.40,
                    "RTG_SP": "AA-", "RTG_MOODY": "Aa2",
                    "INDUSTRY_SECTOR": "Government", "STATE_CODE": "CA",
                    "MUNI_TAX_PROV": "Exempt",
                },
            ]

        for i, cusip in enumerate(cusips):
            results[cusip] = templates[i % len(templates)].copy()

        return results


    # ── Live Bloomberg data ────────────────────────────────────────────────────

    def _get_live_data(self, cusips: list[str]) -> dict:
        ref_data_service = self.session.getService("//blp/refdata")
        request = ref_data_service.createRequest("ReferenceDataRequest")

        # 1. Batch-append all CUSIPs, with optional yellow-key suffix
        for cusip in cusips:
            if self.security_type:
                request.append("securities", f"/cusip/{cusip} {self.security_type}")
            else:
                # No yellow key — Bloomberg auto-detects the security class.
                # Reliable for mixed Corp / Govt / Agency portfolios.
                request.append("securities", f"/cusip/{cusip}")

        # 2. Append the type-appropriate field list
        for field in self.fields:
            request.append("fields", field)

        self.session.sendRequest(request)

        results = {}

        # 3. Parse the asynchronous event loop
        while True:
            event = self.session.nextEvent(500)

            if event.eventType() == blpapi.Event.REQUEST_STATUS:
                break

            for msg in event:
                if msg.hasElement("responseError"):
                    print("Bloomberg blocked this request (Limit or Permission error).")
                    continue

                if msg.hasElement("securityData"):
                    security_data_array = msg.getElement("securityData")
                    for i in range(security_data_array.numValues()):
                        security = security_data_array.getValueAsElement(i)
                        ticker = security.getElementAsString("security")

                        # Extract raw CUSIP from "/cusip/<CUSIP> Corp" etc.
                        raw_cusip = ticker.split("/")[2].split()[0]
                        bond_data = {}

                        field_data = security.getElement("fieldData")
                        for field in field_data.elements():
                            bond_data[str(field.name())] = field.getValue()

                        results[raw_cusip] = bond_data

            if event.eventType() == blpapi.Event.RESPONSE:
                break

        return results
