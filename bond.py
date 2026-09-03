class Bond:
    """
    A simple, clean data container for a single holding.
    It takes the raw CUSIP, Face Value, and the dictionary returned from Bloomberg,
    and maps everything to readable Python variables.
    """
    def __init__(self, cusip: str, face_value: float, bbg_data: dict = None):
        self.cusip = cusip
        self.face_value = float(face_value)
        
        # Default to empty dict if Bloomberg failed to return data for this CUSIP
        if bbg_data is None:
            bbg_data = {}

        # Identifiers & Pricing
        self.security_name = "Unknown"
        for field in ["SHORT_NAME", "DSPLY_NAME", "ISSUER", "SECURITY_DES", "SECURITY_NAME"]:
            val = bbg_data.get(field)
            # If the value exists and isn't a Bloomberg error string, use it and stop searching!
            if val and isinstance(val, str) and not val.startswith("#N/A"):
                self.security_name = val
                break
                
        self.current_price = bbg_data.get("PX_LAST") or 100.0 # Default to Par
        
        # Helper function to get valid floats
        def get_valid_float(fields, default=0.0):
            for f in fields:
                val = bbg_data.get(f)
                if val is not None and not (isinstance(val, str) and val.startswith("#N/A")):
                    try:
                        return float(val)
                    except ValueError:
                        pass
            return default
            
        # Helper function to get valid ratings (ignores N.S. and #N/A)
        def get_valid_rating(fields, default="NR"):
            for f in fields:
                val = bbg_data.get(f)
                if val and isinstance(val, str) and not val.startswith("#N/A") and val != "N.S.":
                    return val
            return default

        # Yields & Coupon
        self.coupon = get_valid_float(["CPN"])
        
        # Pull the actual Yield to Worst (using YAS calculated fields first)
        self.yield_to_worst = get_valid_float([
            "YLD_YTW", "YAS_BOND_YLD", "MUNI_YLD_TO_WORST", 
            "YLD_YTW_MID", "YLD_YTW_BID", "YLD_TO_WORST"
        ])
        
        self.yield_to_maturity = get_valid_float(["YLD_YTM_MID", "YLD_YTM_BID", "YLD_TO_MATURITY"])
        
        # Maturity & Calls
        self.maturity_date = bbg_data.get("MATURITY", "N/A")
        self.next_call_date = bbg_data.get("NXT_CALL_DT", "NC") # NC = Non-Callable
        self.is_callable = self.next_call_date != "NC"
        
        # Risk (Duration & Convexity)
        self.effective_duration = get_valid_float(["DUR_ADJ_MID"])
        self.convexity = get_valid_float(["CNVX_MID"])
        
        # Credit Ratings (Prioritizes Underlying ratings, then Enhanced/Insured, ignores N.S.)
        self.composite_rating = get_valid_rating(["BB_COMPOSITE"])
        self.sp_rating = get_valid_rating(["RTG_SP_UNDERLYING", "RTG_SP_ENHANCED", "RTG_SP_INSURED", "RTG_SP_MUNI_LONG_TERM", "RTG_SP"])
        self.moodys_rating = get_valid_rating(["RTG_MOODY_UNDERLYING", "RTG_MOODY_ENHANCED", "RTG_MOODY_INSURED", "RTG_MDY_MUNI_LONG_TERM", "RTG_MOODY"])
        self.fitch_rating = get_valid_rating(["RTG_FITCH_ENHANCED", "RTG_FITCH_INSURED", "RTG_FITCH"])
        
        # Sector, State, and Tax Status
        self.industry_sector = bbg_data.get("INDUSTRY_SECTOR", "Unknown")
        self.state_code = bbg_data.get("STATE_CODE", "Unknown")
        self.tax_provision = bbg_data.get("MUNI_TAX_PROV", "Unknown")
        
        # ---------------------------------------------------------
        # Computed Metrics (Calculated later by Python analytics)
        # ---------------------------------------------------------
        self.market_value = (self.face_value / 100) * self.current_price
        self.annual_income = self.face_value * (self.coupon / 100)
        self.portfolio_weight = 0.0
        self.taxable_equivalent_yield = 0.0

    def __repr__(self):
        return f"<Bond {self.cusip} | {self.security_name} | YTW: {self.yield_to_worst}%>"
