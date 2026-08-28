class bondObjects:
    def __init__(self, cusip: str, face_value: float, bbg_data: dict = None):
        self.cusip = cusip
        self.face_value = face_value
        
        bbg_data = bbg_data or {}
        # Map Bloomberg fields to the class attributes expected by the print statement
        self.security_name = bbg_data.get("SECURITY_NAME", "N/A")
        self.current_price = bbg_data.get("PX_LAST", "N/A") # Note: PX_LAST wasn't in original fields, defaulting to N/A
        self.maturity_date = bbg_data.get("MATURITY", "N/A")
        self.next_call_date = bbg_data.get("NXT_CALL_DT", "N/A")
        self.yield_to_worst = bbg_data.get("YLD_YTW_MID", "N/A")
        self.yield_to_maturity = bbg_data.get("YLD_YTM_MID", "N/A")
        self.duration = bbg_data.get("DUR_ADJ_MID", "N/A")
        self.convexity = bbg_data.get("CNVX_MID", "N/A")
        self.sp_rating = bbg_data.get("RTG_SP", "N/A")
        self.moodys_rating = bbg_data.get("RTG_MOODY", "N/A")
        self.industry_sector = bbg_data.get("INDUSTRY_SECTOR", "N/A")
        self.state_code = bbg_data.get("STATE_CODE", "N/A")
        self.municipal_tax_provision = bbg_data.get("MUNI_TAX_PROV", "N/A")

    def __hash__(self):
        return hash(self.cusip)

    def __eq__(self, other):
        if isinstance(other, bondObjects):
            return self.cusip == other.cusip
        return False

    def calculate_analytics(self, fedTaxRate: int = 0, stateTaxRate: int = 0) -> None:
        '''
        Calculates the analytics for the bond object based on the provided tax rates
        '''
        # Placeholder for actual analytics calculations
        # For example, you might calculate after-tax yield, duration adjustments, etc.
        pass