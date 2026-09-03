import blpapi

class BloombergClient:
    """
    A standalone client to handle batch requests to the Bloomberg Terminal.
    Includes a 'use_mock' flag to test the app without hitting Bloomberg limits.
    """
    def __init__(self, host: str = "localhost", port: int = 8194, use_mock: bool = True):
        self.host = host
        self.port = port
        self.use_mock = use_mock
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
        Takes a list of CUSIPs and returns a dictionary mapped by CUSIP.
        Example Return: {"123456789": {"YLD_YTW_MID": 4.5, "CPN": 3.0}}
        """
        if not cusips:
            return {}

        if self.use_mock:
            return self._get_mock_data(cusips)
        else:
            return self._get_live_data(cusips)

    def _get_mock_data(self, cusips: list[str]) -> dict:
        print(f"DEBUG: Generating mock Bloomberg data for {len(cusips)} CUSIPs.")
        results = {}
        
        # A few realistic mock templates to cycle through
        templates = [
            {
                "SECURITY_NAME": "NEW YORK ST URBAN DEV", "CPN": 5.000, 
                "PX_LAST": 105.20, "YLD_YTW_MID": 3.45, "YLD_YTM_MID": 3.75,
                "MATURITY": "2035-03-15", "NXT_CALL_DT": "2025-03-15",
                "DUR_ADJ_MID": 7.2, "CNVX_MID": 0.65, "RTG_SP": "AA+", 
                "RTG_MOODY": "Aa1", "INDUSTRY_SECTOR": "Government", 
                "STATE_CODE": "NY", "MUNI_TAX_PROV": "Exempt"
            },
            {
                "SECURITY_NAME": "CALIFORNIA GOVT", "CPN": 4.500, 
                "PX_LAST": 102.10, "YLD_YTW_MID": 3.20, "YLD_YTM_MID": 3.40,
                "MATURITY": "2030-10-01", "NXT_CALL_DT": "2026-10-01",
                "DUR_ADJ_MID": 5.5, "CNVX_MID": 0.40, "RTG_SP": "AA-", 
                "RTG_MOODY": "Aa2", "INDUSTRY_SECTOR": "Government", 
                "STATE_CODE": "CA", "MUNI_TAX_PROV": "Exempt"
            }
        ]

        for i, cusip in enumerate(cusips):
            # Assign a mock template to each CUSIP
            results[cusip] = templates[i % len(templates)].copy()
            
        return results

    def _get_live_data(self, cusips: list[str]) -> dict:
        ref_data_service = self.session.getService("//blp/refdata")
        request = ref_data_service.createRequest("ReferenceDataRequest")
        
        # 1. Batch append all CUSIPs
        for cusip in cusips:
            request.append("securities", f"/cusip/{cusip} Muni")
            
        # 2. Append required fields
        fields = [
            "SECURITY_NAME", "SECURITY_DES", "SHORT_NAME", "DSPLY_NAME", "ISSUER",
            "CPN", "PX_LAST", "MATURITY", "NXT_CALL_DT",
            "YLD_YTW", "MUNI_YLD_TO_WORST", "YAS_BOND_YLD", "YLD_YTW_MID", "YLD_YTW_BID", "YLD_TO_WORST", 
            "YLD_YTM_MID", "YLD_YTM_BID", "YLD_TO_MATURITY",
            "DUR_ADJ_MID", "CNVX_MID", 
            "BB_COMPOSITE", "RTG_SP_UNDERLYING", "RTG_MOODY_UNDERLYING", "RTG_SP_MUNI_LONG_TERM", "RTG_MDY_MUNI_LONG_TERM",
            "RTG_SP_INSURED", "RTG_MOODY_INSURED", "RTG_FITCH_INSURED",
            "RTG_SP_ENHANCED", "RTG_MOODY_ENHANCED", "RTG_FITCH_ENHANCED",
            "RTG_SP", "RTG_MOODY", "RTG_FITCH", "INDUSTRY_SECTOR", "STATE_CODE", "MUNI_TAX_PROV"
        ]
        for field in fields:
            request.append("fields", field)
            
        self.session.sendRequest(request)
        
        results = {}
        
        # 3. Parse asynchronous event loop
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
                        
                        raw_cusip = ticker.split("/")[2].split()[0]
                        bond_data = {}
                        
                        field_data = security.getElement("fieldData")
                        for field in field_data.elements():
                            bond_data[str(field.name())] = field.getValue()
                            
                        results[raw_cusip] = bond_data
            
            if event.eventType() == blpapi.Event.RESPONSE:
                break
                
        return results
