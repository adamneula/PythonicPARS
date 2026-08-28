import blpapi

class bloomberg:
    def __init__(self, host: str = "localhost", port: int = 8194):
        self.host = host
        self.port = port
        self.session = None

    def __enter__(self):
        # MOCKED out blpapi init to avoid connection/limit errors
        '''
        options = blpapi.SessionOptions()
        options.setServerHost(self.host)
        options.setServerPort(self.port)
        self.session = blpapi.Session(options)
        
        if not self.session.start() or not self.session.openService("//blp/refdata"):
            raise RuntimeError("Failed to start BLPAPI session or open refdata service.")
        '''
        print("DEBUG: Using MOCKED Bloomberg connection.")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        '''
        if self.session is not None:
            self.session.stop()
        '''
        pass

    def get_bond_metadata(self, cusips: list[str]) -> list[dict]:
        if not cusips:
            print("Error: The provided CUSIP list is empty. No request was sent to Bloomberg.")
            return []
            
        print(f"DEBUG: Generating MOCKED metadata for {len(cusips)} CUSIPs...")
        bond_data = []

        # Filler info on 3 muni bonds
        muni_templates = [
            {
                "SECURITY_NAME": "NY STATE URBAN DEV 5.00%",
                "CPN": 5.000,
                "MATURITY": "2035-03-15",
                "NXT_CALL_DT": "2025-03-15",
                "YLD_YTW_MID": 3.45,
                "YLD_YTM_MID": 3.75,
                "DUR_ADJ_MID": 7.2,
                "CNVX_MID": 0.65,
                "RTG_SP": "AA+",
                "RTG_MOODY": "Aa1",
                "INDUSTRY_SECTOR": "Government",
                "STATE_CODE": "NY",
                "MUNI_TAX_PROV": "Exempt"
            },
            {
                "SECURITY_NAME": "CALIFORNIA GOVT 4.50%",
                "CPN": 4.500,
                "MATURITY": "2030-10-01",
                "NXT_CALL_DT": "2026-10-01",
                "YLD_YTW_MID": 3.20,
                "YLD_YTM_MID": 3.40,
                "DUR_ADJ_MID": 5.5,
                "CNVX_MID": 0.40,
                "RTG_SP": "AA-",
                "RTG_MOODY": "Aa2",
                "INDUSTRY_SECTOR": "Government",
                "STATE_CODE": "CA",
                "MUNI_TAX_PROV": "Exempt"
            },
            {
                "SECURITY_NAME": "TEXAS WATER DEV BD 4.00%",
                "CPN": 4.000,
                "MATURITY": "2040-08-01",
                "NXT_CALL_DT": "2030-08-01",
                "YLD_YTW_MID": 3.80,
                "YLD_YTM_MID": 3.95,
                "DUR_ADJ_MID": 9.1,
                "CNVX_MID": 0.85,
                "RTG_SP": "AAA",
                "RTG_MOODY": "Aaa",
                "INDUSTRY_SECTOR": "Government",
                "STATE_CODE": "TX",
                "MUNI_TAX_PROV": "Exempt"
            }
        ]

        for i, cusip in enumerate(cusips):
            template = muni_templates[i % len(muni_templates)].copy()
            template["cusip"] = cusip
            bond_data.append(template)
            
        return bond_data

    # ---------------------------------------------------------
    # ORIGINAL BLOOMBERG API LOGIC PRESERVED FOR LATER USE
    # ---------------------------------------------------------
    '''
    def get_bond_metadata_original(self, cusips: list[str]) -> list[dict]:
        if not cusips:
            print("Error: The provided CUSIP list is empty. No request was sent to Bloomberg.")
            return []

        ref_data_service = self.session.getService("//blp/refdata")
        request = ref_data_service.createRequest("ReferenceDataRequest")
        
        # Append all CUSIPs in a single batch
        for cusip in cusips:
            request.append("securities", f"/cusip/{cusip} Muni")
            
        fields = [
            "SECURITY_NAME", "CPN", "MATURITY", "NXT_CALL_DT",
            "YLD_YTW_MID", "YLD_YTM_MID", "DUR_ADJ_MID", "CNVX_MID", 
            "RTG_SP", "RTG_MOODY", "INDUSTRY_SECTOR", "STATE_CODE", "MUNI_TAX_PROV"
        ]
        
        for field in fields:
            request.append("fields", field)
            
        self.session.sendRequest(request)
        bond_data = []
        
        # Parse the asynchronous event loop
        while True:
            event = self.session.nextEvent(500)
            
            # Catch request-level errors (e.g. malformed requests) which would otherwise cause an infinite loop
            if event.eventType() == blpapi.Event.REQUEST_STATUS:
                for msg in event:
                    print(f"Bloomberg Request Error: {msg}")
                break

            for msg in event:
                # Catch overarching response errors (e.g. limits exceeded or workflow review needed)
                if msg.hasElement("responseError"):
                    err = msg.getElement("responseError")
                    err_msg = err.getElementAsString("message")
                    err_cat = err.getElementAsString("category")
                    print(f"\\nCRITICAL BLOOMBERG ERROR ({err_cat}): {err_msg}")
                    print("Bloomberg is actively blocking this request, likely due to data limits or your account needing permissions/workflow review. No data could be retrieved.\\n")
                    continue

                if msg.hasElement("securityData"):
                    security_data_array = msg.getElement("securityData")
                    
                    for i in range(security_data_array.numValues()):
                        security = security_data_array.getValueAsElement(i)
                        ticker = security.getElementAsString("security")
                        
                        # Check for security-level errors (e.g. invalid CUSIPs)
                        if security.hasElement("securityError"):
                            sec_err = security.getElement("securityError")
                            err_msg = sec_err.getElementAsString("message")
                            print(f"Bloomberg Security Error for {ticker}: {err_msg}")
                            continue

                        raw_cusip = ticker.split("/")[2].split()[0]
                        bond_attrs = {"cusip": raw_cusip}
                        
                        field_data = security.getElement("fieldData")
                        for field in field_data.elements():
                            bond_attrs[str(field.name())] = field.getValue()
                            
                        bond_data.append(bond_attrs)
            
            if event.eventType() == blpapi.Event.RESPONSE:
                break
                
        return bond_data
    '''