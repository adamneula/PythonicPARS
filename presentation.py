import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.util import Inches, Pt
from mappings import MUNI_ISSUE_MAP, CORE_INCOME_BUCKETS, CONSOLIDATED_RATINGS_ORDER, MOODYS_TO_SP, RATING_SCORES, SCORE_TO_RATING

# --- HELPER FUNCTIONS ---

def get_shape(slide, shape_name):
    """Searches ONLY the current slide for a shape by name."""
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    return None

def update_text_preserve_format(shape, new_text):
    """
    Overwrites text in a shape while preserving the exact font, size, 
    color, and bold/italic styling set in PowerPoint.
    """
    if not shape.has_text_frame:
        return
    
    # Text in PowerPoint is stored in Paragraphs, which are made of "Runs".
    # The formatting (color, size, bold) is attached to the Run.
    p = shape.text_frame.paragraphs[0]
    
    if p.runs:
        # Overwrite the text of the very first run (keeping its formatting)
        p.runs[0].text = new_text
        
        # Erase the text of any other runs so we don't get leftover letters
        for run in p.runs[1:]:
            run.text = ""
    else:
        # Fallback if the box was completely empty
        shape.text = new_text

def synthesize_rating(row):
    comp = str(row.get('composite_rating', 'NR')).strip()
    if comp != 'NR' and comp != 'nan': return comp
        
    # Fallbacks if composite is NR
    sp = str(row.get('sp_rating', 'NR')).strip()
    if sp != 'NR' and sp != 'nan': return sp
    
    mdy = str(row.get('moodys_rating', 'NR')).strip()
    if mdy != 'NR' and mdy != 'nan': return MOODYS_TO_SP.get(mdy, mdy)
        
    fitch = str(row.get('fitch_rating', 'NR')).strip()
    if fitch != 'NR' and fitch != 'nan': return fitch
    
    return 'NR'

def generate_presentation(advisor_name, client_name="Valued Client", clientTaxRate=0.408):
    # 1. Load & Clean Data
    df = pd.read_excel("Portfolio_Output_20260915_093754.xlsx")
    df['muni_issue_type'] = df['muni_issue_type'].replace(MUNI_ISSUE_MAP)
    
    # 2. Load Template
    prs = Presentation("WorkingTemplate.pptx")
    slides = iter(prs.slides)

    # ========================================================
    # SLIDE 1: Title & Overview
    # ========================================================
    print("Populating Slide 1...")
    current_slide = next(slides)
    client_box = get_shape(current_slide, "ClientName")
    advisor_box = get_shape(current_slide, "AdvisorName")
    money_box = get_shape(current_slide, "ValueBox")
    if client_box:
        update_text_preserve_format(client_box, f"Prepared for: {client_name}")
    if advisor_box:
        update_text_preserve_format(advisor_box, f"Presented by: {advisor_name}")
    if money_box:
        total_value = df['market_value'].sum()
        update_text_preserve_format(money_box, f"${total_value:,.0f}")

    # ========================================================
    # SLIDE 2: Income Projections
    # ========================================================
    print("Populating Slide 2...")
    current_slide = next(slides)
    #TODO: Implement AI insight here appropriately. For now, just a placeholder.

    # ========================================================
    # SLIDE 3: Structure Analysis
    # ========================================================
    print("Populating Slide 3...")
    current_slide = next(slides)
    structure_counts = df.groupby('muni_issue_type')['market_value'].sum()

    structure_chart_data = CategoryChartData()
    structure_chart_data.categories = [str(label).title() for label in structure_counts.index]
    structure_chart_data.add_series('Bond Types', structure_counts.values.tolist())

    structure_chart = get_shape(current_slide, "ClientStructureAnalysis")
    if structure_chart and structure_chart.has_chart:
        structure_chart.chart.replace_data(structure_chart_data)
    #^PASSED VALIDATION^
    
    df['muni_source'] = df['muni_source'].str.title()
    income_counts = df.groupby('muni_source')['market_value'].sum()
    income_chart_data = CategoryChartData()
    income_chart_data.categories = income_counts.index.tolist()
    income_chart_data.add_series('Income Sources', income_counts.values.tolist())

    income_chart = get_shape(current_slide, "ClientIncomeSource")
    if income_chart and income_chart.has_chart:
        income_chart.chart.replace_data(income_chart_data)

    # ========================================================
    # SLIDE 4: Structure Recommendation
    # ========================================================
    print("Populating Slide 4...")
    current_slide = next(slides)
    structure_counts = df.groupby('muni_issue_type')['market_value'].sum()

    structure_chart_data = CategoryChartData()
    structure_chart_data.categories = [str(label).title() for label in structure_counts.index]
    structure_chart_data.add_series('Bond Types', structure_counts.values.tolist())

    structure_chart = get_shape(current_slide, "ClientStructure")
    if structure_chart and structure_chart.has_chart:
        structure_chart.chart.replace_data(structure_chart_data)

    # ========================================================
    # SLIDE 5: Duration & Maturity
    # ========================================================
    print("Populating Slide 5...")
    current_slide = next(slides)
    
    # 1. Calculate Years to Maturity
    # Convert string dates to actual datetime objects, then calculate difference from today
    df['maturity_date'] = pd.to_datetime(df['maturity_date'])
    today = pd.Timestamp.today()
    df['years_to_maturity'] = (df['maturity_date'] - today).dt.days / 365.25
    
    # 2. Bucket the years into the requested ranges
    def bucket_maturity(years):
        if pd.isna(years): return "Unknown"
        if years < 1: return "<1 Year"
        elif years < 3: return "1-3 Years"
        elif years < 5: return "3-5 Years"
        elif years < 7: return "5-7 Years"
        elif years < 10: return "7-10 Years"
        elif years < 15: return "10-15 Years"
        else: return "15+ Years"

    df['maturity_bucket'] = df['years_to_maturity'].apply(bucket_maturity)
    
    # 3. Group and sum market value
    maturity_counts = df.groupby('maturity_bucket')['market_value'].sum()
    
    # 4. Sort chronologically using a hardcoded ordered list
    ordered_maturity = ["<1 Year", "1-3 Years", "3-5 Years", "5-7 Years", "7-10 Years", "10-15 Years", "15+ Years"]
    ordered_labels_mat = [r for r in ordered_maturity if r in maturity_counts.index]
    maturity_counts = maturity_counts.reindex(ordered_labels_mat)
    
    # 5. Build and Inject Chart Data
    chart_data_mat = CategoryChartData()
    chart_data_mat.categories = maturity_counts.index.tolist()
    chart_data_mat.add_series('Maturity Profile', maturity_counts.values.tolist())
    
    mat_chart_shape = get_shape(current_slide, "ClientMaturityProfile")
    if mat_chart_shape and mat_chart_shape.has_chart:
        mat_chart_shape.chart.replace_data(chart_data_mat)
        print("Updated Maturity Profile Chart!")

    # ----------------------------------------------------
    # CALL RATE CHART (Callable vs Non-Callable)
    # ----------------------------------------------------
    # Clean the is_callable column to ensure clean "Callable" and "Non-Callable" labels
    def clean_callable(val):
        val_str = str(val).strip().upper()
        if val_str in ["TRUE", "Y", "YES", "1", "1.0"]:
            return "Callable"
        return "Non-Callable"
        
    df['clean_callable'] = df['is_callable'].apply(clean_callable)
    callable_counts = df.groupby('clean_callable')['market_value'].sum()
    
    chart_data_call = CategoryChartData()
    chart_data_call.categories = callable_counts.index.tolist()
    chart_data_call.add_series('Callability', callable_counts.values.tolist())
    
    call_chart_shape = get_shape(current_slide, "ClientCallRate")
    if call_chart_shape and call_chart_shape.has_chart:
        call_chart_shape.chart.replace_data(chart_data_call)
        print("Updated Call Rate Chart!")

    # ========================================================
    # SLIDE 6: Maturity Recommendation
    # ========================================================
    print("Populating Slide 6...")
    current_slide = next(slides)
    
    # 1. Calculate Years to Effective Maturity
    # Using the new 'effective_maturity_date' column
    df['effective_maturity_date'] = pd.to_datetime(df.get('effective_maturity_date'), errors='coerce')
    df['years_to_eff_maturity'] = (df['effective_maturity_date'] - today).dt.days / 365.25
    
    # 2. Bucket the years (reusing the bucket_maturity function defined in Slide 5)
    df['eff_maturity_bucket'] = df['years_to_eff_maturity'].apply(bucket_maturity)
    
    # 3. Group and sum market value
    eff_maturity_counts = df.groupby('eff_maturity_bucket')['market_value'].sum()
    
    # 4. Sort chronologically (reusing the ordered_maturity list from Slide 5)
    ordered_labels_eff = [r for r in ordered_maturity if r in eff_maturity_counts.index]
    eff_maturity_counts = eff_maturity_counts.reindex(ordered_labels_eff)
    
    # 5. Build and Inject Chart Data
    chart_data_eff = CategoryChartData()
    chart_data_eff.categories = eff_maturity_counts.index.tolist()
    chart_data_eff.add_series('Effective Maturity Profile', eff_maturity_counts.values.tolist())
    
    eff_chart_shape = get_shape(current_slide, "ClientOAM")
    if eff_chart_shape and eff_chart_shape.has_chart:
        eff_chart_shape.chart.replace_data(chart_data_eff)
        print("Updated ClientOAM Chart!")

    # ========================================================
    # SLIDE 7: Coupon Analysis
    # ========================================================
    print("Populating Slide 7...")
    current_slide = next(slides)
    
    # 1. Bucket the coupon rates
    def bucket_coupon(rate):
        if pd.isna(rate): return "Unknown"
        
        if rate < 1: return "<1%"
        elif rate <= 2: return "1%-2%"
        elif rate <= 3: return "2%-3%"
        elif rate <= 4: return "3%-4%"
        elif rate <= 5: return "4%-5%"
        else: return ">5%"
        
    # Convert to numeric just in case it came in as a string
    df['coupon'] = pd.to_numeric(df.get('coupon'), errors='coerce')
    
    # (Safety Check) Bloomberg sometimes exports 4% as 0.04, sometimes as 4.00.
    # If the max coupon in the portfolio is under 0.20, it's definitely in decimal format.
    if df['coupon'].max() < 0.20 and df['coupon'].max() > 0:
        df['coupon_clean'] = df['coupon'] * 100
    else:
        df['coupon_clean'] = df['coupon']
        
    df['coupon_bucket'] = df['coupon_clean'].apply(bucket_coupon)
    
    # 2. Group and sum market value
    coupon_counts = df.groupby('coupon_bucket')['market_value'].sum()
    
    # 3. Sort so the slices stay in numerical order on the pie chart
    ordered_coupons = ["<1%", "1%-2%", "2%-3%", "3%-4%", "4%-5%", ">5%"]
    ordered_labels_cpn = [r for r in ordered_coupons if r in coupon_counts.index]
    coupon_counts = coupon_counts.reindex(ordered_labels_cpn)
    
    # 4. Build and Inject Chart Data
    chart_data_cpn = CategoryChartData()
    chart_data_cpn.categories = coupon_counts.index.tolist()
    chart_data_cpn.add_series('Coupon Rate', coupon_counts.values.tolist())
    
    cpn_chart_shape = get_shape(current_slide, "ClientCouponRate")
    if cpn_chart_shape and cpn_chart_shape.has_chart:
        cpn_chart_shape.chart.replace_data(chart_data_cpn)
        print("Updated ClientCouponRate Chart!")

    # ========================================================
    # SLIDE 8: Credit Quality / Ratings
    # ========================================================
    print("Populating Slide 8...")
    current_slide = next(slides)

    df['syn_rating'] = df.apply(synthesize_rating, axis=1)
    df['clean_rating'] = df['syn_rating'].str.replace('+', '').str.replace('-', '')
    ratings_counts = df.groupby('clean_rating')['market_value'].sum()
    
    ordered_labels = [r for r in CONSOLIDATED_RATINGS_ORDER if r in ratings_counts.index]
    ratings_counts = ratings_counts.reindex(ordered_labels)
    
    chart_data_ratings = CategoryChartData()
    chart_data_ratings.categories = ratings_counts.index.tolist()
    chart_data_ratings.add_series('Credit Quality', ratings_counts.values.tolist())
    
    ratings_chart_shape = get_shape(current_slide, "ClientQualityBreakdown")
    if ratings_chart_shape and ratings_chart_shape.has_chart:
        ratings_chart_shape.chart.replace_data(chart_data_ratings)
        print("Updated Ratings Chart!")

    #^ Going off industry standard, composite rating or only one (composite exists if 2 or more are filled in)

    # ========================================================
    # SLIDE 9: State Breakdown Recommendation
    # ========================================================
    print("Populating Slide 9...")
    current_slide = next(slides)
    
    # Group market value by State Code and sort largest to smallest
    if 'state_code' in df.columns:
        state_counts = df.groupby('state_code')['market_value'].sum().sort_values(ascending=False)
        
        chart_data_state = CategoryChartData()
        chart_data_state.categories = state_counts.index.tolist()
        chart_data_state.add_series('State Breakdown', state_counts.values.tolist())
        
        state_chart_shape = get_shape(current_slide, "ClientStateBreakdown")
        if state_chart_shape and state_chart_shape.has_chart:
            state_chart_shape.chart.replace_data(chart_data_state)
            print("Updated ClientStateBreakdown Chart!")

    # ========================================================
    # SLIDE 10: Portfolio Characteristics
    # ========================================================
    print("Populating Slide 10...")
    current_slide = next(slides)
    
    # Helper for fast weighted averages
    def weighted_avg(col_name):
        valid = df[col_name].notna() & df['market_value'].notna()
        if not valid.any(): return 0.0
        return (df.loc[valid, col_name] * df.loc[valid, 'market_value']).sum() / df.loc[valid, 'market_value'].sum()

    # Numeric Averages (Fully Weighted)
    avg_ytw = weighted_avg('yield_to_worst')
    avg_ytm = weighted_avg('yield_to_maturity')
    avg_coupon = weighted_avg('coupon_clean')
    avg_mat = weighted_avg('years_to_maturity')
    avg_effMat = weighted_avg('years_to_eff_maturity')
    avg_effDur = weighted_avg('effective_duration')
    avg_conv = weighted_avg('convexity')
    
    # Text Box Updates
    genterYTW = get_shape(current_slide, "GenterYTW")
    genterTEY = get_shape(current_slide, "GenterTEY")
    if genterYTW and genterTEY:
        # Assuming GenterYTW text is manually typed in the template and you just calculate the TEY from it
        update_text_preserve_format(genterTEY, f"{float(genterYTW.text.strip('%'))/(1-clientTaxRate):.2f}%")

    coupon = get_shape(current_slide, "ClientAVGCoupon")
    if coupon: update_text_preserve_format(coupon, f"{avg_coupon:.2f}%")

    ytm = get_shape(current_slide, "ClientYTM")
    if ytm: update_text_preserve_format(ytm, f"{avg_ytm:.2f}%")

    ytw = get_shape(current_slide, "ClientYTW")
    if ytw: update_text_preserve_format(ytw, f"{avg_ytw:.2f}%")

    tey = get_shape(current_slide, "ClientTEY")
    if tey: update_text_preserve_format(tey, f"{avg_ytw/(1-clientTaxRate):.2f}%")

    statMat = get_shape(current_slide, "ClientStatMat")
    if statMat: update_text_preserve_format(statMat, f"{avg_mat:.2f} Years")

    effMat = get_shape(current_slide, "ClientEffMat")
    if effMat: update_text_preserve_format(effMat, f"{avg_effMat:.2f} Years")
    
    effDur = get_shape(current_slide, "ClientEffDur")
    if effDur: update_text_preserve_format(effDur, f"{avg_effDur:.2f} Years")

    conv = get_shape(current_slide, "ClientConv")
    if conv: update_text_preserve_format(conv, f"{avg_conv:.2f}")

    # Credit Quality Average (Reverted to Weighted)
    quality = get_shape(current_slide, "ClientQual")
    if quality:
        df_rated = df[df['syn_rating'].isin(RATING_SCORES.keys())].copy()
        df_rated['rating_score'] = df_rated['syn_rating'].map(RATING_SCORES)
        total_rated_value = df_rated['market_value'].sum()
        
        if total_rated_value > 0:
            avg_score_raw = (df_rated['rating_score'] * df_rated['market_value']).sum() / total_rated_value
            avg_quality = SCORE_TO_RATING[round(avg_score_raw)]
        else:
            avg_quality = "NR"
            
        update_text_preserve_format(quality, f"{avg_quality}")
        
    # ========================================================
    # SLIDE 11+: Holdings Detail (Auto-Paginated Table)
    # ========================================================
    from pptx.dml.color import RGBColor
    
    print("Populating Holdings Detail Slides...")
    # Search ALL Slide Masters for the specific layout name
    holdings_layout = None
    for master in prs.slide_masters:
        for layout in master.slide_layouts:
            if layout.name == "HoldingsDetail":
                holdings_layout = layout
                break
        if holdings_layout:
            break
            
    # Absolute fallback if still not found
    if holdings_layout is None:
        holdings_layout = prs.slide_masters[-1].slide_layouts[-1]
        
    df_sorted = df.sort_values(by='maturity_date', na_position='last')
    total_mv = df['market_value'].sum()
    chunk_size = 18
    
    headers = [
        "Face Value", "CUSIP/Ticker", "Name", "Price", "Market Value", "Coupon",
        "Stated Maturity", "Next Call Date", "YTW", "Rating", "Modified Duration", 
        "Effective Duration", "Convexity", "% of Portfolio"
    ]
    
    for i in range(0, len(df_sorted), chunk_size):
        chunk = df_sorted.iloc[i:i+chunk_size]
        new_slide = prs.slides.add_slide(holdings_layout)
        
        # Populate the title placeholder so it becomes solid text that prints
        if new_slide.shapes.title:
            new_slide.shapes.title.text = "Client Holdings"
        
        rows = len(chunk) + 1
        cols = 14
        
        # Add the table (Moved Top from 1.2 inches up to 0.9 inches, Height increased to 6.1 to stretch down one more row)
        table_shape = new_slide.shapes.add_table(rows, cols, Inches(0.2), Inches(0.9), Inches(12.9), Inches(6.1))
        table = table_shape.table
        
        # Explicitly define narrow column widths for data cells and give the rest to the Name column
        col_widths = [
            Inches(0.65), # Face Value (Reduced to ~80% width)
            Inches(1.0),  # CUSIP
            Inches(3.3),  # Name (Absorbed extra space)
            Inches(0.6),  # Price
            Inches(0.65), # Market Value 
            Inches(0.7),  # Coupon (Widened to fit 'n')
            Inches(0.9),  # Stated Maturity
            Inches(0.9),  # Next Call Date
            Inches(0.6),  # YTW
            Inches(0.6),  # Rating
            Inches(0.7),  # Modified Duration 
            Inches(0.8),  # Effective Duration 
            Inches(0.8),  # Convexity
            Inches(0.7)   # % of Portfolio 
        ]
        for col_idx, width in enumerate(col_widths):
            table.columns[col_idx].width = width
        
        # Format Headers
        for col_idx, header in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.text = header
            
            # Set background color to #59002E
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0x59, 0x00, 0x2E)
            
            for p in cell.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)
                    r.font.bold = True
                    r.font.color.rgb = RGBColor(255, 255, 255) # White text for contrast
                    
        # Populate Data
        for row_idx, (_, bond) in enumerate(chunk.iterrows(), start=1):
            face = f"${bond.get('face_value', 0):,.0f}"
            cusip = str(bond.get('cusip', ''))
            name = str(bond.get('security_name', ''))[:28] # Truncate long names slightly
            price = f"${bond.get('current_price', 0):.2f}"
            mv = f"${bond.get('market_value', 0):,.0f}"
            cpn = f"{bond.get('coupon_clean', 0):.3f}%"
            
            mat_dt = bond.get('maturity_date')
            mat = mat_dt.strftime('%m/%d/%Y') if pd.notna(mat_dt) else "N/A"
            
            call_dt = bond.get('next_call_date')
            if pd.notna(call_dt) and call_dt != "NC":
                call = call_dt if isinstance(call_dt, str) else call_dt.strftime('%m/%d/%Y')
            else:
                call = "NC"
                
            ytw = f"{bond.get('yield_to_worst', 0):.2f}%"
            rating = str(bond.get('syn_rating', 'NR'))
            mod_dur = f"{bond.get('modified_duration', 0):.2f}" if pd.notna(bond.get('modified_duration')) else "N/A"
            eff_dur = f"{bond.get('effective_duration', 0):.2f}" if pd.notna(bond.get('effective_duration')) else "N/A"
            conv = f"{bond.get('convexity', 0):.2f}" if pd.notna(bond.get('convexity')) else "N/A"
            pct = f"{(bond.get('market_value', 0) / total_mv) * 100:.1f}%" if total_mv > 0 else "0.0%"
            
            row_data = [face, cusip, name, price, mv, cpn, mat, call, ytw, rating, mod_dur, eff_dur, conv, pct]
            
            for col_idx, val in enumerate(row_data):
                cell = table.cell(row_idx, col_idx)
                cell.text = str(val)
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.size = Pt(8) # 8pt font to fit 14 columns

    # ========================================================
    # SAVE PRESENTATION
    # ========================================================
    output_name = "PAR_Report_Output.pptx"
    prs.save(output_name)
    print(f"\nDone! Presentation saved to {output_name}")

if __name__ == "__main__":
    generate_presentation("Adam Neulander", "Dylan Castillo")
