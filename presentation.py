import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from mappings import MUNI_ISSUE_MAP, CORE_INCOME_BUCKETS, CONSOLIDATED_RATINGS_ORDER, MOODYS_TO_SP

# --- HELPER FUNCTIONS ---

def get_shape(slide, shape_name):
    """Searches ONLY the current slide for a shape by name."""
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    print(f"Warning: Could not find shape '{shape_name}' on this slide.")
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

def generate_presentation(advisor_name, client_name="Valued Client"):
    # 1. Load & Clean Data
    df = pd.read_excel("Par_Portfolio_Output.xlsx")
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
        elif rate <= 5: return "4%-5%" # 5.0 flat now drops into 4%-5%
        else: return "5%+"
        
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
    ordered_coupons = ["<1%", "1%-2%", "2%-3%", "3%-4%", "4%-5%", "5%+"]
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

    # ========================================================
    # SAVE PRESENTATION
    # ========================================================
    output_name = "PAR_Report_Output.pptx"
    prs.save(output_name)
    print(f"\nDone! Presentation saved to {output_name}")

if __name__ == "__main__":
    generate_presentation("Adam Neulander", "Dylan Castillo")
