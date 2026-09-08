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
    # SLIDE 5: Maturity Analysis
    # ========================================================
    print("Populating Slide 5...")
    current_slide = next(slides)

    # ========================================================
    # SLIDE 6: Maturity Recommendation
    # ========================================================
    print("Populating Slide 6...")
    current_slide = next(slides)

    # ========================================================
    # SLIDE 7: Coupon Analysis
    # ========================================================
    print("Populating Slide 7...")
    current_slide = next(slides)

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
    generate_presentation("Robert Hunt", "Chris Wilbricht")
