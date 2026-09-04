import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from mappings import MUNI_ISSUE_MAP, CORE_INCOME_BUCKETS, CONSOLIDATED_RATINGS_ORDER

# --- HELPER FUNCTION ---
def get_shape(slide, shape_name):
    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape
    
    print(f"Warning: Could not find shape '{shape_name}' on this slide.")
    return None

def generate_presentation():
    # 1. Load & Clean Data
    df = pd.read_excel("Par_Portfolio_Output.xlsx")
    
    # We are using the exact column names found in your Excel file!
    # (muni_issue_type, muni_purpose, market_value, etc.)
    
    df['muni_issue_type'] = df['muni_issue_type'].replace(MUNI_ISSUE_MAP)
    
    # 2. Load Template
    prs = Presentation("WorkingTemplate.pptx")
    slides = iter(prs.slides)

    # ========================================================
    # SLIDE 1: Title & Overview
    # ========================================================
    print("Populating Slide 1...")
    current_slide = next(slides)
    name_box = get_shape(current_slide, "ClientName")
    if name_box:
        name_box.text = "Adam Neulander"
    amount_box = get_shape(current_slide, "Strategy + Portfolio Amt")
    if amount_box:
        amount_box.text = "$1,000,000"
    # Example:
    # title_box = get_shape(current_slide, "Client_Name_Box")

    # ========================================================
    # SLIDE 2: Income Projections
    # ========================================================
    print("Populating Slide 2...")
    current_slide = next(slides)

    # ========================================================
    # SLIDE 3: Structure Analysis
    # ========================================================
    print("Populating Slide 3...")
    current_slide = next(slides)
    
    # --- PANDAS MATH ---
    # Group by the lowercase 'muni_issue_type' and sum 'market_value'
    issue_type_counts = df.groupby('muni_issue_type')['market_value'].sum()
    
    # Build chart data
    chart_data_type = CategoryChartData()
    chart_data_type.categories = issue_type_counts.index.tolist()
    chart_data_type.add_series('Bond Types', issue_type_counts.values.tolist())
    
    # Grab the exact chart on THIS slide and inject!
    bond_chart_shape = get_shape(current_slide, "Muni_Issue_Chart") 
    if bond_chart_shape and bond_chart_shape.has_chart:
        bond_chart_shape.chart.replace_data(chart_data_type)

    # ========================================================
    # SLIDE 4: Credit Quality / Ratings
    # ========================================================
    print("Populating Slide 4...")
    current_slide = next(slides)

    # ========================================================
    # SLIDE 5: Duration & Maturity
    # ========================================================
    print("Populating Slide 5...")
    current_slide = next(slides)
    
    # ========================================================
    # SAVE PRESENTATION
    # ========================================================
    output_name = "PAR_Report_Output.pptx"
    prs.save(output_name)
    print(f"\nDone! Presentation saved to {output_name}")

if __name__ == "__main__":
    generate_presentation()
