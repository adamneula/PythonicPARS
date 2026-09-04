import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from mappings import MUNI_ISSUE_MAP, CORE_INCOME_BUCKETS, CONSOLIDATED_RATINGS_ORDER

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
    generate_presentation("Robert Hunt", "Chris Wilbricht")
