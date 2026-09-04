import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData

# (Optional) Import your clean mappings
from mappings import MUNI_ISSUE_MAP, CORE_INCOME_BUCKETS, CONSOLIDATED_RATINGS_ORDER

def generate_presentation(data_path, template_path, output_path):
    print(f"Loading data from {data_path}...")
    
    # 1. Load your Data into Pandas
    df = pd.read_excel(data_path)
    
    # --- PANDAS DATA CLEANING (Optional, do this here) ---
    # df['MUNI_ISSUE_TYP'] = df['MUNI_ISSUE_TYP'].replace(MUNI_ISSUE_MAP)
    # ... aggregate your dataframes here ...
    
    print(f"Loading template from {template_path}...")
    prs = Presentation(template_path)

    # 2. Loop through every slide in the presentation
    for slide_index, slide in enumerate(prs.slides):
        slide_num = slide_index + 1
        print(f"\n======================================")
        print(f" Slide {slide_num}")
        print(f"======================================")

        # 3. Loop through every element (shape) on the current slide
        for shape in slide.shapes:
            print(f"  -> Found Shape: '{shape.name}'")

            # ---------------------------------------------------------
            # A. IF THE SHAPE IS A CHART (Pie, Bar, Line)
            # ---------------------------------------------------------
            if shape.has_chart:
                print(f"     [Chart Type] Ready to replace data.")
                
                # Example Boilerplate for updating a chart:
                '''
                if slide_num == 3 and shape.name == "Bond_Type_Chart":
                    # 1. Prepare the new data
                    chart_data = CategoryChartData()
                    chart_data.categories = ['GO', 'Revenue']
                    chart_data.add_series('Series 1', [50, 50])
                    
                    # 2. Inject it into the chart
                    shape.chart.replace_data(chart_data)
                '''

            # ---------------------------------------------------------
            # B. IF THE SHAPE IS A TEXT BOX
            # ---------------------------------------------------------
            elif shape.has_text_frame:
                # print(f"     [Text] Current text: '{shape.text[:30]}...'")
                
                # Example Boilerplate for updating text:
                '''
                if shape.name == "Client_Name_Box":
                    shape.text = "Genter Capital Management"
                '''
                pass

            # ---------------------------------------------------------
            # C. IF THE SHAPE IS A TABLE
            # ---------------------------------------------------------
            elif shape.has_table:
                print(f"     [Table Type] Ready to edit cells.")
                
                # Example Boilerplate for updating a table:
                '''
                if slide_num == 4 and shape.name == "Holdings_Table":
                    table = shape.table
                    # Overwrite the first cell in the first row
                    table.cell(0, 0).text = "New Value"
                '''

            # ---------------------------------------------------------
            # D. OTHER SHAPES (Pictures, Lines, Groups)
            # ---------------------------------------------------------
            else:
                # E.g., shape.shape_type == MSO_SHAPE_TYPE.PICTURE
                pass

    # 4. Save the populated presentation
    prs.save(output_path)
    print(f"\nDone! Presentation saved to {output_path}")


if __name__ == "__main__":
    # Define your file paths
    DATA_FILE = "Par_Portfolio_Output.xlsx"
    TEMPLATE_FILE = "WorkingTemplate.pptx"
    OUTPUT_FILE = "PAR_Report_Output.pptx"
    
    import os
    if os.path.exists(TEMPLATE_FILE):
        generate_presentation(DATA_FILE, TEMPLATE_FILE, OUTPUT_FILE)
    else:
        print(f"Error: {TEMPLATE_FILE} not found. Please ensure it is in the current directory.")
