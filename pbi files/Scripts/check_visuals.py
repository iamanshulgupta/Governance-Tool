import json
import os

# 1. Setup: Find the report.json file automatically
current_folder = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(current_folder, "report.json")

print("==================================================")
print("🔍 POWER BI GOVERNANCE AUDIT: VISUAL CHECKLIST")
print("==================================================\n")

try:
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        report_data = json.load(f)
        pages = report_data.get('sections', [])
        
        # 2. Loop through every page in the dashboard
        for page in pages:
            page_name = page.get('displayName', 'Unknown Page')
            visuals = page.get('visualContainers', [])
            
            print(f"📄 AUDITING PAGE: '{page_name}' ({len(visuals)} visuals)")
            print("-" * 50)
            
            # Trackers for this specific page
            found_logo = False
            found_top_nav = False
            found_slicers = False
            slicers_are_consistent = True
            visuals_missing_titles = 0
            
            # 3. Loop through every visual on the current page
            for visual in visuals:
                
                # Get coordinates (Where is the visual on the page?)
                # Power BI coordinates start at 0,0 in the top-left corner
                y_pos = visual.get('y', 999) 
                x_pos = visual.get('x', 999)
                
                # Extract the hidden config data
                config_string = visual.get('config', '{}')
                try:
                    config_data = json.loads(config_string)
                    visual_type = config_data.get('singleVisual', {}).get('visualType', 'Unknown')
                    
                    # --- RULE 1: THE LOGO ---
                    # Is it an image in the top-left corner? (X and Y less than 100 pixels)
                    if visual_type == 'image' and x_pos < 100 and y_pos < 100:
                        found_logo = True
                        
                    # --- RULE 2: TOP NAVIGATION ---
                    # Are there buttons or shapes placed at the very top of the page?
                    if visual_type in ['actionButton', 'shape'] and y_pos < 100:
                        found_top_nav = True
                        
                    # --- RULE 3: CONSISTENT FILTERS ---
                    # If it's a filter (slicer), is it on the Left (X<150) or Top (Y<150)?
                    if visual_type == 'slicer':
                        found_slicers = True
                        if x_pos > 150 and y_pos > 150:
                            # Slicer is floating in the middle/right of the page!
                            slicers_are_consistent = False
                            
                    # --- RULE 4: VISUAL TITLES ---
                    # Standard charts (bars, lines, pies) should have titles. 
                    # We skip shapes, images, and cards as they often don't need titles.
                    charts_requiring_titles = ['barChart', 'columnChart', 'lineChart', 'pieChart', 'donutChart', 'tableEx', 'pivotTable']
                    if visual_type in charts_requiring_titles:
                        # Power BI hides titles deep in the vcObjects. If it's not explicitly turned on, it might be missing.
                        # For MVP, we check if 'title' exists in the config string at all.
                        if "'title':" not in config_string and '"title":' not in config_string:
                            visuals_missing_titles += 1

                except Exception:
                    continue # Skip visuals with broken configs
            
            # 4. Print the Audit Results for the Page
            
            # Logo Check
            if found_logo:
                print("  ✅ Logo: Present in top-left.")
            else:
                print("  ❌ Logo: Missing or not in top-left.")
                
            # Navigation Check
            if found_top_nav:
                print("  ✅ Top Nav: Buttons/Shapes detected at the top.")
            else:
                print("  ❌ Top Nav: No navigation elements found at the top.")
                
            # Filter Check
            if found_slicers:
                if slicers_are_consistent:
                    print("  ✅ Filters: Slicers are consistently placed (Top or Left).")
                else:
                    print("  ❌ Filters: Slicers are scattered outside the standard Top/Left zones.")
            else:
                print("  ➖ Filters: No slicers found on this page.")
                
            # Titles Check
            if visuals_missing_titles == 0:
                print("  ✅ Visual Titles: All major charts appear to have titles.")
            else:
                print(f"  ❌ Visual Titles: {visuals_missing_titles} chart(s) might be missing a title.")
                
            # Theme & Tooltips (Conceptual MVP Note)
            print("  ⚠️ Theme & Colors: (Requires cross-checking Theme.json - Passed for MVP)")
            print("  ⚠️ Custom Tooltips: (Requires Deep JSON parsing - Passed for MVP)")
            
            print("\n") # Add a blank line before checking the next page

except FileNotFoundError:
    print(f"❌ ERROR: Could not find report.json. Make sure this script is in the same folder as the report file!")
except Exception as e:
    print(f"❌ An error occurred: {e}")
    
