import streamlit as st
import zipfile
import json
import re
import pandas as pd
from io import BytesIO
import io

# ==========================================
# 🧠 DYNAMIC RULE ENGINE: TEMPLATE GENERATOR
# ==========================================
def create_template():
    df = pd.DataFrame({
        "Rule Name": ["Logo Top Left X", "Logo Top Left Y", "Slicer Left Zone", "Slicer Top Zone", "Charts Must Have Titles"],
        "Target Visual": ["image", "image", "slicer", "slicer", "barChart"],
        "Property to Check": ["x_position", "y_position", "x_position", "y_position", "title_exists"],
        "Condition": ["Less Than", "Less Than", "Greater Than", "Greater Than", "Equals"],
        "Target Value": [100, 100, 150, 150, "True"],
        "Requirement": ["Must Pass", "Must Pass", "Must Pass", "Must Pass", "Must Pass"]
    })
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Governance Rules")
    return output.getvalue()

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Power BI Governance Engine", page_icon="📊", layout="wide")

# --- CUSTOM CSS & HELPERS ---
st.markdown("""
    <style>
           .block-container { padding-top: 2rem; padding-bottom: 2rem; }
           .center-text { text-align: center; }
           .subtitle { text-align: center; color: #555; margin-bottom: 40px; font-size: 1.1rem; display: block; }
           .disclaimer { font-size: 0.8rem; color: #6c757d; text-align: center; margin-top: 15px; font-style: italic; }

           /* PERFECTLY ALIGNED FLOW DIAGRAM */
           .flow-container {
               display: flex;
               align-items: stretch;
               justify-content: space-between;
               gap: 15px;
               width: 100%;
               padding: 10px 0;
           }
           .wf-box {
               flex: 1;
               border: 2px solid #0d6efd;
               border-radius: 8px;
               padding: 15px;
               text-align: center;
               box-shadow: 0 2px 4px rgba(0,0,0,0.05);
               background-color: white;
               display: flex;
               flex-direction: column;
               justify-content: flex-start;
               align-items: center;
           }
           .wf-title { font-weight: bold; margin-bottom: 8px; color: #333; font-size: 15px; }
           .wf-desc { font-size: 12px; color: #555; line-height: 1.4; }
           .arrow-container { display: flex; align-items: center; justify-content: center; }
           .arrow { font-size: 24px; color: #adb5bd; font-weight: bold; transition: transform 0.3s ease; }

           /* MOBILE RESPONSIVENESS FOR FLOW DIAGRAM */
           @media (max-width: 768px) {
               .flow-container { flex-direction: column; gap: 5px; }
               .wf-box { padding: 10px; }
               .arrow { transform: rotate(90deg); margin: 5px 0; }
           }
    </style>
    """, unsafe_allow_html=True)

def custom_header(text):
    st.markdown(f"<h4 style='margin-bottom: 0px;'>{text}</h4><hr style='margin-top: 10px; margin-bottom: 15px; border-top: 1px solid #ddd;'>", unsafe_allow_html=True)

def get_btn_type(filepath):
    return "primary" if st.session_state.active_file == filepath else "secondary"

# --- SESSION STATE INITIALIZATION ---
if "active_file" not in st.session_state:
    st.session_state.active_file = None
if "active_file_obj" not in st.session_state:
    st.session_state.active_file_obj = None
if "audit_completed" not in st.session_state:
    st.session_state.audit_completed = False
if "batch_report_bytes" not in st.session_state:
    st.session_state.batch_report_bytes = None
if "just_audited" not in st.session_state:
    st.session_state.just_audited = False

def set_sample(filename):
    st.session_state.active_file = filename
    st.session_state.active_file_obj = filename
    st.session_state.audit_completed = False

def clear_file():
    st.session_state.active_file = None
    st.session_state.active_file_obj = None
    st.session_state.audit_completed = False
    st.session_state.batch_report_bytes = None

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    st.markdown("### 📊 Governance Engine")
    st.caption("v2.4.0 Compliance (Dynamic Matrix)")
    st.button("➕ Run New Audit", type="primary", use_container_width=True, on_click=clear_file)
    
    st.divider()
    current_view = st.radio("Navigation", ["📊 Audit Workspace", "📖 Audit Documentation"], label_visibility="collapsed")

# ==========================================
# MAIN UI: AUDIT WORKSPACE 
# ==========================================
if current_view == "📊 Audit Workspace":
    
    st.markdown("<h1 class='center-text'>Power BI Governance Engine</h1>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Automated compliance checking for your Power BI reports. Ensure visual consistency, structural integrity, and brand alignment before deployment.</p>", unsafe_allow_html=True)
    
    # --- ROW 0: HOW IT WORKS FLOW DIAGRAM ---
    st.markdown("#### 🔄 How the Audit Works")
    with st.container(border=True):
        html_flow = """
        <div class="flow-container">
            <div class="wf-box">
                <div class="wf-title">Define Your Standard</div>
                <div class="wf-desc">Set your desired layout rules in the configuration panel (e.g., Logo Position).</div>
            </div>
            <div class="arrow-container"><div class="arrow">→</div></div>
            <div class="wf-box">
                <div class="wf-title">Check the Wireframe</div>
                <div class="wf-desc">The Live Wireframe instantly shows where features belong based on your rules.</div>
            </div>
            <div class="arrow-container"><div class="arrow">→</div></div>
            <div class="wf-box">
                <div class="wf-title">Run the Audit</div>
                <div class="wf-desc">Upload your rules matrix and your Power BI file, then click Run New Audit.</div>
            </div>
            <div class="arrow-container"><div class="arrow">→</div></div>
            <div class="wf-box">
                <div class="wf-title">Review Results</div>
                <div class="wf-desc">The dynamic engine evaluates your uploaded file against your custom matrix.</div>
            </div>
        </div>
        """
        st.markdown(html_flow, unsafe_allow_html=True)
    
    # --- ROW 1: CONFIGURE RULES (Static UI only) ---
    with st.container(border=True):
        custom_header("⚙️ Configure Rules (Wireframe Visualizer)")
        
        # Line 1
        r1_c1, r1_c2, r1_c3, r1_c4 = st.columns(4)
        ui_logo = r1_c1.selectbox("1. Logo Position", ["Top Left", "Top Right", "Not Present"])
        ui_ratio = r1_c2.selectbox("2. Aspect Ratio", ["16:9", "16:10", "4:3", "Custom"])
        ui_filter = r1_c3.selectbox("3. Filter Pane Status", ["Available (Top Left)", "Available (Top Right)", "Hidden / Not Available"])
        ui_nav = r1_c4.selectbox("4. Navigation Bar", ["Top", "Left", "Not Available"])

        # Line 2 
        r2_c1, r2_c2, r2_c3, r2_c4 = st.columns(4)
        ui_headers = r2_c1.selectbox("5. Standardized Headers", ["Available", "Not Available"])
        ui_refresh = r2_c2.selectbox("6. Last Refresh Date", ["Top Right", "Bottom Left", "Bottom Right", "Not Present"])
        ui_tooltip = r2_c3.selectbox("7. Tooltip Configuration", ["Custom Report Page", "Standard Default", "Not Evaluated"])

    # --- ROW 2: LIVE WIREFRAME PREVIEW ---
    with st.container(border=True):
        custom_header("👁️ Live Wireframe Preview")
        
        c_top, c_bottom, c_left, c_right = 10, 10, 10, 10
        has_top, has_bottom = False, False

        logo_style = "border: 2px solid #fd7e14; color: #fd7e14; height: 35px; z-index: 10; overflow: hidden;"
        if ui_logo != "Not Present":
            if "Top" in ui_logo: logo_style += "top: 10px;"; has_top = True
            else: logo_style += "bottom: 10px;"; has_bottom = True
            if "Left" in ui_logo: logo_style += "left: 10px;"
            else: logo_style += "right: 10px;"

        refresh_style = "border: 2px solid #0d6efd; color: #0d6efd; height: 35px; z-index: 11; overflow: hidden;"
        if ui_refresh != "Not Present":
            if "Top" in ui_refresh: refresh_style += "top: 10px;"; has_top = True
            else: refresh_style += "bottom: 10px;"; has_bottom = True
            
            offset_l, offset_r = 10, 10
            if ui_logo == ui_refresh:
                if "Left" in ui_refresh: offset_l = 120
                else: offset_r = 120

            if "Left" in ui_refresh: refresh_style += f"left: {offset_l}px;"
            else: refresh_style += f"right: {offset_r}px;"

        header_style = "border: 2px solid #ffc107; color: #ffc107; top: 10px; left: 50%; transform: translateX(-50%); height: 35px; overflow: hidden;"
        if ui_headers == "Available": has_top = True

        if has_top: c_top += 45
        if has_bottom: c_bottom += 45

        nav_style = "border: 2px solid #6c757d; color: #6c757d; z-index: 9; text-align: center;"
        if ui_nav == "Top":
            nav_style += f"top: {c_top}px; left: {c_left}px; right: {c_right}px; height: 35px;"
            c_top += 45
        elif ui_nav == "Bottom":
            nav_style += f"bottom: {c_bottom}px; left: {c_left}px; right: {c_right}px; height: 35px;"
            c_bottom += 45
        elif ui_nav == "Left":
            nav_style += f"left: {c_left}px; top: {c_top}px; bottom: {c_bottom}px; width: 40px; writing-mode: vertical-rl; text-orientation: mixed;"
            c_left += 50
        elif ui_nav == "Right":
            nav_style += f"right: {c_right}px; top: {c_top}px; bottom: {c_bottom}px; width: 40px; writing-mode: vertical-rl; text-orientation: mixed;"
            c_right += 50

        filter_style = "border: 2px solid #0dcaf0; color: #0dcaf0; z-index: 8; writing-mode: vertical-rl; text-orientation: mixed; text-align: center;"
        if ui_filter != "Hidden / Not Available":
            if "Left" in ui_filter:
                filter_style += f"left: {c_left}px; top: {c_top}px; bottom: {c_bottom}px; width: 100px;"
                c_left += 110
            elif "Right" in ui_filter:
                filter_style += f"right: {c_right}px; top: {c_top}px; bottom: {c_bottom}px; width: 100px;"
                c_right += 110

        html = f"""
        <style>
            .wireframe-container {{
                position: relative; width: 100%; height: 450px; 
                background-color: #f1f3f5; border: 2px solid #dee2e6; border-radius: 8px;
                box-shadow: inset 0 0 10px rgba(0,0,0,0.05); overflow: hidden; font-family: sans-serif;
            }}
            .wf-element {{
                position: absolute; display: flex; align-items: center; justify-content: center;
                font-weight: bold; font-size: 13px; border-radius: 4px; background-color: white;
                transition: all 0.3s ease; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            }}
            
            /* DYNAMIC SIZING FOR PERIMETER ELEMENTS (WEB) */
            .wf-header {{ width: 500px; font-size: 16px; }}
            .wf-logo {{ width: 100px; }}
            .wf-refresh {{ width: 120px; }}
            
            .central-zone {{
                position: absolute; background-color: transparent;
                display: flex; flex-direction: column; gap: 10px;
                transition: all 0.3s ease;
            }}
            .kpi-row {{ display: flex; gap: 10px; height: 60px; }}
            .kpi {{ flex: 1; border: 2px dashed #adb5bd; border-radius: 4px; display: flex; align-items: center; justify-content: center; color: #495057; font-weight: bold; background-color: white; font-size: 13px; box-shadow: 0 2px 4px rgba(0,0,0,0.02);}}
            .chart-row {{ display: flex; gap: 10px; flex: 1; }}
            .chart-main {{ flex: 3; border: 2px dashed #adb5bd; border-radius: 4px; display: flex; align-items: center; justify-content: center; color: #495057; font-weight: bold; background-color: white; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.02);}}
            .chart-side {{ flex: 1; border: 2px dashed #adb5bd; border-radius: 4px; display: flex; align-items: center; justify-content: center; color: #495057; font-weight: bold; background-color: white; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.02);}}
            .wf-tooltip {{ background-color: #343a40; color: white; padding: 6px; font-size: 11px; z-index: 20; width: 100px; height: 30px; top: 200px; left: 45%; border: none; box-shadow: 0 4px 8px rgba(0,0,0,0.2);}}
            .wf-tooltip::after {{ content: ''; position: absolute; top: 100%; left: 50%; margin-left: -5px; border-width: 5px; border-style: solid; border-color: #343a40 transparent transparent transparent;}}

            /* MOBILE RESPONSIVENESS FOR WIREFRAME (FIXED OVERLAP & ALIGNMENT) */
            @media (max-width: 768px) {{
                .wireframe-container {{ height: 600px; }} 
                
                /* Force all top elements to match sizes perfectly to prevent overlap on small screens */
                .wf-header, .wf-logo, .wf-refresh {{ width: 95px !important; font-size: 10px !important; }}
                
                .kpi-row {{ flex-wrap: wrap; height: auto; }}
                .kpi {{ min-width: 45%; height: 45px; margin-bottom: 8px; font-size: 11px; }}
                .chart-row {{ flex-direction: column; }}
                .chart-main {{ flex: none; height: 180px; margin-bottom: 10px; font-size: 12px; }}
                .chart-side {{ flex: none; height: 120px; font-size: 12px; }}
                .wf-tooltip {{ left: 5%; top: 100px; width: 80px; font-size: 9px; }}
            }}
        </style>
        
        <div class='wireframe-container'>
            <div class='central-zone' style='top: {c_top}px; bottom: {c_bottom}px; left: {c_left}px; right: {c_right}px;'>
                 <div class='kpi-row'>
                     <div class='kpi'>KPI</div><div class='kpi'>KPI</div><div class='kpi'>KPI</div><div class='kpi'>KPI</div>
                 </div>
                 <div class='chart-row'>
                     <div class='chart-main'>Main Bar Chart</div>
                     <div class='chart-side'>Side Chart</div>
                 </div>
            </div>
        """
        
        if ui_logo != "Not Present": html += f"<div class='wf-element wf-logo' style='{logo_style}'>Logo</div>"
        if ui_nav != "Not Available": html += f"<div class='wf-element' style='{nav_style}'>Navigator</div>"
        if ui_headers == "Available": html += f"<div class='wf-element wf-header' style='{header_style}'>Header</div>"
        if ui_filter != "Hidden / Not Available": html += f"<div class='wf-element' style='{filter_style}'>Filter Pane</div>"
        if ui_refresh != "Not Present": html += f"<div class='wf-element wf-refresh' style='{refresh_style}'>Last Refresh</div>"
        
        if ui_tooltip != "Not Evaluated": html += f"<div class='wf-element wf-tooltip'>Custom Tooltip</div>"
            
        html += "</div>"
        
        st.markdown(html, unsafe_allow_html=True)
        st.caption(f"*Dynamic Dashboard Constraint Engine Active.*")

    st.write("") 

    # --- ROW 3: SAMPLES & AUDIT ZONE ---
    col_samples, col_audit = st.columns(2, gap="large")
    
    with col_samples:
        with st.container(height=380, border=True):
            custom_header("📁 1. Rules Matrix & Samples")
            st.markdown("Download the template, add custom rules, and upload the matrix below to drive the audit.")
            
            st.download_button(
                label="📥 Download Rule Template (Excel)",
                data=create_template(),
                file_name="Dynamic_Rules_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            rules_file = st.file_uploader("Upload Rules Matrix (.xlsx)", type=["xlsx"])
            if rules_file:
                st.success("✅ Dynamic Rules Loaded")

    with col_audit:
        with st.container(height=380, border=True):
            custom_header("🛡️ 2. The Audit Zone")
            
            if st.session_state.active_file:
                display_name = st.session_state.active_file.split('/')[-1] if isinstance(st.session_state.active_file, str) else st.session_state.active_file
                st.success(f"✅ **File Loaded:** {display_name}")
                st.button("✖ Clear File", on_click=clear_file)
            else:
                uploaded_files = st.file_uploader("Drop your .pbix file(s) here", type=["pbix"], accept_multiple_files=True)
                if uploaded_files:
                    st.session_state.active_file_obj = uploaded_files
                    st.session_state.active_file = uploaded_files[0].name if len(uploaded_files) == 1 else f"{len(uploaded_files)} files selected"
                    st.session_state.audit_completed = False
                    st.rerun()
            
            st.write("") 
            
            btn_text = "▶ Run Dynamic Governance Check"
            if st.session_state.audit_completed:
                btn_text = "🔄 Update & Re-Run Check (View Results Below 👇)"
                
            if st.button(btn_text, type="primary", use_container_width=True):
                if st.session_state.active_file_obj and rules_file:
                    with st.spinner("Extracting PBIX Layout Metadata and evaluating matrix..."):
                        try:
                            df_rules = pd.read_excel(rules_file)
                            
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                files_to_process = st.session_state.active_file_obj if isinstance(st.session_state.active_file_obj, list) else [st.session_state.active_file_obj]
                                
                                for uploaded_file in files_to_process:
                                    dashboard_name = uploaded_file.name.replace('.pbix', '')
                                    
                                    try:
                                        with zipfile.ZipFile(uploaded_file, 'r') as pbix_zip:
                                            with pbix_zip.open('Report/Layout') as layout_file:
                                                content = layout_file.read().decode('utf-16-le')
                                                report_data = json.loads(content)
                                                
                                        pages = report_data.get('sections', [])
                                        dashboard_results = [] 
                                        
                                        for page in pages:
                                            page_name = page.get('displayName', 'Unknown Page')
                                            visuals = page.get('visualContainers', [])
                                            
                                            page_rule_stats = {row['Rule Name']: {"evaluated": 0, "failed": 0} for _, row in df_rules.iterrows()}
                                            
                                            for visual in visuals:
                                                y_pos = visual.get('y', 999) 
                                                x_pos = visual.get('x', 999)
                                                config_string = visual.get('config', '{}')
                                                
                                                try:
                                                    config_data = json.loads(config_string)
                                                    v_type = config_data.get('singleVisual', {}).get('visualType', 'Unknown')
                                                    has_title = 'true' if ("'title':" in config_string or '"title":' in config_string) else 'false'
                                                    
                                                    actual_properties = {
                                                        "x_position": x_pos,
                                                        "y_position": y_pos,
                                                        "title_exists": has_title
                                                    }
                                                    
                                                    for _, rule in df_rules.iterrows():
                                                        target_vis = rule['Target Visual']
                                                        prop = rule['Property to Check']
                                                        cond = rule['Condition']
                                                        target_val = str(rule['Target Value']).strip().lower()
                                                        rule_name = rule['Rule Name']
                                                        
                                                        if target_vis.lower() == v_type.lower() or target_vis.lower() == 'all':
                                                            page_rule_stats[rule_name]["evaluated"] += 1
                                                            actual_val = actual_properties.get(prop)
                                                            
                                                            passed = True
                                                            try:
                                                                if cond == 'Less Than': passed = float(actual_val) < float(target_val)
                                                                elif cond == 'Greater Than': passed = float(actual_val) > float(target_val)
                                                                elif cond == 'Equals': passed = str(actual_val).lower() == target_val
                                                            except:
                                                                passed = False 
                                                                
                                                            if not passed:
                                                                page_rule_stats[rule_name]["failed"] += 1

                                                except Exception:
                                                    continue 

                                            row_data = {"Page Name": page_name, "Total Visuals": len(visuals)}
                                            for rule_name, stats in page_rule_stats.items():
                                                if stats["evaluated"] == 0: row_data[rule_name] = "➖ N/A (Visual Not Found)"
                                                elif stats["failed"] == 0: row_data[rule_name] = "✅ Pass"
                                                else: row_data[rule_name] = f"❌ Fail ({stats['failed']} violations)"
                                                
                                            dashboard_results.append(row_data)

                                        if dashboard_results:
                                            df = pd.DataFrame(dashboard_results)
                                            safe_sheet_name = re.sub(r'[\\/*?:\[\]]', '', dashboard_name)[:31]
                                            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

                                    except Exception as e:
                                        st.error(f"❌ Could not process {dashboard_name}: {e}")
                            
                            st.session_state.batch_report_bytes = output.getvalue()
                            st.session_state.audit_completed = True
                            st.session_state.just_audited = True
                            st.rerun()
                            
                        except Exception as e:
                            st.error(f"❌ Failed to evaluate dynamic rules: {e}")
                            
                elif not rules_file:
                    st.error("Please upload a Rules Matrix (.xlsx) in the left panel first.")
                else:
                    st.error("Please upload at least one .pbix file.")
            
            st.markdown("<p class='disclaimer'>*Files uploaded to this tool are processed locally and are not saved or stored anywhere.</p>", unsafe_allow_html=True)

    # --- ROW 4: DYNAMIC RESULTS DASHBOARD ---
    if st.session_state.audit_completed and st.session_state.batch_report_bytes:
        st.divider()
        
        # INVISIBLE ANCHOR FOR AUTO-SCROLLING
        st.markdown("<div id='results-target'></div>", unsafe_allow_html=True)
        
        st.markdown(f"### 📈 Audit Report Generated")
        
        with st.container(border=True):
            st.success(f"🎉 Dynamic Audit Complete! Your multi-tab Excel report is ready.")
            st.markdown("Download the full compliance report to view page-by-page visual violations against your custom matrix.")
            
            st.download_button(
                label="📥 Download Full Audit Log (.xlsx)",
                data=st.session_state.batch_report_bytes,
                file_name="Dynamic_Governance_Audit.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

        # --- ADVANCED AUTO SCROLL MAGIC ---
        if st.session_state.get("just_audited", False):
            st.components.v1.html(
                """
                <script>
                    setTimeout(function() {
                        const target = window.parent.document.getElementById('results-target');
                        if (target) {
                            target.scrollIntoView({ behavior: 'smooth', block: 'start' });
                        } else {
                            const main = window.parent.document.querySelector('.main') || window.parent.document.querySelector('section.main');
                            if (main) {
                                main.scrollTo({ top: main.scrollHeight, behavior: 'smooth' });
                            } else {
                                window.parent.scrollTo({ top: window.parent.document.body.scrollHeight, behavior: 'smooth' });
                            }
                        }
                    }, 500); // Wait 500ms to guarantee results have loaded visually
                </script>
                """,
                height=0
            )
            st.session_state.just_audited = False 

else:
    st.title("📖 Audit Documentation")
    st.info("Documentation and guidelines for the Governance Engine will live here.")