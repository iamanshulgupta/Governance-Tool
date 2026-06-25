import streamlit as st
import zipfile
import json
import re
import pandas as pd
from io import BytesIO
import time

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

# ==========================================
# 🎨 UI: PAGE CONFIG & CUSTOM CSS
# ==========================================
st.set_page_config(page_title="Dynamic PBI Governance", page_icon="⚙️", layout="wide")

st.markdown("""
    <style>
           .block-container { padding-top: 2rem; padding-bottom: 2rem; }
           .center-text { text-align: center; }
           .subtitle { text-align: center; color: #555; margin-bottom: 40px; font-size: 1.1rem; display: block; }
           .disclaimer { font-size: 0.8rem; color: #6c757d; text-align: center; margin-top: 15px; font-style: italic; }

           /* PERFECTLY ALIGNED FLOW DIAGRAM */
           .flow-container {
               display: flex; align-items: stretch; justify-content: space-between;
               gap: 15px; width: 100%; padding: 10px 0;
           }
           .wf-box {
               flex: 1; border: 2px solid #0d6efd; border-radius: 8px;
               padding: 15px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05);
               background-color: white; display: flex; flex-direction: column;
               justify-content: flex-start; align-items: center;
           }
           .wf-title { font-weight: bold; margin-bottom: 8px; color: #333; font-size: 15px; }
           .wf-desc { font-size: 12px; color: #555; line-height: 1.4; }
           .arrow-container { display: flex; align-items: center; justify-content: center; }
           .arrow { font-size: 24px; color: #adb5bd; font-weight: bold; transition: transform 0.3s ease; }

           @media (max-width: 768px) {
               .flow-container { flex-direction: column; gap: 5px; }
               .wf-box { padding: 10px; }
               .arrow { transform: rotate(90deg); margin: 5px 0; }
           }
    </style>
    """, unsafe_allow_html=True)

def custom_header(text):
    st.markdown(f"<h4 style='margin-bottom: 0px;'>{text}</h4><hr style='margin-top: 10px; margin-bottom: 15px; border-top: 1px solid #ddd;'>", unsafe_allow_html=True)

# --- SESSION STATE INITIALIZATION ---
if "batch_audit_completed" not in st.session_state:
    st.session_state.batch_audit_completed = False
if "batch_report_bytes" not in st.session_state:
    st.session_state.batch_report_bytes = None
if "file_count" not in st.session_state:
    st.session_state.file_count = 0

def reset_audit():
    st.session_state.batch_audit_completed = False
    st.session_state.batch_report_bytes = None

# ==========================================
# 🧭 SIDEBAR NAVIGATION
# ==========================================
with st.sidebar:
    st.markdown("### ⚙️ Dynamic Engine")
    st.caption("Batch Compliance Processor")
    st.button("➕ Reset Batch Audit", type="primary", use_container_width=True, on_click=reset_audit)
    
    st.divider()
    current_view = st.radio("Navigation", ["📊 Audit Workspace", "📖 Audit Documentation"], label_visibility="collapsed")

# ==========================================
# MAIN UI: AUDIT WORKSPACE 
# ==========================================
if current_view == "📊 Audit Workspace":
    
    st.markdown("<h1 class='center-text'>Dynamic Power BI Governance</h1>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Upload your custom rules matrix and multiple .pbix files to dynamically batch-process compliance reports.</p>", unsafe_allow_html=True)

    # --- ROW 0: HOW IT WORKS FLOW DIAGRAM ---
    st.markdown("#### 🔄 How the Batch Audit Works")
    with st.container(border=True):
        html_flow = """
        <div class="flow-container">
            <div class="wf-box">
                <div class="wf-title">1. Setup Template</div>
                <div class="wf-desc">Download the dynamic Excel rule template and configure your custom requirements.</div>
            </div>
            <div class="arrow-container"><div class="arrow">→</div></div>
            <div class="wf-box">
                <div class="wf-title">2. Upload Rules</div>
                <div class="wf-desc">Upload your configured Excel Matrix to tell the engine what constraints to look for.</div>
            </div>
            <div class="arrow-container"><div class="arrow">→</div></div>
            <div class="wf-box">
                <div class="wf-title">3. Batch Upload</div>
                <div class="wf-desc">Drop multiple .pbix files into the Audit Zone to process them all simultaneously.</div>
            </div>
            <div class="arrow-container"><div class="arrow">→</div></div>
            <div class="wf-box">
                <div class="wf-title">4. Generate Report</div>
                <div class="wf-desc">Download a comprehensive multi-tab Excel report detailing passes and failures.</div>
            </div>
        </div>
        """
        st.markdown(html_flow, unsafe_allow_html=True)

    # --- ROW 1 & 2: CONFIGURATION AND UPLOADS ---
    col_rules, col_audit = st.columns(2, gap="large")

    with col_rules:
        with st.container(height=350, border=True):
            custom_header("⚙️ 1. Configure Engine Rules")
            st.markdown("Download the template, add your custom constraints, and upload the matrix below.")
            
            st.download_button(
                label="📥 Download Rule Template (Excel)",
                data=create_template(),
                file_name="Dynamic_Rules_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.write("")
            rules_file = st.file_uploader("Upload Configured Rules Matrix (.xlsx)", type=["xlsx"])
            if rules_file:
                st.success("✅ Rules Matrix Loaded")

    with col_audit:
        with st.container(height=350, border=True):
            custom_header("🛡️ 2. The Batch Audit Zone")
            st.markdown("Upload the `.pbix` dashboards you want to evaluate against the rules matrix.")
            
            uploaded_files = st.file_uploader("Upload .pbix files", type="pbix", accept_multiple_files=True, label_visibility="collapsed")
            
            if uploaded_files:
                st.success(f"✅ {len(uploaded_files)} File(s) Queued for Processing")

    st.write("")

    # --- ROW 3: EXECUTION ZONE ---
    if uploaded_files and rules_file:
        btn_text = "⚡ Run Dynamic Batch Audit" if not st.session_state.batch_audit_completed else "🔄 Re-Run Batch Audit"
        
        if st.button(btn_text, type="primary", use_container_width=True):
            with st.spinner(f"Evaluating {len(uploaded_files)} dashboards against dynamic rules..."):
                try:
                    df_rules = pd.read_excel(rules_file)
                    
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for uploaded_file in uploaded_files:
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
                    
                    # Save results to session state
                    st.session_state.batch_report_bytes = output.getvalue()
                    st.session_state.batch_audit_completed = True
                    st.session_state.file_count = len(uploaded_files)
                    
                except Exception as e:
                    st.error(f"❌ Failed to load rules matrix: {e}")

    elif uploaded_files or rules_file:
        st.warning("⚠️ Please upload BOTH a Rules Matrix and at least one .pbix file to run the batch engine.")

    # --- ROW 4: RESULTS DASHBOARD ---
    if st.session_state.batch_audit_completed:
        st.divider()
        st.markdown("### 📈 Batch Audit Results")
        with st.container(border=True):
            c1, c2 = st.columns([3, 1])
            with c1:
                st.success(f"🎉 Successfully audited **{st.session_state.file_count}** dashboard(s) against your custom rules.")
                st.markdown("Your multi-tab Excel report is ready. Each processed dashboard has its own tab detailing page-by-page visual violations.")
            with c2:
                st.download_button(
                    label="📥 Download Full Report (.xlsx)",
                    data=st.session_state.batch_report_bytes,
                    file_name="Dynamic_Governance_Audit.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

# ==========================================
# PAGE 2: AUDIT DOCUMENTATION
# ==========================================
else:
    st.title("📖 Audit Documentation")
    st.info("Documentation and guidelines for the Dynamic Batch Engine.")

    with st.expander("**Overview**", expanded=False):
        st.markdown("""
        The Dynamic Power BI Governance Audit Tool is an automated compliance engine designed to enforce design standards and best practices across Power BI dashboards.
        
        Auditing Power BI reports for structural consistency previously required manually opening every single `.pbix` file, which is a highly tedious process. This tool automates that workflow. By extracting and scanning the underlying JSON metadata of `.pbix` files, it performs batch quality-assurance checks in seconds and generates a detailed compliance report based entirely on your custom Excel parameters.
        """)

    with st.expander("**Key Features**", expanded=False):
        st.markdown("""
        * **In-Memory Batch Processing:** Upload multiple `.pbix` files at once. The tool processes them securely in memory.
        * **Dynamic Rule Configuration:** Download a rule template, add custom rules, adjust pixel boundaries for layout configurations to suit your needs, and upload it back to dynamically drive the audit.
        * **Deep Metadata Parsing:** Treats `.pbix` files as zip archives to unlock the hidden `Report/Layout` structure, scanning exact X/Y coordinates and visual configurations.
        * **Automated Excel Reporting:** Outputs a clean, multi-tab Excel file detailing the exact pass/fail status of every visual, where each tab corresponds to a specific audited dashboard.
        """)

    with st.expander("**How It Works (The Architecture)**", expanded=False):
        st.markdown("""
        This tool leverages Python, Pandas, and Streamlit in the backend.
        
        Power BI `.pbix` files are essentially zipped directories. The engine bypasses the Power BI application entirely by unzipping the `.pbix` file in memory and locating the `Report/Layout` file. Because this file is encoded in UTF-16 LE, the script decodes it into a readable JSON format.
        
        Once the JSON tree is exposed, the engine iterates through every section (page) and visualContainer (chart/slicer/shape). It extracts the configuration string to map coordinates (x, y) and visual types against the active governance checklist.
        """)

    with st.expander("**How to Use the Tool**", expanded=False):
        st.markdown("""
        1. **Download the Template:** Go to the Workspace and click the "Download Rule Template" button to get the standard governance matrix.
        2. **Customize the Rules:** Open the downloaded Excel file. Instead of hardcoding rules, this app reads your Excel file row by row.
           * **Target Visual:** Which visual type does this rule apply to? (e.g., `image`, `slicer`, `barChart`, or `all`)
           * **Property to Check:** What are we measuring? (Currently supports: `x_position`, `y_position`, `title_exists`)
           * **Condition:** `Less Than`, `Greater Than`, or `Equals`
        3. **Upload the Assets:** * Upload your customized Excel checklist into Box 1.
           * Upload one or more `.pbix` files into Box 2.
        4. **Run the Audit:** Click "Run Dynamic Batch Audit."
        5. **Review Results:** Download the generated `Dynamic_Governance_Audit.xlsx` report to instantly see which dashboards meet company standards.
        """)