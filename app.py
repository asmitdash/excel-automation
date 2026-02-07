import streamlit as st
import pandas as pd
import io
import zipfile
import re

# --- CONFIGURATION ---
st.set_page_config(page_title="Town Scorecard Generator", layout="wide")

def clean_column_names(df):
    """
    Standardizes column names to avoid 'KeyError'.
    Removes newlines, extra spaces, and handles variations.
    """
    df.columns = df.columns.str.replace('\n', ' ').str.strip()
    return df

def get_col_by_keyword(df, keywords):
    """
    Finds a column that matches one of the keywords.
    Returns the first match or None.
    """
    for col in df.columns:
        for kw in keywords:
            if kw.lower() in col.lower():
                return col
    return None

def generate_town_excel(town_name, town_df):
    """
    Generates the Azamgarh-style scorecard with strict logic:
    Primary = MD, Dealer, Branch (BR)
    Secondary = ASD
    """
    
    # ==========================================
    # 1. PREPARE DATA & COLUMNS
    # ==========================================
    scorecard_rows = []
    
    # Identify key columns dynamically (Robustness)
    col_strat = get_col_by_keyword(town_df, ['Updated Stratification', 'Stratification'])
    col_bal_type = get_col_by_keyword(town_df, ['BAL Store Type'])
    col_tvs_type = get_col_by_keyword(town_df, ['TVS Store Type'])
    
    # Volumes (Try multiple variations found in your files)
    col_ind_s1 = get_col_by_keyword(town_df, ['S1 Ind - F', 'S1 Ind Vistaar', 'S1 Ind'])
    col_bal_s1 = get_col_by_keyword(town_df, ['BAL S1 Vol', 'BAL S1'])
    col_tvs_s1 = get_col_by_keyword(town_df, ['TVS S1 Vol', 'TVS S1'])
    col_cr = get_col_by_keyword(town_df, ['CR'])
    
    # Intervention
    col_nature = get_col_by_keyword(town_df, ['Nature of Intervention'])
    col_network = get_col_by_keyword(town_df, ['Network Intervention'])
    
    # Filter out 'Closed' stores
    # Look for 'Sub-Location' and 'Highlight' or 'closed'
    clean_df = town_df.copy()
    col_closed = get_col_by_keyword(clean_df, ['Sub-Location', 'Highlight'])
    if col_closed:
        clean_df = clean_df[~clean_df[col_closed].astype(str).str.contains('closed', case=False, na=False)]

    if not col_strat:
        return None # Critical error if stratification missing

    # Get Stratifications
    stratifications = [x for x in clean_df[col_strat].dropna().unique()]

    for strat in stratifications:
        strat_df = clean_df[clean_df[col_strat] == strat]

        # --- CLASSIFICATION LOGIC (Strict) ---
        # Primary: MD, Dealer, Branch, BR
        # Secondary: ASD, AD, Sub, Rep
        # Vacant: Everything else
        
        def classify_row(row):
            val = str(row[col_bal_type]).upper() if col_bal_type else ""
            if any(x in val for x in ['MD', 'DEALER', 'BRANCH', 'BR']):
                return 'Primary'
            elif any(x in val for x in ['ASD', 'AD', 'SUB', 'REP']):
                return 'Secondary'
            else:
                return 'Vacant'

        strat_df = strat_df.copy()
        strat_df['Category'] = strat_df.apply(classify_row, axis=1)

        # Create sub-dataframes
        pri_df = strat_df[strat_df['Category'] == 'Primary']
        sec_df = strat_df[strat_df['Category'] == 'Secondary']
        vac_df = strat_df[strat_df['Category'] == 'Vacant']

        # --- METRIC CALCULATION HELPER ---
        def calc_metrics(df):
            count = len(df)
            
            # TVS Counts (Primary vs Secondary) within this BAL category
            t_pri = 0
            t_sec = 0
            if col_tvs_type:
                t_pri = len(df[df[col_tvs_type].astype(str).str.contains('MD|Dealer|Branch', case=False, na=False)])
                t_sec = len(df[df[col_tvs_type].astype(str).str.contains('ASD|AD|Sub', case=False, na=False)])
            
            # Volumes
            ind = pd.to_numeric(df[col_ind_s1], errors='coerce').sum() if col_ind_s1 else 0
            bal = pd.to_numeric(df[col_bal_s1], errors='coerce').sum() if col_bal_s1 else 0
            tvs = pd.to_numeric(df[col_tvs_s1], errors='coerce').sum() if col_tvs_s1 else 0
            
            ms = (bal / ind) if ind > 0 else 0
            v_gap = tvs - bal
            cr = pd.to_numeric(df[col_cr], errors='coerce').mean() if col_cr else 0
            
            return count, t_pri, t_sec, ind, bal, tvs, ms, v_gap, cr

        # Calculate for each
        p_c, p_tp, p_ts, p_ind, p_bal, p_tvs, p_ms, p_vg, p_cr = calc_metrics(pri_df)
        s_c, s_tp, s_ts, s_ind, s_bal, s_tvs, s_ms, s_vg, s_cr = calc_metrics(sec_df)
        v_c, v_tp, v_ts, v_ind, v_bal, v_tvs, v_ms, v_vg, v_cr = calc_metrics(vac_df)
        
        # --- STORE GAP LOGIC ---
        # Store Gap = BAL - TVS (Positive means BAL is ahead, Negative means Gap)
        # Wait, usually Gap means "Missing". If TVS has 5 and BAL has 3, Gap is 2.
        # Let's stick to standard math: BAL - TVS. 
        # Primary Gap
        gap_p = p_c - p_tp
        # Secondary Gap
        gap_s = s_c - s_ts
        # Vacant Gap (Usually 0 BAL vs 0 TVS, but logic says if it's Vacant row, it's a potential location)
        gap_v = 0 

        # --- TOTALS ---
        tot_bal = p_c + s_c + v_c
        tot_tvs = (p_tp + p_ts) + (s_tp + s_ts) + (v_tp + v_ts)
        tot_gap = gap_p + gap_s + gap_v
        
        # --- BUILD ROWS ---
        # 1. Primary
        scorecard_rows.append([
            strat, "Pri Store", p_c, p_tp, "", gap_p, 0, p_ind, p_bal, p_ms, p_tvs, p_vg, p_cr, 
            "", "", "", "", p_c, 0
        ])
        # 2. Secondary
        scorecard_rows.append([
            "", "ASD", s_c, "", s_ts, gap_s, 0, s_ind, s_bal, s_ms, s_tvs, s_vg, s_cr, 
            "", "", "", "", s_c, 0
        ])
        # 3. Vacant (If any exist)
        scorecard_rows.append([
            "", "Vacant", v_c, "", "", gap_v, v_c, v_ind, v_bal, v_ms, v_tvs, v_vg, "", 
            "", "", "", "", 0, v_c
        ])
        # 4. Total
        scorecard_rows.append([
            "", "Total", tot_bal, tot_tvs, "", tot_gap, v_c, 
            (p_ind+s_ind+v_ind), (p_bal+s_bal+v_bal), "", (p_tvs+s_tvs+v_tvs), (p_vg+s_vg+v_vg), "", 
            "", "", "", "", tot_bal, v_c
        ])
        scorecard_rows.append([""]*19)

    df_scorecard = pd.DataFrame(scorecard_rows)

    # ==========================================
    # SHEET 2: NETWORK PLAN LOGIC
    # ==========================================
    # We use simple aggregation for Sheet 2
    
    # ... (Logic identical to previous, ensuring column existence)
    # Intervention List
    if col_network:
        intervention_df = clean_df[clean_df[col_network].notna()].copy()
        # Keep relevant columns dynamically
        keep_cols = []
        for c in ['Location', 'Stratification', 'TVS Store Type', 'BAL Store Type', 'S1 Ind', 'Nature', 'Network', 'Remarks']:
            found = get_col_by_keyword(clean_df, [c])
            if found: keep_cols.append(found)
        df_intervention = intervention_df[keep_cols] if not intervention_df.empty else pd.DataFrame()
    else:
        df_intervention = pd.DataFrame()


    # ==========================================
    # EXCEL WRITING (XlsxWriter)
    # ==========================================
    output = io.BytesIO()
    workbook = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook_obj = workbook.book

    # STYLES
    fmt_header = workbook_obj.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D9E1F2', 'border': 1})
    fmt_sub = workbook_obj.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D9E1F2', 'border': 1, 'font_size': 9})
    fmt_simple = workbook_obj.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1})
    
    # --- SHEET 1 ---
    if not df_scorecard.empty:
        df_scorecard.to_excel(workbook, sheet_name='Scorecard', startrow=3, header=False, index=False)
        ws1 = workbook.sheets['Scorecard']
        
        ws1.write(0, 1, town_name, workbook_obj.add_format({'bold': True, 'font_size': 14}))
        
        # HEADERS
        headers = ["Stratification", "# Store Count", "BAL", "TVS", "", "Store Gap", "Unique Location Gap", 
                   "IND S1", "S1 BAL Vol", "BAL MS", "S1 TVS Vol", "Vol Gap\n(TVS-BAL)", "CR", 
                   "Addition", "", "Reduction", "", "BAL Network Count\n@ UP 2.0", "Unique Location Gap\npost appointment"]
        
        for i, h in enumerate(headers):
            if h == "TVS": ws1.merge_range(1, 3, 1, 4, h, fmt_header)
            elif h == "Addition": ws1.merge_range(1, 13, 1, 14, h, fmt_header)
            elif h == "Reduction": ws1.merge_range(1, 15, 1, 16, h, fmt_header)
            elif h != "": ws1.write(1, i, h, fmt_header)

        # SUBHEADERS
        for col in [3, 13, 15]: ws1.write(2, col, "Primary", fmt_sub)
        for col in [4, 14, 16]: ws1.write(2, col, "Secondary", fmt_sub)
        for col in [0,1,2,5,6,7,8,9,10,11,12,17,18]: ws1.write(2, col, "", fmt_sub)
        
        ws1.set_column(0, 0, 15)
        ws1.set_column(7, 12, 12)

    # --- SHEET 2 ---
    if not df_intervention.empty:
        ws2_name = 'Network_Plan'
        df_intervention.to_excel(workbook, sheet_name=ws2_name, startrow=1, index=False)
        ws2 = workbook.sheets[ws2_name]
        for i, col in enumerate(df_intervention.columns):
            ws2.write(0, i, col, fmt_simple)
        ws2.set_column(0, len(df_intervention.columns)-1, 15)

    workbook.close()
    return output.getvalue()


# --- MAIN UI ---
st.title("📊 Master Scorecard Generator (Strict Logic)")
st.markdown("""
**Logic Applied:**
* **Primary:** MD, Dealer, Branch (BR)
* **Secondary:** ASD, AD, Sub, Rep
* **Vacant:** All other active rows
* **Closed Stores:** Excluded automatically
""")

uploaded_file = st.file_uploader("Upload 'Accessibility Excel'", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=2) # Try standard header pos
        else:
            df = pd.read_excel(uploaded_file, header=2)
            
        # Clean columns
        df = clean_column_names(df)
        
        # Check if we grabbed the right header. If 'Town' isn't there, reload with header=1
        if 'Town' not in df.columns:
            uploaded_file.seek(0)
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, header=1)
            else:
                df = pd.read_excel(uploaded_file, header=1)
            df = clean_column_names(df)

        # Remove Totals
        clean_df = df[~df['Town'].astype(str).str.contains('Total', case=False, na=False)].copy()
        unique_towns = clean_df['Town'].dropna().unique()
        
        st.success(f"✅ Data Loaded. Found {len(unique_towns)} towns.")
        
        if st.button("Generate Scorecards"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                prog = st.progress(0)
                for i, town in enumerate(unique_towns):
                    t_df = clean_df[clean_df['Town'] == town]
                    data = generate_town_excel(town, t_df)
                    if data:
                        zip_file.writestr(f"Scorecard_{town}.xlsx", data)
                    prog.progress((i + 1) / len(unique_towns))
            
            st.download_button(
                "📥 Download ZIP", 
                zip_buffer.getvalue(), 
                "Scorecards_Final.zip", 
                "application/zip"
            )
            
    except Exception as e:
        st.error(f"Error: {e}")
