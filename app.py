import streamlit as st
import pandas as pd
import io
import zipfile
import xlsxwriter

# --- CONFIGURATION ---
st.set_page_config(page_title="Town Scorecard Generator", layout="wide")

def clean_column_names(df):
    """Standardizes column names to remove newlines and spaces."""
    df.columns = df.columns.str.replace('\n', ' ').str.strip()
    return df

def get_col_by_keyword(df, keywords):
    """Finds a column containing one of the keywords (case-insensitive)."""
    for col in df.columns:
        for kw in keywords:
            if kw.lower() in col.lower():
                return col
    return None

def generate_town_excel(town_name, town_df):
    """
    Generates the Scorecard with:
    - Fixed Stratification Order (Large -> Small -> Rural -> Deep Rural)
    - Zero-filling for missing stratifications
    - Strict BAL - TVS Gap logic
    """
    
    # 1. SETUP & CLEANING
    scorecard_rows = []
    clean_df = town_df.copy()
    
    # Identify Columns
    col_strat = get_col_by_keyword(clean_df, ['Updated Stratification', 'Stratification'])
    col_bal_type = get_col_by_keyword(clean_df, ['BAL Store Type'])
    col_tvs_type = get_col_by_keyword(clean_df, ['TVS Store Type'])
    
    col_ind_s1 = get_col_by_keyword(clean_df, ['S1 Ind - F', 'S1 Ind Vistaar', 'S1 Ind'])
    col_bal_s1 = get_col_by_keyword(clean_df, ['BAL S1 Vol', 'BAL S1'])
    col_tvs_s1 = get_col_by_keyword(clean_df, ['TVS S1 Vol', 'TVS S1'])
    col_cr = get_col_by_keyword(clean_df, ['CR'])
    
    col_nature = get_col_by_keyword(clean_df, ['Nature of Intervention'])
    col_network = get_col_by_keyword(clean_df, ['Network Intervention'])
    
    # Filter Closed Stores
    col_closed = get_col_by_keyword(clean_df, ['Sub-Location', 'Highlight'])
    if col_closed:
        clean_df = clean_df[~clean_df[col_closed].astype(str).str.contains('closed', case=False, na=False)]

    if not col_strat: return None

    # 2. PROCESSING LOOP (Fixed Order)
    # We strictly iterate through this list to ensure Deep Rural always appears
    target_strats = ['Large Town', 'Small Town', 'Rural', 'Deep Rural']

    for strat in target_strats:
        # Filter data for this stratification
        strat_df = clean_df[clean_df[col_strat].astype(str).str.strip().str.lower() == strat.lower()].copy()
        
        # Helper to classify rows (BAL Perspective)
        def classify_row(row):
            val = str(row[col_bal_type]).upper() if col_bal_type else ""
            if any(x in val for x in ['MD', 'DEALER', 'BRANCH', 'BR']): return 'Primary'
            elif any(x in val for x in ['ASD', 'AD', 'SUB', 'REP']): return 'Secondary'
            else: return 'Vacant'

        if not strat_df.empty:
            strat_df['Category'] = strat_df.apply(classify_row, axis=1)
            pri_df = strat_df[strat_df['Category'] == 'Primary']
            sec_df = strat_df[strat_df['Category'] == 'Secondary']
            vac_df = strat_df[strat_df['Category'] == 'Vacant']
        else:
            # Create empty dataframes if stratification is missing
            pri_df = pd.DataFrame()
            sec_df = pd.DataFrame()
            vac_df = pd.DataFrame()

        # METRIC CALCULATOR
        def calc_metrics(df):
            if df.empty:
                return 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
                
            count = len(df) # BAL Count
            
            # TVS Counts (Count BOTH types for this row)
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
            # Volume Gap: Template Header says (TVS - BAL) so we calculate TVS - BAL
            v_gap = tvs - bal 
            cr = pd.to_numeric(df[col_cr], errors='coerce').mean() if col_cr else 0
            
            # Intervention Counts (Addition)
            add_p = 0
            add_s = 0
            if col_nature:
                add_p = len(df[df[col_nature].astype(str).str.contains('Branch|MD', case=False, na=False)])
                add_s = len(df[df[col_nature].astype(str).str.contains('ASD', case=False, na=False)])

            return count, t_pri, t_sec, ind, bal, tvs, ms, v_gap, cr, add_p, add_s

        # Calculate for each Category
        p_c, p_tp, p_ts, p_ind, p_bal, p_tvs, p_ms, p_vg, p_cr, p_ap, p_as = calc_metrics(pri_df)
        s_c, s_tp, s_ts, s_ind, s_bal, s_tvs, s_ms, s_vg, s_cr, s_ap, s_as = calc_metrics(sec_df)
        v_c, v_tp, v_ts, v_ind, v_bal, v_tvs, v_ms, v_vg, v_cr, v_ap, v_as = calc_metrics(vac_df)
        
        # --- STORE GAP LOGIC (Strictly BAL - TVS) ---
        # "calculate BAL - TVS even if... no issue if negative"
        gap_p = p_c - (p_tp + p_ts)
        gap_s = s_c - (s_tp + s_ts)
        gap_v = v_c - (v_tp + v_ts) # Even for vacant, we follow the math. 
        # (Usually 0 BAL - X TVS = Negative gap)
        
        # UNIQUE LOCATION GAP
        ulg_p = 0
        ulg_s = 0
        ulg_v = v_c # Vacant count implies opportunity
        
        # TOTALS
        tot_bal = p_c + s_c + v_c
        tot_tvs_p = p_tp + s_tp + v_tp
        tot_tvs_s = p_ts + s_ts + v_ts
        tot_gap = gap_p + gap_s + gap_v
        tot_ulg = ulg_p + ulg_s + ulg_v
        
        tot_add_p = p_ap + s_ap + v_ap
        tot_add_s = p_as + s_as + v_as
        
        # Network Count Post (Current + Adds)
        net_post_p = p_c + p_ap
        net_post_s = s_c + s_as
        net_post_tot = tot_bal + tot_add_p + tot_add_s
        
        # Unique Loc Gap Post (Current Gap - Adds)
        ulg_post = max(0, tot_ulg - (tot_add_p + tot_add_s))

        # --- APPEND ROWS ---
        # 1. Primary
        scorecard_rows.append([
            strat, "Pri Store", p_c, p_tp, p_ts, gap_p, ulg_p, p_ind, p_bal, p_ms, p_tvs, p_vg, p_cr, 
            p_ap if p_ap > 0 else "", p_as if p_as > 0 else "", "", "", net_post_p, ""
        ])
        # 2. Secondary
        scorecard_rows.append([
            "", "ASD", s_c, s_tp, s_ts, gap_s, ulg_s, s_ind, s_bal, s_ms, s_tvs, s_vg, s_cr, 
            s_ap if s_ap > 0 else "", s_as if s_as > 0 else "", "", "", net_post_s, ""
        ])
        # 3. Vacant
        scorecard_rows.append([
            "", "Vacant", v_c, v_tp, v_ts, gap_v, ulg_v, v_ind, v_bal, v_ms, v_tvs, v_vg, "", 
            v_ap if v_ap > 0 else "", v_as if v_as > 0 else "", "", "", "", ulg_post if ulg_v > 0 else ""
        ])
        # 4. Total
        scorecard_rows.append([
            "", "Total", tot_bal, tot_tvs_p, tot_tvs_s, tot_gap, tot_ulg, 
            (p_ind+s_ind+v_ind), (p_bal+s_bal+v_bal), "", (p_tvs+s_tvs+v_tvs), (p_vg+s_vg+v_vg), "", 
            tot_add_p, tot_add_s, "", "", net_post_tot, ulg_post
        ])
        scorecard_rows.append([""]*19)

    df_scorecard = pd.DataFrame(scorecard_rows)

    # 3. NETWORK PLAN (Sheet 2)
    if col_network:
        intervention_df = clean_df[clean_df[col_network].notna()].copy()
        keep_cols = []
        for c in ['Location', 'Stratification', 'TVS Store Type', 'BAL Store Type', 'S1 Ind', 'Nature', 'Network', 'Remarks']:
            found = get_col_by_keyword(clean_df, [c])
            if found: keep_cols.append(found)
        df_intervention = intervention_df[keep_cols] if not intervention_df.empty else pd.DataFrame()
    else:
        df_intervention = pd.DataFrame()

    # 4. EXCEL WRITING
    output = io.BytesIO()
    workbook = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook_obj = workbook.book

    # Styles
    fmt_header = workbook_obj.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D9E1F2', 'border': 1})
    fmt_sub = workbook_obj.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D9E1F2', 'border': 1, 'font_size': 9})
    fmt_simple = workbook_obj.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1})

    # Sheet 1
    if not df_scorecard.empty:
        df_scorecard.to_excel(workbook, sheet_name='Scorecard', startrow=3, header=False, index=False)
        ws1 = workbook.sheets['Scorecard']
        ws1.write(0, 1, town_name, workbook_obj.add_format({'bold': True, 'font_size': 14}))
        
        headers = ["Stratification", "# Store Count", "BAL", "TVS", "", "Store Gap", "Unique Location Gap", 
                   "IND S1", "S1 BAL Vol", "BAL MS", "S1 TVS Vol", "Vol Gap\n(TVS-BAL)", "CR", 
                   "Addition", "", "Reduction", "", "BAL Network Count\n@ UP 2.0", "Unique Location Gap\npost appointment"]
        
        for i, h in enumerate(headers):
            if h == "TVS": ws1.merge_range(1, 3, 1, 4, h, fmt_header)
            elif h == "Addition": ws1.merge_range(1, 13, 1, 14, h, fmt_header)
            elif h == "Reduction": ws1.merge_range(1, 15, 1, 16, h, fmt_header)
            elif h != "": ws1.write(1, i, h, fmt_header)

        for col in [3, 13, 15]: ws1.write(2, col, "Primary", fmt_sub)
        for col in [4, 14, 16]: ws1.write(2, col, "Secondary", fmt_sub)
        for col in [0,1,2,5,6,7,8,9,10,11,12,17,18]: ws1.write(2, col, "", fmt_sub)
        ws1.set_column(0, 0, 15)
        ws1.set_column(7, 12, 12)

    # Sheet 2
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
st.title("📊 Master Scorecard Generator (Fixed Logic)")
st.markdown("""
**Configuration:**
* **Stratification:** Forces Large -> Small -> Rural -> Deep Rural (even if missing).
* **Gap Calculation:** `BAL - TVS` (Strict).
* **Classification:** MD/Branch = Primary, ASD = Secondary.
""")

uploaded_file = st.file_uploader("Upload 'Accessibility Excel'", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=2)
        else:
            df = pd.read_excel(uploaded_file, header=2)
            
        df = clean_column_names(df)
        if 'Town' not in df.columns:
            uploaded_file.seek(0)
            if uploaded_file.name.endswith('.csv'): df = pd.read_csv(uploaded_file, header=1)
            else: df = pd.read_excel(uploaded_file, header=1)
            df = clean_column_names(df)

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
                    if data: zip_file.writestr(f"Scorecard_{town}.xlsx", data)
                    prog.progress((i + 1) / len(unique_towns))
            
            st.download_button("📥 Download ZIP", zip_buffer.getvalue(), "Scorecards_Final.zip", "application/zip")
            
    except Exception as e:
        st.error(f"Error: {e}")
