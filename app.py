import streamlit as st
import pandas as pd
import io
import zipfile
import xlsxwriter

# --- CONFIGURATION ---
st.set_page_config(page_title="Town Scorecard Generator", layout="wide")

def clean_column_names(df):
    """Standardizes column names: removes newlines, trims spaces."""
    df.columns = df.columns.str.replace('\n', ' ').str.strip()
    return df

def get_col_by_keyword(df, keywords):
    """Finds a column containing one of the keywords (case-insensitive)."""
    for col in df.columns:
        for kw in keywords:
            if kw.lower() in col.lower():
                return col
    return None

def classify_row_bucket(row, col_bal_type):
    """
    Classifies a row into 'Primary', 'Secondary', or 'Vacant' 
    based on BAL Store Type.
    """
    val = str(row[col_bal_type]).upper() if col_bal_type and pd.notna(row[col_bal_type]) else ""
    
    # HARD CODED LOGIC FOR BUCKETS
    if any(x in val for x in ['MD', 'DEALER', 'BRANCH', 'BR']):
        return 'Primary'
    elif any(x in val for x in ['ASD', 'AD', 'SUB', 'REP']):
        return 'Secondary'
    else:
        # If it's not Pri or Sec, it's Vacant (assuming closed stores are already filtered)
        return 'Vacant'

def generate_town_excel(town_name, town_df):
    
    # 1. SETUP & COLUMN MAPPING
    scorecard_rows = []
    
    # Map Columns dynamically to handle slight naming variations
    col_strat = get_col_by_keyword(town_df, ['Updated Stratification', 'Stratification'])
    col_bal_type = get_col_by_keyword(town_df, ['BAL Store Type'])
    col_tvs_type = get_col_by_keyword(town_df, ['TVS Store Type'])
    
    col_ind_s1 = get_col_by_keyword(town_df, ['S1 Ind - F', 'S1 Ind Vistaar', 'S1 Ind'])
    col_bal_s1 = get_col_by_keyword(town_df, ['BAL S1 Vol', 'BAL S1'])
    col_tvs_s1 = get_col_by_keyword(town_df, ['TVS S1 Vol', 'TVS S1'])
    col_cr = get_col_by_keyword(town_df, ['CR'])
    
    col_nature = get_col_by_keyword(town_df, ['Nature of Intervention'])
    col_network = get_col_by_keyword(town_df, ['Network Intervention'])
    col_location = get_col_by_keyword(town_df, ['Location / T2T', 'Location'])
    col_remarks = get_col_by_keyword(town_df, ['Remarks'])
    
    # Filter Closed Stores (HARD CODED SAFETY CHECK)
    col_closed = get_col_by_keyword(town_df, ['Sub-Location', 'Highlight'])
    clean_df = town_df.copy()
    if col_closed:
        clean_df = clean_df[~clean_df[col_closed].astype(str).str.contains('closed', case=False, na=False)]

    if not col_strat:
        return None

    # 2. STRATIFICATION LOOP (Fixed Order)
    # The pivot table shows these categories. We force this order.
    target_strats = ['Large Town', 'Small Town', 'Rural', 'Deep Rural']

    for strat in target_strats:
        # Filter for this stratification
        strat_df = clean_df[clean_df[col_strat].astype(str).str.strip().str.lower() == strat.lower()].copy()
        
        # Add Bucket Column
        if not strat_df.empty:
            strat_df['Bucket'] = strat_df.apply(lambda r: classify_row_bucket(r, col_bal_type), axis=1)
        else:
            strat_df['Bucket'] = [] # Empty

        # Helper to calculate metrics for a specific bucket
        def get_bucket_metrics(bucket_name):
            if strat_df.empty:
                 return [0]*11 # Return zeros if empty
            
            bucket_df = strat_df[strat_df['Bucket'] == bucket_name]
            
            if bucket_df.empty:
                return [0]*11

            # 1. COUNTS
            count_bal = len(bucket_df)
            
            # TVS Counts (Primary vs Secondary logic)
            count_tvs_pri = 0
            count_tvs_sec = 0
            if col_tvs_type:
                count_tvs_pri = len(bucket_df[bucket_df[col_tvs_type].astype(str).str.upper().str.contains('MD|DEALER|BRANCH')])
                count_tvs_sec = len(bucket_df[bucket_df[col_tvs_type].astype(str).str.upper().str.contains('ASD|AD|SUB')])

            # 2. VOLUMES
            vol_ind_s1 = pd.to_numeric(bucket_df[col_ind_s1], errors='coerce').sum() if col_ind_s1 else 0
            vol_bal_s1 = pd.to_numeric(bucket_df[col_bal_s1], errors='coerce').sum() if col_bal_s1 else 0
            vol_tvs_s1 = pd.to_numeric(bucket_df[col_tvs_s1], errors='coerce').sum() if col_tvs_s1 else 0

            # 3. METRICS
            ms = (vol_bal_s1 / vol_ind_s1) if vol_ind_s1 > 0 else 0
            vol_gap = vol_tvs_s1 - vol_bal_s1
            cr = pd.to_numeric(bucket_df[col_cr], errors='coerce').mean() if col_cr else 0
            
            # 4. ADDITIONS (For Network Count)
            add_pri = 0
            add_sec = 0
            if col_nature:
                add_pri = len(bucket_df[bucket_df[col_nature].astype(str).str.upper().str.contains('BRANCH|MD')])
                add_sec = len(bucket_df[bucket_df[col_nature].astype(str).str.upper().str.contains('ASD')])

            return [count_bal, count_tvs_pri, count_tvs_sec, vol_ind_s1, vol_bal_s1, vol_tvs_s1, ms, vol_gap, cr, add_pri, add_sec]

        # GET METRICS FOR EACH ROW
        # Primary Row
        p_c, p_tp, p_ts, p_ind, p_bal, p_tvs, p_ms, p_vg, p_cr, p_ap, p_as = get_bucket_metrics('Primary')
        # Secondary Row
        s_c, s_tp, s_ts, s_ind, s_bal, s_tvs, s_ms, s_vg, s_cr, s_ap, s_as = get_bucket_metrics('Secondary')
        # Vacant Row
        v_c, v_tp, v_ts, v_ind, v_bal, v_tvs, v_ms, v_vg, v_cr, v_ap, v_as = get_bucket_metrics('Vacant')

        # --- CALCULATE GAPS & TOTALS ---
        
        # Store Gap: BAL - (TVS Pri + TVS Sec)
        gap_p = p_c - (p_tp + p_ts)
        gap_s = s_c - (s_tp + s_ts)
        gap_v = v_c - (v_tp + v_ts) # Even for vacant, use math
        
        # Unique Location Gap
        # Logic: Only for Vacant row does this equal the count. For others 0.
        ulg_p = 0
        ulg_s = 0
        ulg_v = v_c 

        # Totals
        tot_bal = p_c + s_c + v_c
        tot_tvs_p = p_tp + s_tp + v_tp
        tot_tvs_s = p_ts + s_ts + v_ts
        tot_gap = gap_p + gap_s + gap_v
        tot_ulg = ulg_p + ulg_s + ulg_v
        
        # Network Count Post
        # Current BAL + Additions. (Reductions assumed 0 as not provided in logic)
        net_post_p = p_c + p_ap
        net_post_s = s_c + s_as
        net_post_tot = tot_bal + (p_ap + p_as) + (s_ap + s_as) + (v_ap + v_as)

        # Unique Loc Gap Post
        # Current Gap - Additions
        ulg_post = max(0, tot_ulg - ((p_ap + p_as) + (s_ap + s_as) + (v_ap + v_as)))

        # --- APPEND ROWS TO LIST ---
        # 1. Pri Store
        scorecard_rows.append([
            strat, "Pri Store", p_c, p_tp, p_ts, gap_p, ulg_p, p_ind, p_bal, p_ms, p_tvs, p_vg, p_cr, 
            p_ap if p_ap > 0 else "", p_as if p_as > 0 else "", "", "", net_post_p, ""
        ])
        # 2. ASD
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
            (p_ap+s_ap+v_ap), (p_as+s_as+v_as), "", "", net_post_tot, ulg_post
        ])
        # Spacer
        scorecard_rows.append([""] * 19)

    df_scorecard = pd.DataFrame(scorecard_rows)

    # 3. NETWORK PLAN (Intervention List)
    # Filter for rows with Network Intervention
    if col_network:
        intervention_df = clean_df[clean_df[col_network].notna()].copy()
        
        # Select specific columns for the output
        desired_cols = {
            col_location: 'Location / T2T',
            col_strat: 'Updated Stratification',
            col_tvs_type: 'TVS Store Type',
            col_bal_type: 'BAL Store Type',
            col_ind_s1: 'S1 Ind Vistaar',
            col_nature: 'Nature of Intervention',
            col_network: 'Network Intervention',
            col_remarks: 'Remarks'
        }
        
        # Build DataFrame dynamically based on available columns
        final_int_df = pd.DataFrame()
        for src_col, dest_col in desired_cols.items():
            if src_col:
                final_int_df[dest_col] = intervention_df[src_col]
            else:
                final_int_df[dest_col] = "" # Fill empty if missing
                
    else:
        final_int_df = pd.DataFrame()

    # 4. EXCEL GENERATION (XlsxWriter)
    output = io.BytesIO()
    workbook = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook_obj = workbook.book

    # Styles
    fmt_header = workbook_obj.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D9E1F2', 'border': 1})
    fmt_sub = workbook_obj.add_format({'bold': True, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D9E1F2', 'border': 1, 'font_size': 9})
    fmt_simple = workbook_obj.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1})

    # SHEET 1: SCORECARD
    if not df_scorecard.empty:
        df_scorecard.to_excel(workbook, sheet_name='Scorecard', startrow=3, header=False, index=False)
        ws1 = workbook.sheets['Scorecard']
        ws1.write(0, 1, town_name, workbook_obj.add_format({'bold': True, 'font_size': 14}))
        
        # Main Headers
        headers = ["Stratification", "# Store Count", "BAL", "TVS", "", "Store Gap", "Unique Location Gap", 
                   "IND S1", "S1 BAL Vol", "BAL MS", "S1 TVS Vol", "Vol Gap\n(TVS-BAL)", "CR", 
                   "Addition", "", "Reduction", "", "BAL Network Count\n@ UP 2.0", "Unique Location Gap\npost appointment"]
        
        for i, h in enumerate(headers):
            if h == "TVS": ws1.merge_range(1, 3, 1, 4, h, fmt_header)
            elif h == "Addition": ws1.merge_range(1, 13, 1, 14, h, fmt_header)
            elif h == "Reduction": ws1.merge_range(1, 15, 1, 16, h, fmt_header)
            elif h != "": ws1.write(1, i, h, fmt_header)

        # Sub Headers
        for col in [3, 13, 15]: ws1.write(2, col, "Primary", fmt_sub)
        for col in [4, 14, 16]: ws1.write(2, col, "Secondary", fmt_sub)
        for col in [0,1,2,5,6,7,8,9,10,11,12,17,18]: ws1.write(2, col, "", fmt_sub)
        
        ws1.set_column(0, 0, 15)
        ws1.set_column(7, 12, 12)

    # SHEET 2: NETWORK PLAN
    if not final_int_df.empty:
        ws2_name = 'Network_Plan'
        final_int_df.to_excel(workbook, sheet_name=ws2_name, startrow=1, index=False)
        ws2 = workbook.sheets[ws2_name]
        for i, col in enumerate(final_int_df.columns):
            ws2.write(0, i, col, fmt_simple)
        ws2.set_column(0, len(final_int_df.columns)-1, 15)

    workbook.close()
    return output.getvalue()


# --- MAIN APP UI ---
st.title("📊 Master Scorecard Generator (Pivot Logic)")
st.markdown("""
**Hard-Coded Logic Applied:**
* **Stratification Order:** Large -> Small -> Rural -> Deep Rural
* **Rows:** Pri Store (MD/Br), ASD (ASD/Sub), Vacant (Rest)
* **Gap Logic:** BAL - TVS
* **Metrics:** Aggregated from Accessibility File
""")

uploaded_file = st.file_uploader("Upload 'Accessibility Excel'", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        # Load File (try standard header rows)
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=2)
        else:
            df = pd.read_excel(uploaded_file, header=2)
            
        df = clean_column_names(df)
        
        # Robust check for Town column
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
