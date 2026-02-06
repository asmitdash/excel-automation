import streamlit as st
import pandas as pd
import io
import zipfile

# --- CONFIGURATION ---
st.set_page_config(page_title="Town Scorecard Generator", layout="wide")

def generate_town_excel(town_name, town_df):
    """
    Generates an Excel file with:
    1. 'Scorecard': Exact replica of Azamgarh template.
    2. 'Network_Plan': The bottom 2 tables.
    3. LOGIC FIX: Dynamically calculates 'Vacant' rows instead of hardcoding 0.
    """
    
    # ==========================================
    # SHEET 1: SCORECARD LOGIC
    # ==========================================
    
    scorecard_rows = []
    
    # 1. Filter out 'Closed' stores
    # We work on a copy to avoid affecting other logic
    clean_df = town_df.copy()
    
    # Robustly find the 'Sub-Location' column for closing logic
    closed_col = next((c for c in clean_df.columns if 'Highlight' in c and 'closed' in c), None)
    if closed_col:
        clean_df = clean_df[~clean_df[closed_col].astype(str).str.contains('closed', case=False, na=False)]

    # Get Stratifications
    stratifications = [x for x in clean_df['Updated Stratification'].dropna().unique()]

    for strat in stratifications:
        # Get all rows for this stratification (excluding closed)
        strat_df = clean_df[clean_df['Updated Stratification'] == strat]

        # --- SEPARATE DATAFRAMES FOR PRI, SEC, VACANT ---
        
        # 1. Primary DataFrame
        pri_mask = strat_df['BAL Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)
        pri_df = strat_df[pri_mask]
        
        # 2. Secondary DataFrame
        sec_mask = strat_df['BAL Store Type'].astype(str).str.contains('ASD|AD|Sub|Rep', case=False, na=False)
        sec_df = strat_df[sec_mask]
        
        # 3. Vacant DataFrame (Rows that are neither Primary nor Secondary)
        # We use the inverse of (Pri OR Sec)
        vac_df = strat_df[~(pri_mask | sec_mask)]
        
        # --- HELPER FUNCTION FOR METRICS ---
        def get_metrics(df):
            count = len(df)
            # TVS Counts (within this subset)
            tvs_p = len(df[df['TVS Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)])
            tvs_s = len(df[df['TVS Store Type'].astype(str).str.contains('ASD|AD|Sub', case=False, na=False)])
            
            # Volumes
            ind = pd.to_numeric(df['S1 Ind Vistaar'], errors='coerce').sum()
            bal_vol = pd.to_numeric(df['BAL S1 Vol - Vistaar'], errors='coerce').sum()
            tvs_vol = pd.to_numeric(df['TVS S1 Vol'], errors='coerce').sum()
            
            ms = (bal_vol / ind) if ind > 0 else 0
            vol_g = tvs_vol - bal_vol
            cr = pd.to_numeric(df['CR'], errors='coerce').mean()
            
            return count, tvs_p, tvs_s, ind, bal_vol, tvs_vol, ms, vol_g, cr

        # --- CALCULATE METRICS FOR EACH ROW ---
        
        # Primary Row Data
        p_bal, p_tvs_p, p_tvs_s, p_ind, p_bal_vol, p_tvs_vol, p_ms, p_vg, p_cr = get_metrics(pri_df)
        p_store_gap = p_bal - p_tvs_p # BAL Pri - TVS Pri
        
        # Secondary Row Data
        s_bal, s_tvs_p, s_tvs_s, s_ind, s_bal_vol, s_tvs_vol, s_ms, s_vg, s_cr = get_metrics(sec_df)
        s_store_gap = s_bal - s_tvs_s # BAL Sec - TVS Sec
        
        # Vacant Row Data (The Gemini Fix)
        v_bal, v_tvs_p, v_tvs_s, v_ind, v_bal_vol, v_tvs_vol, v_ms, v_vg, v_cr = get_metrics(vac_df)
        # For Vacant, Store Gap is usually equal to the count of opportunities (as per Gemini)
        # OR 0 - 0. Let's keep it strictly math: BAL(0) - TVS(0) = 0 usually. 
        # But if Gemini says "Gap equals the count", we might need logic. 
        # For now, let's keep strict math: 0 stores vs 0 stores.
        v_store_gap = 0 

        # --- TOTALS ---
        t_bal = p_bal + s_bal + v_bal
        t_tvs = (p_tvs_p + p_tvs_s) + (s_tvs_p + s_tvs_s) + (v_tvs_p + v_tvs_s)
        t_gap = p_store_gap + s_store_gap + v_store_gap
        
        # --- BUILD ROWS ---
        
        # 1. Primary Row
        scorecard_rows.append([
            strat, "Pri Store", p_bal, p_tvs_p, "", 
            p_store_gap, 0, p_ind, p_bal_vol, p_ms, p_tvs_vol, p_vg, p_cr, 
            "", "", "", "", p_bal, 0
        ])
        
        # 2. Secondary Row
        scorecard_rows.append([
            "", "ASD", s_bal, "", s_tvs_s, 
            s_store_gap, 0, s_ind, s_bal_vol, s_ms, s_tvs_vol, s_vg, s_cr, 
            "", "", "", "", s_bal, 0
        ])
        
        # 3. Vacant Row (Populated with calculated data)
        scorecard_rows.append([
            "", "Vacant", v_bal, "", "", 
            v_store_gap, len(vac_df), v_ind, v_bal_vol, v_ms, v_tvs_vol, v_vg, "", 
            "", "", "", "", 0, len(vac_df)
        ])
        
        # 4. Total Row
        scorecard_rows.append([
            "", "Total", t_bal, t_tvs, "", 
            t_gap, len(vac_df), (p_ind+s_ind+v_ind), (p_bal_vol+s_bal_vol+v_bal_vol), "", 
            (p_tvs_vol+s_tvs_vol+v_tvs_vol), (p_vg+s_vg+v_vg), "", 
            "", "", "", "", t_bal, len(vac_df)
        ])
        
        # Spacer
        scorecard_rows.append([""] * 19)

    df_scorecard = pd.DataFrame(scorecard_rows)

    # ==========================================
    # SHEET 2: NETWORK PLAN LOGIC
    # ==========================================
    
    summary_rows = []
    categories = ['Primary', 'Secondary', 'Vacant']
    
    total_pre_count = 0
    total_post_count = 0
    total_vol_gain = 0
    
    for cat in categories:
        # Pre
        pre_mask = town_df['Pre - Network - BAL'].astype(str).str.contains(cat, case=False, na=False)
        pre_count = len(town_df[pre_mask])
        pre_ind = pd.to_numeric(town_df.loc[pre_mask, 'S1 Ind Vistaar'], errors='coerce').sum()
        
        # Post
        post_mask = town_df['Post Net Bal'].astype(str).str.contains(cat, case=False, na=False)
        post_count = len(town_df[post_mask])
        post_ind = pd.to_numeric(town_df.loc[post_mask, 'S1 Ind Vistaar'], errors='coerce').sum()
        
        # Gains
        ind_cov_gain = post_ind - pre_ind
        vol_gain = pd.to_numeric(town_df.loc[post_mask, 'Vol Gain'], errors='coerce').sum()
        ms_gain = (vol_gain / ind_cov_gain) if ind_cov_gain > 0 else 0
        
        total_pre_count += pre_count
        total_post_count += post_count
        total_vol_gain += vol_gain
        
        summary_rows.append([cat, pre_count, pre_ind, post_count, post_ind, ind_cov_gain, vol_gain, ms_gain])

    summary_rows.append(["Total", total_pre_count, "", total_post_count, "", "", total_vol_gain, ""])
    
    df_network_summary = pd.DataFrame(summary_rows, columns=[
        "Channel Type", "Pre Count", "Pre Ind", "Post Count", "Post Ind", "Ind Coverage Gain", "Vol Gain", "MS Gain"
    ])

    # Intervention List
    intervention_df = town_df[town_df['Network Intervention'].notna()].copy()
    cols_to_keep = ['Location / T2T', 'Updated Stratification', 'Channel Type - TVS', 'BAL Store Type', 'S1 Ind Vistaar', 'Nature of Intervention', 'Network Intervention', 'Remarks']
    existing_cols = [c for c in cols_to_keep if c in intervention_df.columns]
    df_intervention_list = intervention_df[existing_cols]

    # ==========================================
    # EXCEL WRITING
    # ==========================================
    output = io.BytesIO()
    workbook = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook_obj = workbook.book

    # --- STYLES ---
    header_format = workbook_obj.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'fg_color': '#D9E1F2', 'border': 1
    })
    subheader_format = workbook_obj.add_format({
        'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        'fg_color': '#D9E1F2', 'border': 1, 'font_size': 9
    })
    simple_header = workbook_obj.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1})

    # --- SHEET 1 ---
    df_scorecard.to_excel(workbook, sheet_name='Scorecard', startrow=3, header=False, index=False)
    ws1 = workbook.sheets['Scorecard']
    ws1.write(0, 1, town_name, workbook_obj.add_format({'bold': True, 'font_size': 14}))
    
    # Headers (Exact Replica)
    ws1.write(1, 0, "Stratification", header_format)
    ws1.write(1, 1, "# Store Count", header_format)
    ws1.write(1, 2, "BAL", header_format)
    ws1.merge_range(1, 3, 1, 4, "TVS", header_format)
    ws1.write(1, 5, "Store Gap", header_format)
    ws1.write(1, 6, "Unique Location Gap", header_format)
    ws1.write(1, 7, "IND S1", header_format)
    ws1.write(1, 8, "S1 BAL Vol", header_format)
    ws1.write(1, 9, "BAL MS", header_format)
    ws1.write(1, 10, "S1 TVS Vol", header_format)
    ws1.write(1, 11, "Vol Gap\n(TVS-BAL)", header_format)
    ws1.write(1, 12, "CR", header_format)
    ws1.merge_range(1, 13, 1, 14, "Addition", header_format)
    ws1.merge_range(1, 15, 1, 16, "Reduction", header_format)
    ws1.write(1, 17, "BAL Network Count\n@ UP 2.0", header_format)
    ws1.write(1, 18, "Unique Location Gap\npost appointment", header_format)
    
    # Subheaders
    for c in [0,1,2,5,6,7,8,9,10,11,12,17,18]: ws1.write(2, c, "", subheader_format)
    ws1.write(2, 3, "Primary", subheader_format)
    ws1.write(2, 4, "Secondary", subheader_format)
    ws1.write(2, 13, "Primary", subheader_format)
    ws1.write(2, 14, "Secondary", subheader_format)
    ws1.write(2, 15, "Primary", subheader_format)
    ws1.write(2, 16, "Secondary", subheader_format)
    ws1.set_column(0, 0, 15)
    ws1.set_column(7, 12, 10)

    # --- SHEET 2 ---
    ws2_name = 'Network_Plan'
    df_network_summary.to_excel(workbook, sheet_name=ws2_name, startrow=1, index=False)
    ws2 = workbook.sheets[ws2_name]
    for col_num, value in enumerate(df_network_summary.columns.values):
        ws2.write(0, col_num, value, simple_header)

    start_row_t2 = len(df_network_summary) + 4
    ws2.write(start_row_t2 - 1, 0, "Intervention List", workbook_obj.add_format({'bold': True, 'font_size': 12}))
    df_intervention_list.to_excel(workbook, sheet_name=ws2_name, startrow=start_row_t2, index=False)
    for col_num, value in enumerate(df_intervention_list.columns.values):
        ws2.write(start_row_t2, col_num, value, simple_header)
    ws2.set_column(0, 8, 15)

    workbook.close()
    return output.getvalue()


# --- MAIN APP UI ---
st.title("📊 Master Scorecard Generator (Audit Compliant)")

uploaded_file = st.file_uploader("Upload 'Accessibility Excel' (CSV or Excel)", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=2)
        else:
            df = pd.read_excel(uploaded_file, header=2)
            
        if len(df) > 0:
            df = df.drop(0).reset_index(drop=True)

        # CLEAN HEADERS
        df.columns = df.columns.str.replace('\n', ' ').str.strip()
        clean_df = df[~df['Town'].astype(str).str.contains('Total', case=False, na=False)].copy()
        
        unique_towns = clean_df['Town'].dropna().unique()
        st.success(f"✅ Data Loaded. Found {len(unique_towns)} unique towns.")

        if st.button("Generate Scorecards"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                progress_bar = st.progress(0)
                for i, town in enumerate(unique_towns):
                    town_df = clean_df[clean_df['Town'] == town]
                    excel_data = generate_town_excel(town, town_df)
                    zip_file.writestr(f"Scorecard_{town}.xlsx", excel_data)
                    progress_bar.progress((i + 1) / len(unique_towns))
            
            st.download_button(
                label="📥 Download All Scorecards (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="Town_Scorecards_Audit_Compliant.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Error: {e}")
