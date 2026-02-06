import streamlit as st
import pandas as pd
import io
import zipfile

# --- CONFIGURATION ---
st.set_page_config(page_title="Town Scorecard Generator", layout="wide")

def generate_town_excel(town_name, town_df):
    """
    Generates an Excel file for a specific town with exact formatting,
    colors, and merged headers using XlsxWriter.
    Includes logic to EXCLUDE 'Closed' stores from counts.
    """
    
    # 1. PREPARE DATA ROWS
    data_rows = []
    
    # Get Stratifications present in this town
    stratifications = [x for x in town_df['Updated Stratification'].dropna().unique()]

    for strat in stratifications:
        strat_df = town_df[town_df['Updated Stratification'] == strat]

        # --- FILTER LOGIC (NEW): Exclude Closed Stores ---
        # We assume 'Sub-Location' column contains the text "(Highlighted Red - closed)"
        # We filter OUT rows where this column contains "closed" (case insensitive)
        active_strat_df = strat_df[~strat_df['Sub-Location\n(Highlighted Red - closed)'].astype(str).str.contains('closed', case=False, na=False)]

        # --- CALCULATIONS (Using Active Stores Only) ---
        
        # COUNTS
        bal_pri = len(active_strat_df[active_strat_df['BAL Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)])
        bal_sec = len(active_strat_df[active_strat_df['BAL Store Type'].astype(str).str.contains('ASD|AD|Sub|Rep', case=False, na=False)])
        tvs_pri = len(active_strat_df[active_strat_df['TVS Store Type'].astype(str).str.contains('MD|Dealer', case=False, na=False)])
        tvs_sec = len(active_strat_df[active_strat_df['TVS Store Type'].astype(str).str.contains('ASD|AD|Sub', case=False, na=False)])
        
        # VOLUMES (Volumes should likely still be summed for the whole town, or just active? 
        # Usually Strategy is based on TOTAL potential, so we use the FULL strat_df for Volumes)
        ind_s1 = pd.to_numeric(strat_df['S1 Ind Vistaar'], errors='coerce').sum()
        bal_s1_vol = pd.to_numeric(strat_df['BAL S1 Vol - Vistaar'], errors='coerce').sum()
        tvs_s1_vol = pd.to_numeric(strat_df['TVS S1 Vol'], errors='coerce').sum()
        
        # METRICS
        bal_ms = (bal_s1_vol / ind_s1) if ind_s1 > 0 else 0
        vol_gap = tvs_s1_vol - bal_s1_vol
        cr_val = pd.to_numeric(strat_df['CR'], errors='coerce').mean()

        # GAPS
        store_gap_pri = bal_pri - tvs_pri
        store_gap_sec = bal_sec - tvs_sec
        
        # --- ROW GENERATION ---
        
        # 1. Primary Store Row
        data_rows.append([
            strat, "Pri Store", bal_pri, tvs_pri, "", store_gap_pri, 0, 
            ind_s1, bal_s1_vol, bal_ms, tvs_s1_vol, vol_gap, cr_val,
            "", "", "", "", bal_pri, 0
        ])
        
        # 2. Secondary Store Row
        data_rows.append([
            "", "ASD", bal_sec, "", tvs_sec, store_gap_sec, 0, 
            "", "", "", "", "", "",
            "", "", "", "", bal_sec, 0
        ])
        
        # 3. Vacant Row
        data_rows.append([
            "", "Vacant", 0, "", "", 0, 0, 
            "", "", "", "", "", "",
            "", "", "", "", 0, 0
        ])
        
        # 4. Total Row
        data_rows.append([
            "", "Total", (bal_pri + bal_sec), (tvs_pri + tvs_sec), "", (store_gap_pri + store_gap_sec), 0, 
            "", "", "", "", "", "",
            "", "", "", "", (bal_pri + bal_sec), 0
        ])
        
        # Blank row between stratifications
        data_rows.append([""] * 19)

    # Create DataFrame
    df_output = pd.DataFrame(data_rows)

    # 2. WRITE TO EXCEL WITH FORMATTING
    output = io.BytesIO()
    workbook = pd.ExcelWriter(output, engine='xlsxwriter')
    df_output.to_excel(workbook, sheet_name='Scorecard', startrow=3, header=False, index=False)
    
    worksheet = workbook.sheets['Scorecard']
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
    
    # Write Town Name
    worksheet.write(0, 1, town_name, workbook_obj.add_format({'bold': True, 'font_size': 14}))

    # --- WRITE HEADERS ---
    worksheet.write(1, 0, "Stratification", header_format)
    worksheet.write(1, 1, "# Store Count", header_format)
    worksheet.write(1, 2, "BAL", header_format)
    worksheet.merge_range(1, 3, 1, 4, "TVS", header_format)
    worksheet.write(1, 5, "Store Gap", header_format)
    worksheet.write(1, 6, "Unique Location Gap", header_format)
    worksheet.write(1, 7, "IND S1", header_format)
    worksheet.write(1, 8, "S1 BAL Vol", header_format)
    worksheet.write(1, 9, "BAL MS", header_format)
    worksheet.write(1, 10, "S1 TVS Vol", header_format)
    worksheet.write(1, 11, "Vol Gap\n(TVS-BAL)", header_format)
    worksheet.write(1, 12, "CR", header_format)
    worksheet.merge_range(1, 13, 1, 14, "Addition", header_format)
    worksheet.merge_range(1, 15, 1, 16, "Reduction", header_format)
    worksheet.write(1, 17, "BAL Network Count\n@ UP 2.0", header_format)
    worksheet.write(1, 18, "Unique Location Gap\npost appointment", header_format)

    # Row 3 (Sub Headers)
    worksheet.write(2, 0, "", subheader_format)
    worksheet.write(2, 1, "", subheader_format)
    worksheet.write(2, 2, "", subheader_format)
    worksheet.write(2, 3, "Primary", subheader_format)
    worksheet.write(2, 4, "Secondary", subheader_format)
    worksheet.write(2, 5, "", subheader_format)
    worksheet.write(2, 6, "", subheader_format)
    worksheet.write(2, 7, "", subheader_format)
    worksheet.write(2, 8, "", subheader_format)
    worksheet.write(2, 9, "", subheader_format)
    worksheet.write(2, 10, "", subheader_format)
    worksheet.write(2, 11, "", subheader_format)
    worksheet.write(2, 12, "", subheader_format)
    worksheet.write(2, 13, "Primary", subheader_format)
    worksheet.write(2, 14, "Secondary", subheader_format)
    worksheet.write(2, 15, "Primary", subheader_format)
    worksheet.write(2, 16, "Secondary", subheader_format)
    worksheet.write(2, 17, "", subheader_format)
    worksheet.write(2, 18, "", subheader_format)

    # Adjust Column Widths
    worksheet.set_column(0, 0, 15)
    worksheet.set_column(1, 1, 12)
    worksheet.set_column(2, 4, 8)
    worksheet.set_column(7, 12, 10)
    
    workbook.close()
    return output.getvalue()


# --- MAIN APP UI ---
st.title("📊 Master Scorecard Generator (Fixed Closed Stores)")

uploaded_file = st.file_uploader("Upload 'Accessibility Excel' (CSV or Excel)", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=2)
        else:
            df = pd.read_excel(uploaded_file, header=2)
            
        if len(df) > 0:
            df = df.drop(0).reset_index(drop=True)

        # Clean Headers (Remove Newlines and Spaces)
        df.columns = df.columns.str.strip() # Remove leading/trailing spaces
        
        # Clean 'Total' rows
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
                file_name="Town_Scorecards_Final.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"Error: {e}")
