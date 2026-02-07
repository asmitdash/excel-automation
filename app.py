import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile

st.set_page_config(page_title="Network Matrix Generator", layout="wide")

st.title("📂 Multi-Town Network Matrix Generator")
st.markdown("Upload your Master Excel. The app will generate a formatted **Table 1** for every town and provide a ZIP file.")

uploaded_file = st.file_uploader("Upload Master Excel", type=['xlsx'])

def classify_type(val):
    val = str(val).strip().upper()
    if val in ['MD', 'BRAND', 'REP BY BR', 'REP BY MD']: return 'Pri Store'
    if val in ['ASD', 'REP BY ASD']: return 'ASD'
    if val in ['NAN', 'NA', 'CLOSED', 'BLANK', '']: return 'Vacant'
    return 'Vacant'

def generate_table_1(df_town):
    stratifications = ['Large Town', 'Small Town', 'Rural', 'Deep Rural']
    categories = ['Pri Store', 'ASD', 'Vacant']
    
    df_town = df_town.copy()
    df_town['BAL_Cat'] = df_town['BAL Store Type'].apply(classify_type)
    df_town['TVS_Cat'] = df_town['TVS Store Type'].apply(classify_type)
    df_town['Pre_Cat'] = df_town['Pre - Network - BAL'].apply(classify_type)

    rows = []
    for strat in stratifications:
        for cat in categories:
            sub = df_town[(df_town['Updated Stratification'] == strat) & (df_town['BAL_Cat'] == cat)]
            
            # Base Counts
            bal = len(sub[sub['BAL_Cat'] != 'Vacant'])
            tvs_pri = len(sub[sub['TVS_Cat'] == 'Pri Store'])
            tvs_sec = len(sub[sub['TVS_Cat'] == 'ASD'])
            
            # Unique Location Gap (Logic: TVS present, BAL Vacant)
            loc_gap_sub = df_town[(df_town['Updated Stratification'] == strat) & 
                                  (df_town['BAL_Cat'] == 'Vacant') & 
                                  (df_town['TVS_Cat'] != 'Vacant')]
            u_gap = loc_gap_sub['Location / T2T'].nunique() if cat == 'Vacant' else 0

            # Interventions (Logic based on Pre-Network status)
            inter_sub = df_town[(df_town['Updated Stratification'] == strat) & (df_town['Pre_Cat'] == cat)]
            add_p, add_s, red_p, red_s = 0, 0, 0, 0

            for _, r in inter_sub.iterrows():
                act = str(r['Network Intervention']).strip().upper()
                post = classify_type(r['Post Net Bal'])
                pre = classify_type(r['Pre - Network - BAL'])
                if act in ['YES', 'REPLACEMENT']:
                    if post == 'Pri Store': add_p += 1
                    if post == 'ASD': add_s += 1
                    if act == 'REPLACEMENT' or post == 'Vacant':
                        if pre == 'Pri Store': red_p += 1
                        if pre == 'ASD': red_s += 1

            rows.append({
                "Stratification": strat, "# Store Count": cat, "BAL": bal, "TVS: Pri": tvs_pri,
                "TVS: Sec": tvs_sec, "Unique Loc Gap": u_gap, "IND S1": sub['S1 Ind - F Vistaa'].sum(),
                "S1 BAL Vol": sub['BAL S1 Vol - Vistaa'].sum(), "S1 TVS Vol": sub['TVS S1 Vol Basis MS'].sum(),
                "Add: Pri": add_p, "Add: Sec": add_s, "Red: Pri": red_p, "Red: Sec": red_s
            })

    res = pd.DataFrame(rows)
    # Applying Formulas
    res['Store Gap'] = (res['TVS: Pri'] + res['TVS: Sec']) - res['BAL']
    res['BAL MS'] = (res['S1 BAL Vol'] / res['IND S1']).fillna(0)
    res['Vol Gap'] = res['S1 TVS Vol'] - res['S1 BAL Vol']
    res['CR'] = (res['S1 TVS Vol'] / res['S1 BAL Vol']).replace([np.inf, -np.inf], 0).fillna(0)
    res['BAL @ 2.0'] = (res['Add: Pri'] + res['Add: Sec'] + res['BAL']) - (res['Red: Pri'] + res['Red: Sec'])
    res['Post Gap'] = res['Unique Loc Gap'] - res['Store Gap']
    
    return res

if uploaded_file:
    master_df = pd.read_excel(uploaded_file)
    towns = master_df['Town'].unique()
    
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        for town in towns:
            town_data = master_df[master_df['Town'] == town]
            table1 = generate_table_1(town_data)
            
            # Excel formatting using XlsxWriter
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                table1.to_excel(writer, sheet_name='Table 1', index=False, startrow=1)
                workbook = writer.book
                worksheet = writer.sheets['Table 1']
                
                # Formatting Styles
                header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center'})
                strat_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bold': True})
                num_fmt = workbook.add_format({'border': 1, 'align': 'center'})
                pct_fmt = workbook.add_format({'border': 1, 'num_format': '0%', 'align': 'center'})
                
                # Write Headers
                for col_num, value in enumerate(table1.columns.values):
                    worksheet.write(1, col_num, value, header_fmt)
                
                # Merge Stratification Cells (Exactly like Azamgarh file)
                for i in range(0, len(table1), 3):
                    worksheet.merge_range(i+2, 0, i+4, 0, table1.iloc[i, 0], strat_fmt)
                
                # Apply number and percentage formats
                worksheet.set_column('A:B', 15)
                worksheet.set_column('C:S', 10, num_fmt)
                worksheet.set_column('J:J', 10, pct_fmt) # BAL MS column
                
            zip_file.writestr(f"{town}_Table_1.xlsx", output.getvalue())

    st.success(f"Processed {len(towns)} towns!")
    st.download_button(
        label="📥 Download All Towns (ZIP)",
        data=zip_buffer.getvalue(),
        file_name="Network_Matrices.zip",
        mime="application/zip"
    )
    
