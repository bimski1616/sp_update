import streamlit as st
import pandas as pd
import re
import io

# -- SETTING PAGE --
st.set_page_config(page_title="Project Master Control Tower", layout="wide", page_icon="🚀")

# --- SEMUA LOGIC ASLI (HARAM DIRUBAH) ---
def super_clean(val):
    if pd.isna(val): return ""
    return re.sub(r'\s+', '', str(val)).upper()

def clean_period_final(val):
    if pd.isna(val): return ""
    val = super_clean(val)
    match = re.search(r'M-?0?(\d+)', val)
    return f"M{match.group(1)}" if match else val

def convert_points_der(val):
    if pd.isna(val) or val == 0 or val == '0' or val == "": return 0.0
    val_str = str(val).strip()
    if ':' in val_str:
        parts = val_str.split(':')
        try:
            h = float(parts[0]); m = float(parts[1])
            return h + (m / 10.0) # Rumus: 1:08 -> 1.8
        except: return 0.0
    try:
        return float(str(val).replace(',', '.'))
    except: return 0.0

# --- UI STREAMLIT ---
st.title("🚀 Project Master Control Tower")
st.markdown("Update **Total Points (VEB)** dengan Audit Trail & Filter Detail.")

# --- SIDEBAR UNTUK SEARCH & FILTER ---
st.sidebar.header("🔍 Filter & Search")
search_query = st.sidebar.text_input("Cari Uniq_ID / User / Group", "").upper()

# --- SECTION 1: UPLOAD Master & Sources ---
col_master, col_src = st.columns([1, 2])

with col_master:
    st.subheader("1. File Master")
    file_master = st.file_uploader("Upload SP_New.xlsx", type=['xlsx'])

with col_src:
    st.subheader("2. File Update")
    c1, c2, c3 = st.columns(3)
    up_ioh = c1.file_uploader("IOH", type=['xlsx'])
    up_tsel = c2.file_uploader("TSEL", type=['xlsx'])
    up_xls = c3.file_uploader("XLS", type=['xlsx'])

# --- PROSES UTAMA ---
if file_master:
    point_mapping = {}
    source_mapping = {}

    # Konfigurasi sesuai logic asli (HANYA POINT > 0)
    configs = [
        (up_ioh, 'Sitelist', 'Group Name', 'Period SP', 'TL Swap/Integ', 'Points', 'IOH'),
        (up_tsel, 'Site List', 'Group Name', 'Period', 'User Account', 'Points', 'TSEL'),
        (up_xls, 'Sheet2', 'Group Name', 'Period', 'User Account', 'Total Points', 'XLS')
    ]

    for uploaded_file, sheet, g_col, p_col, u_col, pt_col, label in configs:
        if uploaded_file:
            try:
                df_src = pd.read_excel(uploaded_file, sheet_name=sheet)
                df_src.columns = [str(c).strip() for c in df_src.columns]
                
                # Logic Combo Key (ID|User)
                df_src['Combo_Key'] = (df_src[g_col].apply(super_clean) + "-" + 
                                      df_src[p_col].apply(clean_period_final) + "|" + 
                                      df_src[u_col].apply(super_clean))
                
                df_src['Pt_Clean'] = df_src[pt_col].apply(convert_points_der)
                
                for _, row in df_src.iterrows():
                    if row['Pt_Clean'] > 0: # Filter Anti-Tiban-Nol
                        point_mapping[row['Combo_Key']] = row['Pt_Clean']
                        source_mapping[row['Combo_Key']] = label
                st.sidebar.success(f"✅ {label} Loaded")
            except Exception as e:
                st.sidebar.error(f"❌ Error {label}: {e}")

    # Load & Update Master (Jaga semua kolom asli)
    df_target = pd.read_excel(file_master)
    temp_key = df_target['Uniq_ID'].apply(super_clean) + "|" + df_target['User Account'].apply(super_clean)
    
    # Update Status & Source
    df_target['Update Status'] = temp_key.apply(lambda x: "Updated" if x in point_mapping else "No Match/Zero")
    df_target['Update Source'] = temp_key.map(source_mapping).fillna("-")
    
    # Update Points (Hanya update kalo ada di mapping, kalo ngga tetep pake nilai asli)
    df_target['Total Points (VEB)'] = temp_key.map(point_mapping).fillna(df_target['Total Points (VEB)'])

    # --- SIDEBAR ADDITIONAL FILTERS ---
    status_options = ["ALL"] + sorted(df_target['Update Status'].unique())
    selected_status = st.sidebar.selectbox("Filter Status", status_options)

    source_options = ["ALL"] + sorted(df_target['Update Source'].unique())
    selected_source = st.sidebar.selectbox("Filter Source", source_options)

    # --- APLIKASI FILTER KE PREVIEW ---
    df_display = df_target.copy()
    
    if search_query:
        mask = (
            df_display['Uniq_ID'].astype(str).str.upper().str.contains(search_query) |
            df_display['User Account'].astype(str).str.upper().str.contains(search_query) |
            df_display['Group Name'].astype(str).str.upper().str.contains(search_query)
        )
        df_display = df_display[mask]
    
    if selected_status != "ALL":
        df_display = df_display[df_display['Update Status'] == selected_status]
        
    if selected_source != "ALL":
        df_display = df_display[df_display['Update Source'] == selected_source]

    # --- PREVIEW TABEL ---
    st.divider()
    st.subheader(f"3. Preview Update ({len(df_display)} rows)")
    st.dataframe(df_display, use_container_width=True)

    # --- DOWNLOAD SECTION ---
    st.divider()
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_target.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 Download Hasil Update (.xlsx)",
        data=buffer.getvalue(),
        file_name="Master_Control_Tower_Update.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("💡 Ayo bre, upload file SP_New.xlsx dulu buat mulai!")