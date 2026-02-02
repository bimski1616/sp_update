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
        (up_ioh, 'Sitelist', 'Group Name', 'Period SP', 'TL Installation', 'Points', 'IOH'),
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
    
    # --- DETEKSI BARIS BARU BERDASARKAN GROUP NAME ---
    new_rows_by_group = {}
    temp_key_set = set(temp_key.values)  # Convert ke set untuk efficient lookup
    
    for uploaded_file, sheet, g_col, p_col, u_col, pt_col, label in configs:
        if uploaded_file:
            try:
                df_src = pd.read_excel(uploaded_file, sheet_name=sheet)
                df_src.columns = [str(c).strip() for c in df_src.columns]
                
                for _, row in df_src.iterrows():
                    combo_key = (super_clean(row[g_col]) + "-" + 
                                clean_period_final(row[p_col]) + "|" + 
                                super_clean(row[u_col]))
                    
                    points_clean = convert_points_der(row[pt_col])
                    
                    if points_clean > 0 and combo_key not in temp_key_set:
                        group_clean = super_clean(row[g_col])
                        period_clean = clean_period_final(row[p_col])
                        if group_clean not in new_rows_by_group:
                            new_rows_by_group[group_clean] = []
                        # Store HANYA kolom penting dari source file
                        new_row_data = {
                            g_col: row[g_col],           # Group Name (dengan nama asli kolom)
                            p_col: row[p_col],           # Period (dengan nama asli kolom)
                            u_col: row[u_col],           # User Account (dengan nama asli kolom)
                            pt_col: row[pt_col],         # Points asli (dengan nama asli kolom)
                            # Tambah juga dengan nama standard untuk display
                            'Group Name': row[g_col],
                            'Period': row[p_col],
                            'User Account': row[u_col],
                            'Points': row[pt_col],
                            'Period_Clean': period_clean,
                            'Points_Clean': points_clean,
                            'Source': label,
                            'g_col': g_col,
                            'u_col': u_col,
                            'p_col': p_col
                        }
                        new_rows_by_group[group_clean].append(new_row_data)
            except:
                pass
    
    # --- AUTO ADD SEMUA BARIS BARU KE MASTER ---
    if new_rows_by_group:
        for group_name in new_rows_by_group.keys():
            for new_row in new_rows_by_group[group_name]:
                new_entry = {col: "" for col in df_target.columns}
                
                # Copy semua data dari source yang ada di master columns
                for col in df_target.columns:
                    if col in new_row:
                        new_entry[col] = new_row[col]
                
                # Isi kolom penting yang perlu di-generate (OVERRIDE kolom yang sudah di-copy jika perlu)
                g_col = new_row.get('g_col')
                u_col = new_row.get('u_col')
                p_col = new_row.get('p_col')
                period_clean = new_row.get('Period_Clean', '')
                points_clean = new_row.get('Points_Clean', 0)
                
                # Gunakan periode yang sudah di-clean
                if 'Period' in df_target.columns:
                    new_entry['Period'] = period_clean
                
                if g_col:
                    new_entry['Group Name'] = new_row[g_col]
                    # Uniq_ID 1: Group Name + Period
                    new_entry['Uniq_ID'] = super_clean(new_row[g_col]) + "-" + period_clean
                
                if u_col:
                    new_entry['User Account'] = new_row[u_col]
                    # Uniq_ID 2: User Account + Period (sesuai format di Master, tidak uppercase)
                    new_entry['Uniq_ID 2'] = new_row[u_col] + period_clean
                
                # Gunakan Points yang sudah di-clean dan di-convert
                new_entry['Total Points (VEB)'] = points_clean
                new_entry['Update Status'] = "Added (New Row)"
                new_entry['Update Source'] = new_row.get('Source', '')
                
                df_target = pd.concat([df_target, pd.DataFrame([new_entry])], ignore_index=True)
    
    # Inisialisasi kolom Update Status dan Update Source jika belum ada (untuk baris lama)
    if 'Update Status' not in df_target.columns:
        df_target['Update Status'] = ""
    if 'Update Source' not in df_target.columns:
        df_target['Update Source'] = "-"
    
    # Update Status & Source (SKIP baris baru yang sudah punya status)
    mask_old = df_target['Update Status'] != "Added (New Row)"
    df_target.loc[mask_old, 'Update Status'] = temp_key[mask_old].apply(lambda x: "Updated" if x in point_mapping else "No Match/Zero")
    df_target.loc[mask_old, 'Update Source'] = temp_key[mask_old].map(source_mapping).fillna("-")
    
    # Update Points (Hanya update kalo ada di mapping, kalo ngga tetep pake nilai asli)
    # SKIP untuk baris baru (Added New Row) agar points tidak ter-override
    mask_not_new = df_target['Update Status'] != "Added (New Row)"
    df_target.loc[mask_not_new, 'Total Points (VEB)'] = temp_key[mask_not_new].map(point_mapping).fillna(df_target.loc[mask_not_new, 'Total Points (VEB)'])

    # --- SIDEBAR ADDITIONAL FILTERS ---
    status_options = ["ALL"] + sorted(df_target['Update Status'].fillna("").unique().astype(str))
    selected_status = st.sidebar.selectbox("Filter Status", status_options)

    source_options = ["ALL"] + sorted(df_target['Update Source'].fillna("-").unique().astype(str))
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

    # --- SECTION BARIS BARU PER GROUP NAME ---
    st.divider()
    if new_rows_by_group:
        total_new_rows = sum(len(rows) for rows in new_rows_by_group.values())
        st.success(f"✅ {total_new_rows} baris baru SUDAH DITAMBAHKAN ke Master dari {len(new_rows_by_group)} group!")
        st.subheader("📌 Detail Baris Baru yang Ditambahkan")
        
        with st.expander("Lihat Baris Baru", expanded=False):
            for group_name in sorted(new_rows_by_group.keys()):
                st.write(f"**Group: {group_name}** ({len(new_rows_by_group[group_name])} rows)")
                df_new = pd.DataFrame(new_rows_by_group[group_name])
                st.dataframe(df_new[['Group Name', 'Period', 'User Account', 'Points', 'Source']], use_container_width=True)
                st.divider()
    else:
        st.info("ℹ️ Tidak ada baris baru terdeteksi. Semua data update sudah ada di Master.")

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
