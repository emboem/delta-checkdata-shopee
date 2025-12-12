import streamlit as st
import pandas as pd
import io
import warnings

# Abaikan warning style excel
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- Konfigurasi Halaman ---
st.set_page_config(page_title="Audit Keuangan Shopee - Delta", layout="wide")
st.title("💸 Aplikasi Audit Keuangan Shopee")
st.subheader("Cocokkan Data Pesanan dengan Riwayat Saldo/Transaksi Penghasilan")
st.markdown("""
**Logika Baru:**
1.  **Data Acuan:** File Pesanan (Real-time).
2.  **Data Pembanding:** File Riwayat Saldo/Transaksi (Total uang masuk/keluar per pesanan).
3.  **Fitur:** Menggabungkan produk double dalam satu pesanan & mencocokkan waktu transaksi.
""")

# --- Fungsi Pembersih Angka ---
def clean_currency_indo(x):
    """Membersihkan format mata uang Indonesia."""
    if pd.isna(x) or str(x).strip() in ['', '-']: return 0.0
    x = str(x)
    # Hapus Rp, spasi, non-breaking space
    clean = x.replace('Rp', '').replace(' ', '').replace('\xa0', '')
    
    try:
        # Format (1.000,00) -> 1000.0
        if ',' in clean and '.' in clean:
             clean = clean.replace('.', '').replace(',', '.')
        # Format (10.000) -> 10000
        elif '.' in clean: 
            clean = clean.replace('.', '')
        # Format (10,000) -> 10000 (antisipasi format inggris)
        elif ',' in clean:
            clean = clean.replace(',', '.')
            
        return float(clean)
    except:
        return 0.0

# --- Fungsi Smart Loader (Pencari Header Agresif) ---
def load_data_smart(file, keywords):
    """Mencari header di semua sheet berdasarkan keyword dengan jangkauan lebih luas."""
    try:
        # Jika CSV
        if file.name.endswith('.csv'):
            # Baca 100 baris pertama untuk mencari header
            df_temp = pd.read_csv(file, header=None, nrows=100, dtype=str)
            header_row = -1
            
            # Loop baris demi baris
            for idx, row in df_temp.iterrows():
                # Gabungkan baris jadi string huruf kecil semua
                row_str = " ".join(row.astype(str).fillna('').values).lower()
                
                # Cek apakah SEMUA keyword ada di baris tersebut?
                # Tidak perlu semua, cukup salah satu keyword UNIK yang pasti ada di header
                # Misal: "No. Pesanan"
                if any(k.lower() in row_str for k in keywords):
                    header_row = idx
                    break
            
            if header_row != -1:
                file.seek(0)
                df = pd.read_csv(file, skiprows=header_row, dtype=str)
                return df, "CSV"
            return None, "Keyword tidak ditemukan di CSV (Cek 100 baris pertama)"

        # Jika Excel
        else:
            excel_file = pd.ExcelFile(file)
            found_sheet = None
            final_df = None
            
            for sheet in excel_file.sheet_names:
                # Baca 100 baris pertama tiap sheet
                df_temp = pd.read_excel(file, sheet_name=sheet, header=None, nrows=100, dtype=str)
                
                for idx, row in df_temp.iterrows():
                    row_str = " ".join(row.astype(str).fillna('').values).lower()
                    
                    # Logika Pencarian: Cukup temukan salah satu keyword kunci
                    if any(k.lower() in row_str for k in keywords):
                        # Reset pointer dan baca ulang dari baris header tersebut
                        final_df = pd.read_excel(file, sheet_name=sheet, skiprows=idx, dtype=str)
                        found_sheet = sheet
                        break
                
                if final_df is not None:
                    break
            
            if final_df is not None:
                return final_df, found_sheet
            else:
                return None, "Keyword tidak ditemukan di Excel (Cek 100 baris pertama)"

    except Exception as e:
        return None, str(e)

# --- UPLOAD FILE ---
st.divider()
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. File Pesanan (Acuan)")
    st.caption("File Orders (Oktober). Berisi: No. Pesanan, Waktu, Harga Produk.")
    file_order = st.file_uploader("Upload File Pesanan", type=["xlsx", "csv"], key="ord")

with col2:
    st.subheader("2. File Penghasilan (Saldo)")
    st.caption("File Saldo (Okt-Nov). Berisi: Tanggal Transaksi, No. Pesanan, Jumlah.")
    file_income = st.file_uploader("Upload File Penghasilan", type=["xlsx", "csv"], key="inc")

# --- PROSES UTAMA ---
if st.button("🚀 Mulai Analisis", type="primary"):
    if file_order and file_income:
        with st.spinner("Sedang memproses data..."):
            
            # ==========================================
            # 1. PROSES DATA PESANAN (ORDERS)
            # ==========================================
            kw_order = ["Total Harga Produk", "Harga Awal", "Waktu Pesanan Dibuat"]
            df_ord, sheet_ord = load_data_smart(file_order, kw_order)
            
            if df_ord is None:
                st.error("Gagal membaca File Pesanan. Pastikan ada kolom 'Total Harga Produk'.")
                st.stop()

            # Mapping Kolom Pesanan
            col_id_ord = next((c for c in df_ord.columns if "No. Pesanan" in str(c) or "Order ID" in str(c)), None)
            col_harga_ord = next((c for c in df_ord.columns if "Total Harga Produk" in str(c)), None)
            col_waktu_ord = next((c for c in df_ord.columns if "Waktu Pesanan Dibuat" in str(c) or "Time Created" in str(c)), None)
            col_status_ord = next((c for c in df_ord.columns if "Status Pesanan" in str(c)), None)

            if not col_id_ord or not col_harga_ord:
                st.error("Kolom penting (No. Pesanan / Total Harga Produk) hilang di file Pesanan.")
                st.stop()

            # Pembersihan & Konversi Data Pesanan
            df_ord['No_Pesanan_Ref'] = df_ord[col_id_ord].astype(str).str.strip()
            df_ord['Harga_Produk_Clean'] = df_ord[col_harga_ord].apply(clean_currency_indo)
            
            # Handling Waktu & Status
            if col_waktu_ord:
                df_ord['Waktu_Pesanan'] = df_ord[col_waktu_ord].astype(str)
            else:
                df_ord['Waktu_Pesanan'] = "-"
                
            if col_status_ord:
                df_ord['Status_Pesanan'] = df_ord[col_status_ord].astype(str).str.strip()
            else:
                df_ord['Status_Pesanan'] = "Unknown"

            # AGGREGASI PESANAN (Solusi untuk Nomor Pesanan Double)
            df_ord_final = df_ord.groupby('No_Pesanan_Ref').agg({
                'Harga_Produk_Clean': 'sum',
                'Waktu_Pesanan': 'first',
                'Status_Pesanan': 'first'
            }).reset_index()

            # ==========================================
            # 2. PROSES DATA PENGHASILAN (SALDO/TRANSAKSI)
            # ==========================================
            # Keyword pencarian: fokus ke "Saldo Akhir" atau "Jumlah" yang pasti ada di baris header laporan saldo
            kw_income = ["Saldo Akhir", "Jumlah", "Tanggal Transaksi"]
            df_inc, sheet_inc = load_data_smart(file_income, kw_income)

            if df_inc is None:
                st.error("Gagal membaca File Penghasilan. Header kolom 'Saldo Akhir' atau 'Jumlah' tidak ditemukan di 100 baris pertama.")
                st.stop()
            
            # Mapping Kolom Penghasilan
            col_id_inc = next((c for c in df_inc.columns if "No. Pesanan" in str(c)), None)
            col_jumlah_inc = next((c for c in df_inc.columns if "Jumlah" in str(c) or "Amount" in str(c)), None)
            col_waktu_inc = next((c for c in df_inc.columns if "Tanggal Transaksi" in str(c) or "Waktu" in str(c)), None)

            if not col_id_inc or not col_jumlah_inc:
                st.error(f"Header ditemukan di sheet '{sheet_inc}', tapi kolom 'No. Pesanan' atau 'Jumlah' tidak terbaca dengan benar.")
                st.write("Kolom yang terbaca:", list(df_inc.columns))
                st.stop()

            # Filter Baris Valid: Hapus baris yang No. Pesanannya kosong (header kosong/total di bawah)
            df_inc = df_inc.dropna(subset=[col_id_inc])
            
            # Pembersihan Data Penghasilan
            df_inc['No_Pesanan_Ref'] = df_inc[col_id_inc].astype(str).str.strip()
            df_inc['Jumlah_Clean'] = df_inc[col_jumlah_inc].apply(clean_currency_indo)
            
            if col_waktu_inc:
                df_inc['Waktu_Transaksi'] = df_inc[col_waktu_inc].astype(str)
            else:
                df_inc['Waktu_Transaksi'] = "-"

            # AGGREGASI PENGHASILAN (Net Income per Pesanan)
            df_inc_final = df_inc.groupby('No_Pesanan_Ref').agg({
                'Jumlah_Clean': 'sum',
                'Waktu_Transaksi': 'max'
            }).reset_index()

            # ==========================================
            # 3. PENGGABUNGAN DATA (LEFT JOIN)
            # ==========================================
            merged = pd.merge(df_ord_final, df_inc_final, on='No_Pesanan_Ref', how='left')
            
            merged['Jumlah_Clean'] = merged['Jumlah_Clean'].fillna(0)
            merged['Waktu_Transaksi'] = merged['Waktu_Transaksi'].fillna("-")
            merged['Selisih'] = merged['Harga_Produk_Clean'] - merged['Jumlah_Clean']

            # ==========================================
            # 4. LOGIKA OUTPUT
            # ==========================================
            
            def tentukan_status(row):
                status_ord = str(row['Status_Pesanan']).lower()
                income = row['Jumlah_Clean']
                
                if 'batal' in status_ord or 'cancel' in status_ord:
                    return "DIBATALKAN"
                elif income > 0:
                    return "SINKRON (CAIR)"
                else:
                    return "BELUM CAIR / DATA HILANG"

            merged['Status_Analisis'] = merged.apply(tentukan_status, axis=1)

            # Format Tampilan
            display_df = merged.rename(columns={
                'No_Pesanan_Ref': 'No Pesanan',
                'Waktu_Pesanan': 'Waktu Pesanan (Order)',
                'Waktu_Transaksi': 'Waktu Pencairan (Saldo)',
                'Harga_Produk_Clean': 'Total Harga Produk (Awal)',
                'Jumlah_Clean': 'Total Uang Diterima (Net)',
                'Selisih': 'Gap (Biaya/Potongan)',
                'Status_Pesanan': 'Status Shopee'
            })

            cols_order = [
                'No Pesanan', 
                'Status_Analisis',
                'Waktu Pesanan (Order)', 
                'Waktu Pencairan (Saldo)', 
                'Total Harga Produk (Awal)', 
                'Total Uang Diterima (Net)', 
                'Gap (Biaya/Potongan)',
                'Status Shopee'
            ]
            final_df = display_df[cols_order]
            final_df = final_df.sort_values(by=['Status_Analisis', 'Waktu Pesanan (Order)'], ascending=[False, True])

            # --- TAMPILAN DASHBOARD ---
            st.success("✅ Analisis Selesai!")
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Pesanan Sinkron (Cair)", len(final_df[final_df['Status_Analisis'] == "SINKRON (CAIR)"]))
            m2.metric("Pesanan Dibatalkan", len(final_df[final_df['Status_Analisis'] == "DIBATALKAN"]))
            m3.metric("Total Potongan", f"Rp {final_df[final_df['Status_Analisis'] == 'SINKRON (CAIR)']['Gap (Biaya/Potongan)'].sum():,.0f}")

            st.divider()
            
            # Default Filter: Hanya Tampilkan Sinkron & Batal
            default_options = ["SINKRON (CAIR)", "DIBATALKAN"]
            available_options = final_df['Status_Analisis'].unique().tolist()
            # Pastikan default options ada di data yang tersedia
            valid_defaults = [opt for opt in default_options if opt in available_options]

            filter_status = st.multiselect(
                "Filter Status Tampilan:",
                options=available_options,
                default=valid_defaults
            )
            
            df_filtered = final_df[final_df['Status_Analisis'].isin(filter_status)]

            st.dataframe(
                df_filtered.style.format({
                    'Total Harga Produk (Awal)': 'Rp {:,.0f}',
                    'Total Uang Diterima (Net)': 'Rp {:,.0f}',
                    'Gap (Biaya/Potongan)': 'Rp {:,.0f}'
                }), 
                use_container_width=True
            )

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_filtered.to_excel(writer, sheet_name='Hasil Audit', index=False)
            
            st.download_button("📥 Download Excel", buffer.getvalue(), "Hasil_Audit_Shopee_Lengkap.xlsx", "application/vnd.ms-excel")

    else:

        st.info("Silakan upload kedua file untuk memulai.")
