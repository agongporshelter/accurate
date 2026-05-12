import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Consolidator Invoice", layout="wide")
st.title("📊 Aplikasi Penggabung & Pembersih Data Invoice")
st.markdown("Unggah semua file Excel invoice Anda. Aplikasi akan mengambil **Kode PT (SIG, SN, dll)**, mengisi sel yang ter-merge, dan menghapus baris total sehingga siap untuk analisis data.")

uploaded_files = st.file_uploader("Pilih file Excel (.xlsx)", type=['xlsx'], accept_multiple_files=True)

# Daftar kolom yang harus di-forward fill jika kosong (merged cells)
fill_cols = ['Cabang', 'Pelanggan', 'Nama Proyek Utama', 'Tahun Join', 'Sales', 'Mulai PKS', 'Selesai PKS', 'Layanan']

if uploaded_files:
    df_list = []
    progress = st.progress(0)
    
    for i, file in enumerate(uploaded_files):
        try:
            # Membaca excel, header ada di baris ke-3 (index 2)
            df = pd.read_excel(file, header=2)
            
            # 1. Deteksi Otomatis Kode PT dari Header (Kolom yang berisi SN, RCI, GSU, SIG, dll)
            kode_pt = ""
            entity_col_name = None
            for col in df.columns:
                # Cek jika nama kolom adalah huruf kapital semua dan pendek (seperti SN, SIG, RCI)
                if isinstance(col, str) and col.isupper() and len(col) <= 5 and col != 'HEAD':
                    kode_pt = col
                    entity_col_name = col
                    break
            
            # Buat kolom baru Kode_PT di posisi paling kiri
            df.insert(0, 'Kode_PT', kode_pt)
            
            # 2. Hapus kolom yang tidak diperlukan (Unnamed, bintang *, HEAD, dan kolom kode entitas asli)
            cols_to_drop = [col for col in df.columns if 'Unnamed' in str(col) or str(col).strip() == '*' or str(col).strip() == 'HEAD']
            if entity_col_name:
                cols_to_drop.append(entity_col_name) # Hapus kolom asli (misal 'SIG'), karena sudah diganti 'Kode_PT'
                
            df = df.drop(columns=cols_to_drop, errors='ignore')
            
            # 3. Hapus baris kosong
            df = df.dropna(how='all')
            
            # 4. Isi sel kosong (Merged cells) ke bawah (forward fill)
            existing_fill_cols = [col for col in fill_cols if col in df.columns]
            df[existing_fill_cols] = df[existing_fill_cols].ffill()
            
            # 5. Hapus baris Total Cabang
            if 'Cabang' in df.columns:
                df = df[~df['Cabang'].astype(str).str.contains('Total Cabang', na=False)]
                
            # 6. Tambahkan kolom sumber file (opsional, untuk traceability)
            df['Sumber_File'] = file.name
            
            df_list.append(df)
            
        except Exception as e:
            st.error(f"Error memproses {file.name}: {e}")
            
        progress.progress((i + 1) / len(uploaded_files))
    
    if df_list:
        # Gabungkan semua dataframe
        final_df = pd.concat(df_list, ignore_index=True)
        
        # Pastikan Kode_PT ada di paling depan
        cols = list(final_df.columns)
        if 'Kode_PT' in cols:
            cols.insert(0, cols.pop(cols.index('Kode_PT')))
            final_df = final_df[cols]
            
        st.success(f"✅ Berhasil memproses {len(uploaded_files)} file. Total {len(final_df)} baris data siap analisis.")
        
        # Tampilkan data
        st.dataframe(final_df, use_container_width=True)
        
        # Konversi ke Excel untuk diunduh
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Data_Gabungan')
            
            # Auto-adjust column widths
            worksheet = writer.sheets['Data_Gabungan']
            for idx, col in enumerate(final_df.columns):
                max_len = max(final_df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.column_dimensions[worksheet.cell(row=1, column=idx+1).column_letter].width = min(max_len, 30)
                
        output.seek(0)
        
        st.download_button(
            label="📥 Unduh Data Gabungan (Excel)",
            data=output,
            file_name="Data_Invoice_Gabungan_Siap_Analisis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Tidak ada data yang valid untuk digabungkan.")
