import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidator Invoice", layout="wide")
st.title("📊 Aplikasi Penggabung & Pembersih Data Invoice")
st.markdown("Unggah semua file Excel invoice Anda. Aplikasi akan mengambil **Kode PT (SIG, SN, dll)**, mengisi sel yang ter-merge, dan menghapus baris total sehingga siap untuk analisis data.")

uploaded_files = st.file_uploader("Pilih file Excel (.xlsx)", type=['xlsx'], accept_multiple_files=True)

# Daftar kolom yang harus di-forward fill jika kosong (merged cells)
fill_cols = ['Cabang', 'Pelanggan', 'Nama Proyek Utama', 'Tahun Join', 'Sales', 'Mulai PKS', 'Selesai PKS', 'Layanan']
valid_codes = ['SN', 'SNI', 'RCI', 'ION', 'GSU', 'SIG'] # Daftar kode PT yang valid

if uploaded_files:
    df_list = []
    progress = st.progress(0)
    
    for i, file in enumerate(uploaded_files):
        try:
            # 1. Baca file mentah (10 baris pertama) untuk mencari header dan kode PT
            df_raw = pd.read_excel(file, header=None, nrows=10)
            
            # Cari baris header (yang mengandung kata 'Cabang')
            header_row = None
            for idx, row in df_raw.iterrows():
                if row.astype(str).str.contains('Cabang').any():
                    header_row = idx
                    break
            
            if header_row is None:
                st.warning(f"Tidak dapat menemukan baris header 'Cabang' di file {file.name}. File dilewati.")
                continue
                
            # Cari Kode PT (SN, RCI, dll) di baris-baris sebelum header
            kode_pt = ""
            for idx, row in df_raw.iterrows():
                if idx < header_row:
                    for val in row:
                        val_str = str(val).strip().upper()
                        if val_str in valid_codes:
                            kode_pt = val_str
                            break
                if kode_pt:
                    break
            
            if not kode_pt:
                kode_pt = "UNKNOWN" # Jika tidak ketemu, akan ditulis UNKNOWN

            # 2. Baca ulang Excel menggunakan baris header yang sudah ditemukan
            df = pd.read_excel(file, header=header_row)
            
            # Buat kolom baru Kode_PT di posisi paling kiri
            df.insert(0, 'Kode_PT', kode_pt)
            
            # 3. Hapus kolom yang tidak diperlukan (Unnamed, bintang *, HEAD, dll)
            cols_to_drop = [col for col in df.columns if 'Unnamed' in str(col) or str(col).strip() == '*' or str(col).strip() == 'HEAD']
            df = df.drop(columns=cols_to_drop, errors='ignore')
            
            # 4. Hapus baris kosong
            df = df.dropna(how='all')
            
            # 5. Isi sel kosong (Merged cells) ke bawah (forward fill)
            existing_fill_cols = [col for col in fill_cols if col in df.columns]
            df[existing_fill_cols] = df[existing_fill_cols].ffill()
            
            # 6. Hapus baris Total Cabang
            if 'Cabang' in df.columns:
                df = df[~df['Cabang'].astype(str).str.contains('Total Cabang', na=False)]
                
            # 7. Tambahkan kolom sumber file
            df['Sumber_File'] = file.name
            
            df_list.append(df)
            
        except Exception as e:
            st.error(f"Error memproses {file.name}: {e}")
            
        progress.progress((i + 1) / len(uploaded_files))
    
    if df_list:
        # Gabungkan semua dataframe
        final_df = pd.concat(df_list, ignore_index=True)
        
        st.success(f"✅ Berhasil memproses {len(uploaded_files)} file. Total {len(final_df)} baris data siap analisis.")
        
        # Tampilkan data
        st.dataframe(final_df, use_container_width=True)
        
        # Konversi ke Excel untuk diunduh
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Data_Gabungan')
                
        output.seek(0)
        
        st.download_button(
            label="📥 Unduh Data Gabungan (Excel)",
            data=output,
            file_name="Data_Invoice_Gabungan_Siap_Analisis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Tidak ada data yang valid untuk digabungkan.")
