import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Consolidator Invoice", layout="wide")
st.title("📊 Aplikasi Penggabung & Pembersih Data Invoice")
st.markdown("Unggah semua file Excel invoice Anda di sini. Aplikasi akan menghapus kolom kosong, mengisi sel yang ter-merge, dan menghapus baris total sehingga siap untuk analisis data (Pivot Table, dll).")

uploaded_files = st.file_uploader("Pilih file Excel (.xlsx)", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    df_list = []
    progress = st.progress(0)
    
    for i, file in enumerate(uploaded_files):
        try:
            # Membaca excel, header ada di baris ke-3 (index 2)
            df = pd.read_excel(file, header=2)
            
            # 1. Hapus kolom yang tidak diperlukan (Unnamed, bintang *, dan kode entitas)
            cols_to_drop = [col for col in df.columns if 'Unnamed' in str(col) or str(col).strip() == '*' or str(col).strip() in ['SN', 'SNI', 'RCI', 'ION', 'GSU', 'SIG', 'HEAD']]
            df = df.drop(columns=cols_to_drop, errors='ignore')
            
            # 2. Hapus baris kosong
            df = df.dropna(how='all')
            
            # 3. Isi sel kosong (Merged cells) ke bawah (forward fill)
            fill_cols = ['Cabang', 'Pelanggan', 'Nama Proyek Utama', 'Tahun Join', 'Sales', 'Mulai PKS', 'Selesai PKS', 'Layanan']
            existing_fill_cols = [col for col in fill_cols if col in df.columns]
            df[existing_fill_cols] = df[existing_fill_cols].ffill()
            
            # 4. Hapus baris Total Cabang
            if 'Cabang' in df.columns:
                df = df[~df['Cabang'].astype(str).str.contains('Total Cabang', na=False)]
                
            # 5. Tambahkan kolom sumber file
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