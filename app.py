import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Cek Data Individu", layout="wide")

st.title("📋 Aplikasi Cek Validasi Data Individu")
st.markdown("Upload file Excel untuk mengecek kelengkapan data di kolom D s/d Q")

# Upload file
uploaded_file = st.file_uploader(
    "Pilih file Excel (.xlsx)", 
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        # Baca file Excel
        df = pd.read_excel(uploaded_file, sheet_name=0)
        
        st.success(f"✅ File berhasil diupload! Total {len(df)} baris data")
        
        # Tampilkan header
        st.subheader("📌 Header / Kolom yang tersedia:")
        headers = df.columns.tolist()
        
        # Buat tabel header dengan indeks kolom (A-Z)
        header_data = []
        for i, col in enumerate(headers):
            col_letter = chr(65 + i) if i < 26 else f"Col{i+1}"
            header_data.append({"Kolom": col_letter, "Nama Header": col})
        
        st.dataframe(pd.DataFrame(header_data), use_container_width=True, hide_index=True)
        
        # Kolom yang dicek (D s/d Q) = index 3 s/d 16
        kolom_yang_dicek = headers[3:17]  # D sampai Q
        
        st.subheader("🔍 Hasil Pengecekan Data Kosong (Kolom D - Q)")
        
        # ============ BUAT CEKLIS LANGSUNG PER BARIS ============
        
        # Siapkan data untuk ceklis
        ceklis_data = []
        status_data = []
        
        for idx, row in df.iterrows():
            baris = idx + 2  # Karena Excel baris 1 = header
            
            # Data per baris untuk ceklis
            row_ceklis = {"Baris": baris}
            
            # Cek tiap kolom D-Q
            for col in kolom_yang_dicek:
                nilai = row[col]
                kosong = pd.isna(nilai) or str(nilai).strip() == ""
                
                # LOGIKA KHUSUS UNTUK BO
                if col == "BO":
                    pekerjaan = str(row.get("Pekerjaan", "")).strip().upper()
                    is_wajib_bo = any(p in pekerjaan for p in ["PELAJAR/MAHASISWA", "IBU RUMAH TANGGA"])
                    
                    if is_wajib_bo:
                        # WAJIB DIISI
                        row_ceklis[col] = "❌" if kosong else "✅"
                    else:
                        # TIDAK WAJIB - tetap hijau
                        row_ceklis[col] = "✅"  # Anggap aman walaupun kosong
                else:
                    # Kolom lain: wajib diisi semua
                    row_ceklis[col] = "❌" if kosong else "✅"
            
            # Tentukan status akhir
            ada_kosong = any(v == "❌" for k, v in row_ceklis.items() if k != "Baris")
            
            # Cek khusus BO error
            bo_value = row.get("BO")
            bo_kosong = pd.isna(bo_value) or str(bo_value).strip() == ""
            pekerjaan = str(row.get("Pekerjaan", "")).strip().upper()
            is_wajib_bo = any(p in pekerjaan for p in ["PELAJAR/MAHASISWA", "IBU RUMAH TANGGA"])
            error_bo = is_wajib_bo and bo_kosong
            
            if error_bo:
                status = "⚠️ Error BO (Wajib Diisi)"
            elif ada_kosong:
                status = "❌ Data Tidak Sesuai"
            else:
                status = "✅ Sesuai"
            
            row_ceklis.insert(0, "Status", status)
            ceklis_data.append(row_ceklis)
            status_data.append(status)
        
        # Buat DataFrame ceklis
        ceklis_df = pd.DataFrame(ceklis_data)
        
        # Tampilkan statistik
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📊 Total Data", len(df))
        with col2:
            st.metric("✅ Sesuai", sum(1 for s in status_data if s == "✅ Sesuai"))
        with col3:
            st.metric("❌ Data Tidak Sesuai", sum(1 for s in status_data if s == "❌ Data Tidak Sesuai"))
        with col4:
            st.metric("⚠️ Error BO", sum(1 for s in status_data if s == "⚠️ Error BO (Wajib Diisi)"))
        
        # ============ TAMPILKAN CEKLIS LENGKAP ============
        st.subheader("📋 Ceklis Kelengkapan Data per Baris")
        st.markdown("**Keterangan:** ✅ = Terisi | ❌ = Kosong / Bermasalah")
        
        # Styling biar lebih enak dilihat
        st.dataframe(
            ceklis_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Baris": st.column_config.NumberColumn("Baris ke-"),
                "Status": st.column_config.TextColumn("Status Akhir"),
            }
        )
        
        # ============ TAMPILKAN YANG BERMASALAH ============
        st.subheader("🚨 Data yang Bermasalah")
        masalah_df = ceklis_df[ceklis_df["Status"] != "✅ Sesuai"]
        if len(masalah_df) > 0:
            st.dataframe(masalah_df, use_container_width=True, hide_index=True)
            st.warning(f"⚠️ Ada {len(masalah_df)} baris yang bermasalah!")
        else:
            st.success("🎉 Semua data lengkap dan sesuai!")
        
        # ============ DATA ASLI + HASIL (Side by Side) ============
        st.subheader("📊 Data Asli + Hasil Pengecekan")
        
        # Gabungkan data asli dengan hasil ceklis
        df_hasil = df.copy()
        
        # Tambahkan kolom status dan hasil pengecekan
        df_hasil.insert(0, "Status", status_data)
        
        # Tambahkan kolom ceklis per kolom
        for col in kolom_yang_dicek:
            df_hasil[f"Cek_{col}"] = ceklis_df[col].values
        
        st.dataframe(
            df_hasil,
            use_container_width=True,
            hide_index=True
        )
        
        # ============ FILTER DATA ============
        st.subheader("🔎 Filter Data")
        filter_status = st.selectbox(
            "Pilih filter:",
            ["Semua Data", "✅ Sesuai", "❌ Data Tidak Sesuai", "⚠️ Error BO (Wajib Diisi)"]
        )
        
        if filter_status == "✅ Sesuai":
            st.dataframe(df[ceklis_df["Status"] == "✅ Sesuai"], use_container_width=True)
        elif filter_status == "❌ Data Tidak Sesuai":
            st.dataframe(df[ceklis_df["Status"] == "❌ Data Tidak Sesuai"], use_container_width=True)
        elif filter_status == "⚠️ Error BO (Wajib Diisi)":
            st.dataframe(df[ceklis_df["Status"] == "⚠️ Error BO (Wajib Diisi)"], use_container_width=True)
        else:
            st.dataframe(df, use_container_width=True)
        
        # ============ DOWNLOAD HASIL ============
        st.subheader("⬇️ Download Hasil Pengecekan")
        
        # Buat file Excel dengan 3 sheet
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Sheet 1: Data Asli + Hasil Ceklis
            df_hasil.to_excel(writer, sheet_name='Data+Hasil_Ceklis', index=False)
            
            # Sheet 2: Ceklis (ringkasan)
            ceklis_df.to_excel(writer, sheet_name='Ceklis', index=False)
            
            # Sheet 3: Data Bermasalah
            masalah_df = ceklis_df[ceklis_df["Status"] != "✅ Sesuai"]
            if len(masalah_df) > 0:
                masalah_df.to_excel(writer, sheet_name='Data_Bermasalah', index=False)
        
        output.seek(0)
        
        st.download_button(
            label="📥 Download Hasil Pengecekan (Excel)",
            data=output,
            file_name=f"hasil_pengecekan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # ============ TAMPILKAN HEADER YANG DICEK ============
        with st.expander("📋 Lihat Header yang Dicek (Kolom D - Q)"):
            header_ceklis = []
            for i, col in enumerate(kolom_yang_dicek):
                col_letter = chr(65 + 3 + i)  # D = 68
                header_ceklis.append({
                    "Kolom": col_letter,
                    "Nama Header": col,
                    "Wajib Diisi": "✅" if col != "BO" else "✅ (Kecuali BO tidak wajib)"
                })
            st.dataframe(pd.DataFrame(header_ceklis), use_container_width=True, hide_index=True)
        
    except Exception as e:
        st.error(f"❌ Terjadi error: {e}")
        st.info("Pastikan file Excel memiliki format yang benar.")

else:
    st.info("👆 Upload file Excel untuk memulai pengecekan.")
