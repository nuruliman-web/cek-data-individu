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
        kolom_yang_dicek_nama = [f"{chr(65+3+i)}. {kolom_yang_dicek[i]}" for i in range(len(kolom_yang_dicek))]
        
        st.subheader("🔍 Hasil Pengecekan Data Kosong (Kolom D - Q)")
        
        # Siapkan data untuk hasil
        hasil_data = []
        baris_bermasalah = 0
        error_bo_count = 0
        
        for idx, row in df.iterrows():
            baris = idx + 2  # Karena Excel baris 1 = header
            kolom_kosong = []
            status = "✅ Sesuai"
            warning_bo = ""
            is_error_bo = False
            
            # Cek kolom D sampai Q (index 3-16)
            for i, col in enumerate(kolom_yang_dicek):
                nilai = row[col]
                # Cek kosong (None, NaN, atau string kosong)
                if pd.isna(nilai) or str(nilai).strip() == "":
                    kolom_kosong.append(col)
            
            # LOGIKA KHUSUS: BO (kolom Q / index 16) wajib untuk PELAJAR/MAHASISWA dan IBU RUMAH TANGGA
            pekerjaan = str(row.get("Pekerjaan", "")).strip().upper()
            bo_value = row.get("BO")
            bo_kosong = pd.isna(bo_value) or str(bo_value).strip() == ""
            
            # Cek apakah pekerjaan termasuk yang wajib BO
            pekerjaan_wajib_bo = ["PELAJAR/MAHASISWA", "IBU RUMAH TANGGA"]
            is_wajib_bo = any(p in pekerjaan for p in pekerjaan_wajib_bo)
            
            if is_wajib_bo and bo_kosong:
                # Ini error BO
                is_error_bo = True
                error_bo_count += 1
                if "BO" not in kolom_kosong:
                    kolom_kosong.append("BO")
                warning_bo = "⚠️ WAJIB DIISI (Pekerjaan: " + pekerjaan + ")"
            
            # Tentukan status
            if kolom_kosong:
                if is_error_bo and len(kolom_kosong) == 1 and kolom_kosong[0] == "BO":
                    # Hanya error BO, bukan data kosong biasa
                    status = "⚠️ Error BO (Wajib Diisi)"
                else:
                    status = "❌ Data Tidak Sesuai"
                baris_bermasalah += 1
            
            # Simpan hasil per baris
            hasil_data.append({
                "Baris": baris,
                "Status": status,
                "Kolom Kosong": ", ".join(kolom_kosong) if kolom_kosong else "-",
                "Catatan BO": warning_bo if warning_bo else "-",
                "Pekerjaan": row.get("Pekerjaan", ""),
                "BO Terisi": "✅" if not bo_kosong else "❌",
                "Is Wajib BO": "✅" if is_wajib_bo else "❌"
            })
        
        # Tampilkan statistik
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📊 Total Data", len(df))
        with col2:
            st.metric("✅ Data Sesuai", len(df) - baris_bermasalah)
        with col3:
            st.metric("❌ Data Tidak Sesuai", baris_bermasalah - error_bo_count)
        with col4:
            st.metric("⚠️ Error BO", error_bo_count)
        
        # Tampilkan hasil dalam tabel
        st.subheader("📋 Detail Pengecekan per Baris")
        hasil_df = pd.DataFrame(hasil_data)
        st.dataframe(
            hasil_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Baris": st.column_config.NumberColumn("Baris ke-"),
                "Status": st.column_config.TextColumn("Status"),
                "Kolom Kosong": st.column_config.TextColumn("Kolom Kosong"),
                "Catatan BO": st.column_config.TextColumn("Catatan BO"),
                "Pekerjaan": st.column_config.TextColumn("Pekerjaan"),
                "BO Terisi": st.column_config.TextColumn("BO Terisi"),
                "Is Wajib BO": st.column_config.TextColumn("Wajib BO?")
            }
        )
        
        # Tampilkan ceklis per kolom
        st.subheader("📊 Ceklis Kelengkapan Data per Kolom (D - Q)")
        
        # Buat dataframe ceklis
        ceklis_data = []
        for idx, row in df.iterrows():
            baris = idx + 2
            row_data = {"Baris": baris}
            
            for col in kolom_yang_dicek:
                nilai = row[col]
                kosong = pd.isna(nilai) or str(nilai).strip() == ""
                
                # LOGIKA BO: cek pekerjaan
                if col == "BO":
                    pekerjaan = str(row.get("Pekerjaan", "")).strip().upper()
                    if any(p in pekerjaan for p in ["PELAJAR/MAHASISWA", "IBU RUMAH TANGGA"]):
                        # Wajib diisi
                        row_data[col] = "❌" if kosong else "✅"
                    else:
                        # Tidak wajib, selalu ✅ meskipun kosong
                        row_data[col] = "✅ (Tidak Wajib)" if kosong else "✅"
                else:
                    row_data[col] = "❌" if kosong else "✅"
            
            ceklis_data.append(row_data)
        
        ceklis_df = pd.DataFrame(ceklis_data)
        
        # Tambahkan kolom "Status Akhir"
        ceklis_df.insert(0, "Status Akhir", hasil_df["Status"].values)
        
        st.dataframe(
            ceklis_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Baris": st.column_config.NumberColumn("Baris"),
                "Status Akhir": st.column_config.TextColumn("Status Akhir")
            }
        )
        
        # Tampilkan data yang bermasalah aja
        st.subheader("🚨 Data yang Bermasalah")
        masalah_df = hasil_df[hasil_df["Status"] != "✅ Sesuai"]
        if len(masalah_df) > 0:
            st.dataframe(masalah_df, use_container_width=True, hide_index=True)
            st.warning(f"⚠️ Ada {len(masalah_df)} baris yang bermasalah! Perbaiki data yang kosong.")
        else:
            st.success("🎉 Semua data lengkap dan sesuai!")
        
        # ============ SISI KIRI: DATA ASLI, SISI KANAN: HASIL ============
        st.subheader("📊 Data Asli + Hasil Pengecekan (Side by Side)")
        
        # Gabungkan data asli dengan hasil pengecekan
        df_hasil_gabung = df.copy()
        
        # Tambahkan kolom hasil pengecekan ke data asli
        df_hasil_gabung.insert(0, "Status", hasil_df["Status"].values)
        df_hasil_gabung.insert(1, "Kolom Kosong", hasil_df["Kolom Kosong"].values)
        df_hasil_gabung.insert(2, "Catatan BO", hasil_df["Catatan BO"].values)
        df_hasil_gabung.insert(3, "BO Terisi", hasil_df["BO Terisi"].values)
        df_hasil_gabung.insert(4, "Wajib BO", hasil_df["Is Wajib BO"].values)
        
        st.dataframe(
            df_hasil_gabung,
            use_container_width=True,
            hide_index=True
        )
        
        # Filter data berdasarkan status
        st.subheader("🔎 Filter Data")
        filter_status = st.selectbox(
            "Pilih filter:",
            ["Semua Data", "✅ Sesuai", "❌ Data Tidak Sesuai", "⚠️ Error BO (Wajib Diisi)"]
        )
        
        if filter_status == "✅ Sesuai":
            st.dataframe(df[hasil_df["Status"] == "✅ Sesuai"], use_container_width=True)
        elif filter_status == "❌ Data Tidak Sesuai":
            st.dataframe(df[hasil_df["Status"] == "❌ Data Tidak Sesuai"], use_container_width=True)
        elif filter_status == "⚠️ Error BO (Wajib Diisi)":
            st.dataframe(df[hasil_df["Status"] == "⚠️ Error BO (Wajib Diisi)"], use_container_width=True)
        else:
            st.dataframe(df, use_container_width=True)
        
        # ============ DOWNLOAD HASIL ============
        st.subheader("⬇️ Download Hasil Pengecekan")
        
        # Buat file Excel dengan 3 sheet
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Sheet 1: Data Asli + Hasil
            df_hasil_gabung.to_excel(writer, sheet_name='Data+Hasil', index=False)
            
            # Sheet 2: Hasil Pengecekan (ringkasan)
            hasil_df.to_excel(writer, sheet_name='Ringkasan_Hasil', index=False)
            
            # Sheet 3: Ceklis
            ceklis_df.to_excel(writer, sheet_name='Ceklis', index=False)
            
            # Sheet 4: Data Bermasalah
            masalah_df.to_excel(writer, sheet_name='Data_Bermasalah', index=False)
        
        output.seek(0)
        
        st.download_button(
            label="📥 Download Hasil Pengecekan (Excel)",
            data=output,
            file_name=f"hasil_pengecekan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ Terjadi error: {e}")
        st.info("Pastikan file Excel memiliki format yang benar.")

else:
    st.info("👆 Upload file Excel untuk memulai pengecekan.")
