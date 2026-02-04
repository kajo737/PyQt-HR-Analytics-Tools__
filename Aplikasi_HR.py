import sys
import os
import warnings
from datetime import datetime, time, timedelta
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import openpyxl
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                             QFileDialog, QProgressBar, QMessageBox, QGroupBox, 
                             QGridLayout, QStackedWidget, QTableWidget, 
                             QTableWidgetItem, QHeaderView)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QFont, QPalette, QColor

warnings.simplefilter(action='ignore')
def process_logic(file_transaksi, file_shift, output_excel, signal_status, signal_progress):
    
    def clean_nama_series(nama_series):
        return (
            nama_series
            .astype(str)
            .str.replace(r"\bnan\b", "", regex=True)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    # Masukkan Data yang ingin di exclude ke config file, jika blom ada excel config_file maka akan terbuat excel default otomatis 
    CONFIG_FILE = "config_exclude.xlsx"
    DEFAULT_NAMA = [
        
    ]
    DEFAULT_ID = [
       
    ]
    DEFAULT_DEPT_MAIN = []
    DEFAULT_DEPT_EVENT = []

    signal_status.emit("Memuat Konfigurasi Exclude...")
    
    if not os.path.exists(CONFIG_FILE):
        signal_status.emit("Membuat file config_exclude.xlsx...")
        max_len = max(len(DEFAULT_NAMA), len(DEFAULT_ID), len(DEFAULT_DEPT_MAIN), len(DEFAULT_DEPT_EVENT))
        def pad_list(l, length): return l + [None] * (length - len(l))
        df_conf = pd.DataFrame({
            "EXCLUDE_NAMA": pad_list(DEFAULT_NAMA, max_len),
            "EXCLUDE_ID": pad_list(DEFAULT_ID, max_len),
            "EXCLUDE_DEPT_MAIN": pad_list(DEFAULT_DEPT_MAIN, max_len),
            "EXCLUDE_DEPT_EVENT": pad_list(DEFAULT_DEPT_EVENT, max_len)
        })
        df_conf.to_excel(CONFIG_FILE, index=False)
        NAMA_EXCLUDE, ID_EXCLUDE, DEPT_EXCLUDE_MAIN = DEFAULT_NAMA, DEFAULT_ID, DEFAULT_DEPT_MAIN
        DEPT_EXCLUDE_EVENT = DEFAULT_DEPT_EVENT + DEFAULT_DEPT_MAIN 
    else:
        try:
            df_conf = pd.read_excel(CONFIG_FILE)
            NAMA_EXCLUDE = df_conf["EXCLUDE_NAMA"].dropna().astype(str).tolist() if "EXCLUDE_NAMA" in df_conf.columns else []
            ID_EXCLUDE = df_conf["EXCLUDE_ID"].dropna().astype(int).tolist() if "EXCLUDE_ID" in df_conf.columns else []
            DEPT_EXCLUDE_MAIN = df_conf["EXCLUDE_DEPT_MAIN"].dropna().astype(str).tolist() if "EXCLUDE_DEPT_MAIN" in df_conf.columns else []
            dept_event_only = df_conf["EXCLUDE_DEPT_EVENT"].dropna().astype(str).tolist() if "EXCLUDE_DEPT_EVENT" in df_conf.columns else []
            DEPT_EXCLUDE_EVENT = list(set(DEPT_EXCLUDE_MAIN + dept_event_only))
            NAMA_EXCLUDE = [clean_nama_series(pd.Series([n]))[0] for n in NAMA_EXCLUDE]
        except Exception as e:
            print(f"Gagal baca config: {e}")
            NAMA_EXCLUDE, ID_EXCLUDE, DEPT_EXCLUDE_MAIN, DEPT_EXCLUDE_EVENT = [], [], [], []

    # --- MEMBACA FILE (FIX UNTUK XLS & XLSX) ---
    signal_status.emit("Membaca File Input...")
    signal_progress.emit(10)
   
    # 1. LOAD DATA & PENGUBAHAN NAMA MENGIKUTI FILE ATTENDENCE

    df = pd.read_excel(file_transaksi, sheet_name=0)
    df_shift = pd.read_excel(file_shift, sheet_name=0)
    df.columns = df.columns.str.strip()
    df_shift.columns = df_shift.columns.str.strip()
    if "Employee ID" in df_shift.columns:
        df_shift = df_shift.rename(columns={"Employee ID": "Personnel ID"})
    
    if "Employee Name" in df_shift.columns:
        df_shift["Employee Name"] = clean_nama_series(df_shift["Employee Name"])

    if "Personnel ID" in df_shift.columns and "Employee Name" in df_shift.columns:
        kamus_nama_shift = df_shift.dropna(subset=["Personnel ID"]).drop_duplicates("Personnel ID").set_index("Personnel ID")["Employee Name"].to_dict()
    else:
        kamus_nama_shift = {}
        print("PERINGATAN: Kolom 'Employee ID' atau 'Employee Name' tidak ditemukan di File Shift!")
    if "Personnel ID" in df.columns:
        
        df["Nama_Dari_Shift"] = df["Personnel ID"].map(kamus_nama_shift)
    
        mask_ketemu = df["Nama_Dari_Shift"].notna()
        
        if "First Name" in df.columns:
            # 1. Jika ID ketemu di shift, pakai nama dari shift (Timpa First Name)
            df.loc[mask_ketemu, "First Name"] = df.loc[mask_ketemu, "Nama_Dari_Shift"]
            # 2. Jika ID TIDAK ketemu, bersihkan nama aslinya (agar tidak berantakan)
            df.loc[~mask_ketemu, "First Name"] = clean_nama_series(df.loc[~mask_ketemu, "First Name"])
        
        if "Last Name" in df.columns:
            # 1. Jika ID ketemu, kosongkan Last Name (karena First Name sudah berisi nama lengkap dari Shift)
            df.loc[mask_ketemu, "Last Name"] = ""
            # 2. Jika ID TIDAK ketemu, biarkan Last Name ada (hanya dibersihkan)
            df.loc[~mask_ketemu, "Last Name"] = clean_nama_series(df.loc[~mask_ketemu, "Last Name"])
    
        df.drop(columns=["Nama_Dari_Shift"], inplace=True)

    # Rename Tanggal
    if "Attendance Date" in df_shift.columns:
        df_shift = df_shift.rename(columns={"Attendance Date": "Tanggal"})

    # Konversi Waktu
    df_shift["Tanggal"] = pd.to_datetime(df_shift["Tanggal"], errors='coerce').dt.date
    df_shift["Shift In"] = pd.to_datetime(df_shift["Shift In"].astype(str), errors='coerce').dt.time
    df_shift["Shift Out"] = pd.to_datetime(df_shift["Shift Out"].astype(str), errors='coerce').dt.time
    
    # Hapus shift kosong (00:00 - 00:00)
    df_shift = df_shift[
        ~((df_shift["Shift In"] == time(0, 0)) & (df_shift["Shift Out"] == time(0, 0)))
    ].reset_index(drop=True)
    
    # Konversi Waktu Absen Aktual di File Attendence
    if "Attendance Time Out" in df_shift.columns:
        df_shift["Attendance Time Out"] = pd.to_datetime(df_shift["Attendance Time Out"].astype(str), errors='coerce').dt.time
    if "Attendance Time In" in df_shift.columns:
        df_shift["Attendance Time In"] = pd.to_datetime(df_shift["Attendance Time In"].astype(str), errors='coerce').dt.time
    signal_status.emit("Sinkronisasi Departemen...")
    
    dept_ref = df[["Personnel ID", "Department Name"]].drop_duplicates("Personnel ID", keep="first")
    
    if "Department Name" not in df_shift.columns:
        df_shift = df_shift.merge(dept_ref, on="Personnel ID", how="left")
        df_shift["Department Name"] = df_shift["Department Name"].fillna("Unknown")
    
    # Bersihkan nama kolom jika terjadi duplikasi saat merge otomatis
    if "Department Name_x" in df_shift.columns:
        df_shift = df_shift.rename(columns={"Department Name_x": "Department Name"})

    # Fungsi Filter Exclude
    def apply_exclusion(df, exclude_nama=True, exclude_id=True, exclude_dept=None):
        if df is None or df.empty: return df
        df = df.copy()
        if exclude_nama and "Nama" in df.columns:
            df["Nama"] = clean_nama_series(df["Nama"])
            df = df[~df["Nama"].isin(NAMA_EXCLUDE)]
        if exclude_id and "Personnel ID" in df.columns:
            df = df[~df["Personnel ID"].isin(ID_EXCLUDE)]
        if exclude_dept and "Department Name" in df.columns:
            df = df[~df["Department Name"].isin(exclude_dept)]
        return df.reset_index(drop=True)

    
    # SORTING, PAIRING & CLEANING

    signal_status.emit("Memproses Sorting & Pairing Data...")
    signal_progress.emit(25)
    
    # 1. Parsing DateTime & Sorting Wajib
    df["datetime"] = pd.to_datetime(df["Date"].astype(str) + " " + df["Time"].astype(str), errors="coerce")
    df = df.dropna(subset=["datetime"])
    
    # Sort: ID -> Nama -> Tanggal -> Waktu (Paling Kecil)
    sort_columns = ["Personnel ID", "datetime"]
    if "First Name" in df.columns:
        sort_columns = ["Personnel ID", "First Name", "datetime"]
        
    df = df.sort_values(sort_columns).reset_index(drop=True)

    # 2. Filtering State Machine (In -> Out -> In -> Out)
    cleaned_rows = []
    
    for pid, group in df.groupby("Personnel ID"):
        expect_in = True  # State awal: Harus Plant In
        last_valid_time = None
        
        for _, row in group.iterrows():
            event = row["Event Point"]
            curr_time = row["datetime"]
            
            # Cek Clustering 60 Detik
            if last_valid_time and (curr_time - last_valid_time).total_seconds() <= 60:
                continue 
            
            if expect_in:
                if event == "Plant-In":
                    cleaned_rows.append(row)
                    last_valid_time = curr_time
                    expect_in = False 
                # Jika ketemu Plant-Out saat butuh In, abaikan
            
            else: 
                if event == "Plant-Out":
                    cleaned_rows.append(row)
                    last_valid_time = curr_time
                    expect_in = True
                # Jika ketemu Plant-In saat butuh Out, abaikan

    df_clean = pd.DataFrame(cleaned_rows)

    # 3. Membuat DataFrame Pasangan (Durasi)
    signal_status.emit("Menghitung Durasi Kerja...")
    signal_progress.emit(40)
    
    hasil = []
    if not df_clean.empty:
        for pid, group in df_clean.groupby("Personnel ID"):
            group = group.reset_index(drop=True)
            # Loop dengan step 2 (ambil indeks 0 dan 1, 2 dan 3, dst)
            for i in range(0, len(group) - 1, 2):
                row_in = group.iloc[i]
                row_out = group.iloc[i+1]
                
                # Safety check: Pastikan pasangannya benar In dan Out
                if row_in["Event Point"] == "Plant-In" and row_out["Event Point"] == "Plant-Out":
                    hasil.append({
                        "Personnel ID": pid,
                        "Nama": f"{row_in['First Name']} {row_in['Last Name']}",
                        "Tanggal": row_in["datetime"].date(),
                        "Plant-In": row_in["datetime"],
                        "Plant-Out": row_out["datetime"],
                        "Selisih Menit": round((row_out["datetime"] - row_in["datetime"]).total_seconds()/60, 2),
                        "Department Name": row_in["Department Name"]
                    })
    
    df_selisih = pd.DataFrame(hasil)
    
    # 4. Merge dengan Data Shift
    jam_istirahat = [(time(4,0), time(5,0)), (time(12,0), time(13,0)), (time(18,0), time(19,0))]
    cols_merge = ["Personnel ID", "Tanggal", "Shift In", "Shift Out", "Attendance Time Out"]
    
    if df_selisih.empty:
        df_sheet2 = pd.DataFrame()
        df_gabung = pd.DataFrame()
    else:
        df_sheet2 = df_selisih.merge(df_shift[cols_merge], on=["Personnel ID", "Tanggal"], how="left")
        df_gabung = df_selisih.merge(df_shift[cols_merge], on=["Personnel ID", "Tanggal"], how="left")
    
    if df_clean.empty:
        df_event = pd.DataFrame()
    else:
        df_clean["Tanggal"] = df_clean["datetime"].dt.date
        df_event = df_clean.merge(df_shift[cols_merge], on=["Personnel ID", "Tanggal"], how="left")

    signal_status.emit("Menganalisa Pelanggaran (Sheet 3-9)...")
    signal_progress.emit(50)

    # --- LOGIKA SHEET 3: PULANG SEBELUM SHIFT SELESAI ---
    def cek_pulang_sebelum_shift(group):
        res = []
        group = group.sort_values("datetime").reset_index(drop=True)
        for i, row in group.iterrows():
            if row["Event Point"] != "Plant-Out": continue
            if pd.isna(row["Shift In"]) or pd.isna(row["Shift Out"]): continue
            
            s_in = datetime.combine(row["Tanggal"], row["Shift In"])
            s_out = datetime.combine(row["Tanggal"], row["Shift Out"])
            
            # Handle Shift Malam (Lintas Hari)
            if s_out <= s_in: 
                s_out += timedelta(days=1)
            
            plant_out = row["datetime"]
            
            # Cek range waktu: 60 menit sebelum shift berakhir
            if not (s_out - timedelta(minutes=60) <= plant_out < s_out): 
                continue
            
            # Cek Validasi: Apakah ada masuk lagi setelah ini?
            ada_masuk_lagi = False
            for j in range(i + 1, len(group)):
                next_row = group.iloc[j]
                if next_row["Event Point"] == "Plant-In":
                    # Jika ada In lagi dalam rentang toleransi shift out (+3 jam), anggap bukan pulang awal
                    if (s_out - timedelta(minutes=60) <= next_row["datetime"] <= s_out + timedelta(minutes=180)):
                        ada_masuk_lagi = True
                    break
            
            if ada_masuk_lagi: 
                continue
            
            res.append({
                "Tanggal": row["Tanggal"],
                "Personnel ID": row["Personnel ID"],
                "Nama": f"{row['First Name'] or ''} {row['Last Name'] or ''}".strip(),
                "Department Name" : row["Department Name"],
                "Plant-Out": plant_out,
                "Shift In": row["Shift In"],
                "Shift Out": row["Shift Out"],
                "Selisih (menit)": round((s_out - plant_out).total_seconds() / 60, ),
                "Attendance Time Out":row["Attendance Time Out"],
                "Keterangan": "Pulang sebelum shift selesai"
            })
        return pd.DataFrame(res)
    
    # Eksekusi Sheet 3
    if not df_event.empty:
        df_sheet3 = pd.concat([cek_pulang_sebelum_shift(g) for _, g in df_event.groupby("Personnel ID")], ignore_index=True)
    else:
        df_sheet3 = pd.DataFrame()
    
    df_sheet3 = apply_exclusion(df_sheet3, exclude_dept=DEPT_EXCLUDE_MAIN)

    # --- LOGIKA SHEET 4: JEDA SHIFT / ISTIRAHAT ---
    def build_shift_datetime(row):
        if pd.isna(row["Shift In"]): return None, None
        s_in = datetime.combine(row["Tanggal"], row["Shift In"])
        s_out = datetime.combine(row["Tanggal"], row["Shift Out"])
        if s_out <= s_in: s_out += timedelta(days=1)
        return s_in, s_out

    def cek_jeda_shift(group):
        res = []
        group = group.sort_values("datetime").reset_index(drop=True)
        i = 0
        while i < len(group) - 1:
            curr = group.iloc[i]
            
            # Cari pasangan Out -> Inef
            if curr["Event Point"] != "Plant-Out":
                i += 1
                continue
            
            next_row = group.iloc[i+1]
            if next_row["Event Point"] == "Plant-Out":
                i += 1 
                continue
            
            if next_row["Event Point"] == "Plant-In":
                plant_out = curr["datetime"]
                plant_in_next = next_row["datetime"]
                
                s_in, s_out = build_shift_datetime(curr)
                if s_in is None or s_out is None:
                    i += 1
                    continue
                
                # Validasi: Out harus di dalam jam kerja
                if not (s_in <= plant_out < s_out):
                    i += 1
                    continue
                
                # Validasi Waktu
                if plant_in_next <= plant_out:
                    i += 1
                    continue
                
                jeda_menit = (plant_in_next - plant_out).total_seconds() / 60
                
                # Aturan Jumat vs Hari Biasa
                hari = plant_out.weekday() # 4 = Jumat
                # Jika Jumat dan jam istirahat Jumat (08:00 - 16:00 range luas) batas 100 menit
                batas = 100 if (hari == 4 and time(8,0) <= plant_out.time() <= time(16,0)) else 70
                
                if jeda_menit > batas:
                    res.append({
                        "Tanggal": curr["Tanggal"],
                        "Personnel ID": curr["Personnel ID"],
                        "Nama": f"{curr['First Name']} {curr['Last Name']}",
                        "Department Name" : curr.get("Department Name"),
                        "Plant-Out": plant_out,
                        "Plant-In Berikutnya": plant_in_next,
                        "Shift In": curr["Shift In"],
                        "Shift Out": curr["Shift Out"],
                        "Jeda Menit": round(jeda_menit, 2),
                        "Selisih": round(jeda_menit - batas, 2),
                        "Keterangan": f"Jeda melebihi {batas} menit"
                    })
                    i += 2 # Skip next row karena sudah dipasangkan
                    continue
            i += 1
        return pd.DataFrame(res)
    
    # Eksekusi Sheet 4
    if not df_event.empty:
        df_sheet4 = pd.concat([cek_jeda_shift(g) for _, g in df_event.groupby("Personnel ID")], ignore_index=True)
        # Filter tambahan: Jeda aneh di atas 220 menit biasanya bukan istirahat tapi shift split/lembur
        df_sheet4 = df_sheet4[df_sheet4["Jeda Menit"] <= 220].reset_index(drop=True)
    else:
        df_sheet4 = pd.DataFrame()
        
    df_sheet4 = apply_exclusion(df_sheet4, exclude_dept=DEPT_EXCLUDE_MAIN)

    # --- LOGIKA SHEET 5: KELUAR > 2 KALI ---
    def cek_lebih_dari_2_per_shift(group):
        res = []
        group = group.sort_values("datetime").reset_index(drop=True)
        
        si = group.loc[0, "Shift In"]
        so = group.loc[0, "Shift Out"]
        if pd.isna(si) or pd.isna(so): return pd.DataFrame()
        
        tgl = group.loc[0, "Tanggal"]
        s_in = datetime.combine(tgl, si)
        s_out = datetime.combine(tgl, so)
        if s_out <= s_in: s_out += timedelta(days=1)
        
        # Ambil event hanya dalam jam kerja
        evs = group[(group["datetime"] >= s_in) & (group["datetime"] <= s_out) & (group["Event Point"].isin(["Plant-In", "Plant-Out"]))]
        
        jin = (evs["Event Point"] == "Plant-In").sum()
        jout = (evs["Event Point"] == "Plant-Out").sum()
        
        # Jumlah pasangan trip (keluar-masuk)
        jumlah_pasangan = min(jin, jout)
        
        if jumlah_pasangan > 2:
            res.append({
                "Tanggal": tgl,
                "Personnel ID": group.loc[0, "Personnel ID"],
                "Nama": f"{group.loc[0,'First Name']} {group.loc[0,'Last Name']}",
                "Department Name": group.loc[0, "Department Name"],
                "Shift In": si,
                "Shift Out": so,
                "Jumlah Plant-In": jin,
                "Jumlah Plant-Out": jout,
                "Jumlah Pasangan": jumlah_pasangan,
                "Keterangan": "Plant-In & Plant-Out > 2 kali dalam 1 shift"
            })
        return pd.DataFrame(res)
    
    # Eksekusi Sheet 5
    if not df_event.empty:
        df_sheet5 = pd.concat([cek_lebih_dari_2_per_shift(g) for _, g in df_event.groupby(["Personnel ID", "Tanggal", "Shift In", "Shift Out"])], ignore_index=True)
    else:
        df_sheet5 = pd.DataFrame()
        
    df_sheet5 = apply_exclusion(df_sheet5, exclude_dept=DEPT_EXCLUDE_MAIN + DEPT_EXCLUDE_EVENT)

    # --- LOGIKA SHEET 6: KELUAR MASUK SELAMA SHIFT (FILTER ISTIRAHAT) ---
    def cek_event_shift(group):
        res = []
        group = group.sort_values("datetime").reset_index(drop=True)
        si = group.loc[0, "Shift In"]
        so = group.loc[0, "Shift Out"]
        if pd.isna(si): return pd.DataFrame()
        
        tgl = group.loc[0, "Tanggal"]
        win = datetime.combine(tgl, si)
        wout = datetime.combine(tgl, so)
        if wout <= win: wout += timedelta(days=1)
        
        for _, row in group.iterrows():
            if row["Event Point"] not in ["Plant-In", "Plant-Out"]: continue
            et = row["datetime"]
            
            # Hanya cek event DI DALAM jam kerja
            if not (win <= et <= wout): continue
            
            # Filter 1: Jangan anggap pelanggaran jika dekat jam datang/pulang (+- 60 menit)
            if (win - timedelta(minutes=60) <= et <= win + timedelta(minutes=60)) or \
               (wout - timedelta(minutes=60) <= et <= wout + timedelta(minutes=60)): 
                continue
            
            # Filter 2: Jangan anggap pelanggaran jika di jam istirahat
            di = False
            for m, s in jam_istirahat:
                im = datetime.combine(tgl, m)
                isel = datetime.combine(tgl, s)
                # Toleransi +- 60 menit dari jam istirahat
                if (im - timedelta(minutes=60) <= et <= im + timedelta(minutes=60)) or \
                   (isel - timedelta(minutes=60) <= et <= isel + timedelta(minutes=60)): 
                    di = True
                    break
            if di: continue
            
            res.append({
                "Tanggal": tgl,
                "Personnel ID": row["Personnel ID"],
                "Nama": f"{row['First Name']} {row['Last Name']}",
                "Department Name": row["Department Name"],
                "Event": row["Event Point"],
                "Waktu Event": et,
                "Shift In": si,
                "Shift Out": so,
                "Keterangan": "Keluar/Masuk selama shift"
            })
        return pd.DataFrame(res)
    
    # Eksekusi Sheet 6
    if not df_event.empty:
        df_sheet6 = pd.concat([cek_event_shift(g) for _, g in df_event.groupby(["Personnel ID", "Tanggal", "Shift In", "Shift Out"])], ignore_index=True)
    else:
        df_sheet6 = pd.DataFrame()
        
    df_sheet6 = apply_exclusion(df_sheet6, exclude_dept=DEPT_EXCLUDE_MAIN + DEPT_EXCLUDE_EVENT)

    # --- LOGIKA SHEET 9: DETAIL TERLAMBAT ---
    df_sheet9 = df_shift[df_shift["Attendance Code + Name In"].astype(str).str.lower().str.contains("late", na=False)].copy()
    
    if not df_sheet9.empty:
        if "Employee Name" in df_sheet9.columns:
            df_sheet9["Nama"] = df_sheet9["Employee Name"]
        
        # Filter Exclude SEKARANG BISA BEKERJA karena Dept Name sudah ada
        df_sheet9 = apply_exclusion(df_sheet9, exclude_dept=DEPT_EXCLUDE_MAIN)
        
        # Select kolom output
        cols_final = ["Tanggal", "Personnel ID", "Nama", "Attendance Time In", "Attendance Time Out", "Shift In", "Shift Out", "Department Name"]
        # Pastikan kolom ada
        cols_final = [c for c in cols_final if c in df_sheet9.columns]
        
        df_sheet9 = df_sheet9[cols_final].rename(columns={
            "Personnel ID":"ID", "Nama":"Siapa", "Attendance Time In":"Attendance In", "Attendance Time Out":"Attendance Out"
        })
    else:
        df_sheet9 = pd.DataFrame(columns=["ID", "Siapa", "Attendance In", "Attendance Out", "Department Name"])

    # --- LOGIKA SHEET 7: REKAP TOTAL ---
    signal_status.emit("Membuat Rekap...")
    signal_progress.emit(80)
    
    def count_occ(df, col):
        if df.empty: return pd.DataFrame(columns=["Personnel ID", "Nama", "Department Name", col])
        return df.groupby(["Personnel ID", "Nama", "Department Name"]).size().reset_index(name=col)

    s3 = count_occ(df_sheet3, "Pulang sebelum shift selesai")
    s4 = count_occ(df_sheet4, "Terlalu lama Istirahat")
    s5 = count_occ(df_sheet5, "keluar > 2 dalam sehari")
    s6 = count_occ(df_sheet6, "Keluar masuk selama shift")
    
    # Hitung Telat dari Sheet 9 yang SUDAH DIFILTER DEPT-NYA
    if not df_sheet9.empty:
        s_telat = df_sheet9.groupby(["ID", "Siapa", "Department Name"]).size().reset_index(name="Jumlah Telat (Present Late)").rename(columns={"ID": "Personnel ID", "Siapa": "Nama"})
    else:
        s_telat = pd.DataFrame(columns=["Personnel ID", "Nama", "Department Name", "Jumlah Telat (Present Late)"])

    # Merge OUTER untuk menggabungkan semua pelanggaran
    df_sheet7 = s3.merge(s4, on=["Personnel ID", "Nama", "Department Name"], how="outer") \
                  .merge(s5, on=["Personnel ID", "Nama", "Department Name"], how="outer") \
                  .merge(s6, on=["Personnel ID", "Nama", "Department Name"], how="outer") \
                  .merge(s_telat, on=["Personnel ID", "Nama", "Department Name"], how="outer")
    
    # Isi NaN dengan 0
    cols_score = ["Pulang sebelum shift selesai", "Terlalu lama Istirahat", "keluar > 2 dalam sehari", "Keluar masuk selama shift", "Jumlah Telat (Present Late)"]
    for c in cols_score:
        if c not in df_sheet7.columns: df_sheet7[c] = 0
    
    df_sheet7 = df_sheet7.fillna(0)
    df_sheet7 = df_sheet7.sort_values(["Department Name", "Nama"]).reset_index(drop=True)
    
    # Final check exclusion (double safety)
    df_sheet7 = apply_exclusion(df_sheet7, exclude_dept=DEPT_EXCLUDE_MAIN)

    # --- LOGIKA SHEET 8: TELAT & LEMBUR ---
    def cek_tl(row):
        status = str(row.get("Attendance Code + Name In", "")).lower()
        if "late" not in status: return None
        
        # Cek kolom lembur
        ots = ["Overtime Weight  1.5", "Overtime Weight  2", "Overtime Weight  3", "Overtime Weight  4", "Overtime Weight Hour"]
        al = False
        for c in ots:
            v = pd.to_numeric(row.get(c), errors="coerce")
            if v and v != 0:
                # Pengecualian: kadang OT 1.5 yg nilainya 0.5 dianggap auto, bisa di-skip jika perlu.
                # Disini kita anggap semua OT valid jika > 0 (kecuali aturan spesifik 0.5 diabaikan)
                if c == "Overtime Weight  1.5" and v == 0.5: continue 
                al = True
                break
        
        if not al: return None
        
        # Hitung menit telat
        st = None
        try:
            si = row.get("Shift In")
            ai = row.get("Attendance Time In")
            if pd.notna(si) and pd.notna(ai):
                dt1 = datetime.combine(datetime.today(), pd.to_datetime(str(si)).time())
                dt2 = datetime.combine(datetime.today(), pd.to_datetime(str(ai)).time())
                if dt2 > dt1: 
                    st = round((dt2 - dt1).total_seconds()/60, 2)
        except: pass
        
        return {
            "Employee Name": row.get("Employee Name"),
            "Shift In": row.get("Shift In"), "Shift Out": row.get("Shift Out"),
            "Time In": row.get("Attendance Time In"), "Time Out": row.get("Attendance Time Out"),
            "Date Out": row.get("Attendance Date Out"),
            "Selisih Telat (Menit)": st,
            "1.5": row.get("Overtime Weight  1.5"), "2": row.get("Overtime Weight  2"),
            "3": row.get("Overtime Weight  3"), "4": row.get("Overtime Weight  4"),
            "Overtime Weight Hour": row.get("Overtime Weight Hour"),
            "Keterangan": "Telat dan Lembur"
        }

    rows_tl = []
    for _, r in df_shift.iterrows():
        res_tl = cek_tl(r)
        if res_tl: rows_tl.append(res_tl)
    df_sheet8 = pd.DataFrame(rows_tl)

    # CLEANING FINAL & EXPORT EXCEL
    signal_status.emit("Menyimpan File Excel...")
    signal_progress.emit(90)

    # Standardisasi Nama Departemen
    list_df = [df_selisih, df_sheet2, df_sheet3, df_sheet4, df_sheet5, df_sheet6, df_sheet7, df_sheet8, df_sheet9]
    for d in list_df:
        if "Department Name" in d.columns:
            d["Department Name"] = d["Department Name"].replace("PT. Sumber IndahPerkasa", "Quality Control Section")
            d["Department Name"] = d["Department Name"].replace("Pressing Section", "KCP Department")
        if "Nama" in d.columns:
            d["Nama"] = clean_nama_series(d["Nama"])

    # Tulis ke Excel
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        if not df_sheet3.empty: df_sheet3.to_excel(writer, sheet_name="Keluar Sebelum Shift Selesai", index=False)
        if not df_sheet4.empty: df_sheet4.to_excel(writer, sheet_name="Jeda Plant-Out ke Plant-In", index=False)
        if not df_sheet5.empty: df_sheet5.to_excel(writer, sheet_name="Plant-In Out > 2x Per Shift", index=False)
        if not df_sheet6.empty: df_sheet6.to_excel(writer, sheet_name="Keluar masuk selama shift", index=False)
        if not df_sheet8.empty: df_sheet8.to_excel(writer, sheet_name="Telat dan lembur", index=False)
        if not df_sheet9.empty: df_sheet9.to_excel(writer, sheet_name="Detail Terlambat", index=False)
        if not df_sheet7.empty: df_sheet7.to_excel(writer, sheet_name="Rekap Total", index=False)

    # VISUALISASI PDF 
    try:
        signal_status.emit("Membuat Visualisasi PDF...")
        output_pdf = output_excel.replace(".xlsx", ".pdf")
        if output_pdf == output_excel: output_pdf += ".pdf"

        # Daftar kolom pelanggaran yang akan divisualisasikan
        cols_viz = ["Pulang sebelum shift selesai", "Terlalu lama Istirahat", "Keluar masuk selama shift", "Jumlah Telat (Present Late)"]
        # Validasi: Hanya ambil kolom yang benar-benar ada di data
        cols_exist = [c for c in cols_viz if c in df_sheet7.columns]

        if cols_exist and not df_sheet7.empty:
            with PdfPages(output_pdf) as pdf:
                # 1. CHART: Total per Jenis Pelanggaran (Bar Horizontal)
                tp = df_sheet7[cols_exist].sum().sort_values()
                plt.figure(figsize=(10, 6))
                plt.barh(tp.index, tp.values, color="#e74c3c")
                for i, v in enumerate(tp.values): 
                    plt.text(v, i, f" {int(v)}", va="center")
                plt.title("Jenis Pelanggaran Paling Dominan")
                plt.xlabel("Jumlah Kejadian")
                plt.tight_layout()
                pdf.savefig()
                plt.close()

                # 2. CHART: Proporsi (Pie Chart)
                plt.figure(figsize=(7, 7))
                plt.pie(tp.values, labels=tp.index, autopct="%1.1f%%", startangle=90)
                plt.title(f"Proporsi Jenis Pelanggaran\nTotal: {int(tp.sum())}")
                plt.tight_layout()
                pdf.savefig()
                plt.close()

                # 3. CHART: Distribusi Departemen (Stacked Bar) - Top 15
                dj = df_sheet7.groupby("Department Name")[cols_exist].sum()
                dj["Total"] = dj.sum(axis=1)
                dj = dj.sort_values("Total", ascending=True) # Ascending untuk Barh
                
                # Ambil Top 15 agar grafik tidak terlalu padat
                if len(dj) > 15: dj = dj.tail(15)
                
                if not dj.empty:
                    ax = dj.drop(columns="Total").plot(kind="barh", stacked=True, figsize=(12, 8), colormap="viridis")
                    plt.title("Distribusi Pelanggaran per Departemen (Top 15)")
                    plt.legend(bbox_to_anchor=(1.02, 1))
                    for i, (idx, row) in enumerate(dj.iterrows()):
                        if row["Total"] > 0:
                            plt.text(row["Total"], i, f" {int(row['Total'])}", va="center", fontsize=9)
                            
                    plt.tight_layout()
                    pdf.savefig()
                    plt.close()

                # 4. CHART: Top 10 Karyawan (STACKED BAR - RINCIAN)
                df_sheet7["Total Pelanggaran"] = df_sheet7[cols_exist].sum(axis=1)
                
                # Ambil Top 10, urutkan Ascending agar yang terbanyak ada di atas chart
                top10 = df_sheet7.sort_values("Total Pelanggaran", ascending=False).head(10).sort_values("Total Pelanggaran", ascending=True)
                if not top10.empty:
                    top10_stack = top10.copy()
                    top10_stack["Label"] = top10_stack["Nama"] + " (" + top10_stack["Department Name"] + ")"
                    top10_stack = top10_stack.set_index("Label")
                    ax = top10_stack[cols_exist].plot(kind="barh", stacked=True, figsize=(12, 7), colormap="coolwarm")
                    
                    plt.title("Rincian Pelanggaran Top 10 Karyawan")
                    plt.xlabel("Jumlah Pelanggaran")
                    plt.ylabel("")
                    plt.legend(bbox_to_anchor=(1.02, 1), title="Jenis Pelanggaran")
                    for i, v in enumerate(top10_stack["Total Pelanggaran"]):
                        plt.text(v, i, f" {int(v)}", va="center", fontweight='bold')
                        
                    plt.tight_layout()
                    pdf.savefig()
                    plt.close()

                # 5. LOOPS: Detail per Jenis Pelanggaran (Dept & Karyawan)
                for col in cols_exist:
                    # --- A. Top 10 Departemen per Jenis Pelanggaran ---
                    dept_sum = df_sheet7.groupby("Department Name")[col].sum().sort_values(ascending=True)
                    dept_sum = dept_sum[dept_sum > 0] 
                    if not dept_sum.empty:
                        top10_dept_jenis = dept_sum.tail(10)
                        plt.figure(figsize=(10, 6))
                        bars = plt.barh(top10_dept_jenis.index, top10_dept_jenis.values, color="#2ecc71")
                        plt.title(f"Top 10 Departemen - {col}")
                        plt.xlabel("Jumlah Kejadian")
                        for bar in bars:
                            width = bar.get_width()
                            plt.text(width, bar.get_y() + bar.get_height()/2, f" {int(width)}", va="center")
                        plt.tight_layout()
                        pdf.savefig()
                        plt.close()

                    # --- B. Top 10 Karyawan per Jenis Pelanggaran ---
                    # Filter data > 0 dan sort
                    emp_jenis = df_sheet7[df_sheet7[col] > 0].sort_values(col, ascending=True)
                    
                    if not emp_jenis.empty:
                        top10_emp = emp_jenis.tail(10)
                        labels_emp = top10_emp["Nama"] + " (" + top10_emp["Department Name"] + ")"
                        plt.figure(figsize=(10, 6))
                        bars = plt.barh(labels_emp, top10_emp[col], color="orange")
                        plt.title(f"Top 10 Karyawan - {col}")
                        plt.xlabel("Jumlah Kejadian")
                        for bar in bars:
                            width = bar.get_width()
                            plt.text(width, bar.get_y() + bar.get_height()/2, f" {int(width)}", va="center")
                            
                        plt.tight_layout()
                        pdf.savefig()
                        plt.close()

                # 6. Heatmap Detail per Departemen
                for dept in sorted(df_sheet7["Department Name"].unique()):
                    d_f = df_sheet7[df_sheet7["Department Name"] == dept].set_index("Nama")[cols_exist]
                    # Skip jika departemen bersih tanpa pelanggaran
                    if d_f.empty or d_f.sum().sum() == 0: continue
                    # Atur tinggi gambar dinamis berdasarkan jumlah karyawan
                    h = max(5, len(d_f) * 0.4)
                    plt.figure(figsize=(12, h))
                    im = plt.imshow(d_f.values, aspect="auto", cmap="Reds")
                    plt.colorbar(im, label="Jumlah")
                    plt.yticks(range(len(d_f)), d_f.index)
                    plt.xticks(range(len(cols_exist)), cols_exist, rotation=45, ha="right")
                    # Tampilkan angka di dalam kotak heatmap
                    for i in range(len(d_f.index)):
                        for j in range(len(cols_exist)):
                            val = d_f.iloc[i, j]
                            if val > 0: 
                                plt.text(j, i, int(val), ha="center", va="center")
                    plt.title(f"Detail Pelanggaran: {dept}")
                    plt.tight_layout()
                    pdf.savefig()
                    plt.close()
    except Exception as e:
        print(f"Error PDF: {e}")
# BAGIAN 2: LOGIKA MODUL 2 (PEMERIKSA LEMBUR)
def clean_time_l(jam):
    if pd.isna(jam): return None
    return str(jam).split()[0]

def to_dt_l(tgl, jam):
    jam = clean_time_l(jam)
    if jam is None: return None
    return pd.to_datetime(f"{tgl} {jam}", errors="coerce")

def deteksi_lembur(row):
    try:
        am = to_dt_l(row["Tanggal"], row["Absen Masuk"])
        ak = to_dt_l(row["Tanggal"], row["Absen Keluar"])
        tin = to_dt_l(row["Tanggal"], row["In"])
        tout = to_dt_l(row["Tanggal"], row["Out"])
        
        if None in (am, ak, tin, tout): return None, None
        
        # Handling Lintas Hari
        if am > tout: am -= timedelta(days=1)
        if ak < tin: ak += timedelta(days=1)
        if tout < tin: tout += timedelta(days=1)

        # Cek Lembur Awal
        if tin < am: 
            return "Lembur Awal, Telat Masuk", int((am - tin).total_seconds() / 60)
        
        # Cek Lembur Akhir
        if tout > ak: 
            return "Lembur Akhir, Pulang Cepat", int((tout - ak).total_seconds() / 60)
            
    except: pass
    return None, None

# BAGIAN 3: GUI & THREADING (INTERFACE UTAMA)
class Modul1Worker(QThread):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, t, s, o):
        super().__init__()
        self.t, self.s, self.o = t, s, o

    def run(self):
        try:
            process_logic(self.t, self.s, self.o, self.status, self.progress)
            self.progress.emit(100)
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

class Modul2Worker(QThread):
    progress = pyqtSignal(int)
    data_ready = pyqtSignal(list)

    def __init__(self, files):
        super().__init__()
        self.files = files

    def run(self):
        REQUIRED_COLS = ["Tanggal", "Absen Masuk", "Absen Keluar", "In", "Out"]
        all_data_list = []
        
        for i, file in enumerate(self.files):
            try:
                temp = pd.read_excel(file, header=None)
                target_df = None
                
                # Cari header di 20 baris pertama
                for idx in range(min(len(temp), 20)):
                    row_vals = [str(val).strip() for val in temp.iloc[idx].values]
                    if all(col in row_vals for col in REQUIRED_COLS):
                        target_df = pd.read_excel(file, header=idx)
                        break
                
                if target_df is not None:
                    target_df.columns = target_df.columns.astype(str).str.strip().str.replace('\n', ' ')
                    
                    for _, row in target_df.iterrows():
                        status, selisih = deteksi_lembur(row)
                        if status:
                            all_data_list.append([
                                file.split("/")[-1], 
                                str(row.get("Tanggal", "")).split()[0], 
                                str(row.get("Nama", "-")), 
                                str(row.get("Absen Masuk", "-")), 
                                str(row.get("Absen Keluar", "-")), 
                                str(row.get("In", "-")), 
                                str(row.get("Out", "-")), 
                                status, 
                                str(selisih)
                            ])
                
                self.progress.emit(int((i + 1) / len(self.files) * 100))
            except Exception as e:
                print(f"Error: {e}")

        self.data_ready.emit(all_data_list)

# GUI & THREADING
# C1: #383642 (Dark Grey/Blue) -> Background Aplikasi
# C2: #f0cba8 (Light Peach)    -> Teks, Icon, Hover Button
# C3: #b89372 (Tan Brown)      -> Tombol Utama (Action)
# C4: #706258 (Taupe Brown)    -> Input Field, Tombol Menu

STYLESHEET = """
    /* 1. GLOBAL SETTING */
    QWidget {
        font-family: 'Segoe UI', sans-serif;
        background-color: #383642; /* C1: Background Utama */
        color: #f0cba8;            /* C2: Warna Teks Utama */
    }

    /* 2. TEMPAT MENULIS (INPUT FIELD) */
    QLineEdit {
        background-color: #706258; /* C4: Background Input */
        color: #ffffff;            /* Putih agar kontras di coklat tua */
        border: 2px solid #b89372; /* C3: Border */
        border-radius: 10px;
        padding: 10px 15px;
        font-size: 14px;
        selection-background-color: #f0cba8; /* C2 */
        selection-color: #383642;
    }
    QLineEdit:focus {
        border: 2px solid #f0cba8; /* C2: Border menyala saat diketik */
        background-color: #7a6b61; /* Sedikit lebih terang */
    }
    QLineEdit::placeholder {
        color: #d0c0b0; /* Warna placeholder samar */
        font-style: italic;
    }

    /* 3. TOMBOL BIASA (Browse/Simpan) */
    QPushButton {
        background-color: #b89372; /* C3: Tombol Utama */
        color: #383642;            /* Teks Gelap agar terbaca */
        border-radius: 10px;
        padding: 10px 20px;
        font-weight: bold;
        font-size: 13px;
        border: none;
    }
    QPushButton:hover {
        background-color: #f0cba8; /* C2: Berubah jadi terang saat hover */
        color: #383642;
        margin-top: -2px;
    }
    QPushButton:pressed {
        background-color: #706258; /* C4: Gelap saat ditekan */
        color: #f0cba8;
        margin-top: 2px;
    }

    /* 4. TOMBOL MENU BESAR (Di Halaman Awal) */
    QPushButton#BtnMenu {
        background-color: #706258; /* C4 */
        color: #f0cba8;            /* C2 */
        border: 2px solid #b89372; /* C3 */
        border-radius: 15px;
        padding: 25px;
        font-size: 16px;
        text-align: left;
    }
    QPushButton#BtnMenu:hover {
        background-color: #b89372; /* C3 */
        color: #383642;            /* C1 */
        border: 2px solid #f0cba8; /* C2 */
    }

    /* 5. TOMBOL AKSI UTAMA (Mulai Proses) - Dibuat Paling Terang */
    QPushButton#BtnAction {
        background-color: #f0cba8; /* C2: Paling mencolok */
        color: #383642;            /* C1: Teks gelap */
        font-size: 15px;
        padding: 15px;
        border-radius: 10px;
    }
    QPushButton#BtnAction:hover {
        background-color: #ffffff; /* Putih saat hover */
    }
    QPushButton#BtnAction:pressed {
        background-color: #b89372; /* C3 */
    }

    /* 6. GROUP BOX (Kotak Pembungkus) */
    QGroupBox {
        border: 2px solid #706258; /* C4 */
        border-radius: 10px;
        margin-top: 25px;
        font-weight: bold;
        color: #b89372; /* C3: Judul Group */
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 20px;
        padding: 0 5px;
    }

    /* 7. TABEL */
    QTableWidget {
        background-color: #383642; /* C1 */
        gridline-color: #706258;   /* C4 */
        color: #f0cba8;            /* C2 */
        border: 1px solid #706258;
    }
    QHeaderView::section {
        background-color: #706258; /* C4 */
        color: #f0cba8;            /* C2 */
        padding: 8px;
        border: 1px solid #383642;
        font-weight: bold;
    }
    
    /* 8. PROGRESS BAR */
    QProgressBar {
        border: 2px solid #706258;
        border-radius: 5px;
        text-align: center;
        background-color: #383642;
    }
    QProgressBar::chunk {
        background-color: #b89372; /* C3 */
        border-radius: 3px;
    }
    
    /* 9. TOMBOL KEMBALI */
    QPushButton#BtnBack {
        background-color: transparent;
        color: #706258; /* C4: Samar */
        border: 1px solid #706258;
        text-align: left;
    }
    QPushButton#BtnBack:hover {
        color: #f0cba8; /* C2 */
        border: 1px solid #f0cba8;
        background-color: #383642;
    }
"""

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tools Analisis HR - kajo-_-")
        self.setGeometry(100, 100, 1100, 750)
        self.setStyleSheet(STYLESHEET) 
        
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)
        
        self.init_menu_halaman()
        self.init_modul_1_ui()
        self.init_modul_2_ui()
        
    def init_menu_halaman(self):
        halaman = QWidget()
        layout = QVBoxLayout(halaman)
        layout.setContentsMargins(70, 70, 70, 70)
        layout.setSpacing(25)
        
        # Header
        title = QLabel("HR ANALYTICS")
        title.setFont(QFont("Segoe UI", 36, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # Warna Judul Pakai C2 (Peach) agar kontras
        title.setStyleSheet("color: #f0cba8; letter-spacing: 3px; margin-bottom: 5px;")
        layout.addWidget(title)
        
        subtitle = QLabel("Pusat Kendali & Otomatisasi Data")
        subtitle.setFont(QFont("Segoe UI", 14))
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # Warna Subjudul Pakai C3 (Tan)
        subtitle.setStyleSheet("color: #b89372; margin-bottom: 40px;")
        layout.addWidget(subtitle)
        
        # Tombol Menu (Style BtnMenu)
        btn_modul1 = QPushButton("  📂  ANALISIS ABSENSI & KELUAR MASUK PLAN\n       (Generate PDF Report & Excel Rekapitulasi)")
        btn_modul1.setObjectName("BtnMenu") 
        btn_modul1.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_modul1.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(1))
        
        btn_modul2 = QPushButton("  🔍  SCANNER ANOMALI DATA LEMBUR\n       (Deteksi Otomatis & Live Preview Table)")
        btn_modul2.setObjectName("BtnMenu")
        btn_modul2.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_modul2.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(2))
        
        layout.addWidget(btn_modul1)
        layout.addWidget(btn_modul2)
        layout.addStretch()
        
        lbl_credit = QLabel("Developed by: kajo-_-")
        lbl_credit.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        lbl_credit.setStyleSheet("color: #706258;") # Warna C4 samar
        lbl_credit.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_credit)
        
        self.stacked_widget.addWidget(halaman)

    def init_modul_1_ui(self):
        halaman = QWidget()
        layout = QVBoxLayout(halaman)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        # Tombol Kembali
        btn_back = QPushButton("← KEMBALI")
        btn_back.setFixedWidth(120)
        btn_back.setObjectName("BtnBack")
        btn_back.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_back.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(0))
        layout.addWidget(btn_back)

        title = QLabel("Analisis Absensi Plant")
        title.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        title.setStyleSheet("color: #f0cba8; margin-bottom: 10px;")
        layout.addWidget(title)

        # Group Box Input
        grp_in = QGroupBox(" UPLOAD DATA ")
        grid_in = QGridLayout(grp_in)
        grid_in.setVerticalSpacing(20)
        grid_in.setHorizontalSpacing(15)
        grid_in.setContentsMargins(20, 35, 20, 25)
        
        # Komponen Input
        self.m1_txt_trans = QLineEdit(); self.m1_txt_trans.setPlaceholderText("File Transaksi Keluar Masuk Plan (.xlsx)...")
        btn_br1 = QPushButton("PILIH FILE")
        btn_br1.clicked.connect(lambda: self.browse_file(self.m1_txt_trans))

        self.m1_txt_shift = QLineEdit(); self.m1_txt_shift.setPlaceholderText("File Attendence (.xlsx)...")
        btn_br2 = QPushButton("PILIH FILE")
        btn_br2.clicked.connect(lambda: self.browse_file(self.m1_txt_shift))

        self.m1_txt_out = QLineEdit(); self.m1_txt_out.setPlaceholderText("Simpan hasil sebagai (xlsx & pdf)...")
        btn_br3 = QPushButton("LOKASI SIMPAN")
        btn_br3.clicked.connect(self.browse_save)

        # Label (Warna C2)
        l1 = QLabel("Data Transaksi:"); l1.setStyleSheet("color: #f0cba8;")
        l2 = QLabel("Data Shift:"); l2.setStyleSheet("color: #f0cba8;")
        l3 = QLabel("Output Hasil:"); l3.setStyleSheet("color: #f0cba8;")

        grid_in.addWidget(l1, 0, 0); grid_in.addWidget(self.m1_txt_trans, 0, 1); grid_in.addWidget(btn_br1, 0, 2)
        grid_in.addWidget(l2, 1, 0); grid_in.addWidget(self.m1_txt_shift, 1, 1); grid_in.addWidget(btn_br2, 1, 2)
        grid_in.addWidget(l3, 2, 0); grid_in.addWidget(self.m1_txt_out, 2, 1); grid_in.addWidget(btn_br3, 2, 2)
        
        layout.addWidget(grp_in)

        # Tombol Proses Besar (Style BtnAction)
        self.m1_btn_run = QPushButton("MULAI PROSES ANALISIS")
        self.m1_btn_run.setObjectName("BtnAction")
        self.m1_btn_run.setCursor(Qt.CursorShape.PointingHandCursor)
        self.m1_btn_run.clicked.connect(self.start_modul_1)
        layout.addWidget(self.m1_btn_run)

        # Progress & Status
        self.m1_pbar = QProgressBar()
        self.m1_pbar.setFixedHeight(20)
        self.m1_pbar.setTextVisible(False)
        layout.addWidget(self.m1_pbar)
        
        self.m1_lbl_stat = QLabel("Menunggu input user...")
        self.m1_lbl_stat.setStyleSheet("color: #706258; font-style: italic;")
        self.m1_lbl_stat.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.m1_lbl_stat)
        layout.addStretch()

        self.stacked_widget.addWidget(halaman)

    def init_modul_2_ui(self):
        halaman = QWidget()
        layout = QVBoxLayout(halaman)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)

        btn_back = QPushButton("← KEMBALI")
        btn_back.setFixedWidth(120)
        btn_back.setObjectName("BtnBack")
        btn_back.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_back.clicked.connect(lambda: self.stacked_widget.setCurrentIndex(0))
        layout.addWidget(btn_back)

        title = QLabel("Pemeriksa Data Lembur")
        title.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        title.setStyleSheet("color: #f0cba8; margin-bottom: 5px;")
        layout.addWidget(title)

        # Tombol Select (Style BtnAction)
        self.m2_btn_select = QPushButton("📂  PILIH FILE EXCEL & SCAN SEKARANG")
        self.m2_btn_select.setObjectName("BtnAction")
        self.m2_btn_select.setCursor(Qt.CursorShape.PointingHandCursor)
        self.m2_btn_select.clicked.connect(self.start_modul_2)
        layout.addWidget(self.m2_btn_select)

        self.m2_pbar = QProgressBar()
        self.m2_pbar.setFixedHeight(8)
        self.m2_pbar.setTextVisible(False)
        layout.addWidget(self.m2_pbar)

        # Tabel
        self.m2_table = QTableWidget()
        self.m2_table.setAlternatingRowColors(True)
        # Warna alternatif baris tabel (C4 gelap dan C1)
        self.m2_table.setStyleSheet("QTableWidget { alternate-background-color: #454252; }")
        headers = ["Sumber File", "Tanggal", "Nama", "Abs. Masuk", "Abs. Keluar", "In", "Out", "Status", "Menit"]
        self.m2_table.setColumnCount(len(headers))
        self.m2_table.setHorizontalHeaderLabels(headers)
        self.m2_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.m2_table)

        self.stacked_widget.addWidget(halaman)

    # --- LOGIC FUNCTIONS ---
    def browse_file(self, edit):
        f, _ = QFileDialog.getOpenFileName(self, "Pilih File", "", "Excel Files (*.xlsx *.xls)")
        if f: edit.setText(f)

    def browse_save(self):
        f, _ = QFileDialog.getSaveFileName(self, "Simpan Ke", "", "Excel Files (*.xlsx)")
        if f:
            if not f.endswith(".xlsx"): f += ".xlsx"
            self.m1_txt_out.setText(f)

    def start_modul_1(self):
        if not all([self.m1_txt_trans.text(), self.m1_txt_shift.text(), self.m1_txt_out.text()]):
            QMessageBox.warning(self, "Peringatan", "Mohon lengkapi semua file input dan output!")
            return
        
        self.m1_btn_run.setEnabled(False)
        self.m1_btn_run.setText("⏳ SEDANG MEMPROSES...")
        self.m1_lbl_stat.setText("Sedang menganalisis data, mohon tunggu...")
        self.m1_lbl_stat.setStyleSheet("color: #f0cba8; font-weight: bold;")
        
        self.m1_worker = Modul1Worker(self.m1_txt_trans.text(), self.m1_txt_shift.text(), self.m1_txt_out.text())
        self.m1_worker.progress.connect(self.m1_pbar.setValue)
        self.m1_worker.status.connect(self.m1_lbl_stat.setText)
        self.m1_worker.finished.connect(self.modul1_done)
        self.m1_worker.error.connect(self.modul1_fail)
        self.m1_worker.start()

    def modul1_done(self):
        self.m1_btn_run.setEnabled(True)
        self.m1_btn_run.setText("MULAI PROSES ANALISIS")
        self.m1_lbl_stat.setText("✅ Selesai! File berhasil disimpan.")
        self.m1_lbl_stat.setStyleSheet("color: #b89372; font-weight: bold;") # C3
        QMessageBox.information(self, "Sukses", "Analisis Selesai!\nFile Excel dan PDF telah dibuat.")

    def modul1_fail(self, msg):
        self.m1_btn_run.setEnabled(True)
        self.m1_btn_run.setText("MULAI PROSES ANALISIS")
        self.m1_lbl_stat.setText("❌ Terjadi Kesalahan.")
        self.m1_lbl_stat.setStyleSheet("color: #f0cba8;") # C2
        QMessageBox.critical(self, "Error", f"Terjadi kesalahan:\n{msg}")

    def start_modul_2(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Pilih File", "", "Excel Files (*.xlsx *.xls)")
        if files:
            self.m2_table.setRowCount(0)
            self.m2_worker = Modul2Worker(files)
            self.m2_worker.progress.connect(self.m2_pbar.setValue)
            self.m2_worker.data_ready.connect(self.display_modul_2)
            self.m2_worker.start()

    def display_modul_2(self, data):
        self.m2_table.setRowCount(len(data))
        for r, row_data in enumerate(data):
            for c, val in enumerate(row_data):
                self.m2_table.setItem(r, c, QTableWidgetItem(str(val)))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())