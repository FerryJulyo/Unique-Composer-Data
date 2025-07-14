import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from rapidfuzz import fuzz
import os
import time

# --- GUI untuk pilih file ---
print("📂 Menampilkan dialog pilih file Excel...")
Tk().withdraw()
file_path = askopenfilename(title="Pilih file Excel", filetypes=[("Excel Files", "*.xlsx *.xls")])
if not file_path:
    print("❌ Tidak ada file dipilih.")
    exit()

print(f"📄 File dipilih: {file_path}")

# --- Baca sheet All Data ---
print("📥 Membaca sheet 'All Data'...")
df = pd.read_excel(file_path, sheet_name="All Data")

# Ambil hanya kolom ID dan Csong
print("🔍 Membersihkan data dan mengambil kolom ID & Csong...")
df = df[['ID', 'Csong']].dropna()
df['Csong'] = df['Csong'].str.strip()

total_rows = len(df)
print(f"📊 Total baris dibaca: {total_rows}")

# --- Fuzzy Unique Matching (≥ 96) ---
print("⚙️ Memulai proses pencocokan fuzzy (≥ 96)...\n")
unique_rows = []
seen = []

start_time = time.time()
for i, (_, row) in enumerate(df.iterrows(), 1):
    csong = row['Csong']
    is_duplicate = False
    for existing in seen:
        if fuzz.token_sort_ratio(csong, existing) >= 96:
            is_duplicate = True
            break
    if not is_duplicate:
        unique_rows.append(row)
        seen.append(csong)
    
    percent = (i / total_rows) * 100
    print(f"   🔄 Baris {i}/{total_rows} -> {'✔️ unik' if not is_duplicate else '⏩ duplikat'}  ({percent:.2f}%)")

elapsed = time.time() - start_time

# --- Hasil akhir sebagai DataFrame ---
result_df = pd.DataFrame(unique_rows)
print(f"\n✅ Total unik ditemukan: {len(result_df)} dari {total_rows} baris")
print(f"⏱️ Waktu proses: {elapsed:.2f} detik")

# --- Simpan ke file Excel ---
output_path = os.path.join(os.path.dirname(file_path), "unique_csong_output.xlsx")
result_df.to_excel(output_path, index=False)
print(f"💾 Hasil disimpan ke: {output_path}")
