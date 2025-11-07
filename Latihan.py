import pandas as pd
import numpy as np
import os

# Pastikan script berjalan di folder yang sama dengan file Excel
os.chdir(os.path.dirname(__file__))

# -----------------------------------------------------
# 1. Import Data
# -----------------------------------------------------
data = pd.read_excel("Data Wisudawan.xlsx")

# -----------------------------------------------------
# 2. Normalisasi Nama Kolom
# -----------------------------------------------------
data.columns = data.columns.str.strip().str.lower().str.replace(r'\s+', ' ', regex=True)
print("Nama kolom terdeteksi:", data.columns.tolist())

# -----------------------------------------------------
# 3. Bersihkan Data
# -----------------------------------------------------
data['nama mahasiswa'] = data['nama mahasiswa'].astype(str).str.title().str.strip()
data['program studi'] = data['program studi'].replace('', np.nan)
data['ipk'] = data['ipk'].fillna(0)

kolom_lama_studi = [c for c in data.columns if 'lama studi' in c][0]
data[kolom_lama_studi] = data[kolom_lama_studi].fillna(0)

before = len(data)
data = data.dropna(subset=['program studi'])
after = len(data)
print(f"Data tanpa Program Studi dihapus: {before - after}")

# -----------------------------------------------------
# 4. Bersihkan Typo dan Filter IPK
# -----------------------------------------------------
before = len(data)
data = data[~data['program studi'].str.contains("TRLP", case=False, na=False)]
after = len(data)
print(f"Data dengan Program Studi 'TRLP' dihapus: {before - after}")

data['program studi'] = data['program studi'].replace({'TPPLL': 'TPPL'})

before = len(data)
data = data[data['ipk'] != 0]
after = len(data)
print(f"Data dengan IPK = 0 dihapus: {before - after}")

data = data[(data['ipk'] > 0.0) & (data['ipk'] <= 4.0)]
data.loc[(data[kolom_lama_studi] < 4) | (data[kolom_lama_studi] > 14), kolom_lama_studi] = np.nan

# -----------------------------------------------------
# 5. Aturan D3 dan D4
# -----------------------------------------------------
before = len(data)
data = data[~((data['program studi'].str.contains("D3", case=False)) & (data[kolom_lama_studi] > 8))]
after = len(data)
print(f"Data D3 yang dihapus karena >8 semester: {before - after}")

before = len(data)
data = data[~((data['program studi'].str.contains("D4", case=False)) & (data[kolom_lama_studi] < 8))]
after = len(data)
print(f"Data D4 yang dihapus karena <8 semester: {before - after}")

# -----------------------------------------------------
# 6. Hapus Duplikat
# -----------------------------------------------------
before = len(data)
data = data.drop_duplicates(subset=['nim', 'nama mahasiswa'], keep='first')
after = len(data)
print(f"Data duplikat yang dihapus: {before - after}")

# -----------------------------------------------------
# 7. Tambah Kolom Grade, Predikat, Tahun Wisuda
# -----------------------------------------------------
data['Grade'] = [
    'A' if IPK >= 3.75 else
    'B+' if IPK >= 3.5 else
    'B' if IPK >= 3.0 else
    'C' if IPK >= 2.5 else
    'D'
    for IPK in data['ipk']
]

data['Predikat'] = [
    'Cumlaude' if (IPK >= 3.75 and study <= 8)
    else 'Sangat Memuaskan' if (IPK >= 3.5 and study <= 9)
    else 'Memuaskan' if IPK >= 3.0
    else 'Cukup'
    for IPK, study in zip(data['ipk'], data[kolom_lama_studi])
]

data['tahun wisuda'] = 2025

# -----------------------------------------------------
# 8. Urutkan Kolom Sesuai Format Akhir
# -----------------------------------------------------
Kolom_Urutan = [
    'nim',
    'nama mahasiswa',
    'program studi',
    'ipk',
    kolom_lama_studi,
    'Grade',
    'Predikat',
    'tahun wisuda'
]

# -----------------------------------------------------
# 9. Simpan ke Excel
# -----------------------------------------------------
data = data[Kolom_Urutan]
data.columns = [
    'NIM',
    'Nama Mahasiswa',
    'Program Studi',
    'IPK',
    'Lama Studi (Semester)',
    'Grade',
    'Predikat',
    'Tahun Wisuda'
]

data.to_excel("Data_Wisudawan_Final.xlsx", index=False, columns=data.columns)
print("\n✅ File akhir berhasil dibuat: Data_Wisudawan_Final.xlsx")
print(data)