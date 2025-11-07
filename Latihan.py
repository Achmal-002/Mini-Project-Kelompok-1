import pandas as pd
import numpy as np
import os

# Pastikan bekerja di folder tempat file ini berada
os.chdir(os.path.dirname(__file__))

# -----------------------------------------------------
# 1. Import Data
# -----------------------------------------------------
df = pd.read_excel("Data Wisudawan.xlsx")

# -----------------------------------------------------
# 2. Bersihkan Data Awal
# -----------------------------------------------------
# Rapikan nama kolom dan teks
df.columns = df.columns.str.strip().str.title()
df['Nama Mahasiswa'] = df['Nama Mahasiswa'].astype(str).str.title().str.strip()

# Tangani nilai kosong
df['Program Studi'] = df['Program Studi'].replace('', np.nan)
df['IPK'] = df['IPK'].fillna(0)
df['Lama Studi (Semester)'] = df['Lama Studi (Semester)'].fillna(0)

# Hapus baris tanpa Program Studi
before = len(df)
df = df.dropna(subset=['Program Studi'])
after = len(df)
print(f"\nData tanpa Program Studi dihapus: {before - after}")

# -----------------------------------------------------
# 3. Bersihkan Typo Program Studi dan Filter IPK
# -----------------------------------------------------
# Hapus baris dengan Program Studi 'TRLP' (typo)
before = len(df)
df = df[~df['Program Studi'].str.contains("TRLP", case=False, na=False)]
after = len(df)
print(f"Data dengan Program Studi 'TRLP' dihapus: {before - after}")

# Perbaiki typo 'TPPLL' menjadi 'TPPL'
df['Program Studi'] = df['Program Studi'].replace({'TPPLL': 'TPPL'})

# Hapus data dengan IPK = 0 (tidak valid)
before = len(df)
df = df[df['IPK'] != 0]
after = len(df)
print(f"Data dengan IPK = 0 dihapus: {before - after}")

# Pastikan IPK valid (0–4)
df = df[(df['IPK'] > 0.0) & (df['IPK'] <= 4.0)]

# Semester di luar batas (4–14) diganti NaN
df.loc[(df['Lama Studi (Semester)'] < 4) | (df['Lama Studi (Semester)'] > 14),
       'Lama Studi (Semester)'] = np.nan

# -----------------------------------------------------
# 4. Aturan Tambahan Khusus D3 & D4
# -----------------------------------------------------
# D3 tidak boleh lebih dari 8 semester
before = len(df)
df = df[~((df['Program Studi'].str.contains("D3", case=False)) &
          (df['Lama Studi (Semester)'] > 8))]
after = len(df)
print(f"Data D3 yang dihapus karena >8 semester: {before - after}")

# D4 tidak boleh kurang dari 8 semester
before = len(df)
df = df[~((df['Program Studi'].str.contains("D4", case=False)) &
          (df['Lama Studi (Semester)'] < 8))]
after = len(df)
print(f"Data D4 yang dihapus karena <8 semester: {before - after}")

# -----------------------------------------------------
# 5. Hapus Data Duplikat
# -----------------------------------------------------
before = len(df)
df = df.drop_duplicates(subset=['Nim', 'Nama Mahasiswa'], keep='first')
after = len(df)
print(f"Data duplikat yang dihapus: {before - after}")

# -----------------------------------------------------
# 6. Simpan ke Excel (Hanya Data Bersih)
# -----------------------------------------------------
output_file = "Data_Wisudawan_Cleansheet.xlsx"

df.to_excel(output_file, index=False, sheet_name='Data_Bersih')

print("\n✅ File berhasil dibuat:")
print(f"→ {output_file}")

# -----------------------------------------------------
# 7. Preview Data Setelah Dibersihkan
# -----------------------------------------------------
print("\n=== Data Setelah Dibersihkan (10 baris pertama) ===")
print(df.head(10))


# Penambahan Untuk Predikat, Grade
data = pd.read_excel("Data_Wisudawan_Cleansheet.xlsx")
data['Grade'] = ['A' if IPK >= 3.75 else 'B+' if IPK >= 3.5 else 'B' if IPK >= 3.0 else 'C' if IPK >= 2.5 else 'D' for IPK in data['IPK']]
data['Predikat'] = ['Cumlaude' if (IPK >= 3.75 and study <= 8) else 'Sangat Memuaskan' if (IPK >= 3.5 and study <= 9)else 'Memuaskan' if IPK >= 3.0 else 'Cukup' for IPK, study in zip(data['IPK'], data['Lama Studi (Semester)'])]
Kolom_Urutan = ['NIM','Nama Mahasiswa','Program Studi','IPK','Lama Studi (Semester)','Grade','Predikat','Tahun Wisuda']
print(data[Kolom_Urutan])
data.to_excel("Achmal_.xlsx", index=False, columns=Kolom_Urutan)