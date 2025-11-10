import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

# -----------------------------------------------------
# Pastikan script berjalan di folder yang sama dengan file Excel
# -----------------------------------------------------
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
# 8. Tambahkan kolom Rata-rata IPK per Prodi
# -----------------------------------------------------
data['Rata-rata IPK Prodi'] = data.groupby('program studi')['ipk'].transform('mean').round(2)

# -----------------------------------------------------
# 9. Urutkan Kolom Sesuai Format Akhir
# -----------------------------------------------------
Kolom_Urutan = [
    'nim',
    'nama mahasiswa',
    'program studi',
    'ipk',
    kolom_lama_studi,
    'Grade',
    'Predikat',
    'Rata-rata IPK Prodi',
    'tahun wisuda'
]

data = data[Kolom_Urutan]

# Ganti nama kolom untuk tampilan akhir
data.columns = [
    'NIM',
    'Nama Mahasiswa',
    'Program Studi',
    'IPK',
    'Lama Studi (Semester)',
    'Grade',
    'Predikat',
    'Rata-rata IPK Prodi',
    'Tahun Wisuda'
]

# -----------------------------------------------------
# 10. Simpan ke Excel
# -----------------------------------------------------
data.to_excel("Data_Wisudawan_Final.xlsx", index=False)
print("\n✅ File akhir berhasil dibuat: Data_Wisudawan_Final.xlsx")

# -----------------------------------------------------
# 11. Tampilkan Data di Terminal
# -----------------------------------------------------
print("\n📄 Data Wisudawan:")
print(data)

# -----------------------------------------------------
# 12. Analisis Cumlaude
# -----------------------------------------------------
cumlaude_per_prodi = data[data['Predikat'] == 'Cumlaude']['Program Studi'].value_counts()
if not cumlaude_per_prodi.empty:
    prodi_terbanyak = cumlaude_per_prodi.idxmax()
    jumlah_terbanyak = cumlaude_per_prodi.max()
    print(f"\nProgram Studi dengan Cumlaude terbanyak: {prodi_terbanyak} ({jumlah_terbanyak} mahasiswa)")
else:
    print("\nTidak ada mahasiswa Cumlaude pada data ini.")

# -----------------------------------------------------
# 13. Tampilkan Rata-rata IPK per Prodi di Terminal
# -----------------------------------------------------
rata_ipk_per_prodi = data.groupby('Program Studi', as_index=False)['Rata-rata IPK Prodi'].mean().round(2)
print("\nRata-rata IPK per Program Studi:")
print(rata_ipk_per_prodi)

# -----------------------------------------------------
# 14. Visualisasi Data
# -----------------------------------------------------
# Grafik 1: Jumlah wisudawan per program studi
plt.figure(figsize=(10, 6))
data['Program Studi'].value_counts().plot(kind='bar', color='red', edgecolor='black')
plt.title('Jumlah Wisudawan per Program Studi', fontsize=14, fontweight='bold')
plt.xlabel('Program Studi')
plt.ylabel('Jumlah Wisudawan')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

# Grafik 2: Distribusi Predikat Kelulusan
plt.figure(figsize=(8, 8))
data['Predikat'].value_counts().plot(kind='pie', labels=data['Predikat'].value_counts().index,
                                     autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
plt.title('Distribusi Predikat Kelulusan', fontsize=14, fontweight='bold')
plt.axis('equal')
plt.show()

# Grafik 3: Perbandingan rata-rata IPK antar Prodi
plt.figure(figsize=(10, 6))
data.groupby('Program Studi')['Rata-rata IPK Prodi'].mean().sort_values(ascending=False).plot(
    kind='bar', color='lightgreen', edgecolor='black')
plt.title('Perbandingan Rata-rata IPK antar Program Studi', fontsize=14, fontweight='bold')
plt.xlabel('Program Studi')
plt.ylabel('Rata-rata IPK')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

# Grafik 4: Sebaran IPK seluruh wisudawan
plt.figure(figsize=(8, 5))
plt.hist(data['IPK'], bins=10, color='salmon', edgecolor='black')
plt.title('Sebaran IPK Seluruh Wisudawan', fontsize=14, fontweight='bold')
plt.xlabel('IPK')
plt.ylabel('Jumlah Mahasiswa')
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()

# Grafik 5: Jumlah Cumlaude per Program Studi
if not cumlaude_per_prodi.empty:
    plt.figure(figsize=(10,6))
    cumlaude_per_prodi.plot(kind='bar', color='gold', edgecolor='black')
    plt.title("Jumlah Mahasiswa Cumlaude per Program Studi", fontsize=14, fontweight='bold')
    plt.xlabel("Program Studi")
    plt.ylabel("Jumlah Cumlaude")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()

print("\n✅ Semua grafik berhasil ditampilkan.")
