import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

# -----------------------------------------------------
# 1. Pastikan file dan direktori
# -----------------------------------------------------
os.chdir(os.path.dirname(__file__))

# -----------------------------------------------------
# 2. Import Data
# -----------------------------------------------------
data = pd.read_excel("Data Wisudawan.xlsx")

# -----------------------------------------------------
# 3. Normalisasi Nama Kolom
# -----------------------------------------------------
data.columns = data.columns.str.strip().str.lower().str.replace(r'\s+', ' ', regex=True)
print("Nama kolom terdeteksi:", data.columns.tolist())

# -----------------------------------------------------
# 4. Bersihkan Data
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
# 5. Bersihkan Typo dan Filter IPK
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
# 6. Aturan D3 dan D4
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
# 7. Hapus Duplikat
# -----------------------------------------------------
before = len(data)
data = data.drop_duplicates(subset=['nim', 'nama mahasiswa'], keep='first')
after = len(data)
print(f"Data duplikat yang dihapus: {before - after}")

# -----------------------------------------------------
# 8. Tambah Kolom Grade, Predikat, Tahun Wisuda
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

data['Tahun Wisuda'] = 2025
data['Rata-rata IPK Prodi'] = data.groupby('program studi')['ipk'].transform('mean').round(2)

# -----------------------------------------------------
# 9. Tambah Rata-rata IPK ke bagian bawah file
# -----------------------------------------------------
rata_per_prodi = (
    data.groupby('program studi', as_index=False)['ipk']
    .mean()
    .round(2)
    .rename(columns={'ipk': 'Rata-rata IPK'})
)

# Baris kosong pemisah
baris_kosong = pd.DataFrame([[''] * len(data.columns)], columns=data.columns)

# Gabungkan data utama + pemisah + ringkasan rata-rata
rata_per_prodi.rename(columns={'program studi': 'Program Studi'}, inplace=True)
rata_per_prodi['Keterangan'] = 'Rata-rata IPK per Prodi'
rata_per_prodi = rata_per_prodi[['Program Studi', 'Rata-rata IPK', 'Keterangan']]

# Simpan gabungan ke Excel
with pd.ExcelWriter("Data_Wisudawan_Final.xlsx", engine='openpyxl') as writer:
    data.to_excel(writer, index=False, sheet_name="Data Wisudawan")
    rata_per_prodi.to_excel(writer, index=False, sheet_name="Rata-rata IPK Prodi")

print("\n✅ File 'Data_Wisudawan_Final.xlsx' berhasil dibuat dan sudah termasuk ringkasan rata-rata IPK per prodi.")

# -----------------------------------------------------
# 10. Analisis Cumlaude
# -----------------------------------------------------
cumlaude_per_prodi = data[data['Predikat'] == 'Cumlaude']['program studi'].value_counts()
if not cumlaude_per_prodi.empty:
    prodi_terbanyak = cumlaude_per_prodi.idxmax()
    jumlah_terbanyak = cumlaude_per_prodi.max()
    print(f"🎓 Program Studi dengan Cumlaude terbanyak: {prodi_terbanyak} ({jumlah_terbanyak} mahasiswa)")
else:
    print("Tidak ada mahasiswa Cumlaude pada data ini.")

# -----------------------------------------------------
# 11. Visualisasi Data
# ----------------------------------------------------
# --- Grafik 1: Jumlah wisudawan per program studi ---
# --- Revisi Menjadi Warna Merah ( Diagram Batang ) ---
plt.figure(figsize=(10, 6))
jumlah_wisudawan = data['Program Studi'].value_counts()
jumlah_wisudawan.plot(kind='bar', color='red', edgecolor='black')
plt.title('Jumlah Wisudawan per Program Studi', fontsize=14, fontweight='bold')
plt.xlabel('Program Studi')
plt.ylabel('Jumlah Wisudawan')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

plt.figure(figsize=(8, 8))
predikat_counts = data['Predikat'].value_counts()
plt.pie(predikat_counts, labels=predikat_counts.index, autopct='%1.1f%%', startangle=90, colors=plt.cm.Paired.colors)
plt.title('Distribusi Predikat Kelulusan', fontsize=14, fontweight='bold')
plt.axis('equal')
plt.show()

plt.figure(figsize=(10, 6))
rata_ipk = data.groupby('program studi')['ipk'].mean().sort_values(ascending=False)
rata_ipk.plot(kind='bar', color='lightgreen', edgecolor='black')
plt.title('Perbandingan Rata-rata IPK antar Program Studi', fontsize=14, fontweight='bold')
plt.xlabel('Program Studi')
plt.ylabel('Rata-rata IPK')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

plt.figure(figsize=(8, 5))
plt.hist(data['ipk'], bins=10, color='salmon', edgecolor='black')
plt.title('Sebaran IPK Seluruh Wisudawan', fontsize=14, fontweight='bold')
plt.xlabel('IPK')
plt.ylabel('Jumlah Mahasiswa')
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()
plt.show()

plt.figure(figsize=(10, 6))
cumlaude_per_prodi.plot(kind='bar', color='gold', edgecolor='black')
plt.title("Jumlah Mahasiswa Cumlaude per Program Studi", fontsize=14, fontweight='bold')
plt.xlabel("Program Studi")
plt.ylabel("Jumlah Cumlaude")
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

print("\n✅ Semua grafik berhasil ditampilkan.")
