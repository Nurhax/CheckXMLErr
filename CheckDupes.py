import pandas as pd
import os
import json
from collections import defaultdict

# Ganti dengan path ke file Excel kamu
excel_path = r''

# Cek file
if not os.path.exists(excel_path):
    raise FileNotFoundError(f"File tidak ditemukan: {excel_path}")

# Baca sheet SCADAErr
try:
    df = pd.read_excel(excel_path, sheet_name='SCADAErr')
except ValueError as e:
    raise ValueError(f"Gagal membaca sheet 'SCADAErr': {e}")

# Validasi kolom
if 'Instances' not in df.columns or 'File' not in df.columns:
    raise ValueError("Kolom 'Instances' dan 'File' harus ada di sheet")

# Buat struktur untuk menyimpan hasil
hashmap = {}
duplicates = defaultdict(list)

# Iterasi setiap baris
for _, row in df.iterrows():
    instance = row['Instances']
    file = row['File']

    if instance in hashmap:
        # Duplikat ditemukan
        duplicates[instance].append(file)
    else:
        hashmap[instance] = file

# Tampilkan hasil utama
print("=== Hashmap unik (Instances → File) ===")
for key, value in hashmap.items():
    print(f"{key}: {value}")

# Tampilkan duplikat
if duplicates:
    print("\n=== Duplikat ditemukan ===")
    for key, dupe_files in duplicates.items():
        total = len(dupe_files) + 1  # +1 untuk yang pertama
        print(f"{key}: {total} kemunculan | File = [{hashmap[key]}] + {dupe_files}")
else:
    print("\nTidak ada duplikat ditemukan.")

# Simpan ke JSON
output_data = {
    "unique_instances": hashmap,
    "duplicates": {k: {"count": len(v)+1, "files": [hashmap[k]] + v} for k, v in duplicates.items()}
}

output_path = os.path.splitext(excel_path)[0] + '_hashmap_with_duplicates.json'
with open(output_path, 'w') as f:
    json.dump(output_data, f, indent=4)

print(f"\nHashmap dan duplikat disimpan ke: {output_path}")
