import pandas as pd
import os
import xml.etree.ElementTree as ET
import copy
from collections import defaultdict

# === PATH SETUP ===
excel_path = r'C:\KP\Files\PLNSHIZ\Summary_IFS_Points.xlsx'
xml_folder = r'C:\KP\Files\CheckErr\imported_file'
output_xml = r'C:\KP\Files\PLNSHIZ\extracted_IfsPoints.xml'

# === BACA EXCEL ===
if not os.path.exists(excel_path):
    raise FileNotFoundError(f"File tidak ditemukan: {excel_path}")

df = pd.read_excel(excel_path, sheet_name='SCADAErr')

if 'Instances' not in df.columns or 'File' not in df.columns:
    raise ValueError("Kolom 'Instances' dan 'File' harus ada di sheet 'SCADAErr'")

# Set PathB untuk pencocokan cepat
valid_pathbs = set(df['Instances'].dropna().str.strip())
print(f"Total PathB valid dari Excel: {len(valid_pathbs)}")

# === DICTIONARY: Parent Path -> List of IfsPoint ===
grouped_by_parent_path = defaultdict(list)
count_total = 0

# === LOOP FILE XML ===
for filename in os.listdir(xml_folder):
    if filename.lower().endswith('.xml'):
        file_path = os.path.join(xml_folder, filename)
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            # Loop semua Parent
            for parent in root.findall(".//Parent"):
                parent_path = parent.get("Path")
                if not parent_path:
                    continue

                for ifs_point in parent.findall("IfsPoint"):
                    # Gabungkan dua jenis link
                    links = ifs_point.findall("Link_IfsPointLinksToInfo") + ifs_point.findall("Link_IfsPointSlaveLinksToInfo")

                    for link in links:
                        path_b = link.get('PathB')
                        if path_b and path_b.strip() in valid_pathbs:
                            grouped_by_parent_path[parent_path].append(copy.deepcopy(ifs_point))
                            count_total += 1
                            break  # cukup satu link match

        except ET.ParseError as e:
            print(f"Gagal parse file {filename}: {e}")
        except Exception as e:
            print(f"Error di file {filename}: {e}")

# === BANGUN OUTPUT XML ===
root_output = ET.Element("IFPointList")
for parent_path, ifspoints in grouped_by_parent_path.items():
    parent_elem = ET.Element("Parent")
    parent_elem.set("Path", parent_path)
    for ip in ifspoints:
        parent_elem.append(ip)
    root_output.append(parent_elem)

# === TULIS FILE XML ===
tree_output = ET.ElementTree(root_output)
tree_output.write(output_xml, encoding='utf-8', xml_declaration=True)

print(f"\n✅ Total IfsPoint valid: {count_total}")
print(f"✅ Total Parent Path unik: {len(grouped_by_parent_path)}")
print(f"📁 Disimpan ke: {output_xml}")
