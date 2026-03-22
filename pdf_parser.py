import pdfplumber
import re
import pandas as pd
import os
from datetime import datetime

def extract_correct_percent(line_text):
    """Finds the number associated with %, <, or <=."""
    matches = re.findall(r'(?:<|<=|=|\-\s?)?\s*([\d\.]+)\s*%', line_text)
    if not matches:
        matches = re.findall(r'\b(\d+(?:\.\d+)?)\b', line_text)
    
    valid_numbers = []
    for n in matches:
        try:
            val = float(n)
            if 0.01 <= val <= 100:
                valid_numbers.append(val)
        except:
            continue
    return max(valid_numbers) if valid_numbers else None

def clean_chemical_name(name):
    """Standardizes names for better display."""
    name = re.sub(r'^[\d\.\-\s•*]+', '', name)
    name = name.split(';')[0].split('(')[0]
    return name.strip().upper()

def process_all_sds(folder_path):
    all_data = []
    output_name = "Master_SDS_Matrix.xlsx"
    backup_name = f"Master_SDS_Matrix_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    
    if not os.path.exists(folder_path):
        print(f"Folder not found: {folder_path}")
        return

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            product_name = filename.rsplit('.', 1)[0]
            pdf_path = os.path.join(folder_path, filename)
            
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    full_text = "".join([p.extract_text() or "" for p in pdf.pages])
                    s = "Y" if any(x in full_text for x in ["Solid", "Powder", "Granules"]) else "N"
                    v = "Y" if any(x in full_text for x in ["Volatile", "Vapor", "Evaporat"]) else "N"
                    
                    for page in pdf.pages:
                        text = page.extract_text()
                        if not text: continue
                        for line in text.split('\n'):
                            cas_match = re.search(r'(\d{2,7}-\d{2}-\d)', line)
                            if cas_match:
                                cas = cas_match.group(1).strip()
                                raw_name = line.split(cas)[0].strip()
                                percent = extract_correct_percent(line)
                                
                                if percent is not None:
                                    all_data.append({
                                        "CAS Number": cas,
                                        "Contaminant Name": clean_chemical_name(raw_name),
                                        "Solids (Y/N)": s,
                                        "Volatile (Y/N)": v,
                                        "Product": product_name,
                                        "Percentage": percent
                                    })
            except Exception as e:
                print(f"Error reading {filename}: {e}")

    if not all_data:
        print("No data found.")
        return

    df = pd.DataFrame(all_data)

    # --- CAS-BASED MERGE ---
    # 1. Pick the best name for each CAS (the shortest one is usually the cleanest)
    name_map = df.groupby('CAS Number')['Contaminant Name'].min().to_dict()
    df['Contaminant Name'] = df['CAS Number'].map(name_map)

    # 2. Pivot: Use CAS as the anchor to prevent duplicates
    matrix = df.pivot_table(
        index=["Contaminant Name", "CAS Number", "Solids (Y/N)", "Volatile (Y/N)"],
        columns="Product",
        values="Percentage",
        aggfunc='max'
    ).reset_index()

    # 3. Sort and Save
    matrix = matrix.sort_values(by="Contaminant Name")
    
    # Save the main file (Overwrite)
    matrix.to_excel(output_name, index=False)
    # Save a dated backup
    matrix.to_excel(backup_name, index=False)
    
    print(f"✅ Success! Matrix updated. Backup saved as {backup_name}")

# Run
sds_folder = os.path.join(os.getcwd(), "sds_files")
process_all_sds(sds_folder)
