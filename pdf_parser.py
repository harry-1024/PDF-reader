import pdfplumber
import re
import pandas as pd
import os

def get_conservative_percentage(line_text):
    # Find all numbers on the line
    numbers = re.findall(r'(\d+(?:\.\d+)?)', line_text)
    # Filter for numbers likely to be percentages
    valid_percents = [float(n) for n in numbers if 0.1 <= float(n) <= 100]
    return max(valid_percents) if valid_percents else 0.0

def check_physical_properties(text):
    props = {"is_solid": "N", "is_volatile": "N"}
    if re.search(r'\b(Solid|Powder|Dust|Crystal|Granules|Flakes)\b', text, re.I):
        props["is_solid"] = "Y"
    if re.search(r'\b(Volatile|Vapor Pressure|Evaporation)\b', text, re.I):
        props["is_volatile"] = "Y"
    return props

def process_all_sds(folder_path, output_name="Master_SDS_Matrix.xlsx"):
    all_data = []
    
    if not os.path.exists(folder_path):
        print(f"Error: Folder {folder_path} not found.")
        return

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            product_name = filename.rsplit('.', 1)[0]
            
            try:
                print(f"🔍 Analyzing: {filename}...")
                with pdfplumber.open(pdf_path) as pdf:
                    full_text = "".join([p.extract_text() or "" for p in pdf.pages])
                    phys = check_physical_properties(full_text)
                    
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if not page_text: continue
                        
                        for line in page_text.split('\n'):
                            cas_match = re.search(r'(\d{2,7}-\d{2}-\d)', line)
                            if cas_match:
                                cas = cas_match.group(1).strip()
                                
                                # CLEANING: Remove extra spaces and leading symbols from name
                                name_part = line.split(cas)[0].strip()
                                name = re.sub(r'^[\d\.\-\s•*]+', '', name_part).strip().upper()
                                if not name: name = "UNKNOWN CHEMICAL"
                                
                                percent_val = get_conservative_percentage(line)

                                all_data.append({
                                    "Contaminant Name": name,
                                    "CAS Number": cas,
                                    "Solids (Y/N)": phys["is_solid"],
                                    "Volatile (Y/N)": phys["is_volatile"],
                                    "Product": product_name,
                                    "Percentage": percent_val
                                })
            except Exception as e:
                print(f"Skipping {filename}: {e}")

    if not all_data:
        print("❌ No data extracted.")
        return

    # 1. Convert to DataFrame
    df = pd.DataFrame(all_data)

    # 2. Pivot: This merges all 'TOLUENE' into one row
    matrix = df.pivot_table(
        index=["Contaminant Name", "CAS Number", "Solids (Y/N)", "Volatile (Y/N)"],
        columns="Product",
        values="Percentage",
        aggfunc='max'
    ).reset_index()

    # 3. Sort Alphabetically
    matrix = matrix.sort_values(by="Contaminant Name")

    # 4. Remove the 0s: Replace 0 with an empty string so it shows as a blank cell
    # We use a copy to avoid formatting errors
    matrix_clean = matrix.replace(0, "")

    # Save to Excel
    matrix_clean.to_excel(output_name, index=False)
    print(f"\n✅ Master Matrix Created with blank cells for missing values!")
    print(f"📂 Location: {os.path.abspath(output_name)}")

# Run
sds_folder = os.path.join(os.getcwd(), "sds_files")
if os.path.exists(sds_folder):
    process_all_sds(sds_folder)
else:
    os.makedirs(sds_folder)
    print(f"Put PDFs in: {sds_folder}")
