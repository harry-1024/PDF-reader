import pdfplumber
import re
import pandas as pd
import os

def get_conservative_percentage(percent_string):
    numbers = re.findall(r'[\d\.]+', percent_string)
    return max(float(n) for n in numbers) if numbers else 0.0

def check_physical_properties(text):
    props = {"is_solid": "N", "is_volatile": "N"}
    if re.search(r'\b(Solid|Powder|Dust|Crystal|Granules)\b', text, re.I):
        props["is_solid"] = "Y"
    if "volatile" in text.lower() or "vapor pressure" in text.lower():
        props["is_volatile"] = "Y"
    return props

def process_all_sds(folder_path, output_name="Master_SDS_Matrix.xlsx"):
    all_data = []
    
    # Loop through every PDF in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, filename)
            product_name = filename.replace(".pdf", "") # Use filename as Product Name
            
            try:
                print(f"Processing: {filename}...")
                with pdfplumber.open(pdf_path) as pdf:
                    full_text = "".join([p.extract_text() or "" for p in pdf.pages])
                    phys = check_physical_properties(full_text)
                    
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if not page_text: continue
                        for line in page_text.split('\n'):
                            cas_match = re.search(r'(\b\d{2,7}-\d{2}-\d\b)', line)
                            if cas_match:
                                cas = cas_match.group(1)
                                name_part = line.split(cas)[0].strip()
                                name = re.sub(r'^[\d\.\-\s•*]+', '', name_part) or "Unknown"
                                
                                percent_match = re.search(r'([\d\.]+\s*-\s*[\d\.]+\s*%|[\d\.]+\s*%|[\d\.]+\s*percent)', line, re.I)
                                percent_val = get_conservative_percentage(percent_match.group(1)) if percent_match else 0.0

                                all_data.append({
                                    "Contaminant Name": name.upper(),
                                    "CAS Number": cas,
                                    "Solids (Y/N)": phys["is_solid"],
                                    "Volatile (Y/N)": phys["is_volatile"],
                                    "Product": product_name,
                                    "Percentage": percent_val
                                })
            except Exception as e:
                print(f"Error with {filename}: {e}")

    if not all_data:
        print("No data found!")
        return

    # 1. Create a DataFrame
    df = pd.DataFrame(all_data)

    # 2. "Pivot" the data: Products become columns, Contaminants stay as rows
    # This puts the Percentage in the cells
    matrix = df.pivot_table(
        index=["Contaminant Name", "CAS Number", "Solids (Y/N)", "Volatile (Y/N)"],
        columns="Product",
        values="Percentage",
        aggfunc='max' # If a contaminant appears twice in one SDS, take the max
    ).reset_index()

    # 3. Sort by Contaminant Name alphabetically
    matrix = matrix.sort_values(by="Contaminant Name")

    # 4. Save to Excel
    matrix.to_excel(output_name, index=False)
    print(f"✅ Master Matrix saved to: {output_name}")

# --- RUNNING IT ---
# 1. Create a folder named 'sds_files' on your Desktop
# 2. Put all your SDS PDFs in there
# 3. Update the path below to your folder
sds_folder = "/Users/zhangyuhao/Desktop/SDS_Project/sds_files"

# Make the folder if it doesn't exist yet
if not os.path.exists(sds_folder):
    os.makedirs(sds_folder)
    print(f"Please put your PDFs in the new folder: {sds_folder}")
else:
    process_all_sds(sds_folder)
