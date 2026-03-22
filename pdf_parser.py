import pdfplumber
import re
import pandas as pd
import os

def get_conservative_percentage(line_text):
    """
    Improved: Searches the entire line for percentage-like patterns
    and returns the highest number found.
    """
    # Look for patterns like "10-30", "10 - 30", "< 5", "80%"
    # This regex is broader to catch numbers even if the % sign is far away
    numbers = re.findall(r'(\d+(?:\.\d+)?)', line_text)
    
    # We only want numbers that are likely percentages (usually 0.1 to 100)
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
            product_name = filename.rsplit('.', 1)[0] # Strip .pdf extension
            
            try:
                print(f"🔍 Analyzing: {filename}...")
                with pdfplumber.open(pdf_path) as pdf:
                    full_text = "".join([p.extract_text() or "" for p in pdf.pages])
                    phys = check_physical_properties(full_text)
                    
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if not page_text: continue
                        
                        for line in page_text.split('\n'):
                            # Find CAS Number
                            cas_match = re.search(r'(\d{2,7}-\d{2}-\d)', line)
                            if cas_match:
                                cas = cas_match.group(1)
                                
                                # Extract Name (Everything before the CAS)
                                name_part = line.split(cas)[0].strip()
                                name = re.sub(r'^[\d\.\-\s•*]+', '', name_part) or "Unknown Chemical"
                                
                                # Extract Percentage from the same line
                                percent_val = get_conservative_percentage(line)

                                all_data.append({
                                    "Contaminant Name": name.upper(),
                                    "CAS Number": cas,
                                    "Solids (Y/N)": phys["is_solid"],
                                    "Volatile (Y/N)": phys["is_volatile"],
                                    "Product": product_name,
                                    "Percentage": percent_val
                                })
            except Exception as e:
                print(f"Skipping {filename} due to error: {e}")

    if not all_data:
        print("❌ No data was extracted. Please check if PDFs are readable.")
        return

    # Create DataFrame
    df = pd.DataFrame(all_data)

    # 1. Pivot Table: Contaminants as Rows, Products as Columns
    # index: First 4 columns you requested
    # columns: The Product Names (from filenames)
    matrix = df.pivot_table(
        index=["Contaminant Name", "CAS Number", "Solids (Y/N)", "Volatile (Y/N)"],
        columns="Product",
        values="Percentage",
        aggfunc='max'
    ).reset_index()

    # 2. Sort Alphabetically by Name
    matrix = matrix.sort_values(by="Contaminant Name")

    # 3. Final polish: Fill empty cells with 0 (if a product doesn't have that chemical)
    matrix = matrix.fillna(0)

    # Save to Excel
    matrix.to_excel(output_name, index=False)
    print(f"\n✅ Master Matrix Created!")
    print(f"📂 Location: {os.path.abspath(output_name)}")

# --- AUTO-RUN ---
# This looks for a folder called 'sds_files' in your project directory
sds_folder = os.path.join(os.getcwd(), "sds_files")

if not os.path.exists(sds_folder):
    os.makedirs(sds_folder)
    print(f"Empty folder created at: {sds_folder}")
    print("Please put your SDS PDFs in that folder and run this script again.")
else:
    process_all_sds(sds_folder)
