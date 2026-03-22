import pdfplumber
import re
import pandas as pd
import os

def extract_correct_percent(line_text):
    """
    Finds the number associated with %, <, or <=. 
    Example: '< 10%' -> 10.0, '0.1 - 5%' -> 5.0
    """
    # Look for numbers specifically near % signs or inequality symbols
    # Pattern: finds numbers that have <, <=, or - before them, or % after them
    matches = re.findall(r'(?:<|<=|=|\-\s?)?\s*([\d\.]+)\s*%', line_text)
    
    if not matches:
        # Fallback: if no % sign, look for any decimal number on the line
        matches = re.findall(r'\b(\d+(?:\.\d+)?)\b', line_text)
    
    valid_numbers = []
    for n in matches:
        try:
            val = float(n)
            # Only keep if it looks like a concentration (0.01 to 100)
            if 0.01 <= val <= 100:
                valid_numbers.append(val)
        except:
            continue
            
    return max(valid_numbers) if valid_numbers else ""

def clean_chemical_name(name):
    """Removes junk characters and standardizes names to prevent duplicates."""
    # Remove leading dots, dashes, numbers, and bullet points
    name = re.sub(r'^[\d\.\-\s•*]+', '', name)
    # Remove common extra words that cause duplicates
    name = name.split(';')[0].split('(')[0] 
    return name.strip().upper()

def process_all_sds(folder_path, output_name="Master_SDS_Matrix.xlsx"):
    all_data = []
    
    if not os.path.exists(folder_path):
        print(f"Folder not found: {folder_path}")
        return

    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".pdf"):
            product_name = filename.rsplit('.', 1)[0]
            pdf_path = os.path.join(folder_path, filename)
            
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    # Get physical properties from the whole file
                    full_text = "".join([p.extract_text() or "" for p in pdf.pages])
                    if "Solid" in full_text or "Powder" in full_text: s, v = "Y", "N"
                    else: s, v = "N", "Y"
                    
                    for page in pdf.pages:
                        text = page.extract_text()
                        if not text: continue
                        for line in text.split('\n'):
                            cas_match = re.search(r'(\d{2,7}-\d{2}-\d)', line)
                            if cas_match:
                                cas = cas_match.group(1)
                                raw_name = line.split(cas)[0].strip()
                                
                                all_data.append({
                                    "Contaminant Name": clean_chemical_name(raw_name),
                                    "CAS Number": cas.strip(),
                                    "Solids (Y/N)": s,
                                    "Volatile (Y/N)": v,
                                    "Product": product_name,
                                    "Percentage": extract_correct_percent(line)
                                })
            except Exception as e:
                print(f"Error reading {filename}: {e}")

    if not all_data:
        print("No data found.")
        return

    # Create DataFrame
    df = pd.DataFrame(all_data)
    
    # Remove rows where Percentage is empty string
    df = df[df["Percentage"] != ""]

    # --- THE MAGIC STEP: PIVOT ---
    # This combines all "TOLUENE" rows into one, using the MAX percentage found per product
    matrix = df.pivot_table(
        index=["Contaminant Name", "CAS Number", "Solids (Y/N)", "Volatile (Y/N)"],
        columns="Product",
        values="Percentage",
        aggfunc='max'
    ).reset_index()

    # Sort Alphabetically
    matrix = matrix.sort_values(by="Contaminant Name")

    # Save to Excel - index=False removes the row numbers
    # We leave NaNs (empty values) as they are, which Excel renders as blank cells
    matrix.to_excel(output_name, index=False)
    print(f"✅ Success! Master Matrix created with {len(matrix)} unique chemicals.")

# Set path and run
sds_folder = os.path.join(os.getcwd(), "sds_files")
process_all_sds(sds_folder)
