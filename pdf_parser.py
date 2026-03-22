import pdfplumber
import re
import pandas as pd

def get_conservative_percentage(percent_string):
    """Extracts the highest numerical value from a range string."""
    numbers = re.findall(r'[\d\.]+', percent_string)
    if not numbers:
        return 0.0
    return max(float(n) for n in numbers)

def check_physical_properties(text):
    """Detects physical state and volatility."""
    props = {"is_solid": "N", "is_volatile": "N"}
    if re.search(r'\b(Solid|Powder|Dust|Crystal|Granules)\b', text, re.I):
        props["is_solid"] = "Y"
    if "volatile" in text.lower() or "vapor pressure" in text.lower():
        props["is_volatile"] = "Y"
    return props

def process_sds_to_excel(pdf_path, output_name="SDS_Results.xlsx"):
    all_data = []
    
    try:
        print(f"--- Analyzing: {pdf_path} ---")
        with pdfplumber.open(pdf_path) as pdf:
            full_text = "".join([p.extract_text() or "" for p in pdf.pages])
            
            # Global properties
            density_match = re.search(r'(?:Relative Density|Specific Gravity|Density):\s*([\d\.]+)', full_text, re.I)
            density = density_match.group(1) if density_match else "N/A"
            phys = check_physical_properties(full_text)
            
            for page in pdf.pages:
                page_text = page.extract_text()
                if not page_text: continue
                
                lines = page_text.split('\n')
                for line in lines:
                    # Look for CAS number
                    cas_match = re.search(r'(\b\d{2,7}-\d{2}-\d\b)', line)
                    if cas_match:
                        cas = cas_match.group(1)
                        
                        # 1. Extract Name (Text before the CAS number)
                        # We take everything before the CAS and clean up extra spaces/symbols
                        name_part = line.split(cas)[0].strip()
                        # Clean up common leading characters like bullet points or numbers
                        name = re.sub(r'^[\d\.\-\s•*]+', '', name_part) or "Unknown Name"
                        
                        # 2. Extract Percentage
                        percent_match = re.search(r'([\d\.]+\s*-\s*[\d\.]+\s*%|[\d\.]+\s*%|[\d\.]+\s*percent)', line, re.I)
                        percent_val = get_conservative_percentage(percent_match.group(1)) if percent_match else "Check PDF"

                        # 3. Build row in your specific order
                        all_data.append({
                            "Contaminant Name": name,
                            "CAS Number": cas,
                            "Solids (Y/N)": phys["is_solid"],
                            "Volatile (Y/N)": phys["is_volatile"],
                            "% Composition": percent_val,
                            "Relative Density": density
                        })

        if not all_data:
            print("❌ No ingredients found.")
            return

        # Create DataFrame and Reorder Columns explicitly
        df = pd.DataFrame(all_data).drop_duplicates()
        
        column_order = [
            "Contaminant Name", 
            "CAS Number", 
            "Solids (Y/N)", 
            "Volatile (Y/N)", 
            "% Composition",
            "Relative Density"
        ]
        df = df[column_order]

        df.to_excel(output_name, index=False)
        print(f"✅ Success! Results saved to: {output_name}")
        
    except Exception as e:
        print(f"An error occurred: {e}")

# Run it
process_sds_to_excel("test_sds.pdf")
