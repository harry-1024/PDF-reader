import pdfplumber
import re
import pandas as pd

def get_conservative_percentage(percent_string):
    numbers = re.findall(r'[\d\.]+', percent_string)
    return max(float(n) for n in numbers) if numbers else 0.0

def check_physical_properties(text):
    props = {"is_solid": "No", "is_volatile": "No"}
    if re.search(r'\b(Solid|Powder|Dust|Crystal)\b', text, re.I):
        props["is_solid"] = "Yes"
    if "volatile" in text.lower() or "vapor pressure" in text.lower():
        props["is_volatile"] = "Yes"
    return props

def process_sds_to_excel(pdf_path, output_name="SDS_Results.xlsx"):
    all_data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "".join([p.extract_text() or "" for p in pdf.pages])
        
        # Section 9 extraction
        density_match = re.search(r'(?:Relative Density|Specific Gravity):\s*([\d\.]+)', full_text, re.I)
        density = density_match.group(1) if density_match else "N/A"
        phys = check_physical_properties(full_text)
        
        # Section 3 extraction
        matches = re.findall(r'(\d{2,7}-\d{2}-\d).*?([\d\.]+\s*-\s*[\d\.]+\s*%|[\d\.]+\s*%)', full_text)
        
        for cas, percent_str in matches:
            all_data.append({
                "CAS Number": cas,
                "Max Concentration (%)": get_conservative_percentage(percent_str),
                "Density": density,
                "Solid": phys["is_solid"],
                "Volatile": phys["is_volatile"],
                "Source File": pdf_path
            })

    # Save to Excel
    df = pd.DataFrame(all_data)
    df.to_excel(output_name, index=False)
    print(f"✅ Success! Data saved to {output_name}")

# To run it:
process_sds_to_excel("your_test_file.pdf")
