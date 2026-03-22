import pdfplumber
import re

def get_conservative_percentage(percent_string):
    """提取百分比范围中的最高值 (Conservatism)"""
    # Find all numbers (including decimals)
    numbers = re.findall(r'[\d\.]+', percent_string)
    if not numbers:
        return 0.0
    # Convert to floats and return the maximum
    return max(float(n) for n in numbers)

def check_physical_properties(text):
    """判断物理状态与挥发性"""
    props = {"is_solid": "Unknown", "is_volatile": "No"}
    
    # 1. Solid check
    if re.search(r'\b(Solid|Powder|Dust|Crystal)\b', text, re.I):
        props["is_solid"] = "Yes"
    elif re.search(r'\b(Liquid|Viscous)\b', text, re.I):
        props["is_solid"] = "No"
        
    # 2. Volatility check (Based on Vapor Pressure > 0.1 mmHg or keyword)
    vp_match = re.search(r'Vapor Pressure:\s*([\d\.]+)', text, re.I)
    if vp_match:
        if float(vp_match.group(1)) > 0.1:
            props["is_volatile"] = "Yes"
    elif "highly volatile" in text.lower():
        props["is_volatile"] = "Yes"
        
    return props

def process_sds(pdf_path):
    print(f"--- Processing: {pdf_path} ---")
    results = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if not page_text: continue
                full_text += page_text
                
                # Section 3: Extraction (CAS and Percentage)
                # Matches patterns like: 1333-86-4  10-30% or 64-17-5 <5%
                matches = re.findall(r'(\d{2,7}-\d{2}-\d).*?([\d\.]+\s*-\s*[\d\.]+\s*%|[\d\.]+\s*%)', page_text)
                
                for cas, percent_str in matches:
                    results.append({
                        "CAS": cas,
                        "Max_%": get_conservative_percentage(percent_str)
                    })

            # Section 9: Extraction (Density & Properties)
            density_match = re.search(r'(?:Specific Gravity|Relative Density):\s*([\d\.]+)', full_text, re.I)
            density = density_match.group(1) if density_match else "Not Found"
            phys_props = check_physical_properties(full_text)

            # Final Printout
            print(f"Relative Density: {density}")
            print(f"Is Solid: {phys_props['is_solid']}")
            print(f"Is Volatile: {phys_props['is_volatile']}")
            print("-" * 30)
            print("Ingredients Found:")
            for item in results:
                print(f"CAS: {item['CAS']} | Max Content: {item['Max_%']}%")
                
    except Exception as e:
        print(f"Error processing PDF: {e}")

# --- HOW TO RUN ---
# 1. Replace 'test_sds.pdf' with your actual filename
# 2. Make sure the PDF is in the same folder as this script
process_sds("test_sds.pdf")
