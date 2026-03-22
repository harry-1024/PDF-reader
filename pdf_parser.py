import pdfplumber
import re
import pandas as pd

def get_conservative_percentage(percent_string):
    """提取百分比范围中的最高值 (Conservatism)"""
    numbers = re.findall(r'[\d\.]+', percent_string)
    if not numbers:
        return 0.0
    return max(float(n) for n in numbers)

def check_physical_properties(text):
    """判断物理状态与挥发性"""
    props = {"is_solid": "No", "is_volatile": "No"}
    # Check for solid keywords
    if re.search(r'\b(Solid|Powder|Dust|Crystal|Granules)\b', text, re.I):
        props["is_solid"] = "Yes"
    # Check for volatility (Vapor pressure or keywords)
    if "volatile" in text.lower() or "vapor pressure" in text.lower():
        props["is_volatile"] = "Yes"
    return props

def process_sds_to_excel(pdf_path, output_name="SDS_Results.xlsx"):
    all_data = []
    
    try:
        print(f"--- Analyzing: {pdf_path} ---")
        with pdfplumber.open(pdf_path) as pdf:
            # Combine text from all pages for physical property search
            full_text = "".join([p.extract_text() or "" for p in pdf.pages])
            
            # 1. Section 9: Density & Properties
            density_match = re.search(r'(?:Relative Density|Specific Gravity|Density):\s*([\d\.]+)', full_text, re.I)
            density = density_match.group(1) if density_match else "N/A"
            phys = check_physical_properties(full_text)
            
            # 2. Section 3: Ingredients (Line-by-Line Smart Match)
            for page in pdf.pages:
                page_text = page.extract_text()
                if not page_text: continue
                
                lines = page_text.split('\n')
                for line in lines:
                    # Look for CAS (Pattern: 1-7 digits, dash, 2 digits, dash, 1 digit)
                    cas_match = re.search(r'(\b\d{2,7}-\d{2}-\d\b)', line)
                    if cas_match:
                        cas = cas_match.group(1)
                        # Look for Percentage in the same line
                        percent_match = re.search(r'([\d\.]+\s*-\s*[\d\.]+\s*%|[\d\.]+\s*%|[\d\.]+\s*percent)', line, re.I)
                        
                        if percent_match:
                            percent_val = get_conservative_percentage(percent_match.group(1))
                        else:
                            percent_val = "Detected (Check Manual)"

                        all_data.append({
                            "CAS Number": cas,
                            "Max Concentration (%)": percent_val,
                            "Relative Density": density,
                            "Is Solid": phys["is_solid"],
                            "Is Volatile": phys["is_volatile"],
                            "Source File": pdf_path
                        })

        if not all_data:
            print("❌ No ingredients found. Please ensure the PDF is not a scanned image.")
            # Debug: Print first bit of text to see what Python sees
            print("Debug Sample Text:", full_text[:300])
            return

        # 3. Clean up and Save
        df = pd.DataFrame(all_data).drop_duplicates(subset=['CAS Number', 'Max Concentration (%)'])
        df.to_excel(output_name, index=False)
        
        print(f"✅ Success! Found {len(df)} unique components.")
        print(f"📊 Results saved to: {output_name}")
        print(df[['CAS Number', 'Max Concentration (%)']].to_string(index=False))
        
    except Exception as e:
        print(f"An error occurred: {e}")

# --- START THE PROGRAM ---
# Ensure your PDF file is named exactly "test_sds.pdf" and is in the same folder
process_sds_to_excel("test_sds.pdf")
