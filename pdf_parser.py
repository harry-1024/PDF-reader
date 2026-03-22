import pdfplumber
import re
import pandas as pd
import os
from datetime import datetime

def extract_correct_percent(text_block):
    """
    Looks for the most conservative percentage in a block of text.
    Handles '<10%', '10-30%', '>= 1.0 - < 5.0 %', etc.
    """
    if not text_block:
        return None
        
    # 1. Clean out the CAS numbers so we don't accidentally extract their digits
    clean_text = re.sub(r'\d{2,7}-\d{2}-\d', '', text_block)
    
    # 2. Look for numbers attached to %, <, <=, or within a range (e.g., 10 - 30)
    # This regex looks for numbers following symbols or preceding a % sign
    matches = re.findall(r'(?:<|<=|>=|=|\-\s?)?\s*([\d\.]+)\s*%', clean_text)
    
    # 3. Fallback: If no % sign, look for numbers near concentration keywords
    if not matches:
        matches = re.findall(r'\b(\d+(?:\.\d+)?)\b', clean_text)

    valid_numbers = []
    for n in matches:
        try:
            val = float(n)
            # Filter for realistic concentration values (0.01% to 100%)
            if 0.01 <= val <= 100:
                valid_numbers.append(val)
        except:
            continue
            
    return max(valid_numbers) if valid_numbers else None

def clean_chemical_name(name):
    """Standardizes names to prevent duplicates."""
    # Remove leading numbers/bullets (e.g., "1.2-Toluene" -> "Toluene")
    name = re.sub(r'^[\d\.\-\s•*]+', '', name)
    # Take the first part before common separators
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
                print(f"🔍 Processing: {filename}")
                with pdfplumber.open(pdf_path) as pdf:
                    full_text = "".join([p.extract_text() or "" for p in pdf.pages])
                    # Determine global properties
                    s = "Y" if any(x in full_text for x in ["Solid", "Powder", "Granules"]) else "N"
                    v = "Y" if any(x in full_text for x in ["Volatile", "Vapor", "Evaporat"]) else "N"
                    
                    for page in pdf.pages:
                        lines = (page.extract_text() or "").split('\n')
                        for i, line in enumerate(lines):
                            cas_match = re.search(r'(\d{2,7}-\d{2}-\d)', line)
                            if cas_match:
                                cas = cas_match.group(1).strip()
                                
                                # NEIGHBOR SEARCH: Look at the line above, current line, and line below
                                # This catches cases where the table is misaligned
                                context_start = max(0, i - 1)
                                context_end = min(len(lines), i + 2)
                                search_context = " ".join(lines[context_start:context_end])
                                
                                percent = extract_correct_percent(search_context)
                                raw_name = line.split(cas)[0].strip() or "UNKNOWN"

                                if percent is not None:
                                    all_data.append({
                                        "CAS Number": cas,
                                        "Contaminant Name": clean_chemical_name(raw_name),
                                        "Solids (Y/N)":
