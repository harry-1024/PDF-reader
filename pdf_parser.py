import pdfplumber
import re
import pandas as pd
import os
from datetime import datetime


# ----------------------------
# Patterns
# ----------------------------
CAS_PATTERN = r'\b\d{2,7}-\d{2}-\d\b'


# ----------------------------
# Helpers
# ----------------------------
def clean_chemical_name(name):
    """Standardizes names to prevent duplicates."""
    if not name:
        return "UNKNOWN"

    name = re.sub(r'^[\d\.\-\s•*]+', '', name)
    name = name.split(';')[0].strip()
    return name.upper()


def parse_conservative_percent(conc_text):
    """
    Return the conservative upper bound from a concentration expression.

    Examples:
    '≤10' -> 10
    '<10' -> 10
    '≥10 - ≤25' -> 25
    '10 - 25' -> 25
    '0.3' -> 0.3
    """
    if not conc_text:
        return None

    text = conc_text.strip()

    # Normalize odd dash spacing
    text = re.sub(r'\s+', ' ', text)

    # Case 1: range, e.g. "≥10 - ≤25" or "10 - 25"
    m = re.search(
        r'(?:[≥>]\s*)?(\d+(?:\.\d+)?)\s*-\s*(?:[≤<]\s*)?(\d+(?:\.\d+)?)',
        text
    )
    if m:
        return float(m.group(2))

    # Case 2: single bound/value, e.g. "≤10", "<10", "10", "≤0.3"
    m = re.search(r'[≤<≥>]?\s*(\d+(?:\.\d+)?)', text)
    if m:
        return float(m.group(1))

    return None


def extract_section_3_text(pdf):
    """
    Extract only Section 3 text from the PDF.
    This avoids grabbing CAS and numbers from Section 8 / Section 16.
    """
    full_text = "\n".join([(p.extract_text() or "") for p in pdf.pages])

    # Try common Section 3 / Section 4 headers
    start_patterns = [
        r'Section\s*3\.\s*Composition\s*/\s*information on ingredients',
        r'Section\s*3\.\s*Composition/information on ingredients',
        r'Section\s*3\b'
    ]
    end_patterns = [
        r'Section\s*4\.\s*First aid measures',
        r'Section\s*4\b'
    ]

    start_idx = None
    end_idx = None

    for pat in start_patterns:
        m = re.search(pat, full_text, flags=re.IGNORECASE)
        if m:
            start_idx = m.start()
            break

    if start_idx is None:
        return ""

    for pat in end_patterns:
        m = re.search(pat, full_text[start_idx:], flags=re.IGNORECASE)
        if m:
            end_idx = start_idx + m.start()
            break

    if end_idx is None:
        section3 = full_text[start_idx:]
    else:
        section3 = full_text[start_idx:end_idx]

    return section3


def extract_ingredient_from_line(line):
    """
    Extract one ingredient record from a single line in Section 3.

    Expected examples:
    Titanium Dioxide ≥10 - ≤25 13463-67-7
    Crystalline Silica, respirable powder ≤10 14808-60-7
    Heavy Paraffinic Oil ≤0.3 64742-65-0
    """
    if not line:
        return None

    line = " ".join(line.split())  # normalize whitespace

    cas_match = re.search(CAS_PATTERN, line)
    if not cas_match:
        return None

    cas = cas_match.group(0)
    left = line[:cas_match.start()].strip()

    # Concentration should be the LAST concentration-like token before CAS
    conc_match = re.search(
        r'((?:[≥>]\s*)?\d+(?:\.\d+)?\s*-\s*(?:[≤<]\s*)?\d+(?:\.\d+)?|[≤<≥>]?\s*\d+(?:\.\d+)?)\s*$',
        left
    )
    if not conc_match:
        return None

    conc_text = conc_match.group(1).strip()
    name = left[:conc_match.start()].strip()

    percent = parse_conservative_percent(conc_text)

    # Basic sanity check
    if not name or percent is None:
        return None

    if not (0 <= percent <= 100):
        return None

    return {
        "CAS Number": cas,
        "Contaminant Name": clean_chemical_name(name),
        "Percentage": percent
    }


def extract_section_3_records(section3_text):
    """
    Parse Section 3 line by line.
    Only lines that look like ingredient rows will be returned.
    """
    records = []
    lines = [ln.strip() for ln in section3_text.split("\n") if ln.strip()]

    for line in lines:
        result = extract_ingredient_from_line(line)
        if result:
            records.append(result)

    return records


def detect_solids_flag(full_text):
    """
    Simple global solids flag.
    You can refine this later if your internal logic changes.
    """
    text = full_text.lower()

    solid_keywords = [
        "solid",
        "powder",
        "granules",
        "granular",
        "pellet",
        "crystal",
        "crystalline"
    ]

    return "Y" if any(k in text for k in solid_keywords) else "N"


def detect_volatile_flag(full_text):
    """
    Simple global volatile flag.
    Kept intentionally conservative/simple for now.
    """
    text = full_text.lower()

    volatile_keywords = [
        "volatile",
        "vapor",
        "vapour",
        "evaporat",
        "voc"
    ]

    return "Y" if any(k in text for k in volatile_keywords) else "N"


# ----------------------------
# Main processing
# ----------------------------
def process_all_sds(folder_path):
    all_data = []
    output_name = "Master_SDS_Matrix.xlsx"
    backup_name = f"Master_SDS_Matrix_{datetime.now().strftime('%Y-%m-%d')}.xlsx"

    if not os.path.exists(folder_path):
        print(f"Folder not found: {folder_path}")
        return

    for filename in os.listdir(folder_path):
        if not filename.lower().endswith(".pdf"):
            continue

        product_name = filename.rsplit('.', 1)[0]
        pdf_path = os.path.join(folder_path, filename)

        try:
            print(f"🔍 Processing: {filename}")

            with pdfplumber.open(pdf_path) as pdf:
                full_text = "\n".join([(p.extract_text() or "") for p in pdf.pages])

                # Global product flags
                s = detect_solids_flag(full_text)
                v = detect_volatile_flag(full_text)

                # Restrict parsing to Section 3 only
                section3_text = extract_section_3_text(pdf)

                if not section3_text.strip():
                    print(f"⚠️ Could not find Section 3 in {filename}")
                    continue

                ingredient_rows = extract_section_3_records(section3_text)

                if not ingredient_rows:
                    print(f"⚠️ No ingredient rows found in Section 3 for {filename}")
                    continue

                for row in ingredient_rows:
                    all_data.append({
                        "CAS Number": row["CAS Number"],
                        "Contaminant Name": row["Contaminant Name"],
                        "Solids (Y/N)": s,
                        "Volatile (Y/N)": v,
                        "Product": product_name,
                        "Percentage": row["Percentage"]
                    })

        except Exception as e:
            print(f"❌ Error reading {filename}: {e}")

    if not all_data:
        print("❌ No data found. Please check your PDF readability.")
        return

    df = pd.DataFrame(all_data)

    # Pick one consistent name for each CAS number
    name_map = (
        df.groupby("CAS Number")["Contaminant Name"]
        .apply(lambda x: max(x, key=len))
        .to_dict()
    )
    df["Contaminant Name"] = df["CAS Number"].map(name_map)

    # Pivot to master matrix
    matrix = df.pivot_table(
        index=["Contaminant Name", "CAS Number", "Solids (Y/N)", "Volatile (Y/N)"],
        columns="Product",
        values="Percentage",
        aggfunc="max"
    ).reset_index()

    matrix = matrix.sort_values(by="Contaminant Name")

    matrix.to_excel(output_name, index=False)
    matrix.to_excel(backup_name, index=False)

    print(f"✅ Master Matrix updated. Found {len(matrix)} unique CAS entries.")
    print(f"✅ Saved: {output_name}")
    print(f"✅ Backup saved: {backup_name}")


# ----------------------------
# Run
# ----------------------------
sds_folder = os.path.join(os.getcwd(), "sds_files")
process_all_sds(sds_folder)
