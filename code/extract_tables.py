import os
import pandas as pd
import pdfplumber

# ==========================
# Paths
# ==========================
input_dir = r"D:\RA\YM\1394\pdf\honar"     # Folder containing PDF files
output_dir = r"D:\RA\YM\1394\Excel"         # Folder to save Excel outputs
os.makedirs(output_dir, exist_ok=True)      # Create output folder if it doesn't exist

# ==========================
# Helper Functions
# ==========================

def reverse_strings(df):
    """
    Reverse strings in a DataFrame.
    Useful for Persian text extracted from PDFs which may appear reversed.
    """
    return df.applymap(lambda x: x[::-1] if isinstance(x, str) else x)


def extract_header_above_table(page, table_bbox, height_above=60):
    """
    Extract the two lines of header above a table.
    
    Args:
        page (pdfplumber.page.Page): PDF page object
        table_bbox (tuple): Table bounding box (x0, y0, x1, y1)
        height_above (int): Height above the table to scan for header
    
    Returns:
        first_line (str), second_line (str)
    """
    x0, y0, x1, y1 = table_bbox
    page_width = page.width

    header_bbox = (0, max(y0 - height_above, 0), page_width, y0)
    text = page.within_bbox(header_bbox).extract_text()

    if not text:
        return "", ""

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    first_line = lines[0][::-1] if len(lines) > 0 else ""
    second_line = lines[1][::-1] if len(lines) > 1 else ""
    return first_line, second_line


def extract_lines_below_table(page, table_bbox, height_scan=70):
    """
    Extract lines below a table (lines 3,4,5) and combine them into a single string.
    
    Args:
        page (pdfplumber.page.Page): PDF page object
        table_bbox (tuple): Table bounding box (x0, y0, x1, y1)
        height_scan (int): Height below the table to scan
    
    Returns:
        combined (str): Combined lines as a single string
    """
    x0, y0, x1, y1 = table_bbox
    page_width = page.width
    page_height = page.height

    footer_bbox = (0, y1, page_width, min(y1 + height_scan, page_height))
    text = page.within_bbox(footer_bbox).extract_text()

    if not text:
        return ""

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    third_line = lines[2][::-1] if len(lines) > 2 else ""
    fourth_line = lines[3][::-1] if len(lines) > 3 else ""
    fifth_line = lines[4][::-1] if len(lines) > 4 else ""

    combined = third_line
    for l in [fourth_line, fifth_line]:
        if l:
            combined += "\n" + l  # You can change separator to space or "|" if preferred

    return combined


# ==========================
# Main PDF Processing Function
# ==========================

def process_pdf(pdf_path, reshte):
    """
    Process a single PDF file: extract tables, reverse Persian text,
    extract headers/footers, and save each table/page as Excel.

    Args:
        pdf_path (str): Path to PDF file
        reshte (str): Name of the subject/major (e.g., 'هنر')
    """
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            print(f"📄 Processing page {page_num} of {os.path.basename(pdf_path)}")

            # Extract tables using pdfplumber
            tables = page.extract_tables()
            table_objs = page.find_tables()

            all_dfs = []

            # Ensure tables and table objects match
            if tables and table_objs and len(tables) == len(table_objs):
                for idx, (table, table_obj) in enumerate(zip(tables, table_objs)):
                    if not table:
                        continue

                    df = pd.DataFrame(table)
                    df = reverse_strings(df)  # Fix Persian text orientation

                    # Extract two header lines above the table
                    header_line1, header_line2 = extract_header_above_table(page, table_obj.bbox)

                    # Extract lines 3-5 below the table
                    footer_combined = extract_lines_below_table(page, table_obj.bbox)

                    # Add metadata columns
                    df["شماره صفحه"] = page_num
                    df["ردیف"] = range(1, len(df) + 1)
                    df["رشته دبیرستان"] = reshte
                    df["عنوان جدول"] = header_line1
                    df["عنوان دوم"] = header_line2
                    df["خط سوم تا پنجم پایین جدول"] = footer_combined

                    all_dfs.append(df)

            # Save concatenated tables for this page
            if all_dfs:
                result_df = pd.concat(all_dfs, ignore_index=True)
                excel_filename = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_page_{page_num}.xlsx"
                result_df.to_excel(os.path.join(output_dir, excel_filename), index=False)
                print(f"✅ Saved: {excel_filename}")


# ==========================
# Run Processing for All PDFs in Input Folder
# ==========================

for pdf_file in os.listdir(input_dir):
    if pdf_file.endswith(".pdf"):
        pdf_path = os.path.join(input_dir, pdf_file)
        process_pdf(pdf_path, reshte="هنر")
