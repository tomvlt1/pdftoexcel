import os
import re
import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def extract_table_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)
    all_data = []
    exclude_phrases = ["FIN DE LA LISTE", "**************************", "-----------------------"]

    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")

        lines = text.split('\n')
        table_started = False
        page_data = []
        for line in lines:
            if table_started:
                if line.strip() == '' or any(phrase in line for phrase in exclude_phrases):
                    continue
                else:
                    page_data.append(line)
            elif '--------------------' in line:
                table_started = True

        if page_data:
            all_data.extend(page_data)

    return all_data

def split_line(line):
    # Split based on more than two spaces
    split_data = re.split(r'\s{3,}', line)

    # Keep only the number from the first element
    if split_data:
        first_element = split_data[0]
        number = first_element.split(' ', 1)[0]  # Get only the number
        split_data = [number] + split_data[1:]

    # Remove any "???????" from the last element if present
    if split_data and len(split_data) > 1:
        split_data[-1] = re.sub(r'\?+', '', split_data[-1])
    
    return split_data

def process_files_in_folder(input_folder, output_excel_path):
    combined_data = []
    wb = Workbook()
    ws_combined = wb.active
    ws_combined.title = "Combined Data"
    
    for filename in os.listdir(input_folder):
        file_path = os.path.join(input_folder, filename)
        
        if filename.endswith(".pdf"):
            table_data = extract_table_from_pdf(file_path)
            if table_data:
                combined_data.append([filename])  # Title related to the name of the PDF
                for line in table_data:
                    combined_data.append([line])  # Original line in one cell
                    split_data = split_line(line)
                    combined_data.append([''] * 9 + split_data)  # Add split data 10 columns to the right
                combined_data.append([''] * 16)  # Add two empty lines (10 + 6 for split columns)
                combined_data.append([''] * 16)

        elif filename.endswith(".xlsx"):
            df = pd.read_excel(file_path)
            sheet_name = os.path.splitext(filename)[0]
            ws = wb.create_sheet(title=sheet_name)
            for row in dataframe_to_rows(df, index=False, header=True):
                ws.append(row)

    # Write the combined PDF data to the first sheet
    df_combined = pd.DataFrame(combined_data)
    for row in dataframe_to_rows(df_combined, index=False, header=False):
        ws_combined.append(row)

    # Save the workbook
    wb.save(output_excel_path)
    print(f"Combined table and Excel sheets saved to {output_excel_path}")

# Example usage
input_folder = "input"  # Change this to the path of your folder containing PDFs and Excel files
output_excel_path = "output/combined_output.xlsx"  # Change this to the desired path of the combined Excel file

process_files_in_folder(input_folder, output_excel_path)
