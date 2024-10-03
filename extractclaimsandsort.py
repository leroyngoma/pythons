import os
import re
import shutil
import PyPDF2
import openpyxl

# Folder containing the PDFs
pdf_folder = r'C:\folder'

# Folder where the organized PDFs will be stored
output_folder = r'C:\folder'

# Path to the Excel file where claim numbers will be recorded
excel_file = r'C:\folder\name.xlsx'

# Ensure the output folder exists
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Create a new Excel workbook and sheet
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Claim Numbers"
sheet.append(["PDF File Name", "Claim Number"])  # Adding header row

# Function to extract claim number from PDF text
def extract_claim_number(pdf_text):
    # Define regular expressions to search for claim details
    patterns = [
        r'claim number[:\s]+(\w+)',       # e.g., "Claim Number: XYZ123"
        r'claim no[:\s]+(\w+)',           # e.g., "Claim No: XYZ123"
        r'claim reference[:\s]+(\w+)'     # e.g., "Claim Reference: XYZ123"
    ]
    
    # Search the text for each pattern
    for pattern in patterns:
        match = re.search(pattern, pdf_text, re.IGNORECASE)
        if match:
            return match.group(1)
    
    return None

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text()
            return text
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return None

# Process each PDF in the source folder
for filename in os.listdir(pdf_folder):
    if filename.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, filename)
        print(f"Processing: {filename}")

        # Extract text from the PDF
        pdf_text = extract_text_from_pdf(pdf_path)

        if pdf_text:
            # Extract the claim number from the PDF text
            claim_number = extract_claim_number(pdf_text)
            
            if claim_number:
                # Create a folder named after the claim number
                claim_folder = os.path.join(output_folder, claim_number)
                if not os.path.exists(claim_folder):
                    os.makedirs(claim_folder)
                
                # Copy the PDF into the claim folder
                shutil.copy(pdf_path, os.path.join(claim_folder, filename))
                print(f"Copied '{filename}' to folder: {claim_folder}")
                
                # Add the extracted claim number and file name to the Excel sheet
                sheet.append([filename, claim_number])
            else:
                print(f"No claim number found in {filename}")
        else:
            print(f"Could not extract text from {filename}")

# Save the Excel file
wb.save(excel_file)

print(f"Processing complete! Claim numbers saved to {excel_file}")
