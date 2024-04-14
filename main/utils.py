import os
import re
import docx2txt
from openpyxl import Workbook
from PyPDF2 import PdfReader

def extract_text_from_docx(file_path):
    return docx2txt.process(file_path)

def extract_text_from_pdf(file_path):
    text = ""
    with open(file_path, 'rb') as f:
        reader = PdfReader(f)
        for page in reader.pages:
            text += page.extract_text()
    return text

def extract_info_from_cv(cv_file):
    file_ext = cv_file.split('.')[-1]
    text = ""
    if file_ext == 'docx':
        text = extract_text_from_docx(cv_file)
    elif file_ext == 'pdf':
        text = extract_text_from_pdf(cv_file)

    emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
    
    contact_numbers = re.findall(r'(\+91|0)?(\d{10})', text)
    cleaned_text = text
    for email in emails:
        cleaned_text = cleaned_text.replace(email, '')
    for phone_tuple in contact_numbers:
        cleaned_text = cleaned_text.replace(phone_tuple[1], '')
    
    return {'Email': emails, 'Contact Number': contact_numbers, 'Text': cleaned_text.strip()}

def extract_info_and_generate_excel(file_path):
    # Extract information from CV file
    cv_info = extract_info_from_cv(file_path)

    # Create a new Excel workbook and worksheet
    wb = Workbook()
    ws = wb.active

    # Write headers
    ws.append(['CV File', 'Email', 'Contact Number', 'Text'])

    # Write extracted information to the worksheet
    ws.append([os.path.basename(file_path), ', '.join(cv_info['Email']), ', '.join([phone[1] for phone in cv_info['Contact Number']]), cv_info['Text']])

    # Save the Excel file
    excel_file_path = 'CV_Info.xlsx'
    wb.save(excel_file_path)

    return excel_file_path
