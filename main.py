import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
import logging

EXCEL_FILE = "data/data.xlsx"
TEMPLATE_FILE = "templates/template.docx"
OUTPUT_DIR = "output/"
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

logging.basicConfig(level=logging.INFO, format='%(message)s')

def start_automation():
    df = pd.read_excel(EXCEL_FILE)
    
    df.columns = [c.replace(' ', '_') for c in df.columns]
    
    for index, row in df.iterrows():
        try:
            doc = DocxTemplate(TEMPLATE_FILE)
            
            context = row.to_dict()
            doc.render(context)

            client_name = row['Full_Name']
            
            word_path = os.path.join(OUTPUT_DIR, f"Certificate_{client_name}.docx")

            doc.save(word_path)
            
            print(f"Generating PDF for: {client_name}...")
            convert(word_path) 
            
            logging.info(f"Success for: {client_name}")
            
        except Exception as e:
            logging.error(f"Error at row {index}: {e}")

if __name__ == "__main__":
    start_automation()
