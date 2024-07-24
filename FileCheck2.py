import sys
import pdfplumber
from PyPDF2 import PdfReader, PdfWriter
import re

def extract_serial_number(text):
    try:
        matches = re.findall(r'\b([A-Z]+)(\d{3,})', text, re.IGNORECASE)
        for match in matches:
            prefix, number = match
            if len(number) >= 3:  
                serial_number = f"{prefix}{number}"
                return serial_number
    except Exception as e:
        print(f"Error extracting serial number: {e}")
    return None

def categorize_and_order_pages(input_path):
    try:
        with pdfplumber.open(input_path) as pdf:
            product_pages = []
            certificate_pages = []
            
            for page_number, page in enumerate(pdf.pages):
                text = page.extract_text() or ''
                serial_number = extract_serial_number(text)
                
                if serial_number:
                    print(f"Certificate page found: Page {page_number + 1} with serial number {serial_number}")
                    certificate_pages.append((page_number, serial_number))
                else:
                    print(f"Product page found: Page {page_number + 1}")
                    product_pages.append(page_number)
                    
            return product_pages, certificate_pages
    except Exception as e:
        print(f"Error categorizing pages: {e}")
    return [], []

def rearrange_pdf(input_path, output_path):
    try:
        reader = PdfReader(input_path)
        writer = PdfWriter()

        product_pages, certificate_pages = categorize_and_order_pages(input_path)
        cert_iter = iter(sorted(certificate_pages, key=lambda x: x[1])) 
        for product_page in product_pages:
            writer.add_page(reader.pages[product_page])
            print(f"Adding product page: {product_page + 1}")

            try:
                cert_page = next(cert_iter)
                writer.add_page(reader.pages[cert_page[0]])
                print(f"Adding certificate page: {cert_page[0] + 1}")
            except StopIteration:
                print("No more certificate pages to add.")

        if writer.pages:
            with open(output_path, 'wb') as output_pdf:
                writer.write(output_pdf)
            print(f"PDF saved to: {output_path}")
        else:
            print("No pages were added to the PDF. Output PDF will be empty.")
    except Exception as e:
        print(f"Error processing PDF: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python script.py <input_pdf_path> <output_pdf_path>")
        sys.exit(1)
    input_pdf_path = sys.argv[1]
    output_pdf_path = sys.argv[2]
    rearrange_pdf(input_pdf_path, output_pdf_path)
