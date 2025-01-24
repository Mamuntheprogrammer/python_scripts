import re
from pdf2image import convert_from_path
import pytesseract


# Download Link : https://github.com/UB-Mannheim/tesseract/wiki

# Path to your Tesseract executable (adjust as needed)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_text_from_pdf(pdf_path):
    # Convert PDF pages to images
    images = convert_from_path(pdf_path)
    text = ""
    for image in images:
        text += pytesseract.image_to_string(image)
    return text

def extract_info(text):
    # Regex patterns for email, phone numbers, and names
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    phone_pattern = r'\b(?:\+?\d{1,3}[-.\s]?)?(?:\(?\d{3}\)?[-.\s]?)?\d{3}[-.\s]?\d{4}\b'
    name_pattern = r'\b[A-Z][a-z]+\s[A-Z][a-z]+(?:\s[A-Z][a-z]+)?\b'  # Example for full names
    
    emails = re.findall(email_pattern, text)
    phones = re.findall(phone_pattern, text)
    names = re.findall(name_pattern, text)
    
    return {
        "emails": emails,
        "phones": phones,
        "names": names
    }

def main():
    pdf_path = "one.pdf"  # Update with your PDF file path
    print("Extracting text from PDF...")
    text = extract_text_from_pdf(pdf_path)
    
    print("Extracting information...")
    extracted_info = extract_info(text)
    
    print("Extraction complete:")
    print("Names:", extracted_info["names"])
    print("Phone Numbers:", extracted_info["phones"])
    print("Emails:", extracted_info["emails"])

if __name__ == "__main__":
    main()
