import os
import time
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from docx import Document
from zipfile import ZipFile
from dotenv import load_dotenv
import openai
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Load environment variables from .env file
load_dotenv()
openai.api_key = os.getenv('OPENAI_API_KEY')

def test_openai_api():
    """Test the OpenAI API to verify that the setup is correct."""
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",  # Use "gpt-4" if you have access
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": "Test if OpenAI API is working correctly."}
            ]
        )
        if response and response.choices:
            print("OpenAI API is working correctly. Test response:", response.choices[0].message['content'])
        else:
            raise Exception("OpenAI API test failed: No response or invalid response structure.")
    except Exception as e:
        print(f"An error occurred while testing the OpenAI API: {e}")
        raise

def measure_execution_time(func):
    """Measure the execution time of a function."""
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        execution_time = end_time - start_time
        print(f"Total execution time: {execution_time:.2f} seconds")
        return result
    return wrapper

def correct_text(text):
    """Correct text using OpenAI GPT"""
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",  # Use "gpt-4" if you have access
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"Edit the following text for grammar, spelling, punctuation, and clarity. Make necessary adjustments to improve sentence structure and coherence without altering the original meaning:\n\n{text}\n\nProvide the revised version below."}
            ]
        )
        corrected_text = response.choices[0].message['content']
        return corrected_text.strip()
    except Exception as e:
        print(f"An error occurred during text correction: {e}")
        return text  # Return the original text if correction fails

@measure_execution_time
def main():
    # Paths
    pdf_path = 'file.pdf'  # Input PDF file
    image_folder = 'images'  # Folder to store images
    output_docx = 'output.docx'  # Output Word document
    output_pdf = 'output.pdf'  # Output PDF file
    output_txt = 'output.txt'  # Output TXT file
    zip_file_path = 'images.zip'  # Output ZIP file

    # Create a folder to store images if it doesn't exist
    os.makedirs(image_folder, exist_ok=True)

    # Step 1: Convert PDF to Images
    print("Converting PDF pages to images...")
    images = convert_from_path(pdf_path, dpi=300)

    # Step 2: Perform OCR on Images and Correct Text
    print("Performing OCR and correcting text...")
    corrected_texts = []
    for i, image in enumerate(images):
        image_path = os.path.join(image_folder, f'page_{i + 1}.jpg')
        image.save(image_path, 'JPEG')

        # Extract text using Tesseract
        text = pytesseract.image_to_string(image, lang='eng')
        print("Original text:")
        print(text)

        # Correct text using OpenAI GPT
        corrected_text = correct_text(text)
        print("Corrected text:")
        print(corrected_text)

        # Append both original and corrected text
        combined_text = f"Original Text:\n{text}\n\nCorrected Text:\n{corrected_text}\n"
        corrected_texts.append(combined_text)

    # Step 3: Create and Format Document
    print("Creating and formatting Word document...")
    doc = Document()

    # Set default font style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = 10  # 10 pt

    # Configure paragraph formatting
    paragraph_format = style.paragraph_format
    paragraph_format.line_spacing = 1.5

    # Add original and corrected text to the document, each on a new page
    for combined_text in corrected_texts:
        doc.add_paragraph(combined_text)
        doc.add_page_break()  # Add a page break after each page's content

    # Save the document in DOCX format
    doc.save(output_docx)
    print(f"Document saved as {output_docx}.")

    # Step 4: Create PDF with ReportLab
    print("Creating PDF file...")
    pdf = canvas.Canvas(output_pdf, pagesize=letter)
    pdf.setFont("Times-Roman", 10)

    for combined_text in corrected_texts:
        # Split text into lines and draw each line
        lines = combined_text.split('\n')
        y_position = 750  # Start position for text on the page
        for line in lines:
            pdf.drawString(30, y_position, line)
            y_position -= 15  # Move to next line with spacing

            if y_position < 40:  # Start a new page if the text reaches the bottom
                pdf.showPage()
                pdf.setFont("Times-Roman", 10)
                y_position = 750

        # Ensure a new page for the next combined text
        pdf.showPage()

    pdf.save()
    print(f"PDF saved as {output_pdf}.")

    # Step 5: Save the corrected text to a TXT file
    with open(output_txt, 'w') as file:
        for combined_text in corrected_texts:
            file.write(combined_text)
            file.write('\n\n--- End of Page ---\n\n')  # Separate pages in TXT file
    print(f"Text saved as {output_txt}.")

    # Step 6: Compress Images to ZIP
    print("Creating ZIP file of images...")
    with ZipFile(zip_file_path, 'w') as zipf:
        for i, image in enumerate(images):
            img_path = os.path.join(image_folder, f'page_{i + 1}.jpg')
            zipf.write(img_path, os.path.basename(img_path))
    print(f"Images compressed into {zip_file_path}.")

    print("All processes completed successfully.")

# Run the OpenAI API test
test_openai_api()

# Run the main function with timing
main()
