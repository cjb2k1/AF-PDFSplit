import logging
import os
import tempfile
from PyPDF2 import PdfFileReader, PdfFileWriter
import azure.functions as func

def split_pdf(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    # Get the binary PDF data from the request body
    pdf_binary = req.get_body()

    # Create a temporary directory to store the split PDF pages
    temp_dir = tempfile.mkdtemp()

    try:
        # Split PDF file into individual pages
        pdf_reader = PdfFileReader(pdf_binary)
        num_pages = pdf_reader.getNumPages()

        # Create a list to store the binary data of the individual PDF pages
        page_data = []

        # Split the PDF into pages and save each page as a separate PDF file
        for page_num in range(num_pages):
            # Create a new PDF file with a single page
            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(pdf_reader.getPage(page_num))

            # Generate a unique filename for the page
            page_filename = f'page_{page_num + 1}.pdf'
            page_filepath = os.path.join(temp_dir, page_filename)

            # Save the page as a separate PDF file
            with open(page_filepath, 'wb') as page_file:
                pdf_writer.write(page_file)

            # Read the binary data of the page file
            with open(page_filepath, 'rb') as page_file:
                page_binary_data = page_file.read()

            # Add the binary data to the list
            page_data.append(page_binary_data)

        # Return the list of individual PDF page binary data as the response
        return func.HttpResponse(body=page_data, status_code=200, mimetype='application/pdf')

    finally:
        # Clean up the temporary directory
        for filename in os.listdir(temp_dir):
            filepath = os.path.join(temp_dir, filename)
            os.remove(filepath)
        os.rmdir(temp_dir)
