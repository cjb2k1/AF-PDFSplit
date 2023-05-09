import logging
import os
import tempfile
from urllib.parse import urlparse

import azure.functions as func
import requests
from PyPDF2 import PdfFileReader, PdfFileWriter


def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    # Parse the input URI to get the SharePoint site URL and file path
    uri = req.params.get('uri')
    uri_parts = urlparse(uri)
    site_url = f'{uri_parts.scheme}://{uri_parts.netloc}'
    file_path = uri_parts.path.strip('/')
    logging.info(f'Site URL: {site_url}, File path: {file_path}')

    # Authenticate to Microsoft Graph API
    auth_url = f'https://login.microsoftonline.com/{os.environ["TenantId"]}/oauth2/v2.0/token'
    auth_data = {
        'grant_type': 'client_credentials',
        'client_id': os.environ['ClientId'],
        'client_secret': os.environ['ClientSecret'],
        'scope': 'https://graph.microsoft.com/.default'
    }
    auth_response = requests.post(auth_url, data=auth_data)
    auth_response.raise_for_status()
    access_token = auth_response.json()['access_token']
    headers = {'Authorization': f'Bearer {access_token}'}

    # Read PDF file from SharePoint Online
    with tempfile.NamedTemporaryFile(delete=False) as f:
        file_url = f'{site_url}/_api/web/GetFileByServerRelativeUrl(\'/{file_path}\')/$value'
        file_response = requests.get(file_url, headers=headers)
        f.write(file_response.content)
        pdf_file_path = f.name

    # Split PDF file into individual pages and save each page to SharePoint Online
    with open(pdf_file_path, 'rb') as pdf_file:
        pdf_reader = PdfFileReader(pdf_file)
        file_name, file_ext = os.path.splitext(os.path.basename(file_path))
        file_name_suffix = f'_part_{pdf_reader.getNumPages()}'
        for page_num in range(pdf_reader.getNumPages()):
            # Create a new PDF file with a single page
            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(pdf_reader.getPage(page_num))
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as f:
                pdf_writer.write(f)
                page_file_path = f.name
            
            # Save the new PDF file to SharePoint Online with the same metadata and filename
            with open(page_file_path, 'rb') as page_file:
                page_file_name = f'{file_name}{file_name_suffix}_{page_num+1}{file_ext}'
                logging.info(f'Saving page {page_num+1} to {page_file_name}')
                file_url = f'{site_url}/_api/v2.0/drives/{os.environ["DriveId"]}/root:/{file_path[:-len(os.path.basename(file_path))]}{page_file_name}:/content'
                file_response = requests.put(file_url, headers=headers, data=page_file)
                file_response.raise_for_status()

            # Clean up temp file
            os.unlink(page_file_path)

    # Clean up temp file
    os.unlink(pdf_file_path)

    return func.HttpResponse(f"{pdf_reader.getNumPages()} pages split and saved successfully to SharePoint.")
``
