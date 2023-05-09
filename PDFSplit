import logging
import os
import tempfile
from urllib.parse import urlparse

import azure.functions as func
from PyPDF2 import PdfFileReader, PdfFileWriter
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.client_request import ClientRequest
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.file_creation_information import FileCreationInformation


def main(req: func.HttpRequest, context: func.Out[func.InputStream]) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    # Parse the input URI to get the SharePoint site URL and file path
    uri = req.params.get('uri')
    uri_parts = urlparse(uri)
    site_url = f'{uri_parts.scheme}://{uri_parts.netloc}'
    file_path = uri_parts.path.strip('/')
    logging.info(f'Site URL: {site_url}, File path: {file_path}')

    # Authenticate to SharePoint Online
    auth_ctx = AuthenticationContext(site_url)
    auth_ctx.acquire_token_for_app(client_id=req.headers['client-id'], client_secret=req.headers['client-secret'])
    ctx = ClientContext(site_url, auth_ctx)
    ctx.request_timeout = 1000000
    request = ClientRequest(ctx)

    # Read PDF file from SharePoint Online
    with tempfile.NamedTemporaryFile(delete=False) as f:
        file_url = f'{site_url}/_api/web/GetFileByServerRelativeUrl(\'/{file_path}\')/$value'
        response = request.execute_request_direct(url=file_url)
        f.write(response.content)
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
                file_info = FileCreationInformation()
                file_info.content = page_file.read()
                file_info.url = page_file_name
                file_info.overwrite = True
                folder_url = f'/_catalogs/masterpage/{os.path.dirname(file_path)}'
                folder = ctx.web.get_folder_by_server_relative_url(folder_url)
                folder.files.add(file_info)
                ctx.execute_query()

            # Clean up temp file
            os.unlink(page_file_path)

    # Clean up temp file
    os.unlink(pdf_file_path)

    return func.HttpResponse(f"{pdf_reader.getNumPages()} pages split and saved successfully to SharePoint.")
