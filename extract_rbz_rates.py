import requests
from requests.adapters import HTTPAdapter, Retry
import pdfplumber
import pandas as pd
import os
from datetime import datetime, timedelta

def download_pdf_for_date(year, month, day):
    try:
        url = f"https://www.rbz.co.zw/documents/Exchange_Rates/{year}/{month}/RATES_{day}_{month}_{year}.pdf"
        
        # Setup retry strategy
        retry_strategy = Retry(
            total=5,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["HEAD", "GET", "OPTIONS"],
            backoff_factor=1
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        http = requests.Session()
        http.mount("https://", adapter)
        http.mount("http://", adapter)
        
        response = http.get(url, verify=False, timeout=10)  # Disable SSL verification, set timeout
        response.raise_for_status()  # Raise HTTPError for bad responses

        pdf_path = f"RATES_{day}_{month}_{year}.pdf"
        with open(pdf_path, 'wb') as file:
            file.write(response.content)
        return pdf_path
    except requests.exceptions.RequestException as e:
        print(f"Error downloading PDF for {day} {month} {year}: {e}")
        return None

def extract_exchange_rates(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            table = page.extract_table()
            data = table[1:]  # Extract data rows (excluding headers)
        return data
    except Exception as e:
        print(f"Error extracting data from PDF {pdf_path}: {e}")
        return None

def clean_data(data):
    # Define the fixed headers
    headers = ["CURRENCY", "INDICES", "BID", "ASK", "Mid Rate", "BID", "ASK", "Mid Rate"]

    # Filter out rows with invalid length or containing the date string
    cleaned_data = [row for row in data if len(row) == len(headers) and not any("Wednesday" in cell for cell in row if cell)]

    # Automatically rename duplicate headers to make them unique
    headers = make_column_names_unique(headers)

    return headers, cleaned_data

def make_column_names_unique(headers):
    seen = {}
    unique_headers = []
    for col in headers:
        if col in seen:
            seen[col] += 1
            unique_headers.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            unique_headers.append(col)
    return unique_headers

def save_to_formats(headers, data, pdf_path, date_str):
    try:
        # Ensure headers are valid XML tags
        headers = [header.replace(" ", "_").replace(".", "_") for header in headers]
        
        df = pd.DataFrame(data, columns=headers)
        
        # Create Archive folder if it does not exist
        if not os.path.exists('Archive'):
            os.makedirs('Archive')

        # Create subfolders for each format
        formats = ["pdf", "excel", "json", "csv", "xml", "html", "markdown"]
        for fmt in formats:
            fmt_folder = os.path.join('Archive', fmt)
            if not os.path.exists(fmt_folder):
                os.makedirs(fmt_folder)

        # Define file paths for each format
        file_paths = {
            "pdf": f"Archive/pdf/exchange_rates_{date_str}.pdf",
            "excel": f"Archive/excel/exchange_rates_{date_str}.xlsx",
            "json": f"Archive/json/exchange_rates_{date_str}.json",
            "csv": f"Archive/csv/exchange_rates_{date_str}.csv",
            "xml": f"Archive/xml/exchange_rates_{date_str}.xml",
            "html": f"Archive/html/exchange_rates_{date_str}.html",
            "markdown": f"Archive/markdown/exchange_rates_{date_str}.md"
        }

        # Save the downloaded PDF if it doesn't already exist
        if not os.path.exists(file_paths["pdf"]):
            os.rename(pdf_path, file_paths["pdf"])
        
        # Save to various formats if the file doesn't already exist
        if not os.path.exists(file_paths["excel"]):
            df.to_excel(file_paths["excel"], index=False)
        if not os.path.exists(file_paths["json"]):
            df.to_json(file_paths["json"], orient='records')
        if not os.path.exists(file_paths["csv"]):
            df.to_csv(file_paths["csv"], index=False)
        if not os.path.exists(file_paths["xml"]):
            df.to_xml(file_paths["xml"], index=False)
        if not os.path.exists(file_paths["html"]):
            df.to_html(file_paths["html"], index=False)
        if not os.path.exists(file_paths["markdown"]):
            df.to_markdown(file_paths["markdown"], index=False)

        return True
    except Exception as e:
        print(f"Error converting data to formats for {date_str}: {e}")
        return False

def update_archive_for_year(year):
    start_date = datetime(year, 1, 1)
    end_date = datetime(year, 12, 31)
    delta = timedelta(days=1)

    current_date = start_date
    while current_date <= end_date:
        day = current_date.strftime("%d")
        month = current_date.strftime("%B").capitalize()
        date_str = current_date.strftime("%Y_%m_%d")

        pdf_path = download_pdf_for_date(year, month, day)
        if pdf_path:
            data = extract_exchange_rates(pdf_path)
            if data:
                headers, cleaned_data = clean_data(data)
                success = save_to_formats(headers, cleaned_data, pdf_path, date_str)
                if success:
                    print(f"Successfully archived exchange rates for {date_str}")
                else:
                    print(f"Failed to archive exchange rates for {date_str}")
            else:
                print(f"No data extracted for {date_str}")
        current_date += delta

if __name__ == '__main__':
    update_archive_for_year(2024)
