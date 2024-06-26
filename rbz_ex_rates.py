import streamlit as st
import requests
from requests.adapters import HTTPAdapter, Retry
import pdfplumber
import pandas as pd
from datetime import datetime

def download_latest_pdf():
    try:
        # Construct the URL based on the current date
        today = datetime.today()
        url = f"https://www.rbz.co.zw/documents/Exchange_Rates/{today.year}/{today.strftime('%B').capitalize()}/RATES_{today.day}_{today.strftime('%B').upper()}_{today.year}.pdf"
        
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

        pdf_path = "latest_exchange_rates.pdf"
        with open(pdf_path, 'wb') as file:
            file.write(response.content)
        return pdf_path
    except requests.exceptions.RequestException as e:
        st.error(f"Error downloading PDF: {e}")
        print(f"Error downloading PDF: {e}")
        return None

def extract_exchange_rates(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page = pdf.pages[0]
            table = page.extract_table()
            headers = table[0]
            data = table[1:]
        return headers, data
    except Exception as e:
        st.error(f"Error extracting data from PDF: {e}")
        print(f"Error extracting data from PDF: {e}")
        return None, None

def clean_data(headers, data):
    # Ensure headers are valid column names
    valid_headers = []
    for i, header in enumerate(headers):
        if header is None or not header.isalnum():
            valid_headers.append(f"column_{i}")
        else:
            valid_headers.append(header)

    # Filter out rows with invalid length or containing the date string
    cleaned_data = [row for row in data if len(row) == len(valid_headers) and not any("Wednesday" in cell for cell in row if cell)]

    return valid_headers, cleaned_data

def make_column_names_unique(headers):
    seen = {}
    for i, col in enumerate(headers):
        if col in seen:
            seen[col] += 1
            headers[i] = f"{col}_{seen[col]}"
        else:
            seen[col] = 0
    return headers

def convert_to_formats(headers, data):
    try:
        headers = make_column_names_unique(headers)
        df = pd.DataFrame(data, columns=headers)
        # Save to various formats
        df.to_excel("exchange_rates.xlsx", index=False)
        df.to_json("exchange_rates.json", orient='records')
        df.to_csv("exchange_rates.csv", index=False)
        df.to_xml("exchange_rates.xml", index=False)
        return df
    except Exception as e:
        st.error(f"Error converting data to formats: {e}")
        print(f"Error converting data to formats: {e}")
        return None

def display_exchange_rates():
    st.title("RBZ Exchange Rates")
    pdf_path = download_latest_pdf()
    if pdf_path:
        headers, data = extract_exchange_rates(pdf_path)
        if headers and data:
            headers, data = clean_data(headers, data)
            df = convert_to_formats(headers, data)
            if df is not None:
                st.write(df)
                
                st.download_button(
                    label="Download as Excel",
                    data=open("exchange_rates.xlsx", "rb").read(),
                    file_name="exchange_rates.xlsx"
                )
                
                st.download_button(
                    label="Download as JSON",
                    data=open("exchange_rates.json", "rb").read(),
                    file_name="exchange_rates.json"
                )
                
                st.download_button(
                    label="Download as CSV",
                    data=open("exchange_rates.csv", "rb").read(),
                    file_name="exchange_rates.csv"
                )
                
                st.download_button(
                    label="Download as XML",
                    data=open("exchange_rates.xml", "rb").read(),
                    file_name="exchange_rates.xml"
                )
    else:
        st.error("Failed to download the latest exchange rates PDF.")

if __name__ == '__main__':
    display_exchange_rates()
