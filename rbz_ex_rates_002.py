import streamlit as st
import requests
from requests.adapters import HTTPAdapter, Retry
import pdfplumber
import pandas as pd
import os
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
            data = table[1:]  # Extract data rows (excluding headers)
        return data
    except Exception as e:
        st.error(f"Error extracting data from PDF: {e}")
        print(f"Error extracting data from PDF: {e}")
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

def convert_to_formats(headers, data):
    try:
        # Ensure headers are valid XML tags
        headers = [header.replace(" ", "_").replace(".", "_") for header in headers]
        
        df = pd.DataFrame(data, columns=headers)
        
        # Create Archive folder if it does not exist
        if not os.path.exists('Archive'):
            os.makedirs('Archive')
        
        # Use the current date and time as part of the file name to avoid overwriting
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save to various formats in the Archive folder
        df.to_excel(f"Archive/exchange_rates_{timestamp}.xlsx", index=False)
        df.to_json(f"Archive/exchange_rates_{timestamp}.json", orient='records')
        df.to_csv(f"Archive/exchange_rates_{timestamp}.csv", index=False)
        df.to_xml(f"Archive/exchange_rates_{timestamp}.xml", index=False)
        df.to_html(f"Archive/exchange_rates_{timestamp}.html", index=False)
        df.to_markdown(f"Archive/exchange_rates_{timestamp}.md", index=False)
        
        # Also save the current files for immediate download
        df.to_excel("exchange_rates.xlsx", index=False)
        df.to_json("exchange_rates.json", orient='records')
        df.to_csv("exchange_rates.csv", index=False)
        df.to_xml("exchange_rates.xml", index=False)
        df.to_html("exchange_rates.html", index=False)
        df.to_markdown("exchange_rates.md", index=False)

        return df
    except Exception as e:
        st.error(f"Error converting data to formats: {e}")
        print(f"Error converting data to formats: {e}")
        return None

def display_exchange_rates():
    st.title("RBZ Exchange Rates")
    pdf_path = download_latest_pdf()
    if pdf_path:
        data = extract_exchange_rates(pdf_path)
        if data:
            headers, data = clean_data(data)
            
            # Debug: Print headers and data
            print("Headers:", headers)
            print("Data sample:", data[:5])

            # Extract the day and USD Midrate
            reporting_day = datetime.today().strftime("%A, %d %B %Y")
            usd_midrate = None
            
            # Access the USD Midrate directly using the provided column and row indices
            try:
                usd_midrate = data[3][7]  # column_7 corresponds to the 5th column (0-based index)
            except IndexError as e:
                st.error(f"Error accessing USD Midrate: {e}")
                print(f"Error accessing USD Midrate: {e}")

            # Display the day and USD Midrate
            st.subheader(f"Exchange Rates for {reporting_day}")
            if usd_midrate:
                st.write(f"**USD Midrate:** {usd_midrate}")
            else:
                st.write("USD Midrate not found.")

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

                st.download_button(
                    label="Download as HTML",
                    data=open("exchange_rates.html", "rb").read(),
                    file_name="exchange_rates.html"
                )

                st.download_button(
                    label="Download as Markdown",
                    data=open("exchange_rates.md", "rb").read(),
                    file_name="exchange_rates.md"
                )
    else:
        st.error("Failed to download the latest exchange rates PDF.")

if __name__ == '__main__':
    display_exchange_rates()
