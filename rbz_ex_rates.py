import streamlit as st
import requests
from bs4 import BeautifulSoup
import pdfplumber
import pandas as pd
import os

def download_latest_pdf():
    url = "https://www.rbz.co.zw/documents/Exchange_Rates/2024/June/"
    response = requests.get(url, verify=False)  # Disable SSL verification
    soup = BeautifulSoup(response.text, 'html.parser')
    links = soup.find_all('a', href=True)
    pdf_link = None

    for link in links:
        if link['href'].endswith('.pdf') and "RATES" in link['href']:
            pdf_link = link['href']
            break

    if pdf_link:
        pdf_url = f"https://www.rbz.co.zw{pdf_link}"
        pdf_path = "latest_exchange_rates.pdf"
        response = requests.get(pdf_url, verify=False)  # Disable SSL verification
        with open(pdf_path, 'wb') as file:
            file.write(response.content)
        return pdf_path
    return None

def extract_exchange_rates(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        table = page.extract_table()
        headers = table[0]
        data = table[1:]
    return headers, data

def convert_to_formats(headers, data):
    df = pd.DataFrame(data, columns=headers)

    # Save to various formats
    df.to_excel("exchange_rates.xlsx", index=False)
    df.to_json("exchange_rates.json", orient='records')
    df.to_csv("exchange_rates.csv", index=False)
    df.to_xml("exchange_rates.xml", index=False)

    return df

def display_exchange_rates():
    st.title("RBZ Exchange Rates")
    pdf_path = download_latest_pdf()
    if pdf_path:
        headers, data = extract_exchange_rates(pdf_path)
        df = convert_to_formats(headers, data)
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
