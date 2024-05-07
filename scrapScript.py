import tkinter as tk
from tkinter import messagebox
import requests
from bs4 import BeautifulSoup
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import urllib3
import os
import sys

# Suppress SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Determine the base path
if hasattr(sys, "_MEIPASS"):
    # If running as a bundled executable
    base_path = sys._MEIPASS
else:
    # If running as a regular Python script, use the current working directory
    base_path = os.path.abspath(".")

# Construct the path to the credentials file directly
creds_path = os.path.join(base_path, "scrap-salvagesalvage-data-71cd1942426e.json")

# Check if the credentials file exists at the constructed path
if not os.path.isfile(creds_path):
    print(f"Error: Credentials file not found at {creds_path}")
else:
    print(f"Credentials file found at {creds_path}")

# Function to set up Google Sheets API credentials and return a worksheet object
def get_google_sheet(sheet_id, sheet_name):
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
        return sheet
    except gspread.exceptions.WorksheetNotFound:
        messagebox.showerror("Google Sheets Error", f"Worksheet '{sheet_name}' not found.")
        raise
    except gspread.exceptions.SpreadsheetNotFound:
        messagebox.showerror("Google Sheets Error", "Spreadsheet not found. Check the spreadsheet ID.")
        raise
    except Exception as e:
        messagebox.showerror("Google Sheets Error", f"Unable to access Google Sheet.\nDetails: {e}")
        raise

# Function to fetch car details from a given URL
def fetch_car_data(url):
    response = requests.get(url, verify=False)
    soup = BeautifulSoup(response.content, 'html.parser')

    car_title = soup.select_one('div.product-title').text.strip()
    lot_number = soup.select_one('div.product-sku').text.strip()
    car_description = soup.select_one('div.product-summary > div').text.strip()
    return car_title, lot_number, car_description

# Function to find the next empty row in the Google Sheet
def find_next_empty_row(sheet):
    str_list = list(filter(None, sheet.col_values(1)))
    return len(str_list) + 1

# Function to add car data to the Google Sheet
def add_car_data_to_google_sheet(sheet_id, sheet_name, url):
    try:
        car_title, lot_number, car_description = fetch_car_data(url)
    except Exception as e:
        messagebox.showerror("Error", f"Unable to fetch data from the URL.\nDetails: {e}")
        return

    try:
        sheet = get_google_sheet(sheet_id, sheet_name)
    except Exception as e:
        messagebox.showerror("Error", f"Unable to access Google Sheet.\nDetails: {e}")
        return

    next_row = find_next_empty_row(sheet)
    today = datetime.today().strftime('%Y,%m,%d')
    sheet.update_cell(next_row, 1, today)  # Date
    sheet.update_cell(next_row, 2, car_title)  # Car Info
    sheet.update_cell(next_row, 3, lot_number)  # Lot Number
    sheet.update_cell(next_row, 4, "Salvage")  # Car Status
    sheet.update_cell(next_row, 5, car_description)  # Car Description

    messagebox.showinfo("Success", "Car data has been added successfully!")

# GUI Setup using tkinter
def create_gui():
    root = tk.Tk()
    root.title("Car Data Scraper")

    # URL input field
    tk.Label(root, text="Car Data Page URL:").grid(row=0, column=0, padx=10, pady=10)
    url_entry = tk.Entry(root, width=60)
    url_entry.grid(row=0, column=1, padx=10, pady=10)

    # Function to handle data scraping and updating the Google Sheet
    def on_scrape():
        url = url_entry.get().strip()
        if not url:
            messagebox.showwarning("Input Error", "Please enter a valid URL.")
            return

        # my Google Sheet ID and sheet name
        google_sheet_id = "1zAeHbVYi2NHvx3r8tiWQEtQaKpHgL2Gp3lT4u42G3lY"
        sheet_name = "SalvageSalvage"
        add_car_data_to_google_sheet(google_sheet_id, sheet_name, url)

    # Scrape Button which is pretty simple atm
    scrape_button = tk.Button(root, text="Scrape Data", command=on_scrape)
    scrape_button.grid(row=1, column=1, pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_gui()