import tkinter as tk
from tkinter import filedialog, messagebox
from fillpdf import fillpdfs
from datetime import datetime
from state_mapping import zip_dict
import csv
import re
import os
import sys
import ftfy

# Function to read CSV as a list of dictionaries


def read_csv_as_dict(filename):
    data = []
    with open(filename, mode='r', newline='', encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row)
    return data

# Function to extract rate information from a string


def get_rate(rate):
    rate_dict = {}
    rate_dict['percentage'] = re.search(r'(\d.*)%', rate).group(1)
    rate_dict['price'] = re.search(r'\$(\d\.?\d?\d?)', rate).group(1)
    rate_dict['monthly'] = re.search(r'Monthly\s*\$(\d+)', rate).group(1)
    return rate_dict

# Function to map state names to state codes


def zip_mapping(dict, full_state):
    try:
        for i in dict:
            n = len(i['name'])
            if i['name'] == full_state[:n]:
                return i['code']
        messagebox.showerror("Error", "Zip Error")
        return ''
    except Exception:
        messagebox.showerror("Error", "Zip Error")
        return
    


def get_legal_name_suffix(business_name):
    try:
        match = re.search(r'\s(\w+)$', business_name)
    except Exception:
        messagebox.showerror("Error", "Legal Name Error")
        return
    return match.group(1) if match else ''

# Function to check required fields


def check_required_fields(data, required_fields):
    missing_fields = []
    for field in required_fields:
        if field not in data or not data[field].strip():
            missing_fields.append(field)
    return missing_fields

# Function to preprocess data for PDF filling


def prejob(data):
    edited_data = {}
    required_fields = [
        'Business Phone', 'Date of Birth', 'Mobile', 'State', 'Home State',
        'State Issued', 'Rate', 'Pricing Type', 'Legal Name of Business',
        'DBA', 'Tax ID', 'Owner Name', 'Street', 'City', 'ZIP', 'Bank Name',
        'Bank Routing', 'Bank Account', 'Social Security Number', 'Home Street',
        'Home City', 'Home ZIP', 'Driver License Number'
    ]
    missing_fields = check_required_fields(data, required_fields)
    if missing_fields:
        message = f"Missing or empty required fields: {
            ', '.join(missing_fields)}"
        user_response = messagebox.askyesno(
            "Missing Fields", message + "\nDo you want to continue?")
        if not user_response:
            raise ValueError(message)
    try:    
        edited_data['_Area_Code'] = re.search(
            r'(?:\+1\s*)?(\d{3})-(\d{3}-\d{4})', data['Business Phone']).group(1)
        edited_data['_Telephone_Num'] = re.search(
            r'(?:\+1\s*)?(\d{3})-(\d{3}-\d{4})', data['Business Phone']).group(2)
        edited_data['_State'] = zip_mapping(zip_dict, data['State'])
    except Exception:
        messagebox.showerror("Error", "Business Information Error")
        return
    try:
        edited_data['_Email'] = edited_data['_Area_Code'] + \
            edited_data['_Telephone_Num'].replace('-', '') + '@ZBSSOLUTIONS.COM'
        edited_data['_dob_mm'] = data['Date of Birth'].split('-')[1]
        edited_data['_dob_dd'] = data['Date of Birth'].split('-')[2]
        edited_data['_dob_yy'] = data['Date of Birth'].split('-')[0]
        edited_data['_Mobile_Ph_Area'] = re.search(
            r'(?:\+1\s*)?(\d{3})-(\d{3}-\d{4})', data['Mobile']).group(1)
        edited_data['_Mobile_Phone'] = re.search(
            r'(?:\+1\s*)?(\d{3})-(\d{3}-\d{4})', data['Mobile']).group(2)
        edited_data['_State_reside'] = zip_mapping(zip_dict, data['Home State'])
        edited_data['_state_issued'] = zip_mapping(zip_dict, data['State Issued'])
    except Exception:
        messagebox.showerror("Error", "Owner Information Error")
        return
    try:
        rate_dict = get_rate(data['Rate'])
        edited_data['_monthly'] = f"{
        float(rate_dict.get('monthly', 0)) - 10:.2f}"
        match data['Pricing Type']:
            case 'Cash Discount (by Percentage %)':
                edited_data['_percentage'] = f"{
                    (1 - (100 / (100 + float(rate_dict['percentage']))))*100:.2f}"
                edited_data['_price'] = ''
            case 'Cash Discount (by Flat Fee $)':
                edited_data['_price'] = f"${float(rate_dict['price']):.2f}"
                edited_data['_percentage'] = ''
            case 'Flat Rate':
                if float(rate_dict['percentage']) == 0:
                    edited_data['_price'] = f"${float(rate_dict['price']):.2f}"
                    edited_data['_percentage'] = ''
                else:
                    edited_data['_percentage'] = f"{
                        float(rate_dict['percentage']):.2f}"
                    edited_data['_price'] = ''
            case _:
                edited_data['_percentage'] = 'Pricing Type Error'
                edited_data['_price'] = 'Pricing Type Error'
    except Exception:
        messagebox.showerror("Error", "Pricing Error")
        return
    
    current_datetime = datetime.now()
    edited_data['_year'] = f"{current_datetime.year}"
    edited_data['_month'] = f"{current_datetime.month:02d}"
    edited_data['_day'] = f"{current_datetime.day:02d}"
    edited_data['_date'] = edited_data['_month'] + '/' + \
        edited_data['_day'] + '/' + edited_data['_year']    
    return edited_data

# Function to fill the PDF


def filling(edited_data, data, output_path):
    insert_date = {}
    match get_legal_name_suffix(data['Legal Name of Business']):
        case 'LLC':
            insert_date['LLC'] = 'On'
            insert_date['þÿ\x00c\x001\x00_\x000\x001\x00[\x005\x00]'] = '6'
        case 'INC':
            insert_date['Corporation'] = 'On'
            insert_date['þÿ\x00c\x001\x00_\x000\x001\x00[\x002\x00]'] = '3'
        case 'CORP':
            insert_date['Corporation'] = 'On'
            insert_date['þÿ\x00c\x001\x00_\x000\x001\x00[\x002\x00]'] = '3'
        case 'CORPORATION':
            insert_date['Corporation'] = 'On'
            insert_date['þÿ\x00c\x001\x00_\x000\x001\x00[\x002\x00]'] = '3'
        case 'ENTERPRISE':
            insert_date['Corporation'] = 'On'
            insert_date['þÿ\x00c\x001\x00_\x000\x001\x00[\x002\x00]'] = '3'
        case 'LTD':
            insert_date['Corporation'] = 'On'
            insert_date['þÿ\x00c\x001\x00_\x000\x001\x00[\x002\x00]'] = '3'    
        case _:
            insert_date['Sole Proprietor or Single Member LLC'] = 'On'
            insert_date['þÿ\x00c\x001\x00_\x000\x001\x00[\x000\x00]'] = '1'
    insert_date['Corporate or Legal Name'] = data['Legal Name of Business']
    insert_date['Doing Business As'] = data['DBA']
    insert_date['Federal Tax ID Nine Digits'] = data['Tax ID']
    insert_date['Authorized Contact Person'] = data['Owner Name']
    insert_date['Corporate Address'] = data['Street']
    insert_date['City'] = data['City']
    insert_date['Zip'] = data['ZIP']
    insert_date['Bankcard Dep Bank'] = data['Bank Name']
    insert_date['Routing #'] = data['Bank Routing']
    insert_date['Account'] = data['Bank Account']
    insert_date['Name 1'] = data['Owner Name']
    insert_date['SSN'] = data['Social Security Number']
    insert_date['Residential Address'] = data['Home Street']
    insert_date['City_3'] = data['Home City']
    insert_date['Zip_3'] = data['Home ZIP']
    insert_date['Drivers License or State ID No'] = data['Driver License Number']
    insert_date['Text52'] = data['Legal Name of Business'] + \
        '(' + data['Owner Name'] + ')'
    insert_date['MERCHANT NAME'] = data['DBA']
    insert_date['Merchant Name Print'] = data['Legal Name of Business'] + \
        '(' + data['Owner Name'] + ')'
    insert_date['Its'] = data['Legal Name of Business'] + \
        '(' + data['Owner Name'] + ')'
    insert_date['MERCHANT DBA'] = data['DBA']
    insert_date['Text62'] = data['Owner Name']
    insert_date['Area Code'] = edited_data['_Area_Code']
    insert_date['Telephone Number'] = edited_data['_Telephone_Num']
    insert_date['Business Email Address'] = edited_data['_Email']
    insert_date['Email Address'] = edited_data['_Email']
    insert_date['MERCHANT EMAIL'] = edited_data['_Email']
    insert_date['Month'] = edited_data['_dob_mm']
    insert_date['Day'] = edited_data['_dob_dd']
    insert_date['Year'] = edited_data['_dob_yy']
    insert_date['Mobile Ph Area'] = edited_data['_Mobile_Ph_Area']
    insert_date['Mobile Phone'] = edited_data['_Mobile_Phone']
    insert_date['Residence Phone Area Code'] = edited_data['_Mobile_Ph_Area']
    insert_date['Residence Telephone'] = edited_data['_Mobile_Phone']
    insert_date['State'] = edited_data['_State']
    insert_date['State_3'] = edited_data['_State_reside']
    insert_date['State of Issue'] = edited_data['_state_issued']
    insert_date['state_issued'] = edited_data['_state_issued']
    insert_date['Date Month'] = edited_data['_month']
    insert_date['Date Day'] = edited_data['_day']
    insert_date['Date Year'] = edited_data['_year']
    insert_date['checklist month'] = edited_data['_month']
    insert_date['checklist day'] = edited_data['_day']
    insert_date['checklist year'] = edited_data['_year']
    insert_date['Date57_af_date'] = edited_data['_date']
    insert_date['Date61_af_date'] = edited_data['_date']
    insert_date['Date66_af_date'] = edited_data['_date']
    insert_date['Date28_af_date'] = edited_data['_date']
    insert_date['Date29_af_date'] = edited_data['_date']
    insert_date['Text26'] = edited_data['_percentage']
    insert_date['Text35'] = edited_data['_price']
    insert_date['Text40'] = edited_data['_monthly']
    insert_date['þÿ\x00f\x001\x00_\x000\x001\x00_\x000\x00_\x00[\x000\x00]'] = data['Legal Name of Business']
    insert_date['þÿ\x00f\x001\x00_\x000\x002\x00_\x000\x00_\x00[\x000\x00]'] = data['DBA']
    insert_date['þÿ\x00f\x001\x00_\x000\x004\x00_\x000\x00_\x00[\x000\x00]'] = data['Street']
    insert_date['þÿ\x00f\x001\x00_\x000\x005\x00_\x000\x00_\x00[\x000\x00]'] = data['City'] + \
        ',' + edited_data['_State'] + ',' + data['ZIP']
    insert_date['W9date'] = edited_data['_date']
    if len(data['Tax ID'].split('-')[0]) == 2:
        insert_date['þÿ\x00T\x00e\x00x\x00t\x00F\x00i\x00e\x00l\x00d\x002\x00[\x002\x00]'] = data['Tax ID'].split(
            '-')[0]
        insert_date['þÿ\x00T\x00e\x00x\x00t\x00F\x00i\x00e\x00l\x00d\x002\x00[\x003\x00]'] = data['Tax ID'].split(
            '-')[1]
    else:
        insert_date['þÿ\x00T\x00e\x00x\x00t\x00F\x00i\x00e\x00l\x00d\x001\x00[\x000\x00]'] = data['Tax ID'].split(
            '-')[0]
        insert_date['þÿ\x00T\x00e\x00x\x00t\x00F\x00i\x00e\x00l\x00d\x002\x00[\x000\x00]'] = data['Tax ID'].split(
            '-')[1]
        insert_date['þÿ\x00T\x00e\x00x\x00t\x00F\x00i\x00e\x00l\x00d\x002\x00[\x001\x00]'] = data['Tax ID'].split(
            '-')[2]

    fillpdfs.write_fillable_pdf(
        addr[0], f"{output_path + data['DBA'] + '.pdf'}", insert_date)

# GUI function to get file paths and process the data


def process_files():
    global addr, csv_file

    # Select CSV file
    csv_file = filedialog.askopenfilename(
        title="Select CSV File", filetypes=[("CSV files", "*.csv")])
    if not csv_file:
        messagebox.showerror("Error", "CSV file not selected")
        return

    # Select PDF file
    if getattr(sys, 'frozen', False):
        pdf_file = os.path.join(sys._MEIPASS, 'EMS_Merchant_Application.pdf')
    else:
        pdf_file = 'EMS_Merchant_Application.pdf'

    # Set PDF file paths
    addr = [pdf_file, os.path.join(os.path.dirname(
        csv_file), 'EMS_Merchant Application_')]

    # Read data and process it
    try:
        raw_data = read_csv_as_dict(csv_file)[0]
        for key in raw_data:
            # remove space at head & tail
            raw_data[key] = raw_data[key].strip()
            # convert special characters
            raw_data[key] = ftfy.fix_text(raw_data[key])
            raw_data[key] = raw_data[key].replace(
                '\xa0', ' ')  # convert non-breaking space
        pre = prejob(raw_data)
        filling(pre, raw_data, addr[1])
        messagebox.showinfo("Success", "PDF filled and saved successfully!")
        # Open the directory containing the new PDF file
        output_directory = os.path.dirname(addr[1])
        if os.name == 'nt':  # For Windows
            os.startfile(output_directory)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Main function to create the GUI


def main():
    root = tk.Tk()
    root.title("PDF Filler")
    root.geometry("320x180")
    btn_process = tk.Button(
        root, text="Select CSV and Process", command=process_files)
    btn_process.pack(pady=20)
    root.mainloop()


if __name__ == "__main__":
    main()
