import os
from email.message import EmailMessage
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from enum import Enum
import time
from datetime import datetime
import ssl
import smtplib
import json

# Redirect print statements directly to log file (i.e. a different stdio)

import sys
import logging

class LogWriter:
    def __init__(self):
        self.logger = logging.getLogger()
    
    def write(self, message):
        if message.strip():  # Ignore empty lines
            self.logger.info(message.strip())
    
    def flush(self):
        # This method is required for Python 3 compatibility, idk why though
        pass

logging.basicConfig(  # Configure the logging
    filename='app.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
)

sys.stdout = LogWriter() # Redirect print statements to the log file

# --------------------------------------------------------------------


# Set up date info

first_e = 15
second_e = 20
thrid_e = 25
test_e = 18

full_date = datetime.now()

days_to_send = {"First email": first_e, "Second Email": second_e, "Third Email": thrid_e, "Test Email": test_e} # Send emails on the 15th, 20th, and 25th 
day_of_month = full_date.day
current_month_name = full_date.strftime("%B")
print(current_month_name)

def update_json_month(current_month_name: str) -> int:
    jsonData = None
    new_month_num = None
    new_month_name = None
    worksheet_number = None

    with open("months.json", "r") as months_json:
        jsonData = json.load(months_json)

    stored_month_num = jsonData.get("month_num") 
    stored_month_name = jsonData.get("month_name")

    if current_month_name != stored_month_name: # If the current month has changed since the last time we checked, increment month number by 1
        print("MONTHS HAVE CHANGED! It is now {}".format(current_month_name)) 
        new_month_num = stored_month_num + 1
        new_month_name = current_month_name
    elif current_month_name == stored_month_name:
        print("Month status is unchanged: currently {}".format(stored_month_name))

    writeData = {
        "month_name": new_month_name if new_month_name is not None else stored_month_name,
        "month_num": new_month_num if new_month_num is not None else stored_month_num,
        "full_date": ""
    }

    with open("months.json", "w") as months_json:
        json.dump(writeData, months_json, indent=4)

    worksheet_number = new_month_num if new_month_num is not None else stored_month_num # Only set this value if new_month is not None 
    return worksheet_number
    

worksheet_num = update_json_month(current_month_name=current_month_name)  
print("Current worksheet number is... {}".format(worksheet_num))

# --------------------------------------------------------------------


# Authentication with the Google API service account (note: 'scope' defines what we have access to via our account)

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name("C:/Users/maxfi/PythonProjects/BillingReminderService/AutoEmails/credentials.json", scope)
client = None

def connect_and_fetch_data():
    global client
    while True:
        try:
            client = gspread.authorize(creds)
            # Try to open a test spreadsheet
            spreadsheet = client.open("(House) Green Circle Payment Tracking ") # Note: The extra whitespace at the end is needed as it is the exact name of the spreadsheet 
            worksheet = spreadsheet.get_worksheet(worksheet_num)
            
            # Fetch a row to confirm data retrieval
            print("Data successfully retrieved!")
            return worksheet  # Return the worksheet if successful
        
        except (gspread.exceptions.APIError, gspread.exceptions.SpreadsheetNotFound) as e:
            print(f"Error occurred: {e}")
            print("Retrying in 5 seconds...")
            time.sleep(5)  # Wait before retrying

# Connect to Google Sheets and fetch data

worksheet = connect_and_fetch_data()
column_index = None
all_values = worksheet.get_all_values()
print(all_values)

# --------------------------------------------------------------------

# Getting info from our Google Sheet, (where n = name)

'''

Column Associations:

Trash = 3-6
Electric = 8-11
Water = 13-16 
Internet = 18-21
Rent = 23-26

'''
# Column number associated with each bill type and column definition
class BillColumns(Enum):
    TRASH_AMNT_DUE = 3
    TRASH_DATE_DUE = 4
    TRASH_AMNT_PAID = 5
    TRASH_DATE_PAID = 6

    ELECTRIC_AMNT_DUE = 8
    ELECTRIC_DATE_DUE = 9
    ELECTRIC_AMNT_PAID = 10
    ELECTRIC_DATE_PAID = 11

    WATER_AMNT_DUE = 13
    WATER_DATE_DUE = 14
    WATER_AMNT_PAID = 15
    WATER_DATE_PAID = 16

    INTERNET_AMNT_DUE = 18
    INTERNET_DATE_DUE = 19
    INTERNET_AMNT_PAID = 20
    INTERNET_DATE_PAID = 21

    RENT_AMNT_DUE = 23
    RENT_DATE_DUE = 24
    RENT_AMNT_PAID = 25
    RENT_DATE_PAID = 26

bill_cols = {
    "Trash": {
        "Amnt_Due": BillColumns.TRASH_AMNT_DUE.value,
        "Date_Due": BillColumns.TRASH_DATE_DUE.value,
        "Amnt_Paid": BillColumns.TRASH_AMNT_PAID.value,
        "Date_Paid": BillColumns.TRASH_DATE_PAID.value
    },
    "Electric": {
        "Amnt_Due": BillColumns.ELECTRIC_AMNT_DUE.value,
        "Date_Due": BillColumns.ELECTRIC_DATE_DUE.value,
        "Amnt_Paid": BillColumns.ELECTRIC_AMNT_PAID.value,
        "Date_Paid": BillColumns.ELECTRIC_DATE_PAID.value
    },
    "Water": {
        "Amnt_Due": BillColumns.WATER_AMNT_DUE.value,
        "Date_Due": BillColumns.WATER_DATE_DUE.value,
        "Amnt_Paid": BillColumns.WATER_AMNT_PAID.value,
        "Date_Paid": BillColumns.WATER_DATE_PAID.value
    },
    "Internet": {
        "Amnt_Due": BillColumns.INTERNET_AMNT_DUE.value,
        "Date_Due": BillColumns.INTERNET_DATE_DUE.value,
        "Amnt_Paid": BillColumns.INTERNET_AMNT_PAID.value,
        "Date_Paid": BillColumns.INTERNET_DATE_PAID.value
    },
    "Rent": {
        "Amnt_Due": BillColumns.RENT_AMNT_DUE.value,
        "Date_Due": BillColumns.RENT_DATE_DUE.value,
        "Amnt_Paid": BillColumns.RENT_AMNT_PAID.value,
        "Date_Paid": BillColumns.RENT_DATE_PAID.value
    }
}

class PatronRows(Enum): 
    MAX = 3
    MADDIE = 4
    GRAHAM = 5
    NICO = 6
    JAKE = 7 
    ETHAN = 8

# --------------------------------------------------------------------

# Email receive list

email_receive_list = {
    'Maddie': 'madelinedominguez16@gmail.com',
    'Graham': 'CaseyMears2@gmail.com',
    'Nico': 'nicosommer8901@gmail.com',
    'Jake': 'macymoodaboo@gmail.com',
    'Ethan': 'ethanhoofard@gmail.com',
    'Max': 'maxfisherbear@gmail.com', 
}

ssl_context = ssl.create_default_context()

# --------------------------------------------------------------------

# Defining Patron class

def check_string_empty(s: str) -> bool:
    return s == "" or s.isspace() # Retruns true if the string is empty

class Patron:
    def __init__(self, name=None, 
             bill_amounts=None, 
             due_dates=None,
             paid_status=None,
             row_num=None,
             patron_email : str = None):
        
        if bill_amounts is None:
            bill_amounts = {"Water": None, "Electric": None, "Trash": None, "Internet": None, "Rent": None}
        if due_dates is None:
            due_dates = {"Water": None, "Electric": None, "Trash": None, "Internet": None, "Rent": None}
        if paid_status is None:
            paid_status = {"Water": False, "Electric": False, "Trash": False, "Internet": False, "Rent": False}

        self.name = name
        self.bill_amounts = bill_amounts
        self.due_dates = due_dates
        self.row_num = row_num
        self.paid_status = paid_status
        self.all_bills_paid = False
        self.formatted_strings = None
        self.patron_email = patron_email

        self.echo_init() # These functions are called upon intialization
        self.get_gsinfo()
        self.has_paid()
    
    def echo_init(self):
        print("initialized tenet {} for this month's bill...".format(self.name))

    def get_gsinfo(self): # Accepts a class instance of type Patron
        row_values = None
        selected_row = self.row_num
        try:
            row_values = worksheet.row_values(selected_row)
            # Determine the maximum column index needed
            max_column_index = max(
                BillColumns.TRASH_DATE_DUE.value,
                BillColumns.TRASH_AMNT_DUE.value,
                BillColumns.ELECTRIC_DATE_DUE.value,
                BillColumns.ELECTRIC_AMNT_DUE.value,
                BillColumns.WATER_DATE_DUE.value,
                BillColumns.WATER_AMNT_DUE.value,
                BillColumns.INTERNET_DATE_DUE.value,
                BillColumns.INTERNET_AMNT_DUE.value,
                BillColumns.RENT_DATE_DUE.value,
                BillColumns.RENT_AMNT_DUE.value
            )

            # Ensure row_values has enough columns
            if not row_values or len(row_values) < max_column_index:
                print(f"Row {selected_row} does not have enough columns.")
                return

            # Extracting data
            self.due_dates["Water"] = row_values[BillColumns.WATER_DATE_DUE.value - 1]
            self.bill_amounts["Water"] = row_values[BillColumns.WATER_AMNT_DUE.value - 1]

            self.due_dates["Trash"] = row_values[BillColumns.TRASH_DATE_DUE.value - 1]
            self.bill_amounts["Trash"] = row_values[BillColumns.TRASH_AMNT_DUE.value - 1]

            self.due_dates["Electric"] = row_values[BillColumns.ELECTRIC_DATE_DUE.value - 1]
            self.bill_amounts["Electric"] = row_values[BillColumns.ELECTRIC_AMNT_DUE.value - 1]

            self.due_dates["Internet"] = row_values[BillColumns.INTERNET_DATE_DUE.value - 1]
            self.bill_amounts["Internet"] = row_values[BillColumns.INTERNET_AMNT_DUE.value - 1]

            self.due_dates["Rent"] = row_values[BillColumns.RENT_DATE_DUE.value - 1]
            self.bill_amounts["Rent"] = row_values[BillColumns.RENT_AMNT_DUE.value - 1]


        except gspread.exceptions.APIError as e:
            print(f"API Error: {e}")



    def has_paid(self) -> None:
        selected_row = self.row_num
        row_values = worksheet.row_values(selected_row)
        num_bills_paid = 0
        for bill_type in bill_cols:  
            amnt_paid_col = bill_cols[bill_type]["Amnt_Paid"] # Get the amnt_paid column of the specified bill type
            amount_paid = row_values[amnt_paid_col - 1]  # Subtract 1 for 0-based index

            if check_string_empty(amount_paid) == False: # Check if cell is empty, if check_string_empty returns False then string is NOT empty, i.e. the person had paid
                self.paid_status[bill_type] = True # If the cell is NOT empty, we say the person had paid
                num_bills_paid += 1
            else:
                self.paid_status[bill_type] = False
        
        if num_bills_paid == 5:
            self.all_bills_paid = True
        
        
    def make_fstrings(self):
        ammount_due = None
        due_date = None
        unpaid_bills = []
        for bill_type in self.paid_status:
            if self.paid_status[bill_type] == False:  # Identifying unpaid bills
                unpaid_bills.append(bill_type)
        
        self.formatted_strings = []
        for item in unpaid_bills:
            if item == "Water":
                ammount_due = self.bill_amounts["Water"]
                due_date = self.due_dates["Water"]
                self.formatted_strings.append(f"Bill: Water\n Amount Due: {ammount_due}, Due Date: {due_date}")
            elif item == "Trash":
                ammount_due = self.bill_amounts["Trash"]
                due_date = self.due_dates["Trash"]
                self.formatted_strings.append(f"Bill: Trash\n Amount Due: {ammount_due}, Due Date: {due_date}")
            elif item == "Electric":
                ammount_due = self.bill_amounts["Electric"]
                due_date = self.due_dates["Electric"]
                self.formatted_strings.append(f"Bill: Electric\n Amount Due: {ammount_due}, Due Date: {due_date}")
            elif item == "Internet":
                ammount_due = self.bill_amounts["Internet"]
                due_date = self.due_dates["Internet"]
                self.formatted_strings.append(f"Bill: Internet\n Amount Due: {ammount_due}, Due Date: {due_date}")
            elif item == "Rent":
                ammount_due = self.bill_amounts["Rent"]
                due_date = self.due_dates["Rent"]
                self.formatted_strings.append(f"Bill: Rent\n Amount Due: {ammount_due}, Due Date: {due_date}")
            else:
                print(f"No bills due for tenet: {self.name}")
    

n_maddie = Patron(name="Maddie", row_num=4, patron_email=email_receive_list["Maddie"])
n_max = Patron(name="Max", row_num=3, patron_email=email_receive_list["Max"])
n_ethan = Patron(name="Ethan", row_num=8, patron_email=email_receive_list["Ethan"])
n_graham = Patron(name="Graham", row_num=5, patron_email=email_receive_list["Graham"])
n_jake = Patron(name="Jake", row_num=7, patron_email=email_receive_list["Jake"])
n_nico = Patron(name="Nico", row_num=6, patron_email=email_receive_list["Nico"])

patron_structs = []
patron_structs = [n_maddie, n_max, n_ethan, n_graham, n_nico, n_jake]

# --------------------------------------------------------------------

# Email logic -> Python App-Password will be stored in my environemnt variables

def get_email_body(tenet_struct: Patron) -> str:
    f_strings = tenet_struct.formatted_strings
    if tenet_struct.all_bills_paid == True:
        print("Tenet {} has paid all of their bills".format(tenet_struct.name))
    elif tenet_struct.all_bills_paid == False:  
        body_txt = "\n\n".join(f_strings) if f_strings else "No bills to display."
        body = '''
        |Reminder| 

The following bill(s) are due --->

{body_txt}

From,
Max Emrich (AKA TwizMaster100, AKA White-Lightning, AKA GreenCircleGuy)

'''.format(body_txt=body_txt)
        print(body)
        return body
    return None

email_sender = "maxrivers777@gmail.com"
email_password = os.environ.get("PYTHON_APP_EMAIL_PASSWORD")

def send_email(tenent_struct: Patron, e_password: str, ssl_context, email_body):
    em = EmailMessage()
    email_subject = 'Green Circle Bill Reminder'
    em['From'] = email_sender
    em['To'] = tenent_struct.patron_email
    em['Subject'] = email_subject
    e_receiver = tenent_struct.patron_email
    em.set_content(email_body)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ssl_context) as smtp:
        smtp.login(email_sender, e_password)
        smtp.sendmail(email_sender, e_receiver, em.as_string()) 
        print("Email sent to: {}".format(tenent_struct.name))

# Check the current date function

def check_date(current_day: str) -> bool: # Returns True if the current day matches one of the selected days to send out an email, False otherwise
    if type(current_day) != str:
        print(ValueError)

    for day in days_to_send.values():
        if current_day == str(day):
            return True

    return False

# --------------------------------------------------------------------

n_max.make_fstrings() # Make formatted strings
email_body = get_email_body(n_max) # Make an email body from those strings 
send_email(n_max, email_password, ssl_context=ssl_context, email_body=email_body) # Creat an email object and send it to tenets

# Check date to see we if we need to send an email today, then send it

if check_date(str(day_of_month)) == True: 
    for patron in patron_structs:
        f_strings = patron.make_fstrings() # Make formatted strings
        email_body = get_email_body(patron) # Make an email body from those strings 
        send_email(patron, email_password, ssl_context=ssl_context, email_body=email_body) # Creat an email object and send it to tenets

# --------------------------------------------------------------------