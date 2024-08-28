# BillingReminderService
This is an automatic emailing service I built for me and my roommates. This app updates us a few times a month on unpaid bills.


## Structure
After giving the program access to a Google API Service account for Google Spreadsheets, I can access a shared spreadsheet using the *gspread* library for Python.

The data is retrieved from the spreadsheet and given to a variable that contains all the rows of the data that can be accessed.

I set a **`Patron`** (AKA anyone who is paying bills) class that reads, parses, and stratifies the data from the spreadsheet based on the tenet's name. The classes then act on that data and define a few structures that hold data about their bill statuses, i.e., their paid status, their due dates, bill amounts, etc.

A simple email body is then constructed based on the **`Patron's`** paid status for each bill.

## Scheduling 
This script is now in my Windows Task Scheduler, where I have the trigger(s) set to be every day at 8:00am. At this time every day, the program will run, check the month and day against the json file I have set up to count the months, and if the conditions are met (the day and month are correct) to send the emails, the emails will be sent to each tenent based on their bill statuses.

## Notes
There is a lot of verbose code in this file, especially in the "Extracting Data" section of the get_gsinfo() method. This is because I was having trouble getting *gspread* to read the name of the index that was accessing the row data in a for-loop, so I opted to make the structure more clear for debugging purposes. 