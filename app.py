import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) 

while True: 
    try: 
        directory = single_input("Please input the directory where the file is located.")+"\\"
        filename = os.listdir(directory)
        if ( 'ProductData.xlsx' in filename) & ('MasterInvoice.pdf' in filename) & ( 'Order1.xlsx'in filename) & ( 'Order2.xlsx' in filename) & ('Order3.xlsx' in filename) & ('Order4.xlsx' in filename):
            break
    except:
        show_message("Error","Please re-enter Directory.")
filename 

product_descriptions = {}
product_prices = {} 
descriptions = pd.read_excel(directory+'ProductData.xlsx', sheet_name = 'Descriptions')
prices = pd.read_excel(directory+'ProductData.xlsx', sheet_name = 'Prices')
for r, row in enumerate(descriptions['Unit number']):   
    product_descriptions[row] = descriptions['Description'][r]
for r, row in enumerate(prices['Unit number']):
    product_prices[row] = float(prices['Price per unit'][r])
product_prices
product_prices['U85392']

df = get_pdf(directory+'MasterInvoice.pdf')
df2 = pd.read_excel('Order2.xlsx')
def get_order_data(invoice_num, filename, directory):
    file = pd.read_excel(directory+filename)
    PDFdictionary = {}
    PDFdictionary['InvoiceNum'] = invoice_num
    for r, row in enumerate(file['Unnamed: 1']):
        if row == "Order Date": 
            date = file['Unnamed: 2'][r].date()
            day = date.day
            month = date.month  
            year = date.year
            PDFdictionary['Date'] = f'{month}/{day}/{year}'
            date_output = f'{month}/{day}/{year}'
        elif row == "Email":
            emailaddress = file['Unnamed: 2'][r]
        elif row == "Name":
            custname = file["Unnamed: 2"][r]
            PDFdictionary['CustomerName'] = file["Unnamed: 2"][r]
        elif row == "Company Name": 
            PDFdictionary['CustomerCompany'] = file["Unnamed: 2"][r]
        elif row == "Address":
            PDFdictionary['CustomerAddress1'] = f'{file["Unnamed: 2"][14]}{file["Unnamed: 2"][15]} {file["Unnamed: 2"][16]} {file["Unnamed: 2"][17]}'
        elif row == "Unit Number": 
            PDFdictionary['Description0'] = file["Unnamed: 1"][22]
            PDFdictionary['Quantity0'] = file["Unnamed: 2"][22]
            PDFdictionary['Price0'] = product_prices[file["Unnamed: 1"][22]] 
            num0 = PDFdictionary["Quantity0"]*PDFdictionary["Price0"]
            totals = num0
            numstr = num0
            PDFdictionary['Amount0'] = numstr
            num1 = 0 
            num2 = 0 
            num3 = 0 
            PDFdictionary['Price0'] = f'${PDFdictionary['Price0']:,.2f}'
            try: 
                PDFdictionary['Description1'] = file["Unnamed: 1"][23]
                PDFdictionary['Quantity1'] = file["Unnamed: 2"][23]
                PDFdictionary['Price1'] = product_prices[file["Unnamed: 1"][23]] 
                num1 = PDFdictionary["Quantity1"]*PDFdictionary["Price1"]
                PDFdictionary['Price1'] = f'${PDFdictionary['Price1']:,.2f}'
                totals = num0 + num1 + num2 + num3
                numstr = num1
                PDFdictionary['Amount1'] = numstr
                PDFdictionary['Description2'] = file["Unnamed: 1"][24]
                PDFdictionary['Quantity2'] = file["Unnamed: 2"][24]
                PDFdictionary['Price2'] = product_prices[file["Unnamed: 1"][24]] 
                num2 = PDFdictionary["Quantity2"]*PDFdictionary["Price2"]
                PDFdictionary['Price2'] = f'${PDFdictionary['Price2']:,.2f}'
                totals = num0 + num1 + num2
                numstr = num2
                PDFdictionary['Amount2'] = (numstr)
                PDFdictionary['Description3'] = file["Unnamed: 1"][25]
                PDFdictionary['Quantity3'] = file["Unnamed: 2"][25]
                PDFdictionary['Price3'] = product_prices[file["Unnamed: 1"][25]] 
                num3 = PDFdictionary["Quantity3"]*PDFdictionary["Price3"]
                PDFdictionary['Price3'] = f'${PDFdictionary['Price3']:,.2f}'
                totals = num0 + num1 + num2 + num3
                numstr = num3
                PDFdictionary['Amount3'] = numstr
            except: 
                num = totals
                numstr = num
                PDFdictionary['Total'] = numstr
                break 
    return PDFdictionary, custname, emailaddress, date_output
PDFdictionary, custname, emailaddress, date_output = get_order_data(8515,'Order1.xlsx', directory)
print(PDFdictionary)

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+ 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf',order_data)
create_invoice(invoice_num =8515,order_data = PDFdictionary,directory = directory)    

def send_email(invoice_num,name,email,date,directory):
    gmail_user = f'{email}'
    gmail_password = open('gmail_password.txt').read()
    recipients = email 
    subject = f'Your Order from PythonTech Co. - Invoice{invoice_num}'
    body = f"""Dear {name},
    
    Attached is your invoice dated {date}.
    
    Thank you for your business. 

    PythonTech Co.
    10 Python Drive 
    Pleasent Grove, UT 84062 
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user 
    message['To'] = recipients 
    message['Subject'] = subject 
    message.set_content(body)

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()
    return
send_email(PDFdictionary['InvoiceNum'], custname, emailaddress,date_output,directory)

ilenames = filename
filenum = 8515
for item in filenames:
    if ('Order' in item) and ("xlsx" in item):
        PDFdictionary, custname,emailaddress, date_output = get_order_data(filenum, item, directory)
        create_invoice(filenum,PDFdictionary, directory)
        send_email(filenum,custname,emailaddress,date_output, directory)
        filenum = filenum + 1

filenames 