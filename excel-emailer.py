import openpyxl # dl pip3 install openpyxl==2.6.2 so that code in Automate the Boring Stuff works
from random import choice
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def read_excel():
    wb = openpyxl.load_workbook('testexcelwb.xlsx')
    sheet = wb['Sheet1']
    html_options = [sheet['I2'].value, sheet['I3'].value]
    #print(html_options)

    recipients = []
    companies = []
    for row in range(2, sheet.max_row + 1): #sheet.max_row
        fn = sheet['A' + str(row)].value
        ln = sheet['B' + str(row)].value
        company = sheet['C' + str(row)].value
        email = sheet['D' + str(row)].value

        company_info = {}
        company_info['company'] = company
        company_info['first_name'] = fn
        company_info['last_name'] = ln
        company_info['email'] = email
        if company not in companies:
            companies.append(company)
            company_info['html'] = html_options[0]
        else:
            company_info['html'] = html_options[1]
        recipients.append(company_info)

    return recipients

def send_email(recipients):
    from_email = os.environ.get('EMAIL_USER')
    password = os.environ.get('EMAIL_PASS')
    port = 587
    server = 'smtp.gmail.com'
    message = MIMEMultipart("alternative")
    message["From"] = from_email

    for recipient in recipients:
        to_email = recipient['email']
        first_name = recipient['first_name']
        last_name = recipient['last_name']
        company = recipient['company']
        html = recipient['html']

        message["To"] = to_email
        message["Subject"] = f"An Email for {first_name}"
        html = html.format(last_name=last_name, first_name=first_name, company=company, to_email=to_email)
        text = html.format(last_name=last_name, first_name=first_name, company=company, to_email=to_email)

        # Turn these into plain/html MIMEText objects
        part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        message.attach(part1)
        message.attach(part2)
        print(f"Preparing to send email to {first_name}")
        try:
            smtpObj = smtplib.SMTP(server, port)
            smtpObj.ehlo()
            smtpObj.starttls()
            smtpObj.login(from_email, password)
            print("Sending Email...")
            smtpObj.sendmail(from_email, to_email, message.as_string()) #must send msg as.string() when using HTML and plaintext options
            print('Email Sent!')
        except Exception as e:
            print(e)
        finally:
            #must include these del lines so that email "headers" go back to empty. Otherwise they won't update in the loop...
            del message["Subject"]
            del message["To"] 
            smtpObj.quit()
    

recipients = read_excel()
send_email(recipients)

    
