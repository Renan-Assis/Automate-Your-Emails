import pandas as pd
from datetime import datetime, timedelta, date
import openpyxl
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

day = date.today() # This is a Global Variable to get the date from TODAY's date everytime the code starts
smtp_port = 587 # Default GMAIL SMTP Port
smtp_server = "smtp.gmail.com" # Default GMAIL SMTP Server
simple_email_context = ssl.create_default_context()
email_from = "rcutrimassis@gmail.com" #<<--- Should be Your E-mail or your Company's
email_to = ["rcutrimassis@gmail.com","renanassis.dev@gmail.com"] #<<--- This should be a list of people you want to send your e-mail to, i put 2 of my emails to test
pswd = 'ucjfvtzzvznegamx' #<<--- This should be your personal E-Mail Password you get from opening permissions to other apps

def read_file():
    file = pd.read_excel('Absenteism Report.xlsx')
    return file

# This function is disabled, you should only use it if you want to get registers from a specific date.
# If you want to use it, you should disable the global variable day on line 11.
# def get_data():
#     day = input("Would you like to see registers from a specific date?")
#     day = turn_to_datetime(day) (These two functions are nested because one is directly used with the other)
#       return day

def turn_to_datetime(data):
    
    format_string = "%d/%m/%Y"
    data_aux = datetime.strptime(data, format_string)
    return data_aux

# This function just makes a quick formatting to a date so it looks more appealing
# Instead of 1/1/2023 it'll be 01/01/2023, for example
def data_config(day):
    if day.day<10 and day.month<10:
        date = "0{}/0{}/{}".format(day.day,day.month,day.year)
    elif day.day>=10 and day.month<10:
        date = "{}/0{}/{}".format(day.day,day.month,day.year)
    else:
        date = "{}/{}/{}".format(day.day,day.month,day.month)
    return date

def get_workers(df, day, today):

    workers = df.loc[df['DATE'] == '{}'.format(day)] # This line gets only the name of the workers who didn't go to work that day
    absent_workers = [] 
    df_workers = workers
    for i in df_workers['NAME']:
        absent_workers.append(df['NAME'].value_counts()[i])
        #Checking how many times on that month these Workers didn't come to labor
    df_workers['THIS MONTH'] = absent_workers
    df_workers['DATE'] = today
    
    return df_workers

def isweekend(data):
    data = turn_to_datetime(data)
    weekday = data.weekday()
    if weekday <5:
        return False
    return True
#Checking if the current day is a Weekend Day, so the e-mail won't be sent to addressees.

def money(df_column):
    return '$ {:,.2f}'.format(df_column)
# Formatting money values that usually comes as "General" in Excel files as a "Money format"

def create_csv(df):    
    
    writer = pd.ExcelWriter('ABSENT TEAMS.xlsx',engine='openpyxl')
    df.to_excel(writer,sheet_name='Sheet1',index=False)
    #This creates an Excel file with the name ABSENT TEAMS - DAY: and places the date gathered from the input or the global variable on the {}
      
    worksheet = writer.sheets['Sheet1']
    # This is a quick editting on the CSV File with some colors,size and bold text so it doesn't look so Raw
    #Setting columns width so the text can fit in.
    worksheet.column_dimensions["A"].width = 4.29
    worksheet.column_dimensions["B"].width = 10
    worksheet.column_dimensions["C"].width = 36
    worksheet.column_dimensions["D"].width = 13.30
    worksheet.column_dimensions["E"].width = 16.70
    worksheet.column_dimensions["F"].width = 40
    worksheet.column_dimensions["G"].width = 14.45
    worksheet.column_dimensions["H"].width = 20
    worksheet.column_dimensions["I"].width = 14.45
    worksheet.column_dimensions["J"].width = 12.60
    worksheet.column_dimensions["K"].width = 11
    worksheet.column_dimensions["L"].width = 12 
    worksheet.column_dimensions["M"].width = 14

    font = openpyxl.styles.Font(color=openpyxl.styles.colors.WHITE,bold=True) #Title Characteristics
    fill = openpyxl.styles.PatternFill(patternType='solid', fgColor=openpyxl.styles.colors.BLUE)

    for cell in worksheet[1]:
        cell.font = font
        cell.fill = fill

    writer.close()

def send_email(email_to,day):
    if isweekend(day) == False:
        message ='''
            UNAVAILABLE WORKING TEAMS ON DAY {}'''.format(day)
            # This is the TITLE of your E-mail.
        
        for person in email_to:
            #This loop makes you send this e-mail message and attachment file to every person you adressed the e-mail in the list on line 16.
            
            body = '''
            Good Morning, Team!
            These are the following teams that couldn't come to work on {}.
            The report follows as an attachment, feel free to contact me if you have any questions.

            Best Regards, Renan Cutrim.'''.format(day)
            #This is the CONTENT of your E-mail.

            msg = MIMEMultipart()
            msg['From'] = email_from
            msg['To'] = person
            msg['Subject'] = message

            msg.attach(MIMEText(body,'plain'))

            filename = 'ABSENT TEAMS.xlsx'
            
            attachment = open(filename, 'rb')

            attachment_package = MIMEBase('application','octet-stream')
            attachment_package.set_payload(attachment.read())
            encoders.encode_base64(attachment_package)
            attachment_package.add_header('Content-Disposition',"Attachment; filename={}".format(filename))
            msg.attach(attachment_package)

            text = msg.as_string()
            print("Connecting to server...")
            TIE_Server = smtplib.SMTP(smtp_server, smtp_port)
            TIE_Server.starttls(context=simple_email_context)
            TIE_Server.login(email_from,pswd)
            print("Conected")

            print(f"Sending mail to -{person}")
            TIE_Server.sendmail(email_from,person,text)
            print(f"Email Sent - {person}")

            TIE_Server.quit()
        else:
            pass

def main():
    
    #day = get_data()
    date = data_config(day)
    df = read_file()
    teams = get_workers(df,day,date)
    create_csv(teams)
    send_email(email_to,date)

if __name__ == "__main__":
    main()