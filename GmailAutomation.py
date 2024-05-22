from email.message import EmailMessage
import ssl
import smtplib
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime


scope=["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]

# cred=ServiceAccountCredentials.from_json_keyfile_name("D:\\college\\MyCodeDiary\\Project\\Gmail_Automation\\Data\\key.json",scopes=scope)
cred=ServiceAccountCredentials.from_json_keyfile_name("Data\\key.json",scopes=scope)

# created the credentials to authorize later

# authorizing
file=gspread.authorize(cred)
workbook=file.open("GmailAutomation")
sheet=workbook.sheet1


for cell in sheet.range("A2:A4"):
    print(cell.value)






sender="ashishbargoti2003@gmail.com"
password="mzpdpdbdkqgxnowz"
receiver="ashishbargoti123@gmail.com"

subject="Gmail Automation Successful!"
body="Hey there gmail automation was successful"

email_object=EmailMessage()

# creating an object of EmailMessage that we will further use to perform actions
email_object["From"]=sender
email_object["To"]=receiver
email_object["Subject"]=subject
email_object.set_content(body)



with open("inbox.txt") as inbox:
    message=inbox.read().split("{}")
    # print(message)
    # print(("ashish").join(message))
    # print(type(message))

def edit_name(message,name):
    return (name).join(message)

all_values=sheet.get_all_values()
# all_values ia list of lists where every inner list is a list of columns 
# [['name', 'email', 'status'], ['ashish', '', ''], ['vinay', '', ''], ['akash', '', '']]
email_index=all_values[0].index("email")
no_of_records=len(all_values)
name_index=all_values[0].index("name")


# print(all_values)



# ssl is a standard security technique used to secure content between any two systems
# here it will an add a layer of security

context=ssl.create_default_context()

# (server,port_number,context)
with smtplib.SMTP_SSL("smtp.gmail.com",465,context=context) as smtp:
    smtp.login(sender,password)
    for i in range(1,no_of_records):
        receiver=all_values[i][email_index]
        company=all_values[i][name_index]

        body=edit_name(message,company)

        # email_object["From"]=sender
        # email_object["To"]=receiver
        # email_object["Subject"]=subject
        email_object.set_content(body)
        # here editing the company name
        

        # defining the elements of this object below:
        

        # smtp.sendmail(sender,receiver,email_object.as_string())
        # now=datetime.now()
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.update_cell(i+1,4,now)
        # update cell (row,col,value)
        print(now)

        smtp.sendmail(sender,receiver,email_object.as_string())
    print("sent successfully!")
