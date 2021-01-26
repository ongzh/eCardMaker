import gspread, sys
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageFont, ImageDraw
import smtplib, ssl , getpass, email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

#STEPS: Get Data -> Make Card -> Send Card -> Update Data

#ACCESS GOOGLE DRIVE
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive","https://www.googleapis.com/auth/gmail.send"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("MTSA eCard registration - Form Responses").sheet1

name_index = []
paid_col = sheet.col_values(7)

for rnum in range(1,len(paid_col)):
    if paid_col[rnum]=="paid" and sheet.cell(rnum+1,8).value =="":
        name_index.append(rnum+1)

if name_index ==[]:
    print ("no new purchases")
    sys.exit()
else:
    row = name_index[0]




#make new card
name = sheet.cell(row,2).value +" " + sheet.cell(row,3).value
card = Image.open("new_card.jpg")
name_font = ImageFont.truetype("Walkway_Bold.ttf",144)
edit_card = ImageDraw.Draw(card)
W, H = name_font.getsize(name)
edit_card.text(((2363-W)/2,1330), name, (255,255,255), font=name_font)
card.save("MTSA e-Card (Front).pdf")

# Email section
port = 465 #for SSL
context = ssl.create_default_context() #SSL context

password = getpass.getpass()
receivermail = sheet.cell(row,4).value 
sendermail = "mtsa.team@gmail.com"

if receivermail == "":
    print("Email not submitted")
    sys.exit()

#Read template email and create message
file = open("mailtemplate.txt", "r")
template = file.read()
file.close()
body = "Dear"+" "+ name + ",\n" + template

# Create a multipart message and set headers
message = MIMEMultipart()
message["From"] = sendermail
message["To"] = receivermail
message["Subject"] = "MTSA eCard"

# Add body to email
message.attach(MIMEText(body, "plain"))

#attach Cards to email
files = ["MTSA e-Card (Front).pdf", "MTSA e-Card (Back).pdf"]
for file in files:
    attachment = open(file, 'rb')
    file_name = file
    part = MIMEBase('application','octet-stream')
    part.set_payload(attachment.read())
    part.add_header('Content-Disposition',
                    'attachment',
                    filename=file_name)
    encoders.encode_base64(part)
    message.attach(part)

#convert message to string
message = message.as_string()


#send the mail
with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
    server.login(sendermail, password)
    server.sendmail(sendermail, receivermail, message)
    server.quit()

#update data
sheet.update_cell(row,8, "yes")

#how many unsent cards left. 
print(len(name_index)-1)
    









