import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageFont, ImageDraw
import sys

#ACCESS GOOGLE DRIVE
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
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

#update sent
sheet.update_cell(row,8, "yes")

#how many more times to run script
print(len(name_index)-1)









