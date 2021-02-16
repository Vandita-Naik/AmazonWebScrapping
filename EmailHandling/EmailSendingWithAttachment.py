import smtplib
from email.message import EmailMessage
msg=EmailMessage()
msg['Subject']="Training Invitaion"
msg['From']="vandunaik04@gmail.com"
msg['To']="vandunaik04@gmail.com"
msg.set_content("This is for practice purpose")

with open("EmailText.txt") as myfile:
    data=myfile.read()
    msg.set_content(data)

with smtplib.SMTP_SSL("smtp.gmail.com",465) as server:# same for all
    server.login("vandunaik04@gmail.com","Vandita@1104")
    server.send_message(msg)

print("Email has been sent")

