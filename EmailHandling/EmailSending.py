# import smtplib
# from email.message import EmailMessage
# msg=EmailMessage()
# msg['Subject']="Training Invitaion"
# msg['From']="vandunaik04@gmail.com"
# msg['To']="vandunaik04@gmail.com"
# msg.set_content("This is for practice purpose")
# server=smtplib.SMTP_SSL("smtp.gmail.com",465) # same for all
# server.login("vandunaik04@gmail.com","Vandita@1104")
# server.send_message(msg)
# # server.sendmail("vandunaik04@gmail.com", "vandunaik04@gmail.com", "hi how are u")
# server.quit()

import smtplib
server=smtplib.SMTP_SSL("smtp.gmail.com",465) # same for all
server.login("vandunaik04@gmail.com","Vandita@1104")
server.sendmail("vandunaik04@gmail.com", "vandunaik04@gmail.com", "hi how are u")
server.quit()