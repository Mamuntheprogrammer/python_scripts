import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# from email import MIMEImage
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
# from email import Encoders



def mail(to,attach):
    # allow either one recipient as string, or multiple as list
    if not isinstance(to,list):
        to = [to]
    # allow either one attachment as string, or multiple as list
    if not isinstance(attach,list):
        attach = [attach]

    gmail_user='*****@aristopharma.com'
    gmail_pwd = "*********"
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = ", ".join(to)
    msg['Subject'] = "Hi mozur vai ! any attachment there "
    body_text="Mail automation for target sheets"
    msg.attach(MIMEText(body_text,'plain'))

    #get all the attachments
    for file in attach:
        with open(file, "rb") as f:
            #attach = email.mime.application.MIMEApplication(f.read(),_subtype="pdf")
            part = MIMEApplication(f.read(),_subtype="pdf")

        # part = MIMEBase('application', 'octet-stream')
        # part.set_payload(open(file, 'rb').read())
        # Encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename= %s' % file)
            msg.attach(part)

    mailServer = smtplib.SMTP("smtp.aristopharma.com", 587)
    # mailServer.ehlo()
    mailServer.starttls()
    # mailServer.ehlo()
    mailServer.login(gmail_user, gmail_pwd)
    text = msg.as_string().encode('ascii', 'ignore')

    mailServer.sendmail(gmail_user, to, text)
    # Should be mailServer.quit(), but that crashes...
    mailServer.close()




# dicm={1310:'monzur.it@aristopharma.com',1311:'almamun.it@aristopharma.com'}

# # M=['monzur.it@aristopharma.com','almamun.it@aristopharma.com']
dicm={
1310:'***@aristopharma.com',

}



dir = input("Enter attachment Directory : ")
os.chdir(dir)
a = os.listdir()
arr=[]

for x in dicm.keys():
    for y in a:
        print(x)
        name,e = y.split('.')
        print(name[:4])
        if x == int(name[:4]):
            arr.append(y)
    if arr:
        mail(dicm[x],arr)
        arr=[]
    else:
        print(str(x)+"-This depot has no file in Directory")


    




# for x in dicm:

#     a = os.listdir()
    


#     mail(M,arr)









