import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
# from email import MIMEImage
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
# from email import Encoders
import datetime


now = datetime.datetime.now()
last_month = now.month-1 if now.month > 1 else 12
if now.month == 1:
    last_year=now.year-1
else:
    last_year=now.year



last="January February March April May June July August September Octobor November December".split()[last_month-1]





def mail(to,attach):
    # allow either one recipient as string, or multiple as list
    if not isinstance(to,list):
        to = [to]
    # allow either one attachment as string, or multiple as list
    if not isinstance(attach,list):
        attach = [attach]

    gmail_user='minshad.it@aristopharma.com'
    gmail_pwd = "minshad.it#978#"
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = ", ".join(to)
    msg['Subject'] = "DCC Repot "+last+" "+last_year
    body_text="Assalamualykum  ,\nPlease check the attached file of DCC Report of "+last+" "+last_year+"\n\nSincerely,\n Minshad"
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
1310:'dhk.it@aristopharma.com',
1311:'dnm.it@aristopharma.com',
1312:'UTR@aristopharma.com',
1313:'ngn@aristopharma.com',
1314:'MYM@aristopharma.com',
1315:'CHT@aristopharma.com',
1316:'SYL@aristopharma.com',
1317:'CML@aristopharma.com',
1318:'NKL@aristopharma.com',
1319:'fen@aristopharma.com',
1320:'BGR@aristopharma.com',
1321:'RJS@aristopharma.com',
1322:'RNG@aristopharma.com',
1323:'DNJ@aristopharma.com',
1324:'JSR@aristopharma.com',
1325:'KHL@aristopharma.com',
1326:'BRS@aristopharma.com',
1327:'FRD@aristopharma.com',
1328:'CXB@aristopharma.com',
1329:'TNG@aristopharma.com',
1330:'PAB@aristopharma.com',
1331:'CHP@aristopharma.com',
1332:'BRB@aristopharma.com',
1333:'MLB@aristopharma.com',
1334:'KIG@aristopharma.com',
1335:'KST@aristopharma.com',
1336:'KNJ@aristopharma.com',
}



dir = input("Enter attachment Directory : ")
os.chdir(dir)
a = os.listdir()
arr=[]

for x in dicm.keys():
    for y in a:
        print(x)
        name,e = y.split('.')
        # print(name[:4])
        if x == int(name[:4]):
            arr.append(y)
    if arr:
        mail(dicm[x],arr)
        print(dicm[x]+"-Ok")
        arr=[]
    else:
        print(str(x)+"-This depot has no file in Directory,Please Check")








