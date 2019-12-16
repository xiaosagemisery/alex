from email.mime.text import MIMEText
import smtplib
from email.mime.multipart import MIMEMultipart
from email.header import Header


def office365(USERID,USERNAME,USERPHONE,SITE):
    mailserver = smtplib.SMTP('mail.medtronic.com')
    EXTENSION = USERPHONE[-5:]
    #mailserver.ehlo()
    #mailserver.starttls()
    #mailserver.login(USERNAME, PASSWORD)
    mail_msg = """
    <p>Hi """ + USERNAME + """(""" + USERID + """)""" + """</p>
    <p>I've assigned  """ + EXTENSION + """ IP phone extension to you. The whole number is """ + USERPHONE + """</p>
    <div>
    <div>
    <div>
    <p><b><span lang="en-US" style="color:#1F497D;font-size:10pt;font-family:Arial,sans-serif;">&nbsp;</span></b><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:#1F497D;font-size:10pt;font-family:Arial,sans-serif;">.&nbsp;</span><span lang="en-AU" style="color:black;font-size:10pt;font-family:Arial,sans-serif;">The instruction is:</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-family:Calibri,sans-serif;">&nbsp;</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-size:10pt;font-family:Arial,sans-serif;">1) press "Services" button and select "Extension Mobility"</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-size:10pt;font-family:Arial,sans-serif;">2) put your userID and PIN.&nbsp; UserID is your normal Windows userID.&nbsp; The PIN is </span><b><span lang="en-AU" style="color:blue;font-size:10pt;font-family:Arial,sans-serif;">12345</span></b><span lang="en-AU" style="color:black;font-size:10pt;font-family:Arial,sans-serif;">.</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-size:10pt;font-family:Arial,sans-serif;">3) press "submit" button, after that you will login to the phone. </span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-size:10pt;font-family:Arial,sans-serif;">4) Once login, you will see your name displayed on the phone screen.&nbsp;</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-family:Calibri,sans-serif;">&nbsp;</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-size:10pt;font-family:Arial,sans-serif;">I've attached a user manual here. </span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-size:10pt;font-family:Arial,sans-serif;">&nbsp;</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-US" style="color:black;font-size:11pt;font-family:等线;">&nbsp;</span><span lang="en-US" style="color:black;"></span></p>
    <p><b><span lang="en-AU" style="color:black;">You should have desk phone on your table.</span></b><span lang="en-US" style="color:black;"></span></p>
    <p><b><span lang="en-AU" style="color:black;">If you have no phone device, please Contact with </span></b><b><span lang="en-AU">Facility(20325987).</span></b><span lang="en-US" style="color:black;"></span></p>
    <p><b><span lang="en-AU" style="color:black;">If you don</span><span style="color:black;">’</span></b><b><span lang="en-AU" style="color:black;">t receive Phone Model:8851,please send me Email.</span></b><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;">&nbsp;</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-size:10pt;font-family:Arial,sans-serif;">Please let us know if you have any questions.</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-size:11pt;font-family:等线;">&nbsp;</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-AU" style="color:black;font-family:Calibri,sans-serif;">&nbsp;</span><span lang="en-US" style="color:black;"></span></p>
    <p><span lang="en-US" style="color:black;font-family:Calibri,sans-serif;">&nbsp;</span><span lang="en-US" style="color:black;"></span></p>
    <p><span style="color:#CCCCCC;">此信息已标记为“美敦力受控”</span><span lang="en-US" style="color:black;"></span></p>
    </div>
    </div>
    </div>
    """

    # 创建一个带附件的实例
    message = MIMEMultipart()
    message.attach(MIMEText(mail_msg, 'html', 'utf-8'))
    subject = 'MDT电话机分机分配'
    acc = ''
    ccp = ''
    ccp2 = ''
    message['Subject'] = Header(subject, 'utf-8')
    if SITE == 'SH':
        acc = 'dora.wu@medtronic.com'
        ccp = 'wut7'
    elif SITE == 'BJ':
        acc = 'sina.li@medtronic.com, yuki.zhao@medtronic.com'
        ccp = 'lix74'
        ccp2 = 'zhaoy49'
    message['Cc'] = acc
    att1 = MIMEText(open('如何登陆IP电话.pptx', 'rb').read(), 'base64', 'utf-8')
    att1["Content-Type"] = 'application/octet-stream'
    att1.add_header('Content-Disposition','attachment', filename="如何登陆IP电话.pptx")
    message.attach(att1)

    att2 = MIMEText(open('8851 使用指南.pptx', 'rb').read(), 'base64', 'utf-8')
    att2["Content-Type"] = 'application/octet-stream'
    att2.add_header('Content-Disposition','attachment', filename="8851 使用指南.pptx")
    message.attach(att2)
    mailserver.sendmail('alex.shen@medtronic.com', [USERID, ccp, ccp2], message.as_string())
    print('邮件发送成功')
    mailserver.quit()


if __name__ == "__main__":
    office365('maj21', 'Luffy Ma', '862199999999', 'SH')
