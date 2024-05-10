import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
import configparser

    
# SMTP 서버를 dictionary로 정의
smtp_info = {
    'hiworks': ('smtps.hiworks.com', 465),
    'naver.com': ('smtp.naver.com', 587)
}

# 메일 보내는 함수 정의
def send_mail(From, To, subject, message, attach_files=(), pw=''):
    smtp_server, port = smtp_info.get(From.split('@')[-1])

    # 메일 객체 생성
    msg = MIMEMultipart()
    msg['From'] = From
    msg['To'] = ', '.join(To)
    msg['Subject'] = Header(subject, 'utf-8')

    # 본문 추가
    msg.attach(MIMEText(message, 'plain', 'utf-8'))

    # 첨부 파일 추가
    for file_path in attach_files:
        file_name = os.path.basename(file_path)
        attachment = MIMEApplication(open(file_path, 'rb').read())
        attachment.add_header('Content-Disposition', 'attachment', filename=(Header(file_name, 'utf-8').encode()))
        msg.attach(attachment)

    # SMTP 서버 연결 및 메일 전송
    try:
        smtp = smtplib.SMTP(smtp_server, port)
        smtp.starttls()  # TLS 보안 연결
        smtp.login(From, pw)
        smtp.sendmail(From, To, msg.as_string())
        smtp.quit()
        print("이메일 전송 성공!")
    except Exception as e:
        print("이메일 전송 실패:", str(e))

def main():
    config = configparser.ConfigParser()
    me = config['MAIL']['FROMMAIL']
    receivers = config['MAIL']['TOMAIL']
    subject = config['MAIL']['SUBJECT']
    message = config['MAIL']['MESSAGE']
    attach_files = config['MAIL']['ATTACHFILE']
    pw = config['MAIL']['PASSWORD']

    send_mail(me, receivers, subject, message, attach_files, pw)

if __name__ == "__main__":
    main()
