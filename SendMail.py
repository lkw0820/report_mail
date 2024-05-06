import smtplib
import re
import os
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#SMTP 서버를 dictionary로 정의
smtp_info ={
    'hiworks':('smtps.hiworks.com',465),
    'naver.com': ('smtp.naver.com', 587)
}
#메일 보내는 함수 정의(발신 메일, 수신메일(여러개 가능), 제목, 본문, 첨부파일 경로, 비밀번호)
def send_mail(From, To, subject, message, attach_files=(),pw='',subtype=''):
    
    #멀티파트로 메일을 만들기 위한 포맷 생성
    form = MIMEBase('multipart','mixed')

    #입력받은 메일 주소와 제목, 본문, 등의 문자열을 인코딩해서 form에 입력
    form['Form'] = form
    form['To'] = ','.join(To)
    form['Subject'] = Header(subject.encode('utf-8'),'utf-8')
    msg = MIMEText(message.encode('utf-8'),_subtype=subtype,_charset='utf-8')
    form.attach(msg)
    #여러개의 파일을 하나씩 첨부
    for fpath in attach_files:
        folder,file = os.path.split(fpath)

        with open(fpath, 'rb') as f:
            body = f.read()
        msg=MIMEApplication(body,_subtype=subtype)

        msg.add_header('Content-Disposition', 'attatchment', filename=(Header(file,'utf-8').encode()))

        form.attach(msg)


    id, host = From.rsplit("@",1)
    smtp_server,port=smtp_info[host]

    #SMTP 접속여부 확인
    if port == 587:
        smtp = smtplib.SMTP(smtp_server,port)
        rcode1,_ = smtp.ehlo()
        rcode2,_ = smtp.starttls()
    else:
        smtp = smtplib.SMTP(smtp_server,port)
        rcode1,_ = smtp.ehlo()
        recode2 = 220
    if rcode1 != 250 or recode2 != 220:
        smtp.quit()
        return '연결 실패'
    smtp.login(From,pw)
    smtp.sendmail(From,To,form.as_string())
    smtp.quit

def main():


if __name__ == "__main__":
    main()
