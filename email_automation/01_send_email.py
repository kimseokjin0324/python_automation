import os
import smtplib #파이썬 내장라이브러리
from email.mime.text import MIMEText


class EmailSender:
    email_addr =None
    password = None
    def __init__(self,email_addr,password):
        print('생성자')
        self.email_addr = email_addr
        self.password = password

    def send_email(self,msg,from_addr,to_addr):
        """
        :param msg: 보낼 메세지
        :param from_addr: 보내는 곳 사람
        :param to_addr: 받는 사람
        :return: 없을 예정
        """
        with smtplib.SMTP('smtp.naver.com', 587) as smtp:
            smtp.starttls()
            smtp.login(self.email_addr,self.password)
            #네이버는 구글과다르게 다른 작업이 필요하다
            msg = MIMEText(msg)
            msg['From'] = from_addr
            msg['To'] =to_addr
            msg['Subject'] = '메일 발송 테스트'
            #로그인 한 이후 이메일 보내기
            smtp.sendmail(from_addr=from_addr,to_addrs=to_addr,msg=msg.as_string())
            smtp.quit()
        print('이메일 전송이 완료되었습니다.')

if __name__ == '__main__':
    # es = EmailSender('kimseokjin0324@gmail.com',os.getenv('MY_GMAIL_PASSWORD'))
    es = EmailSender('sksmstjrwls1@naver.com',os.getenv('MY_NAVER_PASSWORD'))
    es.send_email(' 테스트 입니다 네이버에서 보냄.',from_addr='sksmstjrwls1@naver.com',to_addr='kimseokjin0324@gmail.com')