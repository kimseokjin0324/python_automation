import os

from libs.email_sender_with_attachment import EmailSender

if __name__ == '__main__':
    es = EmailSender('kimseokjin0324@gmail.com',
                     os.getenv('MY_GMAIL_PASSWORD'),
                     manager_name='김지유',
                     template_filename='templates/email_template_1.html'
                     )

    es.send_all_emails('email_list.xlsx')