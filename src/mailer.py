import smtplib
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import base64
from google_auth_oauthlib.flow import InstalledAppFlow
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

EMAIL_ADDRESS = "drakegags@gmail.com"

class Mailer(object):

    def __init__(self):  
        self.token = self.authenticate()

        self.auth_string = self.GenerateOAuth2String(EMAIL_ADDRESS)

        self.smtp_conn = self.smtpAuthentication()

        # self.smtp_conn.quit()

        # self.TestSmtpAuthentication(EMAIL_ADDRESS, self.OAuth2String)

        # self.imap_conn = imaplib.IMAP4_SSL('imap.gmail.com')
        # self.imap_conn.debug = 4
        # self.imap_conn.authenticate('XOAUTH2', lambda x : auth_string)


    def authenticate(self):
        SCOPES = ['https://mail.google.com/']
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('creds/token.json'):
            creds = Credentials.from_authorized_user_file('creds/token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'creds/app_credentials.json', SCOPES)
                creds = flow.run_local_server(port=5001)
            # Save the credentials for the next run
            with open('creds/token.json', 'w') as token:
                token.write(creds.to_json())
        return creds.token

    def GenerateOAuth2String(self, username):
        """Generates an IMAP OAuth2 authentication string.
        See https://developers.google.com/google-apps/gmail/oauth2_overview
        Args:
            username: the username (email address) of the account to authenticate
            access_token: An OAuth2 access token.
        Returns:
            The SASL argument for the OAuth2 mechanism.
        """
        auth_string = 'user=%s\1auth=Bearer %s\1\1' % (username, self.token)
        return auth_string         

    def smtpAuthentication(self):
        print("Otentikasi SMTP...", end="")
        smtp_conn = smtplib.SMTP("smtp.gmail.com", 587)
        # smtp_conn.set_debuglevel(True)
        smtp_conn.ehlo('test')
        smtp_conn.starttls()
        smtp_conn.docmd('AUTH', 'XOAUTH2 ' + base64.b64encode(self.auth_string.encode()).decode())
        print("\rOtentikasi selesai")
        return smtp_conn


    def smtpQuit(self):
        self.smtp_conn.quit()

    def send_mail(self, number, recipient, recipient_name, data, bill_month, bill_year):
        message = MIMEMultipart("alternative")
        message['From'] = EMAIL_ADDRESS
        message['To'] = recipient
        message['Subject'] = f"Tagihan Telepon {number} {bill_month}-{bill_year}"

        # Create the plain-text and HTML version of your message
        text = f"""\
        Yang terhormat,
        {recipient_name}
        Berdasarkan data pemakaian telpon, berikut ini kami sampaikan rincian tagihan dengan nomor {number} bulan {bill_month} tahun {bill_year}."""

        html = f"""
        <html>
        <body>
            <p>Yang terhormat,</p>
            <p>{recipient_name}</p>
            <p>Berdasarkan data pemakaian telpon, berikut ini kami sampaikan rincian tagihan dengan nomor {number} bulan {bill_month} tahun {bill_year}.</p>
            <table style="border-collapse: collapse;">
                <tr>
                """
        for row in data:
            text += f"""{row[0]:20s}: {row[1]}"""
            html += f"""
                    <th style="color:#ffffff;background:#0081b9;padding:5px;border:2px solid #ebebeb">
                        {row[0]}
                    </th>
            """
        html += """
                </tr>
                <tr>
        """        
        for row in data:
            html += f"""
                    <td style="background:white;color:#0D6BB7;padding:5px;border:2px solid #ebebeb;">
                        {row[1]}
                    </td>
            """

        text += """   
            Atas kerja sama Bapak/Ibu, kami mengucapkan terima kasih.
            Salam hormat,
            GEF NT
        """
        html += """   
                </tr>
            </table> 
            <p>Atas kerja sama Bapak/Ibu, kami mengucapkan terima kasih.</p>
            <p>Salam hormat,</p>
            <p>GEF NT</p>
        </body>
        </html>
        """

        # Turn these into plain/html MIMEText objects
        part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        message.attach(part1)
        message.attach(part2)

        self.smtp_conn.sendmail(
            EMAIL_ADDRESS, recipient, message.as_string()
        )

    def _testSmtpAuthentication(self, user, auth_string):
        """Authenticates to SMTP with the given auth_string.
        Args:
            user: The Gmail username (full email address)
            auth_string: A valid OAuth2 string, not base64-encoded, as returned by
                GenerateOAuth2String.
        """
        smtp_conn = smtplib.SMTP("smtp.gmail.com", 587)
        smtp_conn.set_debuglevel(True)
        smtp_conn.ehlo('test')
        smtp_conn.starttls()
        smtp_conn.docmd('AUTH', 'XOAUTH2 ' + base64.b64encode(auth_string.encode()).decode())



        sender_email = user
        receiver_email = "drakegags+" + "fakegags@gmail.com"
        message = """\
        Subject: Hi there

        This message is sent from Python."""

        smtp_conn.sendmail(sender_email, receiver_email, message)

if __name__ == "__main__":
    mailer = Mailer()
    