from pythonzimbra.communication import Communication
from pythonzimbra.tools import auth

url  = 'https://webmail1.angkasapura2.co.id/service/soap'

class Mailer(object):

    def __init__(self, sender_email, sender_password):
        self.email_adress = sender_email

        self.comm = Communication(url)

        self.token = self.authenticate(sender_email, sender_password)

    def authenticate(self, email, password):
        token = auth.authenticate(
            url,
            email,
            password,
            use_password=True
        )
        if token is None:
            print("Gagal login")
            exit()
        return token

    def send_mail(self, extract, bill_month, bill_year):

        html = f"""
        <html>
        <body>
            <p>Yang terhormat,</p>
            <p>Senior Manager {extract["jabatan"]}</p>
            <p>Berdasarkan data pemakaian telpon, berikut ini kami sampaikan rincian tagihan telepon bulan {bill_month} tahun {bill_year}.</p>
            <table style="border-collapse: collapse;">
                <tr>
                """
        for key in extract["data"][0]:
            html += f"""
                    <th style="color:#ffffff;background:#007cc3;padding:5px;border:2px solid #f2f2f2">
                        {key}
                    </th>
            """
        html += """
                </tr>
        """        
        for row in extract["data"]:
            html += """
                    <tr>
            """     
            for key in row:
                html += f"""
                        <td style="background:white;padding:5px;border:2px solid #f2f2f2;">
                            {row[key]}
                        </td>
                """
            html += """
                    </tr>
            """     

        html += """   
            </table> 
            <p>Atas kerja sama Bapak/Ibu, kami mengucapkan terima kasih.</p>
            <p>Salam hormat,</p>
            <p>GEF NT</p>
        </body>
        </html>
        """
        
        info_request = self.comm.gen_request(token=self.token)
        
        info_request.add_request(
            'SendMsgRequest',
            {
                'm': {
                    'su' : f"Tagihan Telepon {bill_month} {bill_year}",
                    'e': {
                        'a':extract["email"],
                        't':'t'
                    },
                    'mp': {
                        'ct' : 'text/html',
                        'content' : html
                    }
                }
            },
            'urn:zimbraMail'
        )


        info_response = self.comm.send_request(info_request)

        if info_response.is_fault():
            print(info_response.get_fault_code())
            print(info_response.get_fault_message())

if __name__ == "__main__":
    mailer = Mailer()