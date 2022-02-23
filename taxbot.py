import os
import shutil
import win32com.client
import fitz

from automation import Automation


class TaxBot(Automation):
    def __init__(self):
        super().__init__()
        self.source_path = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test"
        self.destination_path = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder"

    def run(self):

        # 0. Read in the pdf
        first_name, last_name, sin, email = self.read_taxprep_pdf()

        return
        # 1. Read in the name data
        name = "Holman_Kimberly_A"

        # 2. move the pdf files
        file_1 = name + "_1-Ltr_2020.pdf"
        file_2 = name + "_2-T1_(Toronto)_2020.pdf"
        file_3 = name + "_3-Docs_Needing_Signature_2020.pdf"
        files = [file_1, file_2, file_3]
        self.move_pdfs(self.source_path, self.destination_path, files)

        # 3. Send the email with attachments
        to_address = "david.watson@bakertillywm.com"
        subject = "python_tax_bot Report for:" + name
        html_body = '<h3>This is HTML Body</h3>'
        body = "Dear Kimberly,\n\n " \
               "Please find attached your tax return report \n \n" \
               "Kind Regards"

        self.send_email(to_address=to_address,
                        subject=subject,
                        html_body=html_body,
                        body=body,
                        attachment_files=files,
                        attachment_file_path=self.destination_path
                        )

        return

    @staticmethod
    def move_pdfs(path, destination, files):
        for file in files:
            shutil.move(os.path.join(path, file), destination)

    @staticmethod
    def send_email(**kwargs):

        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = kwargs["to_address"]
        mail.Subject = kwargs['subject']
        mail.HTMLBody = kwargs['html_body']
        # mail.Body = kwargs['body']
        for attachment in kwargs["attachment_files"]:
            mail.Attachments.Add(kwargs["attachment_file_path"] + "/" + attachment)
        mail.Send()

    def read_taxprep_pdf(self):
        with fitz.open(self.source_path + "/00-Taxprep-Holman, Kim.pdf") as doc:
            text = ""
            for page in doc:
                text += page.getText()

        text_list = text.split("\n")
        first_name = text_list[29]
        last_name = text_list[30]
        sin = text_list[31]
        email = text_list[108]
        for idx, line in enumerate(text_list):
            print(f"{idx}: {line}")
        print(f"{first_name}, {last_name}: {sin}, {email}")

        return first_name, last_name, sin, email
