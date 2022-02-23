import os
import fitz
from PyPDF2 import PdfFileWriter, PdfFileReader

from automation import Automation


class TaxBot(Automation):
    def __init__(self):
        super().__init__()

        # Source Path should be static
        self.source_path = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test"

    def run(self):

        # while not self.check_for_00_form():
        #     return

        # self.read_out_pdf()

        # 0. Read in the pdf and set the name variables ####
        first_name, last_name, sin, email = self.read_taxprep_pdf()

        sin = sin.replace(' ', '')
        sin_pw = sin[:-4]  # last 4 digits
        file_name = f"{last_name}_{first_name.replace(' ', '_')}"
        last_init = last_name[0]
        client_folder_path = f"{last_name}, {first_name.split(' ')[0]}"
        year = 2021
        output_folder = "_FINAL T1 DOCS"

        # path_name = "Holman, Kimberly"
        # name = "Holman_Kimberly_A"

        # 1. Create destination path: ####
        destination_path, city = self.select_directory(client_folder_path, year, last_init, output_folder)
        if not destination_path:
            print(f"Can't find the directory for {client_folder_path}")
            return

        print(f"Destination path successfully found for {client_folder_path}")
        # self.destination_path = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder"
        # self.destination_path = "I:/Toronto/_Personal Tax - TOR/H/Holman, Kimberly_94214/2020/CRA"

        # 2. move the pdf files ####

        # file_1 = file_name + f"_1-Ltr_{year}.pdf"
        # file_2 = file_name + f"_2-T1_({city})_{year}.pdf"
        # file_3 = file_name + f"_3-Docs_Needing_Signature_{year}.pdf"
        # files = [file_1, file_2, file_3]

        files = self.get_matching_files(file_name)
        self.move_files(self.source_path, destination_path, files)

        # 3. Decrypt the letter pdf and extract the information
        letter_name = file_name + f"_1-Ltr_{year-1}.pdf"
        # letter_name = file_name + f"_2-T1_({city})_{year -1}.pdf"

        self.decrypt_pdf(destination_path, letter_name, sin_pw)

        partner, total = self.get_letter_info(destination_path, letter_name)
        partner_email = self.get_partner_email(partner)

        # 3. Send the email with attachments
        to_address = "david.watson@bakertillywm.com"
        subject = f"Tax report for: {first_name.split(' ')[0]}"
        html_body = '<h3>This is HTML Body</h3>'
        body = f"Send to {email} from {partner_email}\n\n" \
               "Dear Kimberly,\n\n" \
               f"Please find attached your tax return report. The total amount owing is {total}\n\n" \
               "Kind Regards"
        print(body)
        return
        self.send_email(to_address=to_address,
                        subject=subject,
                        html_body=html_body,
                        body=body,
                        attachment_files=files,
                        attachment_file_path=destination_path
                        )
        print("email sent")
        return

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
        # for idx, line in enumerate(text_list):
        #    print(f"{idx}: {line}")
        print(f"{first_name}, {last_name}: {sin}, {email}")

        return first_name, last_name, sin, email

    def read_out_pdf(self):

        with fitz.open(self.source_path + "/Holman_Kimberly_A_1-Ltr_2020.pdf") as doc:
            text = ""
            for page in doc:
                text += page.getText()

        text_list = text.split("\n")
        for idx, line in enumerate(text_list):
            print(f"{idx}: {line}")

        return

    def select_directory(self, path_name, year, last_init, folder):

        tor = self.check_directory(f"I:/Toronto/_Personal Tax - TOR/{last_init}", path_name)
        van = self.check_directory(f"I:/Vancouver/_Personal Tax - VAN/{last_init}", path_name)
        if van:
            path = f"I:/Vancouver/_Personal Tax - VAN/{last_init}/{van}/{year}/{folder}"
            self.create_directory(path)
            city = "Vancouver"
        elif tor:
            path = f"I:/Toronto/_Personal Tax - TOR/{last_init}/{tor}/{year}/{folder}"
            self.create_directory(path)
            city = "Toronto"
        else:
            return False, None

        return path, city

    def decrypt_pdf(self, path, file, password):

        out = PdfFileWriter()
        if not self.check_directory(path, file):
            print("No Cover Letter Found")
            return

        file = PdfFileReader(f"{path}/{file}")

        # Check if the opened file is actually Encrypted
        if file.isEncrypted:

            file.decrypt(password)
            for idx in range(file.numPages):
                page = file.getPage(idx)
                out.addPage(page)

            # remove the file
            os.remove(f"{path}/{file}")

            with open(f"{path}/{file}", "wb") as f:
                out.write(f)

            print("File decrypted Successfully.")
        else:
            print("File already decrypted.")

    @staticmethod
    def get_letter_info(destination_path, letter_name):
        with fitz.open(f"{destination_path}/{letter_name}") as doc:
            text = ""
            for page in doc:
                text += page.getText()

        text_list = text.split("\n")
        for idx, line in enumerate(text_list):
            print(f"{idx}: {line}")
        total_amount = text_list[51]
        partner_line = text_list[65].split(' ')
        print(partner_line)
        partner = f"{partner_line[1]} {partner_line[2].replace(',','')}"
        print(partner)

        return partner, total_amount

    @staticmethod
    def get_partner_email(partner):

        # set up a read table here:
        if partner == "John Sinclair":
            return 'John.Sinclair@bakertillywm.com'


# ToDo Test the pdf decrypter
# ToDo Read and get the statement information from the letter