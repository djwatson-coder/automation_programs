import time
from datetime import datetime
from PyPDF2 import PdfFileWriter, PdfFileReader
from pikepdf import Pdf
from automation import Automation
import pickle


class TaxBot(Automation):
    """
    Tax Bot Takes Personal Tax Files, stores the releveant files in the correct client location and
    sends an email to the Client partner with the attachments and formatted message to the client.
    """
    def __init__(self):
        super().__init__()

        # Source Path should be static
        self.source_path = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test"
        #self.source_path = "L:\Toronto\T1\T1-2020\PDF"
        #self.source_path =

        self.toronto_personal_client_dir = "I:/Toronto/_Personal Tax - TOR"
        self.vancouver_personal_client_dir = "I:/Vancouver/_Personal Tax - VAN"

        self.completed_entities = [] # stored in a pickle file
        self.tax_prep_string = "00"

        self.slp = 30

    def run(self, store=True):
        """
        Checks the folder to see if there are any tax prep files that have not been stored and sent.
        If there are any it runs the storing and sending process.
        :return: None
        :rtype: None
        """
        if store:
            infile = open("sample.pkl", 'rb')
            self.completed_entities = pickle.load(infile)
            infile.close()
            outfile = open("sample.pkl", 'wb')

        try:
            while True:
                file = self.check_input_folder(self.tax_prep_string)
                if file:
                    self.print_hash_comment("Begin taxbot process")
                    print(f"File Found: {file}")
                    if self.run_process(file):
                        self.completed_entities.append(file)
                        self.print_hash_comment("#####")
                else:
                    # wait 60 seconds
                    print("No New Tax Prep Documents to process")
                    print(f"Completed Ones: {self.completed_entities}")
                    print(f"sleeping for {self.slp} seconds....{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    time.sleep(30)

        except KeyboardInterrupt or Exception:
            if store:
                pickle.dump(self.completed_entities, outfile)
                outfile.close()
            raise

    def run_process(self, document_00):

        # 0. Read in the pdf and set the name variables ####
        first_name, last_name, sin, email, client_code = self.read_taxprep_pdf(document_00)

        sin = sin.replace(' ', '')
        sin_pw = sin
        file_name = f"{last_name}_{first_name.replace(' ', '_')}"
        last_init = last_name[0]
        client_folder_path = f"{last_name}, {first_name.split(' ')[0]}_{client_code}"
        year = 2021
        output_folder = "_FINAL T1 DOCS"

        # 1. Create destination path: ####
        destination_path, city = self.select_directory(client_folder_path, year, last_init, output_folder)
        if not destination_path:
            print(f"Can't find the directory for {client_folder_path}")
            return False

        print(f"Destination path successfully found for {client_folder_path}")

        # 2. move the pdf files ####
        files = self.get_matching_files(file_name)
        self.move_files(self.source_path, destination_path, files)

        print(f"FILES MOVED TO: {destination_path}")

        # 3. Decrypt the letter pdf and extract the information
        letter_name = file_name + f"_1-Ltr_{year - 1}.pdf"
        self.decrypt_pdf(destination_path, letter_name, sin_pw)
        partner, total = self.get_letter_info(destination_path, letter_name)

        # 4. Send the email with attachments
        email_string = "@bakertillywm.com"
        to_address = f"david.watson{email_string}" # deepak.upadhyaya
        subject = f"Tax report for: {first_name.split(' ')[0]}"
        html_body = '<h3>This is HTML Body</h3>'
        body = f"Send to {email} from {partner.replace(' ', '.').lower() + email_string}\n\n" \
               "Dear Kimberly,\n\n" \
               f"Please find attached your tax return report. The total amount owing is {total}\n\n" \
               "Kind Regards"

        self.send_email(to_address=to_address,
                        subject=subject,
                        html_body=html_body,
                        body=body,
                        attachment_files=files,
                        attachment_file_path=destination_path,
                        send=False)

        print("EMAIL SENT TO:")
        print(f"    Partner: {partner} \n"
              f"    Email: {partner.replace(' ', '.').lower() + email_string}")
        print(f"Process Complete for: {first_name} {last_name}")

        return True

    def read_taxprep_pdf(self, document_00):

        text_list = self.get_pdf_as_list(f"{self.source_path}/{document_00}")
        # self.read_out_pdf_list(text_list)
        sin_place, sin = self.find_first_pattern(text_list, "[0-9]{3} [0-9]{3} [0-9]{3}")
        email_place, email = self.find_first_pattern(text_list, "[0-z]*@[0-z]*(.com)")
        client_code_place, client_code = self.find_first_pattern(text_list, "Client code")
        client_code = text_list[client_code_place + 3]
        if email_place == 0:
            email = "No Email"
        if client_code_place == 0:
            client_code = ""
        first_name = text_list[sin_place - 2]
        last_name = text_list[sin_place - 1]

        print(f"    Name:{first_name} {last_name}\n"
              f"    SIN: {sin}\n"
              f"    Client Code: {client_code}\n"
              f"    Email: {email}")

        return first_name, last_name, sin, email, client_code

    def select_directory(self, path_name, year, last_init, folder):

        tor = self.fuzzy_check_directory(f"{self.toronto_personal_client_dir}/{last_init}", path_name)
        van = self.fuzzy_check_directory(f"{self.vancouver_personal_client_dir}/{last_init}", path_name)

        if van:
            path = f"{self.vancouver_personal_client_dir}/{last_init}/{van}/{year}/{folder}"
            self.create_directory(path)
            city = "Vancouver"
        elif tor:
            path = f"{self.toronto_personal_client_dir}/{last_init}/{tor}/{year}/{folder}"
            self.create_directory(path)
            city = "Toronto"
        else:
            raise NotADirectoryError("No File Path for the client at either:\n"
                            f"{self.toronto_personal_client_dir}/{last_init}/{path_name}\n"
                            f"{self.vancouver_personal_client_dir}/{last_init}/{path_name}")

        return path, city

    def decrypt_pdf(self, path, file, password):

        if not self.check_directory(path, file):
            raise NotADirectoryError(f"No File Path for the client at :\n{path}/{file}")

        pdf_file = PdfFileReader(f"{path}/{file}")

        # Check if the opened file is actually Encrypted
        if pdf_file.isEncrypted:

            with Pdf.open(f"{path}/{file}", password=password, allow_overwriting_input=True) as pdf:
                pdf.save(f"{path}/{file}")

            print("File decrypted Successfully.")
        else:
            print("File already decrypted.")

    def get_letter_info(self, path, letter_name):
        text_list = self.get_pdf_as_list(f"{path}/{letter_name}")
        # self.read_out_pdf_list(text_list)
        total_amount = text_list[51]
        idx, partner_line = self.find_first_pattern(text_list, "Per: ")
        partner_line = partner_line.split(' ')
        partner = f"{partner_line[1]} {partner_line[2].replace(',','')}"

        return partner, total_amount

    @staticmethod
    def get_partner_email(partner):

        # set up a read table here:
        if partner == "John Sinclair":
            return 'John.Sinclair@bakertillywm.com'

    def check_input_folder(self, matching_string="00"):
        files = self.get_matching_files(matching_string)
        files_to_process = []

        for file in files:
            if file not in self.completed_entities:
                files_to_process.append(file)

        if len(files_to_process) == 0:
            return False
        else:
            return files_to_process[0]


# ToDo Add client ID to the folder name
# ToDo robustness checks and exception handling
# ToDo get the accurate place of amounts in the letter
# ToDo find the dependencies through print to send an email only to the first name
# ToDO Handle the triaging of fuzzy, no-option and 1 options better between TOR and VAN so the user is only asked
#  for input when necessary
