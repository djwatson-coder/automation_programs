
# Implement before Deployment
# ToDo robustness checks and exception handling - map out different errors that can happen and how to deal with them
# ToDo Change the pickle file to a csv file
# ToDo Add writing to the csv file after every employee has been completed (not on error)
# ToDo find the dependencies through print to send an email only to the first name - if there is one
# ToDo test on 10 Random Samples
# Nice to have
# ToDo implement at GUI to replace the console

import time
from datetime import datetime
from automation import Automation
import pickle
from tbsettings import *


class TaxBot(Automation):
    """
    Tax Bot Takes Personal Tax Files, stores the relevant files in the correct client location and
    sends an email to the Client partner with the attachments and formatted message to the client.
    """
    def __init__(self):
        super().__init__()

        # Source Path should be static
        self.source_path = source_path
        # self.source_path = "L:\Toronto\T1\T1-2020\PDF"
        self.manual_to_do_path = manual_to_do_path

        self.toronto_personal_client_dir = toronto_personal_client_dir
        self.vancouver_personal_client_dir = vancouver_personal_client_dir

        self.completed_entities = completed_entities
        self.tax_prep_string = tax_prep_string

        self.slp = 30  # length of time to sleep for
        self.remove_files = remove_files
        self.enable_emailing = enable_emailing

        # Other variables
        self.year = 2021
        self.output_folder = "_FINAL T1 DOCS"

        self.email_partner = email_partner
        self.email_string = email_string

    def run(self, store=True):
        """
        Checks the folder to see if there are any tax prep files that have not been stored and sent.
        If there are any it runs the storing and sending process.
        :return: None
        :rtype: None
        """
        if store:
            infile = open("../sample.pkl", 'rb')
            self.completed_entities = pickle.load(infile)
            infile.close()
            outfile = open("../sample.pkl", 'wb')
        try:
            while True:
                file = self.check_input_folder(self.tax_prep_string)
                if file:
                    self.print_hash_comment("Begin taxbot process")
                    print(f"File Found: {file}")
                    if self.run_process(file):
                        self.completed_entities.append(file)
                        if store:
                            # ToDo Add to the PICKLE storage
                            print("ToDo Add to the pickle storage")
                        self.print_hash_comment("#####")
                    else:
                        self.move_files(self.source_path, self.manual_to_do_path, [file], True)
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

    def run_process(self, document):

        # 0. Read in the pdf and set the name variables ####
        first_name, last_name, sin, email, client_code = self.read_taxprep_pdf(document)

        file_name = f"{last_name}_{first_name.replace(' ', '_')}"
        last_init = last_name[0]
        client_folder_path = f"{last_name}, {first_name.split(' ')[0]}_{client_code}"

        # 1. Create destination path: ####
        destination_path, city = self.select_directory(client_folder_path, self.year, last_init, self.output_folder)
        if not destination_path:
            print(f"Can't find the directory for {client_folder_path}")
            return False

        print(f"Destination path successfully found for {first_name} {last_name}")

        # 2. move the pdf files ####
        files = self.get_matching_files(self.source_path, file_name)
        if len(files) == 0:
            print(f"-- COULD NOT FIND ANY FILES FOR {first_name} {last_name}")
            return False

        self.move_files(self.source_path, destination_path, files, self.remove_files)
        print(f"FILES MOVED TO: {destination_path}")

        # 3. Decrypt the letter pdf and extract the information
        letter_name = file_name + f"_1-Ltr_{self.year - 1}.pdf"
        self.decrypt_pdf(destination_path, letter_name, sin)
        letter_info = self.get_letter_info(destination_path, letter_name)
        partner = letter_info["partner"]

        # 4. Send the email with attachments
        to_address = "david.watson" + email_string  # "deepak.upadhyaya"
        if self.email_partner:
            to_address = partner.replace(' ', '.').lower() + email_string

        subject = f"Tax report for: {first_name.split(' ')[0]}"
        html_body = '<h3>This is HTML Body</h3>'
        body = f"Send to {email} from {partner.replace(' ', '.').lower() + email_string}\n\n" \
               f"Dear {first_name},\n\n" \
               f"Please find attached your tax return report. \n\n" \
               "Kind Regards"

        self.send_email(to_address=to_address,
                        subject=subject,
                        html_body=html_body,
                        body=body,
                        attachment_files=files,
                        attachment_file_path=destination_path,
                        send=self.enable_emailing)

        print("EMAIL SENT TO:")
        print(f"    Partner: {partner} \n"
              f"    Email: {partner.replace(' ', '.').lower() + email_string}")
        print(f"Process Complete for: {first_name} {last_name}")

        return True

    def read_taxprep_pdf(self, document):

        text_list = self.get_pdf_as_list(f"{self.source_path}/{document}")
        # self.read_out_pdf_list(text_list)
        sin_place, sin = self.find_first_pattern(text_list, "[0-9]{3} [0-9]{3} [0-9]{3}")
        email_place, email = self.find_first_pattern(text_list, "[0-z]*@[0-z]*(.com)")
        start_spot, a = self.find_first_pattern(text_list, "Client code")
        client_code_place, client_code = self.find_first_pattern(text_list[start_spot:start_spot + 10], "^[0-9]*$")
        if email_place == 0:
            email = "No Email"
        if client_code_place == 0:
            client_code = ""
        first_name = text_list[sin_place - 2]
        last_name = text_list[sin_place - 1]
        sin = sin.replace(' ', '')
        # Get All printing Names
        printing_place, printing = self.find_first_pattern(text_list, "Printing Options")
        print(f"Printing ----{text_list[printing_place + 1]}")

        print(f"    Name:{first_name} {last_name}\n"
              f"    SIN: {sin}\n"
              f"    Client Code: {client_code}\n"
              f"    Email: {email}")

        return first_name, last_name, sin, email, client_code

    def select_directory(self, path_name, year, last_init, folder):

        tor_options, tor_close_matches = self.fuzzy_check_directory(f"{self.toronto_personal_client_dir}/"
                                                                      f"{last_init}", path_name, "Toronto ")
        van_options, van_close_matches = self.fuzzy_check_directory(f"{self.vancouver_personal_client_dir}/"
                                                                      f"{last_init}", path_name, "Vancouver ")
        options = tor_options + van_options
        close_matches = tor_close_matches + van_close_matches

        city, city_folder = self.triage_folder_options(options, close_matches)

        if city == "Vancouver":
            path = f"{self.vancouver_personal_client_dir}/{last_init}/{city_folder}/{year}/{folder}"
            self.create_directory(path)
        elif city == "Toronto":
            path = f"{self.toronto_personal_client_dir}/{last_init}/{city_folder}/{year}/{folder}"
            self.create_directory(path)
        else:
            raise NotADirectoryError("No File Path for the client at either:\n"
                            f"{self.toronto_personal_client_dir}/{last_init}/{path_name}\n"
                            f"{self.vancouver_personal_client_dir}/{last_init}/{path_name}")

        return path, city

    def triage_folder_options(self, options, close_matches):
        """
        Takes a dictionary of options and selects the best one for the process. Takes user input when needed
        :return: city, city_path - the city and the path to use for the folder
        :rtype: str, str
        """
        if len(options) == 1:
            return options[0].split(" ", 1)
        elif len(options) > 1:
            return self.get_user_selection(options, "Mulitple Options, which one is the right one?").split(" ", 1)
        if len(close_matches) == 0:
            raise NotADirectoryError("Can't find a matchable directory")
        else:
            return self.get_user_selection(close_matches, "No exact match, would you like to "
                                                          "select one of the close options below ?").split(" ", 1)

    def get_user_selection(self, sel_list, text):

        response = 0
        self.print_hash_comment("fuzzy match input needed")
        print(text)
        "   0 - None (stop the bot)\n"
        for idx, c_match in enumerate(sel_list):
            print(f"    {idx + 1}: {c_match}")

        response = int(input("Enter Here:  "))
        self.print_hash_comment("####")
        if response == 0:
            raise NotADirectoryError("No valid or close directory")
        else:
            return sel_list[response - 1]

    def get_letter_info(self, path, letter_name):
        text_list = self.get_pdf_as_list(f"{path}/{letter_name}")
        # self.read_out_pdf_list(text_list)
        idx, partner_line = self.find_first_pattern(text_list, "Per: ")
        partner_line = partner_line.split(' ')
        letter_info = {"partner": f"{partner_line[1]} {partner_line[2].replace(',','')}"}


        return letter_info

    def check_input_folder(self, matching_string="00"):
        files = self.get_matching_files(self.source_path, matching_string)
        files_to_process = []

        for file in files:
            if file not in self.completed_entities:
                files_to_process.append(file)

        if len(files_to_process) == 0:
            return False
        else:
            return files_to_process[0]

    @staticmethod
    def move_to_manual_folder(name):
        files = []
        # find the files based on the name


