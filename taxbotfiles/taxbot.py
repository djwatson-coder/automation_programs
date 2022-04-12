
# Implement before Deployment
# ToDo robustness checks and exception handling - map out different errors that can happen and how to deal with them
# ToDo Change the pickle file to a csv file
# ToDo Add writing to the csv file after every employee has been completed (not on error)
# ToDo test on 20 Random Samples
# Nice to have
# ToDo implement at GUI to replace the console
import sys
import time
from datetime import datetime
from automation import Automation
from taxbotfiles.tbsettings import *


class TaxBot(Automation):
    """
    Tax Bot Takes Personal Tax Files, stores the relevant files in the correct client location and
    sends an email to the Client partner with the attachments and formatted message to the client.
    """
    def __init__(self, location="TOR"):
        super().__init__()

        self.location = location.upper()
        self.client_info = self.create_client_info(client_info_file)

        # Paths to Transition files
        if location == "TOR":
            self.source_path = source_path_tor
            self.completed_folder = completed_folder_tor
            self.issues_folder = issues_folder_tor
        else:  # location == "VAN":
            self.source_path = source_path_van
            self.completed_folder = completed_folder_van
            self.issues_folder = issues_folder_van

        self.toronto_personal_client_dir = toronto_personal_client_dir
        self.vancouver_personal_client_dir = vancouver_personal_client_dir

        # Storing Data
        self.completed_entities = completed_entities

        # Run-Time Settings
        self.slp = slp
        self.remove_files = remove_files
        self.enable_emailing = enable_emailing
        self.email_partner = email_partner
        self.take_first_selection = take_first_selection
        self.email_contents = email_contents
        self.email_saving = email_saving

        # Run-Time Variables
        self.year = yr
        self.output_folder = output_folder
        self.email_string = email_string
        self.tax_prep_string = tax_prep_string

    def run(self):

        """
        Checks the folder to see if there are any tax prep files that have not been stored and sent.
        If there are any it runs the storing and sending process.
        """

        self.log_info(f'\n\n{self.location} Bot Process Starting at: '
                      f'{time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())}\n\n')

        while True:
            try:
                file = self.check_input_folder([self.tax_prep_string])
                if file:
                    self.print_hash_comment("Begin TaxBot process")
                    self.log_info(f"File Found: {file}")
                    if self.run_process(file):
                        self.completed_entities.append(file)
                        self.move_files(self.source_path, self.completed_folder, [file], remove=False)
                        self.print_hash_comment("")

                    else:
                        self.move_files(self.source_path, self.issues_folder, [file], True)  # Make this
                else:
                    self.pause_bot(seconds=self.slp)
                self.write_to_log()
            except KeyboardInterrupt as e:
                self.log_info(repr(e))
                self.write_to_log()
                sys.exit()

            except Exception as e:
                print(e)
                self.log_info(repr(e))
                self.write_to_log()
                self.move_files(self.source_path, self.issues_folder, [file], True)
                time.sleep(5)

    def run_process(self, document):

        # 0. decrypt the ID document, Read it in and set the client variables ####
        sin = self.get_client_sin(document)
        self.log_info("Decrypting ID File...")
        self.decrypt_pdf(self.source_path, document, sin)
        first_name, last_name, sin_pdf, email, client_code = self.read_taxprep_pdf(document)

        self.log_info(f"Sin Matching Correct: {sin == sin_pdf}")

        file_name = f"{last_name}_{first_name.replace(' ', '_')}"
        last_init = last_name[0]
        # client_folder_path = f"{last_name}, {first_name.split(' ')[0]}_{client_code}"
        client_folder_path = f"{last_name}, {first_name}_{client_code}"

        # 1. Find the destination path based off the clients name and code
        destination_path, city = self.select_directory(client_folder_path, self.year+2, last_init, self.output_folder)
        if not destination_path:  # ToDo Check this before implementation
            self.log_info(f"Can't find the directory for {client_folder_path}")
            return False

        self.log_info(f"Destination path successfully found for {first_name} {last_name}")

        # 2. Move the PDF Files tha match the client string to the new folder
        files = self.get_matching_files(self.source_path, matching_strings=[file_name])
        if len(files) <= 1:
            self.log_info(f"-- COULD NOT FIND ANY FILES FOR {first_name} {last_name}")
            return False

        self.move_files(self.source_path, destination_path, files, self.remove_files)
        self.create_text_file(destination_path, name=sin, sin=sin, path=destination_path)
        self.log_info(f"FILES MOVED TO: {destination_path}")
        for idx, file in enumerate(files):
            self.log_info(f"    {idx}: {file} ")

        # 3. Decrypt the letter pdf and extract the information
        # letter_name = file_name + f"_1-Ltr_{self.year-1}.pdf"  # ToDo Check this before implementation
        letter_name = self.get_matching_files(destination_path, matching_strings=["_1-Ltr_"])[0]
        self.log_info("Decrypting Letter File...")
        self.decrypt_pdf(destination_path, letter_name, sin)
        letter_info = self.get_letter_info(destination_path, letter_name)
        partner = letter_info["partner"]

        # 4. Create and Save/Send Email
        to_address = self.get_email_address(partner)
        subject = f"Tax report for: {first_name.split(' ')[0]}"
        html_body = self.email_contents
        attachment_files = self.get_matching_files(destination_path, matching_strings=["_1-", "_2-", "_3-", "_7-"])

        email_sent = self.create_email(to_address="",
                                       subject=subject, html_body=html_body,  # body=body,
                                       attachment_files=attachment_files, attachment_file_path=destination_path,
                                       send=self.enable_emailing,
                                       save=self.email_saving, save_path=destination_path, save_name=file_name)

        self.print_email_complete(email_sent, to_address=to_address)
        self.log_info(f"TAXBOT PROCESS SUCCESSFULLY COMPLETED FOR: {first_name} {last_name}")

        return True

    def read_taxprep_pdf(self, document):

        text_list = self.get_pdf_as_list(f"{self.source_path}/{document}")
        # self.read_out_pdf_list(text_list)
        sin_place, sin = self.find_first_pattern(text_list, "[0-9]{3} [0-9]{3} [0-9]{3}")
        email_place, email = self.find_first_pattern(text_list, "[0-z]*@[0-z]*(.com)")
        start_spot, a = self.find_first_pattern(text_list, "Client code")
        client_code_place, client_code = self.find_first_pattern(text_list[start_spot:start_spot + 10], "^[0-9]*$")
        # delivery_type_1, delivery_type_2 = self.find_delivery_type(text_list)
        # print(f"Delivery 1: {delivery_type_1}")
        # print(f"Delivery 1: {delivery_type_2}")
        if email_place == 0:
            email = "No Email"
        if client_code_place == 0:
            client_code = ""
        first_name = text_list[sin_place - 2]
        last_name = text_list[sin_place - 1]
        sin = sin.replace(' ', '')
        # Get All printing Names
        # printing_place, printing = self.find_first_pattern(text_list, "Printing Options")
        # print(f"Printing ----{text_list[printing_place + 1]}")

        self.log_info(f"    Name:{first_name} {last_name}\n"
                      f"    SIN: {sin}\n"
                      f"    Client Code: {client_code}\n"
                      f"    Email: {email}")

        return first_name, last_name, sin, email, client_code

    def select_directory(self, path_name, year, last_init, folder):

        if self.location == "TOR":
            options, close_matches = self.fuzzy_check_directory(f"{self.toronto_personal_client_dir}/"
                                                                f"{last_init}", path_name, "Toronto ")
        else:
            options, close_matches = self.fuzzy_check_directory(f"{self.vancouver_personal_client_dir}/"
                                                                f"{last_init}", path_name, "Vancouver ")
        # options = van_options + tor_options
        # close_matches = van_close_matches + tor_close_matches

        self.log_info(f"looking for a directory that matches: {path_name}")
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

        self.print_hash_comment("INPUT NEEDED -- File Location")
        print(text)
        print("    0: None (stop the bot)")
        for idx, c_match in enumerate(sel_list):
            print(f"    {idx + 1}: {c_match}")

        response = 1
        if not self.take_first_selection:
            response = int(input("Enter Here:  "))
        self.print_hash_comment()
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

    def check_input_folder(self, matching_strings):
        files = self.get_matching_files(self.source_path, matching_strings)
        files_to_process = []
        for file in files:
            if file not in self.completed_entities:
                files_to_process.append(file)

        if len(files_to_process) == 0:
            return False
        else:
            return files_to_process[0]

    @staticmethod
    def find_delivery_type(text_list):
        breaking_text = "/"
        # correct_types = ["paper", "e-mail", "pdf"]
        correct_types = ['"ab"-email', 'pape']
        for idx, line in enumerate(text_list):
            if breaking_text in line:
                del_types = line.split(breaking_text)
                print(del_types)
                if del_types[0].lower() in correct_types and del_types[1].lower() in correct_types:
                    return del_types[0], del_types[1]
        return False, False

    def print_email_complete(self, email_sent, **kwargs):
        if email_sent:
            self.log_info(f"EMAIL SENT TO: {kwargs['to_address']}")
        else:
            self.log_info("No Email Sent - setting disabled")

        if self.email_saving:
            self.log_info("Email Saved")
        else:
            self.log_info("No Email Saved - setting disabled")

    def pause_bot(self, seconds):
        try:
            # wait x seconds
            self.log_info("No New Tax Prep Documents to process")
            self.log_info(f"Completed Clients in current instance: {self.completed_entities}")
            self.log_info(f"sleeping for {seconds} seconds....{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            time.sleep(seconds)
        except KeyboardInterrupt as e:
            self.log_info(repr(e))
            self.write_to_log()
            sys.exit()

    def get_email_address(self, name):
        if not self.email_partner:
            name = "david.watson"
        return name.replace(' ', '.').lower() + email_string

    def write_to_log(self):
        with open(log_string, 'a') as f:
            f.write(self.bot_log)
            f.close()

    def create_client_info(self, client_info_file):
        df = self.read_csv_file(client_info_file, keep_cols=['First Name', 'Last Name', 'SIN'])
        df = df.assign(file_name=lambda x: df['Last Name'] + "_" + df['First Name'])
        df['file_name'] = df['file_name'].replace(' ', '_', regex=True)
        df['SIN'] = df['SIN'].replace(' ', '', regex=True)
        return df

    def get_client_sin(self, doc: str):
        # takes a document string and looks up the client SIN from the folder
        name = doc.split(tax_prep_string)[0]
        return self.client_info.loc[self.client_info["file_name"] == name, 'SIN'].iloc[0]
