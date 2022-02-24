import os
import re
from pathlib import Path
import shutil
import win32com.client
import fitz
import difflib


class Automation:
    def __init__(self):
        self.Name = "name"

    def run(self, **kwargs):
        return

    @staticmethod
    def check_directory(directory, path_to_check):
        """ Checks if the folder/file exists in the directory"""
        for folder in os.listdir(directory):
            if folder.startswith(path_to_check):
                return folder
        return False

    @staticmethod
    def fuzzy_check_directory(directory, path_to_check):
        """ Checks if the folder/file exists in the directory,
            if there are multiple it asks which to use,
            if not exact match it gives an option to use another close one"""

        options = []
        for folder in os.listdir(directory):
            if folder.startswith(path_to_check):
                options.append(folder)

        # Exact Match
        if len(options) == 1:
            return options[0]

        # No Matches
        if len(options) == 0:
            close_matches = difflib.get_close_matches(path_to_check, os.listdir(directory))

            if len(close_matches) == 0:
                return False
            print("################### FUZZY MATCH INPUT NEEDED ###################")
            for idx, c_match in enumerate(close_matches):
                print(f"{idx + 1}: {c_match}")
            response = int(input("No exact Match, would you like to select one of the 3 above options? \n"
                                 "  0 - None (stop the bot)\n"
                                 "  1 - First match\n"
                                 "  2 - Second match\n"
                                 "  3 - Third match\n"
                                 "Enter Here:  "))
            print("#################################################################")
            if response == 0:
                return False
            else:
                return close_matches[response - 1]

        # Multiple Matches
        if len(options) > 0:
            for idx, o_match in enumerate(options):
                print(f"{idx + 1}: {o_match}")
            response = int(input("Multiple Options, would you like to select one of the above options? \n"
                                 "0 - None (stop the bot)\n"
                                 "1 - First match\n"
                                 "2 - Second match\n"
                                 "3 - Third match etc.\n"
                                 "Enter Here:  "))
            if response == 0:
                return False
            else:
                return options[response - 1]

        return False

    @staticmethod
    def create_directory(directory):
        Path(directory).mkdir(parents=True, exist_ok=True)

    def get_matching_files(self, matching_string):

        # file_1 = file_name + f"_1-Ltr_{year}.pdf"
        # file_2 = file_name + f"_2-T1_({city})_{year}.pdf"
        # file_3 = file_name + f"_3-Docs_Needing_Signature_{year}.pdf"
        # files = [file_1, file_2, file_3]

        files = []
        for file in os.listdir(self.source_path):
            if file.startswith(matching_string):
                files.append(file)

        return files

    @staticmethod
    def move_files(path, destination, files):
        for file in files:
            if os.path.isfile(f"{destination}/{file}"):
                os.remove(f"{destination}/{file}")
            # shutil.move(os.path.join(path, file), destination)
            shutil.copy(os.path.join(path, file), destination)

    @staticmethod
    def send_email(**kwargs):
        if not kwargs["send"]:
            return
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = kwargs["to_address"]
        mail.Subject = kwargs['subject']
        # mail.HTMLBody = kwargs['html_body']
        mail.Body = kwargs['body']
        for attachment in kwargs["attachment_files"]:
            mail.Attachments.Add(kwargs["attachment_file_path"] + "/" + attachment)
        mail.Send()

    @staticmethod
    def read_out_pdf_list(pdf_list):
        for idx, line in enumerate(pdf_list):
            print(f"{idx}: {line}")

    @staticmethod
    def get_pdf_as_list(path):
        with fitz.open(path) as doc:
            text = ""
            for page in doc:
                text += page.get_text()

        return text.split("\n")

    @staticmethod
    def find_first_pattern(text_list, pattern):
        pattern = re.compile(pattern)
        for idx, line in enumerate(text_list):
            if pattern.match(line):
                # print(f"{idx}: {line}")
                return idx, line
