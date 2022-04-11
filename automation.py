import math
import os
import re
from pathlib import Path
import shutil
import win32com.client
import fitz
import difflib
from PyPDF2 import PdfFileReader
from pikepdf import Pdf


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
    def fuzzy_check_directory(directory, path_to_check, extra_info=""):
        """ Checks if the folder/file exists in the directory,
            if there are multiple possible files it asks which to use,
            if not exact match it gives an option to use another close one"""

        options = []
        for folder in os.listdir(directory):
            if folder.startswith(path_to_check):
                options.append(f"{extra_info}{folder}")

        close_matches = difflib.get_close_matches(path_to_check, os.listdir(directory))
        for idx in range(len(close_matches)):
            close_matches[idx] = f"{extra_info}{close_matches[idx]}"

        return options, close_matches


    @staticmethod
    def create_directory(directory):
        Path(directory).mkdir(parents=True, exist_ok=True)

    def get_matching_files(self, path, matching_strings):

        files = []
        for file in os.listdir(path):
            for string in matching_strings:
                if string in file:
                    files.append(file)

        return files

    @staticmethod
    def move_files(path, destination, files, remove=False):
        for file in files:
            if os.path.isfile(f"{destination}/{file}"):
                os.remove(f"{destination}/{file}")
            if remove:
                shutil.move(os.path.join(path, file), destination)
            else:
                shutil.copy(os.path.join(path, file), destination)

    @staticmethod
    def create_email(**kwargs):

        if not (kwargs["send"] or kwargs["save"]):
            return False
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = kwargs["to_address"]
        mail.Subject = kwargs['subject']
        mail.HTMLBody = kwargs['html_body']
        # mail.Body = kwargs['body']
        for attachment in kwargs["attachment_files"]:
            mail.Attachments.Add(kwargs["attachment_file_path"] + "/" + attachment)

        if kwargs["save"]:
            mail.SaveAs(f'{kwargs["save_path"]}/{kwargs["save_name"]}.msg')

        if kwargs["send"]:
            mail.Send()
            return True
        return False


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

    @staticmethod
    def find_first_pattern(text_list, pattern):
        pattern = re.compile(pattern)
        for idx, line in enumerate(text_list):
            if pattern.match(line):
                # print(f"{idx}: {line}")
                return idx, line
        return 0, False

    @staticmethod
    def print_hash_comment(text=""):

        if text == "":
            print(80*"#")
        else:
            text_length = len(text) + 2
            hash_length = math.floor((80 - text_length) / 2)
            print(f"{hash_length * '#'} {text.upper()} {hash_length * '#'}")

    def create_text_file(self, destination_path, name, **kwargs):
        """ Creates a text file at the destination_path with the name and kwargs as the text"""

        with open(f'{destination_path}/{name}.txt', 'w') as f:
            if kwargs["sin"]:
                f.write(f'SIN: {kwargs["sin"]} \n')
            if kwargs["path"]:
                f.write(f'Folder Path: {kwargs["path"]} \n')
