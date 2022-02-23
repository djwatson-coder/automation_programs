import os
from pathlib import Path
import shutil
import win32com.client


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
    def create_directory(directory):
        Path(directory).mkdir(parents=True, exist_ok=True)

    def get_matching_files(self, matching_string):
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

        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = kwargs["to_address"]
        mail.Subject = kwargs['subject']
        # mail.HTMLBody = kwargs['html_body']
        mail.Body = kwargs['body']
        for attachment in kwargs["attachment_files"]:
            mail.Attachments.Add(kwargs["attachment_file_path"] + "/" + attachment)
        mail.Send()
