
""" Utils for file creating, and reading operations """

import pandas as pd
import win32com.client


def create_text_file(destination_path, name, **kwargs):
    """ Creates a text file at the destination_path with the name and kwargs as the text"""

    with open(f'{destination_path}/{name}.txt', 'w') as f:
        if kwargs["sin"]:
            f.write(f'SIN: {kwargs["sin"]} \n')
        if kwargs["path"]:
            f.write(f'Folder Path: {kwargs["path"]} \n')
        f.close()


def create_gen_tf(destination_path, name, body=""):
    with open(f'{destination_path}/{name}.txt', 'w') as f:
        f.write(body)



def read_csv_file(file_path, keep_cols):
    csv_file = pd.read_csv(file_path)
    return csv_file[keep_cols]


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
