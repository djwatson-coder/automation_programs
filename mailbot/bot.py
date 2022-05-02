""" Bot file that controls that mail bot """

from automation import Automation
import pandas as pd
from datetime import datetime, timedelta
from utils import filehandler as fh
from utils import regextools as rt

class MailBot(Automation):
    def __init__(self):

        super(MailBot, self).__init__()
        self.original_df_name = "12.xlsx"
        self.created_df_name = "Partner_Client_Information.xlsx"

        self.admin_email = "admin@bakertillywm.com"

        # Date

    def create_partner_data(self):

        df = self.read_cch_data(self.original_df_name)
        df = self.select_relevant_info(df)
        df = self.update_partner_row(df)
        df = self.remove_data(df)
        df = self.clean_data(df)

        fh.write_excel(self.created_df_name, df)

    def read_cch_data(self, path: str):
        df = pd.read_excel(path, skiprows=5)
        df.columns = df.columns.str.replace(':', '')
        df.columns = df.columns.str.replace(' ', '_')
        df = df.rename(columns={"Unnamed_0": "Partner"})

        return df

    def select_relevant_info(self, df):
        # Get the relevant information
        df1 = df[df['Partner'].str.contains('Primary Partner: ', na=False)]
        df2 = df[df['Partner'].isin(['C'])]

        df = pd.concat([df1, df2]).sort_index()
        df = df.reset_index()

        return df

    def update_partner_row(self, df):
        # Update the Partner for all rows
        for i, row in df.iterrows():
            val = row['Partner']
            if val == "C":
                val = df.at[i - 1, 'Partner']
            df.at[i, 'Partner'] = val

        return df

    def remove_data(self, df):
        # Remove the Nas in Client Name
        df = df.dropna(subset=['Client_Name'])
        df = df[['Partner', 'Client_Name']]

        return df

    def clean_data(self, df):
        # Remove extra details - Remove form partner: PP, Brackets, commas
        df["Partner"] = df["Partner"].str.replace('Primary Partner: ', '')
        df['Partner'] = df['Partner'].str.replace(r"\(.*\)", "")
        df["Partner"] = df["Partner"].str.replace(',', '')
        df["Partner"] = df["Partner"].str.strip()

        # Remove the non-partner names
        df = df[~df['Partner'].isin(['No Selection', 'zAdministrator'])]

        # Format the Columns
        df["Partner"] = df["Partner"].str.strip()
        df[['Last_Name', 'First_Name']] = df['Partner'].str.split(' ', 1, expand=True)
        df["Last_Name"] = df["Last_Name"].str.replace(' ', '').str.lower()
        df["First_Name"] = df["First_Name"].str.replace(' ', '').str.lower()
        df["Email"] = df["First_Name"] + "." + df["Last_Name"] + "@bakertillywm.com"
        df = df[['Partner', "Email", 'Client_Name']]

        return df

    def get_mail_date(self, text: str):

        mail_date = rt.find_all('\(\d{4}-\d{2}-\d{2} ', text)[0]
        return mail_date.strip().replace("(", "")

    def get_date(self, today_minus=0):
        return (datetime.now() - timedelta(today_minus)).strftime('%Y-%m-%d')

    def get_company_info(self, info: list, break_folder="_INBOX"):

        company = info[info.index(break_folder) - 1]
        folder = "\\" + '\\'.join(info[1:info.index(break_folder) + 1])
        return company, folder

    def get_textfile_info(self, path: str):
        mail = fh.read_text_file(path)
        today_date = self.get_date(1)  # yesterday atm
        info = []
        for mail_line in mail:
            if self.get_mail_date(mail_line) == today_date:
                mail_info = mail_line[0].split('\\')
                company, folder = self.get_company_info(mail_info)
                info.append({"Company": company,
                             "Folder": folder})
        return info

    def get_company_partner(self, company, partner_info):
        for i, row in partner_info.iterrows():
            val = row['Client_Name']
            if val == "Vancouver Cycling Without Age":  # company
                return row['Email']

        return False

    def match_text_file(self, text_file_path: str):
        relevant_info = self.get_textfile_info(text_file_path)
        partner_info = pd.read_excel(self.created_df_name)
        for company_info in relevant_info:
            partner_email = self.get_company_partner(company_info["Company"], partner_info)
            if not partner_email:
                partner_email = self.admin_email

