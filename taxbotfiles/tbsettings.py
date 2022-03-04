
# File Paths
source_path = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder"
# self.source_path = "L:\Toronto\T1\T1-2020\PDF"
manual_to_do_path = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/" \
                         "Test Folder/Manual_to_do"

toronto_personal_client_dir = "I:/Toronto/_Personal Tax - TOR"
vancouver_personal_client_dir = "I:/Vancouver/_Personal Tax - VAN"


completed_entities = []  # stored in an Excel file after each iteration
store = False

tax_prep_string = "00_"
output_folder = "_FINAL T1 DOCS"
year = 2021

slp = 30  # length of time to sleep for
remove_files = True  # Remove the files from the source directory


# Email Settings
enable_emailing = True
email_partner = False
email_string = "@bakertillywm.com"
