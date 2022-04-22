
# FILE PATHS ---
client_info_file = "//tor-fs01/TAXPREP/Vancouver/T1/T1-2021/PDF/David W/T1-Clients-All-Offices.csv"
archive_folder_van = "//van-fs01/TAXPREP/Vancouver/T1/T1-2021/Archive"
archive_folder_tor = "//van-fs01/TAXPREP/Toronto/T1/T1-2021/Archive"

source_path_tor = "//tor-fs01/TAXPREP/Toronto/T1/T1-2021/PDF/BOT"
source_path_van = "//van-fs01/TAXPREP/Vancouver/T1/T1-2021/PDF/BOT"

completed_folder_tor = source_path_tor + "/BOT-completed"
issues_folder_tor = source_path_tor + "/BOT-issues"
completed_folder_van = source_path_van + "/BOT-completed"
issues_folder_van = source_path_van + "/BOT-issues"

toronto_personal_client_dir = "I:/Toronto/_Personal Tax - TOR"
vancouver_personal_client_dir = "I:/Vancouver/_Personal Tax - VAN"

completed_entities = []  # stored in an Excel file after each iteration

tax_prep_string = "_00-ID"
output_folder = "_FINAL T1 DOCS"
yr = 2021

slp = 15  # length of time to sleep for
remove_files = True  # Remove the files from the source directory
take_first_selection = True  # Takes the first folder selection instead of asking for user input

# Email Settings --
enable_emailing = False
email_saving = True
email_partner = False
email_string = "@bakertillywm.com"

# Email contents --
email_subject = ""  # Implement this
email_file = open(f"files/email.txt")
email_contents = email_file.read()
email_file.close()

# ToDo Fix the year in the code before deployment
