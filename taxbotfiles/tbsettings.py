
# FILE PATHS ---
source_path_tor = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder"
source_path_van = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder 2"
# source_path_tor = "//tor-fs01/TAXPREP/Toronto/T1/T1-2021/PDF"
# source_path_van = "//van-fs01/TAXPREP/Vancouver/T1/T1-2021/PDF"

completed_folder_tor = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder/bot-completed"
issues_folder_tor = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder/bot-issues"
completed_folder_van = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder 2/bot-completed"
issues_folder_van = "Q:/Admin - Digital Technology and Risk Advisory/Tax Bot/Tax Test/Test Folder 2/bot-issues"

# completed_folder_tor = "//tor-fs01/TAXPREP/Toronto/T1/T1-2021/PDF/bot-completed"
# issues_folder_van = "//tor-fs01/TAXPREP/Toronto/T1/T1-2021/PDF/bot-issues"
# completed_folder_van = "//tor-fs01/TAXPREP/Toronto/T1/T1-2021/PDF/bot-completed"
# issues_folder_van = "//tor-fs01/TAXPREP/Toronto/T1/T1-2021/PDF/bot-issues"

toronto_personal_client_dir = "I:/Toronto/_Personal Tax - TOR"
vancouver_personal_client_dir = "I:/Vancouver/_Personal Tax - VAN"


completed_entities = []  # stored in an Excel file after each iteration

tax_prep_string = "00_"
output_folder = "_FINAL T1 DOCS"
year = 2021

slp = 30  # length of time to sleep for
remove_files = False  # Remove the files from the source directory
take_first_selection = True  # Takes the first folder selection instead of asking for user input

# Email Settings --
enable_emailing = False
email_saving = True
email_partner = False
email_string = "@bakertillywm.com"

# Email contents --
email_subject = ""  # Implement this
email_file = open(f"Email_Contents.txt")
email_contents = email_file.read()
email_file.close()

# Log-File
log_string = "Bot_Log.txt"

# ToDo Fix the year in the code before deployment
