
from settings import automation_program, run_type
from taxbotfiles.taxbot import TaxBot
from folderdataentry import FolderDataEntry

# Select the automation_Script to Run
def main(**kwargs):
    if automation_program == "FolderDataEntry":
        bot = FolderDataEntry()
        points = {'home': (-1388, 1455), 'selection1': (-1041, 558), 'selection2': (-1041, 558)}
        bot.run(points=points)

    if automation_program == "TaxBot":
        bot = TaxBot(location)
        bot.run()

    return


if __name__ == '__main__':

    location = "VAN"
    main(location=location)

