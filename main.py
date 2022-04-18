
from settings import automation_program
from taxbot.bot import TaxBot
from dataentry.folderdataentry import FolderDataEntry
import sys


# Select the automation_Script to Run
def main(**kwargs):
    if automation_program == "FolderDataEntry":
        bot = FolderDataEntry()
        points = {'home': (-1388, 1455), 'selection1': (-1041, 558), 'selection2': (-1041, 558)}
        bot.run(points=points)

    if automation_program == "TaxBot":
        bot = TaxBot(kwargs["location"])
        bot.run()

    return


if __name__ == '__main__':

    if len(sys.argv) > 1:
        main(location=sys.argv[1])
    else:
        main(location="VAN")
