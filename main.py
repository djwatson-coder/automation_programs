
from settings import automation_program, run_type
from taxbotfiles.taxbot import TaxBot
from folderdataentry import FolderDataEntry
from taxbotfiles.tbot import TaxBotTesting
from taxbotfiles.gui import GUI
import pickle


# Select the automation_Script to Run
def main():
    if automation_program == "FolderDataEntry":
        bot = FolderDataEntry()
        points = {'home': (-1388, 1455), 'selection1': (-1041, 558), 'selection2': (-1041, 558)}
        bot.run(points=points)

    if automation_program == "TaxBot":
        bot = TaxBot()
        bot.run()

    return


def create_pickle():
    outfile = open("sample.pkl", 'wb')
    pickle.dump([], outfile)
    outfile.close()


if __name__ == '__main__':

    if run_type == "testing":
        testing = TaxBotTesting()
        testing.test_something()
    else:
        # create_pickle()
        main()



# ToDo Create access to mouse click
# ToDo Create access to keyboard
# ToDo Create ability to read in Excel files
