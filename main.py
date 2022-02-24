
from settings import automation_program, run_type
from taxbot import TaxBot
from folderdataentry import FolderDataEntry
from tbot import TaxBotTesting
import pickle


# Select the automation_Script to Run
def main():
    if automation_program == "FolderDataEntry":
        bot = FolderDataEntry()
        points = {'home': (-1388, 1455), 'selection1': (-1041, 558), 'selection2': (-1041, 558)}
        bot.run(points=points)

    if automation_program == "TaxBot":
        bot = TaxBot()
        bot.run(False)

    return


if __name__ == '__main__':
    if run_type == "testing":
        testing = TaxBotTesting()
        testing.test_something()
    else:
        #outfile = open("sample.pkl", 'wb')
        #pickle.dump([], outfile)
        #outfile.close()
        main()



# Create access to mouse click
# Create access to keyboard
# Create ability to read in Excel files
