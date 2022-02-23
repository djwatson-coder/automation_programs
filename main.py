
from settings import automation_program, run_type, selection_points
from taxbot import TaxBot
from folderdataentry import FolderDataEntry
import keymouse
import folderdataentry


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


if __name__ == '__main__':
    if  run_type == "sel_extract":
        km = keymouse.KeyMouse()
        km.click_coordinates(selection_points)
    else:
        main()


# Create access to mouse click
# Create access to keyboard
# Create ability to read in Excel files
