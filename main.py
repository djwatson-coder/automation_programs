
import settings
import keymouse
import folderdataentry


# Select the automation_Script to Run
def main():

    if settings.automation_program == "FolderDataEntry":
        bot = folderdataentry.FolderDataEntry()
        points = {'home': (-1388, 1455), 'selection1': (-1041, 558), 'selection2': (-1041, 558)}
        bot.run(points=points)

    return


if __name__ == '__main__':
    if settings.run_type == "sel_extract":
        km = keymouse.KeyMouse()
        km.click_coordinates(settings.selection_points)
    else:
        main()


# Create access to mouse click
# Create access to keyboard
# Create ability to read in Excel files
