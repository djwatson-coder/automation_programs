
# Could be a subclass of the class automation
# Runs the automation project
from time import sleep
from automation import Automation
from dataentry import keymouse
import webbrowser


class FolderDataEntry(Automation):
    def __init__(self):

        super().__init__()
        self.Name = "name"
        self.km = keymouse.KeyMouse()

    def run(self, **kwargs):
        home = kwargs["points"]["home"]
        # self.km.mouse_click(home)
        # sleep(1)
        webbrowser.open('http://google.com', new=2)
        sleep(1)
        self.km.write_text("Python")
        self.km.press_key("enter")

        return
