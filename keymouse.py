from time import sleep
from pynput.mouse import Button, Controller
import win32api
import keyboard


class KeyMouse:
    def __init__(self):
        self.keyboard = "keyboard"
        self.mouse = Controller()
        self.kb = keyboard

    @staticmethod
    # Code to select the null point in the window
    def click_coordinates(positions):
        not_clicked = win32api.GetKeyState(0x01)  # Left button down = 0 or 1. Button up = -127 or -128
        positions_dict = {}
        for coordinate in positions:
            sleep(1)
            print("selecting click point for:", coordinate)
            listen = True
            while listen:
                mouse_click = win32api.GetKeyState(0x01)
                if mouse_click != not_clicked and mouse_click < 0:  # Button state changed
                    positions_dict[coordinate] = win32api.GetCursorPos()
                    print("Coordinate", coordinate, "at:", win32api.GetCursorPos())
                    listen = False
        print(positions_dict)
        return positions_dict

    # Mouse clicks at a specific non-important part of the screen
    def mouse_click(self, point):
        self.mouse.position = point
        self.mouse.press(Button.left)
        self.mouse.release(Button.left)

    # Uses the keyboard to write
    def write_text(self, text):
        self.kb.write(text)

    def press_key(self, key):
        self.kb.press(key)



