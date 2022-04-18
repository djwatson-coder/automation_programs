
""" Automation Base Class for all automation programs"""


import math
import sys


class Automation:
    def __init__(self, log_path='files/log.txt'):
        self.bot_log = ""
        self.log_path = log_path

    def run(self, **kwargs):
        """ Runs the Automation program"""
        return

    # Logging operations -- ToDo Potentially move to logging Util class

    def print_hash_comment(self, text=""):

        if text == "":
            self.log_info(80*"#")
        else:
            text_length = len(text) + 2
            hash_length = math.floor((80 - text_length) / 2)
            self.log_info(f"{hash_length * '#'} {text.upper()} {hash_length * '#'}")

    def log_info(self, text):
        print(text)
        self.bot_log = self.bot_log + text + "\n"
        # save to a log file

    def create_bot_log(self):
        open(self.log_path, 'a').close()
        return ""

    def write_to_log(self):
        with open(self.log_path, 'a') as f:
            f.write(self.bot_log)
            f.close()
        self.bot_log = ""

    def handle_error(self, error, exit_program=False):
        self.log_info(repr(error))
        self.write_to_log()
        if exit_program:
            sys.exit()
