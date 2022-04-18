
""" Utils for OS based operations """

import os
from pathlib import Path
import shutil
import difflib


def check_directory(directory, path_to_check):
    """ Checks if the folder/file exists in the directory"""
    for folder in os.listdir(directory):
        if folder.startswith(path_to_check):
            return folder
    return False


def fuzzy_check_directory(directory, path_to_check, extra_info=""):
    """ Checks if the folder/file exists in the directory,
        if there are multiple possible files it asks which to use,
        if not exact match it gives an option to use another close one"""

    options = []
    for folder in os.listdir(directory):
        if folder.startswith(path_to_check):
            options.append(f"{extra_info}{folder}")

    close_matches = difflib.get_close_matches(path_to_check, os.listdir(directory))
    for idx in range(len(close_matches)):
        close_matches[idx] = f"{extra_info}{close_matches[idx]}"

    return options, close_matches


def create_directory(directory):
    """ Creates a directory if one does not exist"""
    Path(directory).mkdir(parents=True, exist_ok=True)


def get_matching_files(path, matching_strings):
    """ Gets all the files that match any of a list of strings"""
    files = []
    for file in os.listdir(path):
        for string in matching_strings:
            if string in file:
                files.append(file)

    return files


def move_files(path, destination, files, remove=False, new_name=False):
    """ moves or copies files to a destination path"""
    for file in files:
        if os.path.isfile(f"{destination}/{file}"):
            os.remove(f"{destination}/{file}")
        if new_name:
            destination = destination + "/" + new_name
        if remove:
            shutil.move(os.path.join(path, file), destination)
        else:
            shutil.copy(os.path.join(path, file), destination)

