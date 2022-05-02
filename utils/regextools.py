
""" Utils for regex based problems """

import re


def find_first_pattern(text_list, pattern):
    pattern = re.compile(pattern)
    for idx, line in enumerate(text_list):
        if pattern.match(line):
            # print(f"{idx}: {line}")
            return idx, line
    return 0, False


def check_for_instance(check, value, default):
    if check == 0:
        return default
    return value


def find_all(pattern, text):
    return re.findall(pattern, text)
