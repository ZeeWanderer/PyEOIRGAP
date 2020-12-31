import pprint

prettyp = pprint.PrettyPrinter(indent=4)

from enum import Enum

class UserPreset(Enum):
    max = 0
    sergey = 1

def get_mul_sum_string(first: list, second: list):
    str_ = f"{' + '.join(f'({x.__round__(7)} * {y.__round__(7)})' for x, y in zip(first, second))}"
    return str_


def get_sum_string(list_: list):
    str_ = f"{' + '.join(str(x.__round__(10)) for x in list_)}"
    return str_