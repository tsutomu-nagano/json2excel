
import openpyxl
import re
import pandas as pd
from itertools import product


def around_ranges(ranges):
    rows = [r["row"] for r in ranges]
    cols = [r["col"] for r in ranges]

    row_min = min(rows)
    col_min = min(cols)
    row_max = max(rows)
    col_max = max(cols)

    ret = {}
    ret["left"] = [r for r in ranges if r["col"] == col_min]
    ret["right"] = [r for r in ranges if r["col"] == col_max]
    ret["top"] = [r for r in ranges if r["row"] == row_min]
    ret["bottom"] = [r for r in ranges if r["row"] == row_max]

    return(ret)


def convert_auto_range(a1, src, sheet_name):

    ptn = "(?P<FROM>.+):\$auto(\-(?P<DIRECTION>.+))?"
    addr = re.match(ptn, a1)

    if addr:
        addr_from = addr["FROM"]
        addr_direction = addr["DIRECTION"]

        col_str, row = openpyxl.utils.cell.coordinate_from_string(addr_from)
        col = openpyxl.utils.cell.column_index_from_string(col_str)

        df = pd.read_excel(src, sheet_name = sheet_name, skiprows = row - 1, header = None).dropna(axis=1, how='all')

        col_add = 1 if addr_direction == "row" else len(df.columns.tolist()) 
        row_add = 1 if addr_direction == "col" else len(df)

        addr_to = f"{openpyxl.utils.get_column_letter(col + col_add - 1)}{row + row_add - 1}"

        ret = f"{addr_from}:{addr_to}"

    else:
        ret = a1

    return(ret)


def a1_to_range_core(a1):
    col_str, row = openpyxl.utils.cell.coordinate_from_string(a1)
    col = openpyxl.utils.cell.column_index_from_string(col_str)
    return {"row":row, "col":col, "address": a1}

def a1_to_range(a1):


    a1_ = str(a1).split(":")

    if len(a1_) == 1:
        r1c1_ = [a1_to_range_core(a1)]
    else:
        r1c1_ = [a1_to_range_core(a1) for a1 in a1_]

        rows = [r1c1_[0]["row"], r1c1_[1]["row"]]
        cols = [r1c1_[0]["col"], r1c1_[1]["col"]]

        row_min = min(rows)
        row_max = max(rows)

        col_min = min(cols)
        col_max = max(cols)

        r1c1_ = ([{"row":index[0], "col":index[1], "address":f"{openpyxl.utils.get_column_letter(index[1])}{index[0]}"} for index in list(product(
            list(range(row_min, row_max + 1)),
            list(range(col_min, col_max + 1))
            ))])

    return(r1c1_)
