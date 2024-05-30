
import openpyxl
from itertools import product



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
