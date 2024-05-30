import yaml
import openpyxl
import pandas as pd
import re
import json
import sys
import copy
import shutil
import utils

class ExcelDecorator():

    @classmethod
    def decoration(cls, config, dest):
        
        # Hyperlinkなどの装飾処理
        with open(config, 'r') as ymlf:
            deco = yaml.safe_load(ymlf)

        wb = openpyxl.load_workbook(dest)

        for hyperlink in deco["hyperlink"]:

            from_sheet_info = hyperlink["from"]["sheet"]
            from_address = hyperlink["from"]["address"]
            to_sheet_info = hyperlink["to"]["sheet"]
            to_address = hyperlink["to"]["address"]

            for r in utils.a1_to_range(from_address):

                cell = wb[from_sheet_info["text"]][r["address"]]

                if to_sheet_info["text"] == "$.":
                    to_sheet = cell.value


                if to_sheet in wb.sheetnames:

                    cell.hyperlink = f"#{to_sheet}!{to_address}"
                    cell.font = openpyxl.styles.Font(color="0000FF", underline="single")

                    if "direction" in hyperlink:
                        direction = hyperlink["direction"]
                        if direction == "both":
                            
                            cell = wb[to_sheet][to_address]
                            cell.hyperlink = f"#{from_sheet_info['text']}!{r['address']}"
                            cell.font = openpyxl.styles.Font(color="0000FF", underline="single")

        wb.save(dest)
        wb.close()



if __name__ == "__main__":
    args = sys.argv

    src = args[1]
    config = args[2]

    ExcelDecorator.decoration(config, src)



