#!/usr/bin/env python
# Created by Filippo Pisello

import pandas as pd
from os import path, getcwd
from openpyxl import load_workbook


# 1) Main function running the complete process
def format(dataframe, file_name, sheet_name="Sheet 1", check_file_name=True):
    if check_file_name:
        file_name = correct_file_name(file_name=file_name,
                                      required_extension=".xlsx")
    pass


# 2) Building blocks
def correct_file_name(file_name, required_extension):
    """
    Corrects two recurring issues with user-provided file names.

    ---------------
    The two issues are:
    - File name not provided as string (Foo.xlsx vs "Foo.xlsx")
    - File name without file extension ("Foo" vs "Foo.xlsx")
    """
    file_name = str(file_name)  # for cases where user forgot commas
    if not file_name.endswith(required_extension):  # "" forgot file's extension
        file_name = file_name + required_extension
        print(f"The file name provided was modified to {file_name} to avoid ")
        print("errors in the program execution. To deactivate this ")
        print("autocorrection turn to 'False' the default arg 'check_file_name'.")
    return(file_name)


def workbook_sheet_exist(file_name, sheet_name):
    """
    Checks if in the directory exist a workbook and sheet under given names.

    ---------------
    It returns a pair of booleans. The first tells if a workbook called as the
    provided file name exists. The second tells if inside the mentioned workbook
    is present a sheet with the given name.
    """
    if path.exists(getcwd + "/" + file_name):
        workbook = load_workbook(file_name)
        if sheet_name in workbook.sheetnames:
            return True, True, workbook  # both workbook and sheet exist
        else:
            return True, False, workbook
    else:
        return False, False, workbook


def save_df_to_excel(dataframe, file_name, sheet_name, workbook_sheet_exist):
    """
    Saves pandas dataframe to excel. Creates new notebooks or appends to old ones.

    ---------------
    Three outcomes:
    - If no workbook exists under the provided name:
        1) A new workbook is created
    - If it exists and if it already exists a sheet with the name given:
        2) The sheet gets overwritten
    - If it exists and no sheet with the name given exists:
        3) A new sheet is appended to the workbook
    """
    workbook_exist, sheet_exist, workbook = workbook_sheet_exist
    if not workbook_exist:
        dataframe.to_excel(file_name, sheet_name=sheet_name)
    else:
        if sheet_exist: workbook.remove_sheet(sheet_name)
        with pd.ExcelWriter(file_name, mode="a") as writer:
            dataframe.to_excel(writer, sheet_name=sheet_name)
    return
