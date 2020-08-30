#!/usr/bin/env python
# Created by Filippo Pisello

import pandas as pd
from os import path, getcwd
from openpyxl import load_workbook
import string


# 1) Main function running the complete process
def format(dataframe, file_name, sheet_name="Sheet 1", keep_index=False,
           check_file_name=True):
    # Handles frequent user errors in file name
    if check_file_name:
        file_name = correct_file_name(file_name=file_name,
                                      required_extension=".xlsx")

    # Provides to the next function information on the required workbook/sheet
    information = workbook_sheet_exist(file_name=file_name,
                                       sheet_name=sheet_name)

    # Saves the dataframe to excel, it can add a sheet to an existing workbook
    save_df_to_excel(dataframe=dataframe, file_name=file_name,
                     sheet_name=sheet_name, workbook_sheet_exist=information,
                     index=keep_index)

    # Detect the level of depth of index and columns (multiindex, multicolumns)
    index_depth = find_multiindex_multicolumns(dataframe=dataframe)[0] * keep_index
    columns_depth = find_multiindex_multicolumns(dataframe=dataframe)[1]

    # Functions to find the different visual parts of the dataframes
    header = find_header(dataframe=dataframe, multiindex_level=index_depth,
                         multicolumns_level=columns_depth)
    index = find_index(dataframe=dataframe, multiindex_level=index_depth,
                       multicolumns_level=columns_depth)
    body = find_body(dataframe=dataframe, multiindex_level=index_depth,
                     multicolumns_level=columns_depth)
    print(header)
    print(index)
    print(body)
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

    Note that for the sake of the complete program, if both the workbook and the
    sheet exists and the workbook has only one sheet, the function returns false
    and false. This is because, later, overwriting the only sheet will be
    equivalent to directly overwriting the whole workbook.
    """
    if path.exists(getcwd() + "/" + file_name):
        workbook = load_workbook(file_name)
        if sheet_name not in workbook.sheetnames:
            return True, False, workbook
        else:
            if len(workbook.sheetnames) > 1:
                return True, True, workbook  # will overwrite just the sheet
            else:
                return False, False, None  # will replace directly the workbook
    else:
        return False, False, None


def save_df_to_excel(dataframe, file_name, sheet_name, workbook_sheet_exist,
                     index):
    """
    Saves pandas dataframe to excel. Creates new notebooks or appends to old ones.

    ---------------
    Three outcomes:
    - If no workbook exists under the provided name:
        1) A new workbook is created
    - If it exists and if it already exists a sheet with the name given:
        2) The sheet/workbook gets overwritten
    - If it exists and no sheet with the name given exists:
        3) A new sheet is appended to the workbook
    """
    workbook_exist, sheet_exist, workbook = workbook_sheet_exist
    if not workbook_exist:
        dataframe.to_excel(file_name, sheet_name=sheet_name, index=index)
    else:
        if sheet_exist:
            name = workbook.get_sheet_by_name(sheet_name)
            workbook.remove_sheet(name)
            workbook.save(file_name)
        with pd.ExcelWriter(file_name, mode="a") as writer:
            dataframe.to_excel(writer, sheet_name=sheet_name, index=index)
    return


def find_multiindex_multicolumns(dataframe):
    """
    Returns number of multilevels for index and columns of a pandas dataframe.

    ---------------
    Takes a pandas dataframe and returns a list containing respectively the
    number of levels of the multiindex and of the multicolumns. If the index is
    simple the associated dimension is 1. Same is true for columns.
    """
    dimensions_index = []
    for dimension in [dataframe.index, dataframe.columns]:
        if isinstance(dimension, pd.MultiIndex):
            dimensions_index.append(len(dimension[0]))
        else:
            dimensions_index.append(1)
    return dimensions_index


def find_header(dataframe, multiindex_level, multicolumns_level):
    dim_index, dim_columns = dataframe.shape

    header = []    # A list containing pairs in style "A1, B1, ..."
    for number in range(multicolumns_level):
        for letter in range(dim_columns):
            header.append(string.ascii_uppercase[multiindex_level + letter]
                          + str(1 + number))
    return header


def find_index(dataframe, multiindex_level, multicolumns_level):
    dim_index, dim_columns = dataframe.shape

    index = []
    for letter in range(multiindex_level):
        for number in range(dim_index):
            index.append(string.ascii_uppercase[letter]
                         + str(multicolumns_level + 1 + number))
    return index


def find_body(dataframe, multiindex_level, multicolumns_level):
    dim_index, dim_columns = dataframe.shape

    body = []
    for number in range(dim_index):
        for letter in range(dim_columns):
            body.append(string.ascii_uppercase[multiindex_level + letter]
                        + str(multicolumns_level + 1 + number))
    return body
