# Created by Filippo Pisello
from os import path, getcwd

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill

import sys
sys.path.append(r"C:\Users\Filippo Pisello\Desktop\Python\Git Projects\Git_Spreadsheet")
from spreadsheet import Spreadsheet

class Custom_excel(Spreadsheet):
    """
    TBD
    """

    def __init__(self, dataframe: pd.DataFrame, file_name, keep_index: bool=False,
                 skip_rows: int=0, skip_columns: int=0, correct_lists: bool=False,
                 sheet_name="Sheet1", header_style="strong", index_style="light",
                 custom_width=20):
        super().__init__(dataframe, keep_index, skip_rows, skip_columns, correct_lists)
        self.file_name = file_name
        self.sheet_name = sheet_name

        self.workbook_sheet_exist = []
        self.workbook = " "
        self.sheet = " "

        self.header_style = header_style
        self.index_style = index_style

        # Strong
        self.strong_color = "0066cc"
        self.strong_font_color = "ffffff"
        self.strong_font_size = 12
        self.strong_font_bold = True
        self.strong_alignment = "center"
        # Light
        self.light_color = "b2beb5"
        self.light_font_color = "000000"
        self.light_font_size = 11
        self.light_font_bold = False
        self.light_alignment = "left"
        # Plain
        self.plain_font_size = 11
        self.plain_alignment = "center"

        self.custom_width = custom_width

    # -------------------------------------------------------------------------
    # 1) Main function running the complete process
    # -------------------------------------------------------------------------
    def to_custom_excel(self, check_file_name=True):
        # File saving process
        if check_file_name:
            self.correct_file_name()
        self.workbook_sheet_exist, self.workbook = self.check_file_existence()
        self.save_df_to_excel()

        # Table customization
        self.workbook, self.sheet = self.get_workbook_sheet()
        #Loop to apply styles to the different table's components
        components = [self.header, self.index, self.body]
        styles = [self.header_style, self.index_style, "plain"]
        for component, style in zip(components, styles):
            self.apply_style(component, style)
        # Setting columns' width
        self.adjust_all_columns_width()

        # Save file
        self.workbook.save(filename=self.file_name)
        return

    # -------------------------------------------------------------------------
    # 2) Building blocks
    # -------------------------------------------------------------------------
    def correct_file_name(self, required_extension=".xlsx"):
        """
        Adds desired extention if the user did not include it in file name.

        ---------------
        Example:
        - File name without file extension: "Foo" --> "Foo.xlsx"
        """
        if not self.file_name.endswith(required_extension):
            self.file_name = self.file_name + required_extension
            print(f"The file name provided was modified to {self.file_name} to")
            print(" avoid errors in the program execution. To deactivate this ")
            print("autocorrection turn to 'False' the default arg 'check_file_name'.")
        return

    def check_file_existence(self):
        """
        Checks if in the directory exist a workbook and sheet under given names.

        ---------------
        It returns:
        - First bool: tells if a workbook called as the provided file name exists.
        - Second bool: tells if inside the above workbook is present a sheet with the given name.
        - Workbook/None: if the workbook was found it returns it, else None.

        Note that for the sake of the complete program, if both the workbook and the
        sheet exists and the workbook has only one sheet, the function returns false
        and false. This is because, later, overwriting the only sheet will be
        equivalent to directly overwriting the whole workbook.
        """
        if path.exists(getcwd() + "/" + self.file_name):
            workbook = load_workbook(self.file_name)
            if self.sheet_name not in workbook.sheetnames:
                return [True, False], workbook
            else:
                if len(workbook.sheetnames) > 1:
                    return [True, True], workbook  # will overwrite just the sheet
                else:
                    return [False, False], None  # will replace directly the workbook
        else:
            return [False, False], None

    def save_df_to_excel(self):
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
        workbook_exist, sheet_exist = self.workbook_sheet_exist

        if not workbook_exist:
            self.df.to_excel(self.file_name, sheet_name=self.sheet_name,
                            index=self.keep_index)
        else:
            if sheet_exist: # it needs to be removed first
                name = self.workbook.get_sheet_by_name(self.sheet_name)
                self.workbook.remove_sheet(name)
                self.workbook.save(self.file_name)
            with pd.ExcelWriter(self.file_name, mode="a") as writer:
                self.df.to_excel(writer, sheet_name=self.sheet_name,
                                index=self.keep_index)
        return

    def get_workbook_sheet(self):
        """
        Returns the desired workbook and sheet objects based on their name.
        """
        workbook = load_workbook(self.file_name)
        sheet = workbook.get_sheet_by_name(self.sheet_name)
        return workbook, sheet

    def apply_style(self, cell_list, style):
        """
        TBD
        """
        if style == "strong":
            font, fill, alignment = self.strong_formatting()
        elif style == "light":
            font, fill, alignment = self.light_formatting()
        elif style == "plain" or style is None:
            font, fill, alignment = self.plain_formatting()
        else:
            print(f"As {style} is not a recognized style keyword, no action was undertaken.")
            return
        self.apply_formatting_to_cells(cell_list, font, fill, alignment)
        return

    def strong_formatting(self):
        """
        Applies formatting of the type "main" to a range of cells.

        ---------------
        The formatting style includes the font size, font boldness, font color,
        fill color, alignment.
        """
        strong_font = Font(bold=self.strong_font_bold, color=self.strong_font_color,
                           size=self.strong_font_size)
        strong_fill = PatternFill(fill_type="solid", start_color=self.strong_color,
                                  end_color=self.strong_color)
        strong_alignment = Alignment(horizontal=self.strong_alignment,
                                     vertical="center")
        return (strong_font, strong_fill, strong_alignment)

    def light_formatting(self):
        """
        Applies formatting of the type "light" to a range of cells.

        ---------------
        The formatting style includes the font size, font boldness, font color,
        fill color, alignment.
        """
        light_font = Font(bold=self.light_font_bold, color=self.light_font_color,
                          size=self.light_font_size)
        light_fill = PatternFill(fill_type="solid", start_color=self.light_color,
                                 end_color=self.light_color)
        light_alignment = Alignment(horizontal=self.light_alignment,
                                    vertical="center")
        return (light_font, light_fill, light_alignment)

    def plain_formatting(self):
        """
        Applies formatting to body cells.

        ---------------
        The formatting style includes the font size, font boldness, font color,
        fill color, alignment.
        """
        body_font = Font(size=self.plain_font_size)
        body_alignment = Alignment(horizontal=self.plain_alignment,
                                    vertical="center")
        return (body_font, None, body_alignment)

    def adjust_all_columns_width(self):
        """
        Sets the width of all the columns of the sheet to a fixed value
        """
        for value in range(self.table_coordinates[0][0], self.table_coordinates[1][0] + 1):
            self.sheet.column_dimensions[self.letter_from_index(value)].width = self.custom_width
        return

    # -------------------------------------------------------------------------
    # 3) Methods used at their times in building blocks
    # -------------------------------------------------------------------------
    def apply_formatting_to_cells(self, cells_list, font_formatting=None,
                                  fill_formatting=None, alignment_formatting=None):
        """
        Applies a pre-determined formatting if any is provided
        """
        if font_formatting is not None:
            for cell in cells_list:
                self.sheet[cell].font = font_formatting
        if fill_formatting is not None:
            for cell in cells_list:
                self.sheet[cell].fill = fill_formatting
        if alignment_formatting is not None:
            for cell in cells_list:
                self.sheet[cell].alignment = alignment_formatting
        return