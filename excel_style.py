# Created by Filippo Pisello

from openpyxl.styles import Font, Alignment, PatternFill

class ExcelStyle:
    """
    Class to collect all the formatting options for the style of an Excel cell

    Arguments
    ----------------
    fill_color: str, default=None
        Fill color of the cells. If None, no fill color is applied.
    font_color: str, default="000000"
        Font color of the cells. Default color is black.
    font_size: int, default=11
        Size of the cell font.
    font_bold: Bool, default=False
        If True cell text is bold.
    alignment: str, default="center"
        Horizontal alignment of the text content. It can be either "center",
    "right" or "left".
    """
    def __init__(self, fill_color=None, font_color="000000", font_size=11,
                 font_bold=False, alignment="center"):
        self.fill_color = fill_color
        self.font_color = font_color
        self.font_size = font_size
        self.font_bold = font_bold
        self.h_alignment = alignment

    def font(self):
        if self.font_color=="000000" and self.font_size==11 and not self.font_bold:
            return None
        return Font(color=self.font_color, size=self.font_size, bold=self.font_bold)

    def alignment(self):
        return Alignment(horizontal=self.h_alignment, vertical="center")

    def fill(self):
        if self.fill_color is None:
            return None
        return PatternFill(fill_type="solid", start_color=self.fill_color,
                           end_color=self.fill_color)
