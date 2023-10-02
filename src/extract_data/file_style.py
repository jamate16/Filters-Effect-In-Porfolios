from enum import Enum
from dataclasses import dataclass

import openpyxl as opx


class FileStyle(Enum):
    A = "StyleA"
    B = "StyleB"


@dataclass
class FileStyleDetails:
    """
    Dataclass that stores the locations, names, and any other attribute that characterizes a sheet style
    
    Attributes:
        company_name_cell (str): The cell coordinate containing the company name.
        metric_sheet_name (str, optional): The name of the sheet where the metric is located (default is an empty string).
        metric_row_name (str, optional): The substring to search in the index of the rows of metric_sheet_name (default is an empty string).
        sheet_name_base (str): For comparison with sheet_name_compare. Cell coord where the sheet name is located.
        sheet_name_compare (str): For comparison with sheet_name_base. Cell coord where the sheet name is repeated.
    """
    metric_name: str
    metric_sheet_name: str
    company_name_cell: str
    company_name_separator: str
    date_format: str
    sheet_name_base: str
    sheet_name_compare: str
    timestamp_coord: str


# TODO: Implement configurations as a dictionary, not as members
class FileStyleDetailsFactory:
    def __init__(self, date_format: str, company_name_cell: str, company_name_separator: str, sheet_name_base: str, sheet_name_compare: str, timestamp_coord: str):
        self.date_format = date_format
        self.company_name_cell = company_name_cell
        self.company_name_separator = company_name_separator
        self.sheet_name_base = sheet_name_base
        self.sheet_name_compare = sheet_name_compare
        self.timestamp_coord = timestamp_coord

    """ Create FileStyleDetails object with the base values and the two parameters.
    Those two parameters are the ones that change the most when creating a new FileStyleDetails.
      """
    def create(self, metric_name: str, metric_sheet_name: str) -> FileStyleDetails:
        return FileStyleDetails(
            metric_name, metric_sheet_name,
            self.company_name_cell,
            self.company_name_separator,
            self.date_format,
            self.sheet_name_base,
            self.sheet_name_compare,
            self.timestamp_coord
        )


class FileStyleManager:
    def __init__(self, styles : dict):
        self.styles = styles

    def determine_file_style(self, workbook : opx.Workbook) -> FileStyle:
        """
        Compares the contents of two cells (different cells for every file style) that contain the same substring

        Returns:
            FileStyle: If the comparison for the corresponding file style returns true
        """
        active_sheet = workbook.active
        
        for style_name, style in self.styles.items():
            sheet_name_base = active_sheet[style.sheet_name_base].value
            sheet_name_compare = active_sheet[style.sheet_name_compare].value.split("-") #
            sheet_name_compare = sheet_name_compare[0].split("\xa0\xa0")[0].strip() # In style A there a bunch of "\xa0" characters

            if sheet_name_compare in sheet_name_base:
                return style_name
            
        raise Exception("Sheet style not recognized.")
