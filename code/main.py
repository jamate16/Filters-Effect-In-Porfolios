import openpyxl as opx
import os

import pandas as pd
from dataclasses import dataclass
from abc import ABC, abstractmethod
from enum import Enum


@dataclass
class MetricOfCompany:
    company_name: str
    company_ticker: str
    metric_name: str
    metric_data: pd.Series

    def __repr__(self):
        data_status = "Data has been extracted." if not self.metric_data.empty else "Failed to extrac data."
        return f"Company: {self.company_name}, metric name: {self.metric_name}. {data_status}\n"


class OrderOfDate(Enum):
    ASCENDING = 1
    DESCENDING = 2

@dataclass
class FileStructureDetails:
    company_name_cell: str
    order_of_date: OrderOfDate
    metric_sheet_name: str = ""
    metric_row_name: str = ""


class FrequencyOfData(Enum):
    ANNUAL = 1
    QUARTERLY = 2

class FileStyle(Enum):
    A = 1
    B = 2

class IMetricExtractor(ABC):
    def __init__(self, workbook : opx.Workbook, file_structure_details : FileStructureDetails):
        self.workbook = workbook
        self.file_structure_details = file_structure_details

    @abstractmethod
    def get_company_name(self):
        pass

    @abstractmethod
    def get_company_ticker(self):
        pass

    @abstractmethod
    def get_metric_name(self):
        pass

    @abstractmethod
    def get_metric_data(self):
        pass

class MetricExtractorFileStyleA(IMetricExtractor):
    def get_company_name(self):
        pass

    def get_company_ticker(self):
        pass

    def get_metric_name(self):
        pass

    def get_metric_data(self):
        pass

class MetricExtractorFileStyleB(IMetricExtractor):
    def get_company_name(self):
        pass

    def get_company_ticker(self):
        pass

    def get_metric_name(self):
        pass

    def get_metric_data(self):
        pass

class PerMetricExtractor():
    def __init__(self, metric_extractors : dict, data_folder : str):
        self.metric_extractors = metric_extractors
        self.data_folder = data_folder
        self.file_names = os.listdir(data_folder)

    def extract(self, data_frequency : FrequencyOfData):
        for file in self.file_names:
            # TODO: Open the right workbook (based on the frequency)

            # TODO: Get type of file with determine_file_style()

            # TODO: Select the extractor based on the FileStyle

            # TODO: Field by field extract the data for to instantiate MetricOfCompany()

            # TODO: Instantiate MetricOfCompany()

            # TODO: Return MetricOfCompany
            
            pass
        

def determine_file_style(workbook : opx.Workbook) -> FileStyle:
    if "":
        return FileStyle.A
    elif "":
        return FileStyle.B
    else:
        return None  


class DataExtractorConfig:

    def __init__(self, files_path: str, file_a_config: FileStructureDetails, file_b_config: FileStructureDetails):
        self.files_path = files_path
        self.file_a = file_a_config
        self.file_b = file_b_config

    def set_up_file_types(self, sheet_a: str, row_a: str, sheet_b: str, row_b: str, frecuency_of_data: FrequencyOfData = FrequencyOfData.QUARTERLY):
        self.file_a.metric_sheet_name = sheet_a
        self.file_a.metric_row_name = row_a
        self.file_b.metric_sheet_name = sheet_b
        self.file_b.metric_row_name = row_b
        self.frecuency_of_data = frecuency_of_data

    def fully_set_up(self) -> bool:
        return self.file_a.metric_sheet_name != "" and self.file_a.metric_row_name != "" and self.file_b.metric_sheet_name != "" and self.file_b.metric_row_name != ""


class DataExtractor:
    def __init__(self, data_extractor_config: DataExtractorConfig):
        if not data_extractor_config.fully_set_up():
            raise Exception("Data extractor is not fully configured")
        self.config = data_extractor_config
        self.list_of_files = os.listdir(self.config.files_path)

    def extract(self):
        for file in self.list_of_files:
            [company_ticker, frequency, *_] = list(map(str.upper, file.split(".")[0].split("_"))) # Remove the extension of the file, get only the name of the company and the frequency of the data in caps and discard the rest

            if frequency not in FrequencyOfData._member_names_:
                continue
            if FrequencyOfData[frequency] != self.config.frecuency_of_data:
                continue
            
            workbook = opx.load_workbook(os.path.join(self.config.files_path, file))
            metric = self.__extract_metric_of_company(workbook)

            # workbook = openpyxl.load_workbook(file)
            # metric_of_company = self.__extract_metric_of_company(workbook)
            # print(metric_of_company)

    def __extract_metric_of_company(self, workbook: opx.Workbook) -> MetricOfCompany:
        worksheet = None
        file_config = None
        try:
            worksheet = workbook[self.config.file_b.metric_sheet_name] # The sheet names don't need to get stripped
            file_config = self.config.file_b
        except:
            try:
                sheet_with_metric = None

                for sheet_name in workbook.sheetnames:
                    if self.config.file_a.metric_sheet_name in workbook[sheet_name]["A3"].value: # Long company names reach the sheet's name character limit, but, A3 cell always has the type of sheet
                        sheet_with_metric = sheet_name
                        break

                if not sheet_with_metric:
                    raise Exception(f"{self.config.file_a.metric_sheet_name} not found in {workbook.worksheets}. Check if sheet name for style b was correctly typed.")
            except:
                raise Exception("Sheet style not recognized.")

            worksheet = workbook[sheet_with_metric]
            file_config = self.config.file_a

        company_name = workbook[self.config.file_a.company_name_cell].value

        return None
        

        # company_name = workbook[self.config.file_a.company_name_cell].value
        # company_ticker = workbook[self.config.file_b.company_name_cell].value
        # metric_name = workbook[self.config.file_a.sheet_name_of_metric][self.config.file_a.row_name_of_metric].value
        # metric_data = self.__extract_metric_data(workbook)
        # return MetricOfCompany(company_name, company_ticker, metric_name, metric_data)


def main():
    file_type_a_config = FileStructureDetails("A1", OrderOfDate.DESCENDING)
    file_type_b_config = FileStructureDetails("B2", OrderOfDate.ASCENDING)

    data_extractor_config = DataExtractorConfig("companies_data", file_type_a_config, file_type_b_config)
    data_extractor_config.set_up_file_types("Ratios - Key Metric", "Pretax ROA", "Financial Summary", "Pretax ROA - %, TTM", FrequencyOfData.QUARTERLY)

    data_extractor = DataExtractor(data_extractor_config)

    data_extractor.extract()


if __name__ == "__main__":
    main()
