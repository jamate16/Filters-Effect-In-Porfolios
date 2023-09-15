import openpyxl as opx
import os

import pandas as pd
from dataclasses import dataclass
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
class FileConfig:
    company_name_cell: str
    order_of_date: OrderOfDate
    sheet_name_of_metric: str = ""
    row_name_of_metric: str = ""


class FrecuencyOfData(Enum):
    ANNUAL = 1
    QUARTERLY = 2


class DataExtractorConfig:

    def __init__(self, files_path: str, file_a_config: FileConfig, file_b_config: FileConfig):
        self.files_path = files_path
        self.file_a = file_a_config
        self.file_b = file_b_config

    def set_up_file_types(self, sheet_a: str, row_a: str, sheet_b: str, row_b: str, frecuency_of_data: FrecuencyOfData = FrecuencyOfData.QUARTERLY):
        self.file_a.sheet_name_of_metric = sheet_a
        self.file_a.row_name_of_metric = row_a
        self.file_b.sheet_name_of_metric = sheet_b
        self.file_b.row_name_of_metric = row_b
        self.frecuency_of_data = frecuency_of_data

    def fully_set_up(self) -> bool:
        return self.file_a.sheet_name_of_metric != "" and self.file_a.row_name_of_metric != "" and self.file_b.sheet_name_of_metric != "" and self.file_b.row_name_of_metric != ""


class DataExtractor:
    def __init__(self, data_extractor_config: DataExtractorConfig):
        if not data_extractor_config.fully_set_up():
            raise Exception("Data extractor is not fully configured")
        self.config = data_extractor_config
        self.list_of_files = os.listdir(self.config.files_path)

    def extract(self):
        for file in self.list_of_files:
            [company_ticker, frequency, *_] = list(map(str.upper, file.split(".")[0].split("_"))) # Remove the extension of the file, get only the name of the company and the frequency of the data in caps and discard the rest

            if frequency not in FrecuencyOfData._member_names_:
                continue
            if FrecuencyOfData[frequency] != self.config.frecuency_of_data:
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
            worksheet = workbook[self.config.file_b.sheet_name_of_metric] # The sheet names don't need to get stripped
            file_config = self.config.file_b
        except:
            try:
                sheet_with_metric = None

                for sheet_name in workbook.sheetnames:
                    if self.config.file_a.sheet_name_of_metric in workbook[sheet_name]["A3"].value: # Long company names reach the sheet's name character limit, but, A3 cell always has the type of sheet
                        sheet_with_metric = sheet_name
                        break

                if not sheet_with_metric:
                    raise Exception(f"{self.config.file_a.sheet_name_of_metric} not found in {workbook.worksheets}. Check if sheet name for style b was correctly typed.")
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
    file_type_a_config = FileConfig("A1", OrderOfDate.DESCENDING)
    file_type_b_config = FileConfig("B2", OrderOfDate.ASCENDING)

    data_extractor_config = DataExtractorConfig("companies_data", file_type_a_config, file_type_b_config)
    data_extractor_config.set_up_file_types("Ratios - Key Metric", "Pretax ROA", "Financial Summary", "Pretax ROA - %, TTM", FrecuencyOfData.QUARTERLY)

    data_extractor = DataExtractor(data_extractor_config)

    data_extractor.extract()


if __name__ == "__main__":
    main()
