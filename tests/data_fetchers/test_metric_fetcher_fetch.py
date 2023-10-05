import unittest
from pandas.testing import assert_frame_equal

import os
import sys
sys.path.append(os.path.join(os.getcwd()))
from src.data_fetchers.metrics_fetcher import MetricsFetcher
from src.configs.file_style_configs_by_metric import file_style_configs_by_metric


class DataLoadingTestCase(unittest.TestCase):
    def setUp(self):
        self.fetcher = MetricsFetcher(os.path.join("tests", "test_data", "companies_data"), file_style_configs_by_metric)

    def test_load_from_excel_file(self):
        test_metric = "Pretax ROA"
        self.fetcher._load_from_excel_file(test_metric)

        symbols = [metric.company_name for metric in self.fetcher.extracted_data]
        self.assertEqual(symbols, ["Apple Inc", "Abbott Laboratories"])

    def test_save_load_metric_data_pickle_file(self):
        test_metric = "Pretax ROA"
        self.fetcher._load_from_excel_file(test_metric)
        data_to_save = self.fetcher.get_dataframe()
        path_of_file = os.path.join("tests", "test_data", "pickled_data", f"{test_metric}_data.pickle")

        self.fetcher._save_metric_data(path_of_file, self.fetcher.get_dataframe())
        loaded_data = self.fetcher._load_from_pickle_file(path_of_file)

        assert_frame_equal(data_to_save, loaded_data)

        
class FetchingMetricDataTestCase(unittest.TestCase):
    def setUp(self):
        self.fetcher = MetricsFetcher(os.path.join("tests", "test_data", "companies_data"), file_style_configs_by_metric)
        self.test_metric = "Pretax ROA"
        self.path_to_pickled_data = os.path.join("tests", "test_data", "pickled_data")
        self.pickled_data_filename = f"{self.test_metric}_data.pickle"

        os.remove(os.path.join(os.getcwd(), os.path.join(self.path_to_pickled_data, self.pickled_data_filename)))

    def test_fetch_existing_data(self):
        self.fetcher._load_from_excel_file(self.test_metric)
        saved_data = self.fetcher.get_dataframe()
        self.fetcher._save_metric_data(os.path.join(self.path_to_pickled_data, self.pickled_data_filename),
                                       saved_data)

        fetched_data = self.fetcher.fetch(self.test_metric, pickled_data_path=self.path_to_pickled_data)
        assert_frame_equal(fetched_data, saved_data)
        
    def test_fetch_non_existing_data(self):
        self.fetcher._load_from_excel_file(self.test_metric)
        loaded_data = self.fetcher.get_dataframe()

        fetched_data = self.fetcher.fetch(self.test_metric, pickled_data_path=self.path_to_pickled_data)
        assert_frame_equal(fetched_data, loaded_data)


# TODO: write tests to check if dates are being extracted correclty

if __name__ == "__main__":
    unittest.main()