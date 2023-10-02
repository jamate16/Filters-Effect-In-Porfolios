""" Fucntions to get different information about companies given their tickers """
import pandas as pd

class ClassificationFetcher:
    def __init__(self, data_file_path: str):
        self.data_file_path = data_file_path

        self.extracted_data = None

    def fetch(self, tickers: list) -> pd.DataFrame:

        df = pd.read_csv(self.data_file_path, index_col="Symbol")[["GICS Sector", "GICS Sub-Industry"]].T
        df.rename({
            "GICS Sector": "sector",
            "GICS Sub-Industry": "industry"
        }, inplace=True)
        
        self.extracted_data = df[tickers]
        return self.extracted_data.copy()


def main():
    # info_dict = get_classification(['EXC', 'TXT', 'ETN', 'MSI'])
    # print(info_dict)
    pass

if __name__ == "__main__":
    main()