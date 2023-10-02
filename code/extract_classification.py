""" Fucntions to get different information about companies given their tickers """
import os

import pandas as pd


def get_classification(tickers: list) -> pd.DataFrame:
    classification_file_path = os.path.join(os.path.dirname(__file__), "..", "companies_classification", "constituents.csv")

    df = pd.read_csv(classification_file_path, index_col="Symbol")[["GICS Sector", "GICS Sub-Industry"]].T
    df.rename({
        "GICS Sector": "sector",
        "GICS Sub-Industry": "industry"
    }, inplace=True)
    
    return df[tickers]


def main():
    info_dict = get_classification(['EXC', 'TXT', 'ETN', 'MSI'])
    print(info_dict)

if __name__ == "__main__":
    main()