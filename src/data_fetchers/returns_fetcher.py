import pickle
import os

import yfinance as yf

import pandas as pd
import numpy as np

from icecream import ic

class ReturnsFetcher:
    """Calculates stock returns, both quarterly and daily.
    """
    def __init__(self, data_file_path: str):
        self.data_file_path = data_file_path
        self.load_data()

    def download_stock_data(self, symbols, start_date="1999-12-31", end_date="2023-09-29"):
        return yf.download(symbols, start=start_date, end=end_date)

    def calculate_returns(self, symbols: list[str], data: pd.DataFrame, return_type: str="arithmetic") -> [dict, dict]:
        quarterly_returns = {}
        daily_returns = {}

        is_multi_index = isinstance(data.columns, pd.MultiIndex) # When there is only one symbol, columns are not an instance of MultiIndex

        for symbol in symbols:
            df = data.copy()
            if is_multi_index:
                df = df.loc[:, pd.IndexSlice[:, symbol]].copy()

            df["daily_returns"] = df["Adj Close"].pct_change()
            df["quarter"] = df.index.to_period("Q")

            if return_type == "arithmetic":
                quarterly_returns[symbol] = df.groupby("quarter")["daily_returns"].mean()
            elif return_type == "geometric":
                df["daily_growth_multiplier"] = df["daily_returns"] + 1
                quarterly_returns[symbol] = df.groupby("quarter")["daily_growth_multiplier"].apply(np.prod) - 1

                # Convert from daily to quarterly geometric mean
                quarterly_returns[symbol] = (np.power(quarterly_returns[symbol] + 1, 1/63) - 1).rename("quarterly_returns")  # Assuming around 63 trading days in a quarter


            daily_returns[symbol] = df["daily_returns"].copy()

        return quarterly_returns, daily_returns

    def load_data(self):
        try:
            with open(self.data_file_path, "rb") as infile:
                self.returns_data = pickle.load(infile)
        except FileNotFoundError:
            self.returns_data = {"quarterly_returns": {}, "daily_returns": {}}

    def save_data(self):
        with open(self.data_file_path, "wb") as outfile:
            pickle.dump(self.returns_data, outfile)

    def fetch(self, symbols, return_type="arithmetic", refresh_data=False):
        # If saved data is missing symbols, download them and update the saved object
        missing_data_symbols = list(set(symbols) - set(self.returns_data["quarterly_returns"].keys()))

        if refresh_data or return_type != self.returns_data.get("return_type"):
            data = self.download_stock_data(symbols)
            quarterly_returns, daily_returns = self.calculate_returns(symbols, data, return_type)

            self.returns_data = {"quarterly_returns": quarterly_returns, "daily_returns": daily_returns, "return_type": return_type}
            self.save_data()
        
        elif missing_data_symbols:
            data = self.download_stock_data(missing_data_symbols)
            quarterly_returns, daily_returns = self.calculate_returns(missing_data_symbols, data, return_type)

            self.returns_data["quarterly_returns"].update(quarterly_returns)
            self.returns_data["daily_returns"].update(daily_returns)
            self.returns_data["return_type"] = return_type
            self.save_data()

        return self.returns_data["quarterly_returns"], self.returns_data["daily_returns"]


def main():
    quarterly_returns, daily_returns = ReturnsFetcher(os.path.join("data", "pickled_data", "returns_data.pickle")).fetch(["MSFT"], "geometric")
    ic(quarterly_returns)
    ic(daily_returns)

if __name__ == "__main__":
    main()
