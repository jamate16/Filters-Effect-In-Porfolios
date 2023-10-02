import pickle
import os

import yfinance as yf

import pandas as pd
import numpy as np


class ReturnsFetcher:
    """Calculates stock returns, both quarterly and daily.
    """
    def __init__(self, data_file_path: str):
        self.data_file_path = data_file_path
        self.load_data()

    def download_stock_data(self, symbols, start_date="1999-12-31", end_date="2023-09-29"):
        return yf.download(symbols, start=start_date, end=end_date)

    def calculate_returns(self, data, return_type="arithmetic"):
        quarterly_returns = {}
        daily_returns = {}

        for symbol in data.columns.get_level_values(1).unique():
            df = data.loc[:, pd.IndexSlice[:, symbol]].copy()
            df["daily_returns"] = df["Adj Close"].pct_change()
            df["quarter"] = df.index.to_period("Q")

            if return_type == "arithmetic":
                quarterly_returns[symbol] = df.groupby("quarter")["daily_returns"].mean()
            elif return_type == "geometric":
                df["daily_growth_multiplier"] = df["daily_returns"] + 1
                quarterly_returns[symbol] = df.groupby("quarter")["daily_growth_multiplier"].apply(np.prod) - 1

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
        symbol_not_in_sp500 = "ABNB"
        symbols.append(symbol_not_in_sp500) # For the download_stock_data method to work, the argument list hast to have a length of 2+
        # If saved data is missing symbols, download them and update the saved object
        missing_data_symbols = list(set([symbol for symbol in symbols if symbol != symbol_not_in_sp500]) - set(self.returns_data["quarterly_returns"].keys()))

        if refresh_data or return_type != self.returns_data.get("return_type"):
            data = self.download_stock_data(symbols)
            quarterly_returns, daily_returns = self.calculate_returns(data, return_type)
            self.returns_data = {"quarterly_returns": quarterly_returns, "daily_returns": daily_returns, "return_type": return_type}
            self.save_data()
        
        elif missing_data_symbols:
            data = self.download_stock_data(missing_data_symbols)
            quarterly_returns, daily_returns = self.calculate_returns(data, return_type)
            self.returns_data["quarterly_returns"].update(quarterly_returns)
            self.returns_data["daily_returns"].update(daily_returns)
            self.save_data()

        return self.returns_data["quarterly_returns"], self.returns_data["daily_returns"]


def main():
    quarterly_returns, daily_returns = ReturnsFetcher().fetch(["ABT", "MTB", "MSFT"])
    print(type(quarterly_returns))
    print("\n", daily_returns)


if __name__ == "__main__":
    main()
