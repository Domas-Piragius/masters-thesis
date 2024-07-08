import requests
import json
import pandas as pd

API_KEY = "17b581337emsh2c49e9ba4d432a0p139b8ajsn3f1cd6b23669"


def fetch_operating_income(symbol):
    url = "https://yh-finance.p.rapidapi.com/stock/v2/get-financials"
    querystring = {"symbol": symbol, "region": "US"}
    headers = {
        "X-RapidAPI-Key": API_KEY,
        "X-RapidAPI-Host": "yh-finance.p.rapidapi.com"
    }

    try:
        response = requests.request("GET", url, headers=headers, params=querystring)
        data = json.loads(response.text)
        income_statement_history = data['incomeStatementHistory']['incomeStatementHistory']

        operating_income = [
            income_statement_history[3]['operatingIncome']['raw'],
            income_statement_history[2]['operatingIncome']['raw'],
            income_statement_history[1]['operatingIncome']['raw']
        ]

        return operating_income
    except:
        return [None, None, None]


def fetch_total_assets(symbol):
    url = "https://yh-finance.p.rapidapi.com/stock/v2/get-financials"
    querystring = {"symbol": symbol, "region": "US"}
    headers = {
        "X-RapidAPI-Key": API_KEY,
        "X-RapidAPI-Host": "yh-finance.p.rapidapi.com"
    }

    try:
        response = requests.request("GET", url, headers=headers, params=querystring)
        data = json.loads(response.text)
        balance_sheet_history = data['balanceSheetHistory']['balanceSheetStatements']

        total_assets = [
            balance_sheet_history[3]['totalAssets']['raw'],
            balance_sheet_history[2]['totalAssets']['raw'],
            balance_sheet_history[1]['totalAssets']['raw']
        ]

        return total_assets
    except:
        return [None, None, None]


tickers_df = pd.read_excel("company-reviews-test.xlsx")
tickers = tickers_df['Ticker'].tolist()

columns = ["Ticker", "Operating_2019", "Operating_2020", "Operating_2021",
           "T_Assets_2019", "T_Assets_2020", "T_Assets_2021"]
results_df = pd.DataFrame(columns=columns)

for ticker in tickers:
    print(f"Fetching data for {ticker}...")
    operating_income = fetch_operating_income(ticker)
    total_assets = fetch_total_assets(ticker)

    row = pd.Series([ticker, *operating_income, *total_assets], index=columns)
    results_df = pd.concat([results_df, row.to_frame().T], ignore_index=True)

results_df.to_excel("operating_income_and_assets.xlsx", index=False)
