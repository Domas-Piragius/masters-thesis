import requests
import json
import pandas as pd

API_KEY = "17b581337emsh2c49e9ba4d432a0p139b8ajsn3f1cd6b23669"


def fetch_net_income(symbol):
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

        net_income = [
            income_statement_history[3]['netIncome']['raw'],
            income_statement_history[2]['netIncome']['raw'],
            income_statement_history[1]['netIncome']['raw']
        ]

        return net_income
    except:
        return [None, None, None]


def fetch_shareholders_equity(symbol):
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

        shareholders_equity = [
            balance_sheet_history[3]['totalStockholderEquity']['raw'],
            balance_sheet_history[2]['totalStockholderEquity']['raw'],
            balance_sheet_history[1]['totalStockholderEquity']['raw']
        ]

        return shareholders_equity
    except:
        return [None, None, None]


tickers_df = pd.read_excel("company-reviews-test.xlsx")
tickers = tickers_df['Ticker'].tolist()

columns = ["Ticker", "Net_Income_2019", "Net_Income_2020", "Net_Income_2021",
           "S_Equity_2019", "S_Equity_2020", "S_Equity_2021"]
results_df = pd.DataFrame(columns=columns)

for ticker in tickers:
    print(f"Fetching data for {ticker}...")
    net_income = fetch_net_income(ticker)
    shareholders_equity = fetch_shareholders_equity(ticker)

    row = pd.Series([ticker, *net_income, *shareholders_equity], index=columns)
    results_df = pd.concat([results_df, row.to_frame().T], ignore_index=True)

results_df.to_excel("net_income_and_shareholders_equity.xlsx", index=False)
