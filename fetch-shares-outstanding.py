import requests
import json
import pandas as pd

API_KEY = "17b581337emsh2c49e9ba4d432a0p139b8ajsn3f1cd6b23669"


def fetch_financials(symbol):
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
        balance_sheet_history = data['balanceSheetHistory']['balanceSheetStatements']

        total_revenues = [
            income_statement_history[3]['totalRevenue']['raw'],
            income_statement_history[2]['totalRevenue']['raw'],
            income_statement_history[1]['totalRevenue']['raw']
        ]

        shareholders_equity = [
            balance_sheet_history[3]['totalStockholderEquity']['raw'],
            balance_sheet_history[2]['totalStockholderEquity']['raw'],
            balance_sheet_history[1]['totalStockholderEquity']['raw']
        ]

        total_assets = [
            balance_sheet_history[3]['totalAssets']['raw'],
            balance_sheet_history[2]['totalAssets']['raw'],
            balance_sheet_history[1]['totalAssets']['raw']
        ]

        return total_revenues, shareholders_equity, total_assets
    except:
        return ([None, None, None], [None, None, None], [None, None, None])


tickers_df = pd.read_excel("company-reviews-test.xlsx")
tickers = tickers_df['Ticker'].tolist()

columns = ["Ticker",
           "Total_Revenues_2019", "Total_Revenues_2020", "Total_Revenues_2021",
           "S_Equity_2019", "S_Equity_2020", "S_Equity_2021",
           "Total_Assets_2019", "Total_Assets_2020", "Total_Assets_2021"]
results_df = pd.DataFrame(columns=columns)

for ticker in tickers:
    print(f"Fetching data for {ticker}...")
    total_revenues, shareholders_equity, total_assets = fetch_financials(ticker)

    row = pd.Series([ticker, *total_revenues, *shareholders_equity, *total_assets], index=columns)
    results_df = pd.concat([results_df, row.to_frame().T], ignore_index=True)

results_df.to_excel("fAAAAAA.xlsx", index=False)
