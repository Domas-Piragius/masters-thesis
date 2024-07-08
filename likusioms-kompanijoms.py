import requests
import pandas as pd
import openpyxl

# Read tickers from the file
tickers_df = pd.read_excel("company-reviews-test.xlsx")
tickers = tickers_df['Ticker'].tolist()

# RapidAPI credentials
url = "https://apidojo-yahoo-finance-v1.p.rapidapi.com/stock/v2/get-financials"
headers = {
    "X-RapidAPI-Host": "apidojo-yahoo-finance-v1.p.rapidapi.com",
    "X-RapidAPI-Key": "17b581337emsh2c49e9ba4d432a0p139b8ajsn3f1cd6b23669"
}

# Prepare an empty DataFrame for the results
columns = ["Ticker", "Operating_2019", "Operating_2020", "Operating_2021",
           "T_Assets_2019", "T_Assets_2020", "T_Assets_2021"]
results_df = pd.DataFrame(columns=columns)

# Fetch and process data for each ticker
for ticker in tickers:
    print(f"Fetching data for {ticker}...")
    params = {"symbol": ticker, "region": "US"}
    response = requests.get(url, headers=headers, params=params)
    data = response.json()

    try:
        # Extract operating income and total assets
        operating_income = [data['incomeStatementHistory']['incomeStatementHistory'][i]['operatingIncome']['raw'] for i in range(3)]
        total_assets = [data['balanceSheetHistory']['balanceSheetStatements'][i]['totalAssets']['raw'] for i in range(3)]

    except KeyError:
        # Handle missing data
        operating_income = [None, None, None]
        total_assets = [None, None, None]

    # Append the results to the DataFrame
    row = pd.Series([ticker, *operating_income, *total_assets], index=columns)
    results_df = pd.concat([results_df, row.to_frame().T], ignore_index=True)

# Save the results to an Excel file
results_df.to_excel("operating_income_and_assets.xlsx", index=False)
