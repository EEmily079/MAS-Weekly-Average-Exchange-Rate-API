import requests
import pandas as pd

from datetime import date, timedelta




base_url = "https://eservices.mas.gov.sg/apimg-gw/server/monthly_statistical_bulletin_non610ora/exchange_rates_average_for_period_weekly/views/exchange_rates_average_for_period_weekly"
headers = {
    "accept": "application/json; charset=UTF-8",
    "KeyId": "Use_Your_Own_API_Key"
    }
today = date.today()

for i in range(7):
    check_date = today - timedelta(days=i)
    end_of_week_param = check_date.strftime("%Y-%m-%d")

    params = {
        "end_of_week": end_of_week_param
    }

    response = requests.get(base_url, headers=headers, params=params)

    if response.status_code != 200:
        continue

    try:
        data = response.json()
    except:
        continue

    elements = data.get("elements", [])

    if elements:
        print("Latest available date found:", end_of_week_param)
        df = pd.DataFrame(elements)
        selected_columns = [
    "end_of_week",
    "aud_sgd", 
    "chf_sgd", 
    "cny_sgd_100",
    "eur_sgd",
    "gbp_sgd",
    "jpy_sgd_100",
    "usd_sgd",
    ]

        df_selected = df[selected_columns]

        # Convert all rate columns to numeric first
        rate_columns = selected_columns[1:]
        df_selected[rate_columns] = df_selected[rate_columns].apply(pd.to_numeric, errors="coerce")

        # Convert per 100 currency rates to per 1 unit
        df_selected["cny_sgd_100"] = df_selected["cny_sgd_100"] * 0.01
        df_selected["jpy_sgd_100"] = df_selected["jpy_sgd_100"] * 0.01

        print(df_selected)
        break
else:
    raise Exception("No exchange rate data found in the last 7 days.")


df_selected = df_selected.rename(columns={
    "end_of_week": "Date",
    "eur_sgd": "EUR",
    "gbp_sgd": "GBP",
    "usd_sgd": "USD",
    "aud_sgd": "AUD",
    "cny_sgd_100": "CNY",
    "jpy_sgd_100": "JPY",
    "chf_sgd": "CHF"
})

# Insert empty SEK column between JPY and USD
df_selected.insert(
    loc=df_selected.columns.get_loc("USD"),
    column="SEK",
    value=""
)

# Export to Excel
output_file = r"Your_Folder\MAS_Weekly_Exchange_Rates.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_selected.to_excel(writer, sheet_name="Exchange_Rates", index=False)
   

print("Excel file created successfully.")