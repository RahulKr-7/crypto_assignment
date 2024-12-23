import requests
import pandas as pd
from openpyxl import Workbook
import schedule
import time

# Fetch live data from CoinGecko
def fetch_crypto_data():
    url = "https://api.coingecko.com/api/v3/coins/markets"
    params = {
        'vs_currency': 'usd',
        'order': 'market_cap_desc',
        'per_page': 50,
        'page': 1,
        'sparkline': False
    }
    response = requests.get(url, params=params)
    data = response.json()
    
    # DataFrame to process the data
    df = pd.DataFrame(data, columns=[
        'name', 'symbol', 'current_price', 'market_cap',
        'total_volume', 'price_change_percentage_24h'
    ])
    
    # Save to Excel
    save_to_excel(df)
    
    # Perform analysis
    analyze_data(df)

# Save data to Excel
def save_to_excel(df):
    for _ in range(5):  # Retry up to 5 times
        try:
            with pd.ExcelWriter('crypto_data.xlsx', engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, sheet_name='Top 50 Cryptos', index=False)
                print("Excel updated successfully!")
                break  # Exit loop if successful
        except PermissionError:
            print("Excel file is open. Please close it.")
            time.sleep(5)  # Wait for 5 seconds before retrying
def analyze_data(df):
    # Top 5 by market cap
    top_5 = df.nlargest(5, 'market_cap')
    
    # Average price of top 50
    avg_price = df['current_price'].mean()
    
    # Highest and lowest 24h change
    highest_change = df.loc[df['price_change_percentage_24h'].idxmax()]
    lowest_change = df.loc[df['price_change_percentage_24h'].idxmin()]
    
    print("\nAnalysis Report:")
    print("Top 5 Cryptos by Market Cap:\n", top_5[['name', 'market_cap']])
    print(f"Average Price of Top 50: ${avg_price:.2f}")
    print(f"Highest 24h Price Change: {highest_change['name']} ({highest_change['price_change_percentage_24h']:.2f}%)")
    print(f"Lowest 24h Price Change: {lowest_change['name']} ({lowest_change['price_change_percentage_24h']:.2f}%)")

# Schedule to run every 5 minutes
schedule.every(5).minutes.do(fetch_crypto_data)

# Run the schedule
while True:
    schedule.run_pending()
    time.sleep(1)
