import xml.etree.ElementTree as ET
import pandas as pd
import requests
import io
import os
import calendar
from datetime import datetime, timedelta

# Caching to avoid multiple downloads
rate_cache = {}
multipliers_cache = {}
uniform_rate_cache = {}  # New cache for calculated annual averages

def get_cnb_rates(year):
    """Fetches the CNB exchange rate table for a specific year (all currencies)."""
    if year in rate_cache:
        return rate_cache[year], multipliers_cache[year]
    
    print(f"🌐 Fetching daily CNB rates for {year}...")
    url = f"https://www.cnb.cz/en/financial-markets/foreign-exchange-market/central-bank-exchange-rate-fixing/central-bank-exchange-rate-fixing/year.txt?year={year}"
    
    try:
        response = requests.get(url, headers={"User-Agent": "TaxProcessor/1.0"})
        response.raise_for_status()
        
        df = pd.read_csv(io.StringIO(response.text), sep='|')
        new_cols = {}
        multipliers = {}
        
        for col in df.columns:
            if col == 'Date':
                new_cols[col] = 'Date'
                continue
            parts = col.split()
            if len(parts) == 2:
                amt, curr = parts
                new_cols[col] = curr
                multipliers[curr] = float(amt)
        
        df.rename(columns=new_cols, inplace=True)
        df['Date'] = pd.to_datetime(df['Date'], dayfirst=True).dt.strftime('%Y%m%d')
        data = df.set_index('Date')
        
        rate_cache[year] = data
        multipliers_cache[year] = multipliers
        return data, multipliers
    except Exception as e:
        print(f"❌ Could not fetch rates for {year}: {e}")
        return None, None

def get_uniform_rate(year, currency):
    """Calculates the Uniform Rate (Average of month-ends) for a specific year."""
    if not currency or currency == "CZK":
        return 1.0
    
    cache_key = (year, currency)
    if cache_key in uniform_rate_cache:
        return uniform_rate_cache[cache_key]

    rates_df, multipliers = get_cnb_rates(year)
    if rates_df is None or currency not in rates_df.columns:
        return None

    monthly_rates = []
    for month in range(1, 13):
        # Find last day of month
        last_day = calendar.monthrange(int(year), month)[1]
        check_date = datetime(int(year), month, last_day)
        
        # Look back for last available work day if month ends on weekend
        found = False
        for _ in range(7):
            lookup_str = check_date.strftime('%Y%m%d')
            if lookup_str in rates_df.index:
                val = float(str(rates_df.loc[lookup_str, currency]).replace(',', ''))
                monthly_rates.append(val / multipliers.get(currency, 1.0))
                found = True
                break
            check_date -= timedelta(days=1)
    
    if len(monthly_rates) == 12:
        avg_rate = sum(monthly_rates) / 12
        # ROUND TO 5 DECIMALS
        final_rate = round(avg_rate, 5)
        uniform_rate_cache[cache_key] = final_rate
        return final_rate
    else:
        print(f"⚠️ Could not calculate Uniform Rate for {currency} {year} (incomplete data).")
        return None

def get_rate_for_date(date_str, currency):
    """Finds daily rate for any currency."""
    if not currency or currency == "CZK":
        return 1.0
    
    date_short = date_str[:8]
    trade_year = date_short[:4]
    rates_df, multipliers = get_cnb_rates(trade_year)
    
    if rates_df is not None and currency in rates_df.columns:
        check_date = datetime.strptime(date_short, "%Y%m%d")
        for i in range(7):
            lookup_str = check_date.strftime('%Y%m%d')
            if lookup_str in rates_df.index:
                raw_val = rates_df.loc[lookup_str, currency]
                rate_val = float(str(raw_val).replace(',', ''))
                return rate_val / multipliers.get(currency, 1.0)
            check_date -= timedelta(days=1)
    return None

def process_ibkr_xml(input_file, output_file):
    """Processes XML and adds both Daily and Uniform CZK values."""
    if not os.path.exists(input_file):
        print(f"❌ Error: {input_file} not found.")
        return

    print(f"🚀 Processing {input_file}...")
    tree = ET.parse(input_file)
    root = tree.getroot()

    # 1. Process Trades
    trade_count = 0
    for trade in root.findall(".//Trade"):
        date = trade.get('tradeDate')
        curr = trade.get('currency')
        commission_curr = trade.get('ibCommissionCurrency')
        quantity = float(trade.get('quantity'))
        year = date[:4]
        d_comm_rate = None
        u_comm_rate = None
        if commission_curr and commission_curr != curr:
            d_comm_rate= get_rate_for_date(date, commission_curr)
            u_comm_rate = get_uniform_rate(year, commission_curr)
        # Daily Rate
        d_rate = get_rate_for_date(date, curr)
        # Uniform Rate
        u_rate = get_uniform_rate(year, curr)

        price = float(trade.get('tradePrice'))
        if d_comm_rate:
            trade.set('dailyCommissionCZK', f"{float(trade.get('ibCommission')) * d_comm_rate:.2f}")
        if u_comm_rate:
            trade.set('uniformCommissionCZK', f"{round(float(trade.get('ibCommission')) * u_comm_rate, 2):.2f}")
        if d_rate is not None:
            trade.set('dailyRateCZK', f"{d_rate:.4f}")
            trade.set('dailyTradePriceCZK', f"{quantity * d_rate:.4f}")
        
        if u_rate is not None:
            # Force 2 decimal places in the XML attributes
            trade.set('uniformRateCZK', f"{u_rate:.2f}")
            trade.set('uniformTradePriceCZK', f"{round(quantity * u_rate, 2):.2f}")

        trade_count += 1

    # 2. Process Cash Transactions
    cash_count = 0
    for ct in root.findall(".//CashTransaction"):
        date = ct.get('dateTime')
        curr = ct.get('currency')
        year = date[:4]
        
        d_rate = get_rate_for_date(date, curr)
        u_rate = get_uniform_rate(year, curr)
        amount = float(ct.get('amount'))

        if d_rate:
            ct.set('dailyRateCZK', f"{d_rate:.4f}")
            ct.set('amountCZK', f"{amount * d_rate:.2f}")
        
        if u_rate:
            ct.set('uniformRateCZK', f"{u_rate:.2f}")
            ct.set('uniformAmountCZK', f"{round(amount * u_rate, 2):.2f}")
            
        cash_count += 1

    tree.write(output_file, encoding="utf-8", xml_declaration=True)
    print(f"\n✅ DONE! Processed {trade_count} trades and {cash_count} cash transactions.")

# --- START ---
process_ibkr_xml('merged.xml', 'final.xml')