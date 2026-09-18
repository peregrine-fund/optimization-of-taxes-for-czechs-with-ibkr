import pandas as pd
from collections import defaultdict
import os
import sys
from datetime import datetime
import xml.etree.ElementTree as ET
from openpyxl.styles import Font, Color
import yfinance as yf
import glob
import warnings
from cnb_converter import CNBConverter
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", message=".*Timestamp.utcnow is deprecated.*")
import json
# ... your other imports ...

target_year = "2025"
preferred_method = "LIFO"
preferred_rate = "daily"
preferred_rate = preferred_rate.lower()
preferred_method = preferred_method.upper()
SPLITS = {
    'NVDA': [('20240607', 10.0)]
}

suspect_currencies = ["USD", "EUR", "MYR", "AUD", "CAD", "CHF", "HKD", "SGD", "JPY", "CNY", "GBP"]  # Add more as needed


COUNTRY_MAP = {
    "United Kingdom": "UK",
    "United Kingdom of Great Britain and Northern Ireland": "UK",
    "United States": "US",
    "Ireland": "IE",
    "Germany": "DE",
    "France": "FR",
    "Netherlands": "NL",
    "Switzerland": "CH",
    "Canada": "CA",
    "Taiwan": "TW",
    "Malaysia": "MY"
}
# Map of Country Code -> CNB Currency Code
COUNTRY_TO_CURR = {
    # Daily Fixings
    "UK": "GBP",
    "US": "USD",
    "IE": "EUR",
    "DE": "EUR",
    "FR": "EUR",
    "NL": "EUR",
    "CH": "CHF",
    "CA": "CAD",
    "AU": "AUD",
    "JP": "JPY",
    "HK": "HKD",
    "CN": "CNY",
    "PL": "PLN",
    
    # Obscure (Monthly) Fixings
    "TW": "TWD",
    "MY": "MYR",
    "SG": "SGD",
    "TH": "THB",
    "VN": "VND",
    "IN": "INR",
    "KR": "KRW",
    "IL": "ILS",
    "BR": "BRL",
    "MX": "MXN",
    "ZA": "ZAR"
}

EXCHANGE_MAP = {
    "LSE": ".L",      # London
    "LSEETF": ".L", # London ETFs
    "VSE": ".VI",     # Vienna
    "EBS": ".SW",     # SIX Swiss Exchange
    "SBF": ".PA",     # Paris
    "IBIS": ".DE",    # XETRA (Germany)
    "IBIS2": ".DE",   # XETRA
    "AEB": ".AS",     # Amsterdam
    "MIL": ".MI",     # Milan
    "BM": ".MC",      # Madrid
    "OMXH": ".HE",    # Helsinki
    "FRA": ".F",      # Frankfurt
    "TSE": ".T",      # Tokyo
    "HKFE": ".HK",    # Hong Kong
    "ASX": ".AX",     # Australia
    "TSX": ".TO",     # Toronto
}

Broken_Yahoo_Stock_Map={
    "NSOP": "2038.KL"
}

Ignore_Yahoo_Stock_Map=["IBTU", "IBGS"]

moje_dane = {}
country_mismatches = {}
attr_name = ""
parameter_info=""

sum_net_dividends_daily = 0.0
sum_net_dividends_uniform = 0.0
sum_gross_dividends = 0.0

def get_short_country(name):
    return COUNTRY_MAP.get(name, name)

class Tee(object):
    def __init__(self, *files):
        self.files = files
    def write(self, obj):
        for f in self.files:
            f.write(obj)
            f.flush() # Ensure it writes to the file in real-time
    def flush(self):
        for f in self.files:
            f.flush()


class InternetCache:
    def __init__(self, base_dir="./internet"):
        self.base_dir = base_dir
        self.countries_path = os.path.join(base_dir, "countries.json")
        self.cnb_dir = os.path.join(base_dir, "cnb-rates-for-dividends")
        self.overwritten_tickers = []
        self.overwritten_cnb = []
        
        os.makedirs(self.cnb_dir, exist_ok=True)
        
        if os.path.exists(self.countries_path):
            with open(self.countries_path, 'r', encoding='utf-8') as f:
                self.countries_data = json.load(f)
        else:
            self.countries_data = {}

    def get_country(self, symbol, force_update):
        if not force_update and symbol in self.countries_data:
            return self.countries_data[symbol]
        return None

    def save_country(self, symbol, ibkr, isin, yf_c):
        # Check if we are overwriting
        if symbol in self.countries_data:
            self.overwritten_tickers.append(symbol)
            
        self.countries_data[symbol] = {
            "ibkr_country": ibkr, 
            "isin_prefix": isin, 
            "yf_country": yf_c
        }
        with open(self.countries_path, 'w', encoding='utf-8') as f:
            json.dump(self.countries_data, f, indent=4)

    def get_cnb(self, date_str, currency, force_update):
        path = os.path.join(self.cnb_dir, f"{date_str}.json")
        if not force_update and os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if currency in data:
                    return data[currency]
        return None

    def save_cnb(self, date_str, currency, rate):
        path = os.path.join(self.cnb_dir, f"{date_str}.json")
        data = {}
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        
        # Check if we are overwriting this specific currency for this date
        if currency in data:
            self.overwritten_cnb.append(f"{date_str} ({currency})")
            
        data[currency] = rate
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)

    def print_summary(self):
        if self.overwritten_tickers or self.overwritten_cnb:
            print("\n♻️  CACHE UPDATE SUMMARY")
            print("-" * 32)
            if self.overwritten_tickers:
                print(f"Overwritten Tickers: {', '.join(set(self.overwritten_tickers))}")
            if self.overwritten_cnb:
                print(f"Overwritten CNB Rates: {', '.join(set(self.overwritten_cnb))}")
            print("-" * 32)
cache = InternetCache()

def export_dividends_to_xlsx(root, rate_type, target_year, excel_ready):
    global sum_net_dividends_daily
    global sum_net_dividends_uniform
    global sum_gross_dividends
    
    country_totals = defaultdict(lambda: {
        'gross_czk_daily': 0.0, 'tax_czk_daily': 0.0,
        'gross_czk_uniform': 0.0, 'tax_czk_uniform': 0.0,
        'other_czk_daily': 0.0, 'other_czk_uniform': 0.0,
        'original_tax': 0.0,

        'symbols': set()
    })

    # Key: (date, symbol) -> Value: combined data
    stacked_data = defaultdict(lambda: {
        'gross_orig': 0.0, 'tax_orig': 0.0, 
        'gross_daily': 0.0, 'tax_daily': 0.0, 
        'gross_unif': 0.0, 'tax_unif': 0.0,
        'country': ''
    })

    transactions = root.findall(".//CashTransaction")

    for trans in transactions:
        t = trans.attrib
        date_only = t.get('settleDate', '')
        if not date_only.startswith(target_year): 
            continue
            
        type_ = t.get('type', '')
        if "Deposits/Withdrawals" in type_:
            continue
            
        symbol = t.get('symbol', '')
        amount = float(t.get('amount', 0))
        ibkr_country = t.get('issuerCountryCode', '')
        isin_full = t.get('isin', '')
        isin_prefix = isin_full[:2] if isin_full else '' # Země je první 2 znaky ISIN
        exchange = t.get('listingExchange', '')
        code=ibkr_country
        if symbol and symbol not in country_mismatches and symbol not in Ignore_Yahoo_Stock_Map:

            # Získáme zemi z yfinance
            force_refresh = (parameter_info.lower() == "update")
            cached_country = cache.get_country(symbol, force_refresh)
            # Logic to append exchange suffix
            if cached_country and not force_refresh:
                # Fetch directly from JSON
                yf_country = cached_country["yf_country"]
            else:
                suffix = EXCHANGE_MAP.get(exchange, "")
                
                # Special case: If it's a US exchange, we usually don't need a suffix
                if exchange in ["NYSE", "NASDAQ", "ARCA", "AMEX"]:
                    symbol_for_yf = symbol
                else:
                    symbol_for_yf = f"{symbol}{suffix}"
                if symbol in Broken_Yahoo_Stock_Map:
                    symbol_for_yf = Broken_Yahoo_Stock_Map.get(symbol_for_yf, '')
                    yf_info = yf.Ticker(symbol_for_yf).info
                else:
                    yf_info = yf.Ticker(symbol_for_yf).info
                if not yf_info or 'country' not in yf_info:
                    raise ValueError(f"CRITICAL ERROR: Could not fetch country data for {symbol} ({symbol_for_yf}). Stopping execution.")
                yf_country_full = yf_info.get('country', '')
                yf_country = get_short_country(yf_country_full)
                cache.save_country(symbol, ibkr_country, isin_prefix, yf_country)
                # Porovnáme všechny tři zdroje
            if not (ibkr_country == isin_prefix == yf_country):
                # Uložíme řetězec v požadovaném formátu: ticker + issuer + isin + country
                country_mismatches[symbol] = f"{symbol} {ibkr_country} {isin_prefix} {yf_country}"
            code=yf_country
        d_rate = float(t.get('dailyRateCZK', 1.0))
        u_rate = float(t.get('uniformRateCZK', 1.0))

        if parameter_info.upper() in suspect_currencies:
            d_rate, u_rate = 1.0, 1.0

        if symbol:
            country_totals[code]['symbols'].add(symbol)
        
        # --- LOGIC FOR STACKING ---
        # We group by date and symbol to put Tax and Dividend on one line
        key = (date_only, symbol)
        stacked_data[key]['country'] = code

        if "Dividend" in type_:
            stacked_data[key]['gross_orig'] += amount
            stacked_data[key]['gross_daily'] += (amount * d_rate)
            stacked_data[key]['gross_unif'] += (amount * u_rate)
            
            country_totals[code]['gross_czk_daily'] += (amount * d_rate)
            country_totals[code]['gross_czk_uniform'] += (amount * u_rate)
        elif "Withholding Tax" in type_ or "Tax" in type_:
            stacked_data[key]['tax_orig'] += amount
            stacked_data[key]['tax_daily'] += (amount * d_rate)
            stacked_data[key]['tax_unif'] += (amount * u_rate)
            country_totals[code]['tax_czk_daily'] += (amount * d_rate)
            country_totals[code]['tax_czk_uniform'] += (amount * u_rate)
            cnb_currency = COUNTRY_TO_CURR.get(code)
            force_refresh = (parameter_info.lower() == "update")
            
            # 1. Zkusíme načíst z JSON cache
            cnb_rate = cache.get_cnb(date_only, cnb_currency, force_refresh)
            
            if cnb_rate is None:
                # 2. Pokud není v cache, zavoláme API a uložíme
                converter = CNBConverter()
                cnb_rate = converter.get_rate(cnb_currency, date_only, preferred_rate)
                cache.save_cnb(date_only, cnb_currency, cnb_rate)
            if(preferred_rate == "daily"):
                country_totals[code]['original_tax']+=amount*d_rate/cnb_rate
            elif(preferred_rate == "uniform"):
                country_totals[code]['original_tax']+=amount*u_rate/cnb_rate

        else:
            # Handle "Other" (e.g., Lieu of dividends)
            country_totals[code]['other_czk_daily'] += amount * d_rate
            country_totals[code]['other_czk_uniform'] += amount * u_rate

    if not country_totals:
        return

    # --- UPDATED: STACKED INDIVIDUAL TRANSACTION PRINT ---
    print(f"\n📜 NET DIVIDEND LOG BY PAYOUT ({target_year})")
    print("-" * 135)
    print(f"{'Date':<12} | {'Symbol':<8} | {'Gross Orig':>12} | {'Tax rate':>12} | {'Net Orig':>12} | {'Net Daily (CZK)':>18} | {'Net Unif (CZK)':>18}")
    print("-" * 135)
    individual_rows = []
    for (d, s) in sorted(stacked_data.keys()):
        val = stacked_data[(d, s)]
        net_orig = val['gross_orig'] + val['tax_orig']
        tax_rate= (val['tax_orig'] / val['gross_orig'] * 100) if val['gross_orig'] != 0 else 0.0
        net_daily = val['gross_daily'] + val['tax_daily']
        net_unif = val['gross_unif'] + val['tax_unif']
        
        print(f"{d:<12} | {s:<8} | {val['gross_orig']:>12.2f} | {tax_rate:>12.2f} | {net_orig:>12.2f} | {net_daily:>18.2f} | {net_unif:>18.2f}")
        individual_rows.append({
            'Date': d,
            'Symbol': s,
            'Country': val['country'],
            'Gross Original': val['gross_orig'],
            'Tax Original': val['tax_orig'],
            'Net Daily (CZK)': val['gross_daily'] + val['tax_daily'],
            'Net Unif (CZK)': val['gross_unif'] + val['tax_unif']
        })
    print("-" * 135)

    print(f"\n📊 DIVIDEND SUMMARY FOR {target_year}")
    print("=" * 140)
    print(f"{'Country':<12} | {'Symbols':<20} | {'Gross (Daily)':>15} | {'Tax (Daily)':>15} | {'Other (Daily)':>15} | {'Gross (Unif)':>15} | {'Tax (Unif)':>15} | {'Other (Unif)':>15}")
    print("-" * 140)

    summary_data = []
    for country, totals in sorted(country_totals.items()):
        syms_str = ", ".join(sorted(totals['symbols']))
        print(f"{country:<12} | {syms_str[:20]:<20} | {totals['gross_czk_daily']:>15.2f} | {totals['tax_czk_daily']:>15.2f} | {totals['other_czk_daily']:>15.2f} | {totals['gross_czk_uniform']:>15.2f} | {totals['tax_czk_uniform']:>15.2f} | {totals['other_czk_uniform']:>15.2f}")
        
        g_val = totals['gross_czk_daily'] if rate_type == 'daily' else totals['gross_czk_uniform']
        t_val = totals['tax_czk_daily'] if rate_type == 'daily' else totals['tax_czk_uniform']
        o_val = totals['other_czk_daily'] if rate_type == 'daily' else totals['other_czk_uniform']
        original_tax=totals['original_tax']
        summary_data.append({
            'Country': country, 'Symbols': syms_str,
            'Gross Dividend (CZK)': round(g_val, 2),
            'Tax Paid Abroad (CZK)': round(t_val, 2),
            'Net Received (CZK)': round(g_val + t_val, 2),
        })
        moje_dane[country+"-gross"] = round(g_val, 0)
        moje_dane[country+"-tax"] = round(t_val, 0)
        moje_dane[country+"-original-tax"] = round(original_tax, 0)
        sum_net_dividends_daily += round(totals['gross_czk_daily'] + totals['tax_czk_daily'], 2)
        sum_net_dividends_uniform += round(totals['gross_czk_uniform'] + totals['tax_czk_uniform'], 2)
        
    if excel_ready:
        output_file = f"history/dividend_report_{target_year}_{rate_type}.xlsx"
        if not os.path.exists('history'): os.makedirs('history')
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name='Summary')
            pd.DataFrame(individual_rows).to_excel(writer, index=False, sheet_name='Individual_Payments')

def run_tax_simulations(xml_file='final.xml', save_xml=False): # Added save_xml arg here
    global attr_name
    attr_name = f"used-{target_year}"
    excel_detailed_trades = []
    excel_tax_summary = []
    final_inventory_for_excel = {}
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
    except Exception as e:
        print(f"❌ Error: {e}"); return

    trades_elements = [t for t in root.findall(".//Trade") if t.get('assetCategory') != 'CASH']
    trades_elements.sort(key=lambda x: x.get('tradeDate'))

    rates = ['daily', 'uniform']
    methods = ['FIFO', 'LIFO', 'GLOBAL_MAX_LOSS']
    
    print(f"\n🚀 GLOBAL TAX SIMULATION FOR YEAR: {target_year}")
    simulation_results = []

    for rate_type in rates:
        rate_key = f"{rate_type}RateCZK"
        for method in methods:
            inventory = defaultdict(list)
            sales_to_process = []
            detailed_log_entries = [] 
            total_rev, total_cost, total_fees = 0.0, 0.0, 0.0
            
            for t_el in trades_elements:
                # --- PREVENT MULTIPLE DIVISIONS: Work on a copy of the original attributes ---
                t = t_el.attrib.copy()
                # If parameter_info is in suspect_currencies, skip any trade not from the corresponding currency
                if parameter_info.upper() in suspect_currencies:
                    if t.get('currency').upper() != parameter_info:
                        continue
                commission_currency = t.get('ibCommissionCurrency', '')
                currency=t.get('currency', '')

                symbol, date = t['symbol'], t['tradeDate']
                qty = abs(float(t['quantity']))
                price = float(t['tradePrice'])
                rate = float(t.get(rate_key, 1.0))

                if parameter_info.upper() in suspect_currencies:
                    rate=1.0
                
                
                # --- APPLY SPLITS ONCE PER SIMULATION ---
                if symbol in SPLITS:
                    for split_date, ratio in SPLITS[symbol]:
                        if date < split_date:
                            qty *= ratio
                            price /= ratio

                trade_year = int(date[:4])
                comm = abs(float(t.get('ibCommission', 0))) * rate
                if(commission_currency != currency):
                    comm=abs(float(t.get(f'{rate_type}CommissionCZK')))

                tax = abs(float(t.get('taxes', 0))) * rate
                trade_fees = comm + tax  # <--- Define this

                if t['buySell'] == 'BUY':
                    # --- NEW: Check for previous years' usage ---
                    already_used_ratio = 0.0
                    for attr, val in t_el.attrib.items():
                        if attr.startswith("used-"):
                            try:
                                attr_year = int(attr.split("-")[1])
                                # If this lot was used in a year BEFORE our target year
                                if attr_year < int(target_year):
                                    already_used_ratio += float(val)
                            except (ValueError, IndexError):
                                continue
                    
                    # Calculate remaining quantity
                    # If already_used_ratio is 1.0, the lot is fully gone.
                    available_qty = qty * (1.0 - already_used_ratio)
                    
                    if already_used_ratio > 1.0000001: # (The .0000001 allows for tiny float math rounding errors)
                        error_msg = f"⚠️ WARNING: {symbol} CASH lot from {date} is overused! Total used: {already_used_ratio*100:.2f}%. Fixing to 0."
                        sys.exit(error_msg)
                    # ----------------------------------------------------
                    
                    # Only add to inventory if there's actually something left to sell
                    if available_qty > 1e-9:
                        cost = (price * qty * rate) # Total cost remains based on original purchase
                        
                        inventory[symbol].append({
                            'qty': available_qty,       # Use the adjusted quantity
                            'orig_qty': qty,            # Keep original for ratio calculations
                            'unit_cost': cost/qty, 
                            'date': date,
                            'raw_buy_price': price,
                            'purchase_fees': trade_fees,
                            'xml_el': t_el 
                        })

                    if trade_year == int(target_year): 
                        detailed_log_entries.append({
                            'symbol': symbol, 'type': 'BUY', 'qty': available_qty, 'price': price,
                            'price_total': price * available_qty * rate, 'pnl': 0.0, 'is_sell': False, 'date': date,
                            'trade_fees': trade_fees  
                        })
                
                elif t['buySell'] == 'SELL':
                    comm = abs(float(t.get('ibCommission', 0))) * rate 
                    tax = abs(float(t.get('taxes', 0))) * rate
                    sell_price_czk = price * rate
                    rev = (sell_price_czk * qty) 

                    if trade_year < float(target_year):
                        continue
                    elif trade_year == float(target_year):
                        total_rev += rev
                        sales_to_process.append({'symbol': symbol, 'qty': qty, 'sell_price_czk': sell_price_czk, 'date': date, 'trade_fees': trade_fees})
                    else:
                        continue
                else:
                    print("NEW BUY/SELL TYPE ENCOUNTERED!")
            
            # Process target year sales
            for sale in sales_to_process:
                symbol, needed = sale['symbol'], sale['qty']
                trade_fees = sale['trade_fees']
                if method == 'FIFO': 
                    inventory[symbol].sort(key=lambda x: x['date'])
                elif method == 'LIFO': 
                    inventory[symbol].sort(key=lambda x: x['date'], reverse=True)
                elif method == 'GLOBAL_MAX_LOSS':
                    inventory[symbol].sort(key=lambda x: x['unit_cost'], reverse=True)

                while needed > 0 and inventory[symbol]:
                    lot = inventory[symbol][0]
                    current_lot_qty = lot['qty']
                    take = min(needed, current_lot_qty)
                    share_fee = take / sale['qty']
                    partial_trade_fee = trade_fees * share_fee
                    lot_cost = take * lot['unit_cost']
                    lot_rev = take * sale['sell_price_czk']
                    sold_price = sale['sell_price_czk']
                    lot_realized_pnl = lot_rev - lot_cost
                    purchase_fees=lot['purchase_fees'] 
                    partial_purchase_fees = purchase_fees * (take / lot['orig_qty'])
                    total_fees += (partial_purchase_fees + partial_trade_fee)
                    if save_xml and method == preferred_method and rate_type == preferred_rate:
                        usage_ratio = take / lot['orig_qty']
                        if usage_ratio > 0:
                            # This creates the variable name like 'used2024'
                            
                            # Get current usage (if any) and add the new fraction
                            current_val = float(lot['xml_el'].get(attr_name, 0))
                            new_val = round(current_val + usage_ratio, 8)
                            
                            # This is the part that actually WRITES it to the XML object
                            lot['xml_el'].set(attr_name, str(new_val))
                    if lot['date'][:4] < target_year:
                        detailed_log_entries.append({
                            'symbol': symbol, 
                            'type': 'BUY*',  # Marked with * to show it's historical
                            'qty': take, 
                            'price': lot['raw_buy_price'], 
                            'price_total': lot_cost, 
                            'pnl': 0.0, 
                            'is_sell': False, 
                            'date': lot['date'],
                            'trade_fees': partial_purchase_fees
                        })
                    detailed_log_entries.append({
                        'symbol': symbol, 
                        'type': 'SOLD', 
                        'qty': take, 
                        'price': lot['raw_buy_price'], 
                        'price_total': lot_rev, 
                        'pnl': lot_realized_pnl, 
                        'is_sell': True,
                        'sold_price': sold_price,
                        'date': sale['date'],
                        'trade_fees': partial_trade_fee,
                        'lot_cost': lot_cost,
                        'purchase_fees': partial_purchase_fees
                    })

                    total_cost += lot_cost
                    lot['qty'] -= take
                    needed -= take
                    if lot['qty'] <= 1e-9: 
                        inventory[symbol].pop(0)
            alternative_total_rev=0.0
            if rate_type == preferred_rate and method == preferred_method:
                final_inventory_for_excel = {k: [lot.copy() for lot in v] for k, v in inventory.items()}
            
                excel_tax_summary.append({
                'Preferred Method': f"{method} / {rate_type}",
                'Total Revenue (Příjmy) [CZK]': total_rev,
                'Total Purchase Cost (Výdaje) [CZK]': total_cost,
                'Total Fees [CZK]': total_fees,
                'Net Trade PnL [CZK]': total_rev - total_cost - total_fees
                })
                print(f"\n--- DETAILED TICKER REPORT: {rate_type.upper()} / {method} ---")
                # Updated header string:
                print(f"{'Symbol':<8} | {'Type':<6} | {'Date':<10} | {'Qty':>10} | {'Price':>10} | {'Total CZK':>14} | {'Fees CZK':>12} | {'PnL CZK':>14}")
                print("-" * 115)

                grouped = defaultdict(list)
                for entry in detailed_log_entries:
                    grouped[entry['symbol']].append(entry)
                
                for ticker in sorted(grouped.keys()):
                    t_buy_total, t_sell_total, t_pnl_total, t_fees_total, t_buy_sold_shares, t_pre_fees_pnl_total = 0.0, 0.0, 0.0, 0.0, 0.0,0.0
                    print(f"{ticker:<8} | {'Type':<6} | {'Date':<10} | {'Qty':>10} | {'Price':>10} | {'Total CZK':>14} | {'Fees CZK':>12} | {'PnL CZK':>14}")
                    print("-" * 115)

                    for e in grouped[ticker]:
                        excel_detailed_trades.append({
                            'Ticker': ticker,
                            'Type': e['type'],
                            'Date': e['date'],
                            'Qty': e['qty'],
                            'Price': e['price'],
                            'Total CZK': e['price_total'],
                            'Fees CZK': e.get('trade_fees', 0.0),
                            'PnL CZK': e['pnl']
                        })
                        # Get fees for this row (default to 0 for sells if not tracked there)
                        row_fees = e.get('trade_fees', 0.0)
                        print(f"{e['symbol']:<8} | {e['type']:<6} | {e['date']:<10} | {e['qty']:>10.2f} | {e['price']:>10.2f} | {e['price_total']:>14.2f} | {row_fees:>12.2f} | {e['pnl']:>14.2f}")

                        if e['is_sell']:

                            t_fees_total += e['purchase_fees'] + e['trade_fees']  # Add purchase fees for sells
                            t_sell_total += e['price_total']

                            t_pre_fees_pnl_total+= e['pnl']
                            t_buy_sold_shares += e['lot_cost']

                        else:
                            # CHANGE THIS LINE to only sum buys from the target year
                            if e['date'].startswith(target_year):
                                t_buy_total += e['price_total']
                    t_pnl_total =t_pre_fees_pnl_total -t_fees_total
                    

                    alternative_total_rev+= t_sell_total
                    print(f"{'---':<8} | {'SUM':<6} | {'Ticker: ' + ticker:<30}")
                    print(f"{'':<8} | {'':<6} | Total Cost:     {t_buy_total + t_fees_total:>14.2f} CZK") # <--- Purchase + Fees
                    print(f"{'':<8} | {'':<6} | SOLD Purchase: {t_buy_sold_shares:>14.2f} CZK")
                    print(f"{'':<8} | {'':<6} | SOLD Fees:     {t_fees_total:>14.2f} CZK")
                    print(f"{'':<8} | {'':<6} | Total Sold:     {t_sell_total:>14.2f} CZK")
                    print(f"{'':<8} | {'':<6} | Net PnL:        {t_pnl_total:>14.2f} CZK")

                    print("-" * 65)
                print("alternative_total_rev")
                print(alternative_total_rev)
                if alternative_total_rev != total_rev:
                    error_msg = f"❌ ERROR: Mismatch in total revenue calculation for {method} / {rate_type}! Calculated: {total_rev:.2f} CZK, Alternative Calc: {alternative_total_rev:.2f} CZK"
                    #sys.exit(error_msg)
            res_entry = {
            'rate': rate_type.upper(),
            'method': method,
            'rev': total_rev,
            'cost': total_cost,
            'fees': total_fees,
            'net': total_rev - total_cost - total_fees
        }
            simulation_results.append(res_entry)

        # NEW: Only save the TOTAL summary for your chosen method
        
            if rate_type == preferred_rate and method == preferred_method:
                excel_tax_summary.append({
                    'Preferred Method': f"{method} / {rate_type}",
                    'Total Revenue (Příjmy) [CZK]': total_rev,
                    'Total Purchase Cost (Výdaje) [CZK]': total_cost,
                    'Total Fees [CZK]': total_fees,
                    'Net Trade PnL [CZK]': total_rev - total_cost - total_fees
                })
                moje_dane['Prijmy'] = round(total_rev, 0)
                moje_dane['Vydaje'] = round(total_cost+total_fees, 0)
    # ... continue to save_xml logic ...
    if save_xml and parameter_info.upper() not in suspect_currencies: # Only save if we're not filtering by country, otherwise the XML modifications won't make sense:
        output_path = f"history/{target_year}+{preferred_method}.xml"
        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        print(f"✅ Saved simulation XML to: {output_path}")
        export_dividends_to_xlsx(root, preferred_rate, target_year, True)

        # Cleanup the attributes from the 'root' object in memory
        for el in root.findall(".//Trade"):
            if attr_name in el.attrib:
                del el.attrib[attr_name] 
    else:
        export_dividends_to_xlsx(root, preferred_rate, target_year, False)
    if save_xml and parameter_info.upper() not in suspect_currencies:
        intro_data = {
        'Property': ['Tax Payer', 'Target Year', 'Calculation Date', 'Method used', 'Rate used'],
        'Value': [parameter_info, target_year, datetime.now().strftime("%Y-%m-%d %H:%M"), preferred_method, preferred_rate]
    }
        holdings_list = []
        for sym, lots in final_inventory_for_excel.items():
            for lot in lots:
                if lot['qty'] > 1e-7:
                    holdings_list.append({
                        'Symbol': sym,
                        'Owned Qty': lot['qty'],
                        'Acquisition Date': lot['date'],
                        'Unit Cost (CZK)': lot['unit_cost'],
                        'Total Basis (CZK)': lot['qty'] * lot['unit_cost']
                    })
        
        excel_path = f"history/trade_report_{target_year}_.xlsx"
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            # Save Introduction first
            pd.DataFrame(intro_data).to_excel(writer, sheet_name='Introduction', index=False)
            
            # Save the rest
            pd.DataFrame(excel_tax_summary).to_excel(writer, sheet_name='Tax_Summary', index=False)
            pd.DataFrame(excel_detailed_trades).to_excel(writer, sheet_name='Detailed_Trades', index=False)
            pd.DataFrame(holdings_list).to_excel(writer, sheet_name='Current_Holdings', index=False)
            workbook = writer.book
            for sheetname in workbook.sheetnames:
                sheet = workbook[sheetname]
                for col in sheet.columns:
                    max_length = 0
                    column_letter = col[0].column_letter # Get the column name (A, B, C...)
                    
                    for cell in col:
                        try:
                            # Measure the length of the cell value
                            if cell.value is not None:
                                # Add a little extra space (2) so it's not too tight
                                current_length = len(str(cell.value))
                                if current_length > max_length:
                                    max_length = current_length
                        except:
                            pass
                    
                    # Set the width (limit to 50 so it doesn't get crazy wide)
                    adjusted_width = min(max_length + 2, 50)
                    sheet.column_dimensions[column_letter].width = adjusted_width
                for row in sheet.iter_rows():
                    for cell in row:
                        # Setting color to None allows the UI (LibreOffice) 
                        # to choose Black or White based on your background.
                        cell.font = Font(color=None, name='Calibri', size=11)
        
        print(f"✅ Full report (3 sheets) saved to: {excel_path}")
    # Final summary table
    dividends_method=[sum_net_dividends_daily, sum_net_dividends_uniform]
    print("\n" + "=" * 140)
    print(f"{'Rate Type':<10} | {'Method':<15} | {'Revenue (CZK)':>15} | {'Purchase (CZK)':>15} | {'Fees (CZK)':>15} | {'Net trade':>15} | {'Dividends':>15} | {'Total Net':>15}")
    print("-" * 140)
    
    for res in simulation_results:
        if res['rate']=="DAILY":
            res['dividend']=sum_net_dividends_daily
        elif res['rate']=="UNIFORM":
            res['dividend']=sum_net_dividends_uniform
        net_final=res['net']+res['dividend']
        print(f"{res['rate']:<10} | {res['method']:<15} | {res['rev']:>15.2f} | {res['cost']:>15.2f} | {res['fees']:>15.2f} | {res['net']:>15.2f} | {res['dividend']:>15.2f} | {net_final:>15.2f}")

    print("=" * 140)
    if country_mismatches:
        print("\n⚠️ COUNTRY DATA MISMATCH REPORT (Ticker | Issuer | ISIN | YF)")
        print("-" * 60)
        for report_line in country_mismatches.values():
            print(report_line)
        print("-" * 60)
    print("\n" + "═" * 40)
    print("📋 FINAL DATA FOR TAX FILLING")
    print("═" * 32)
    prev_prefix = None

    for key, value in moje_dane.items():
        current_prefix = key[:2]  # Get the first two letters
        
        # If the prefix changed (and it's not the very first item), print separator
        if prev_prefix is not None and current_prefix != prev_prefix:
            print("-" * 32)
        
        # Adjusted widths: key column is 20, value column is 8
        print(f"{key:<20}: {int(value):>8d}")
        
        prev_prefix = current_prefix

    print("═" * 32)
def run_cash_tax_simulation(xml_file='final.xml', save_xml=False):
    global attr_name
    attr_name = f"used-{target_year}"
    excel_detailed_trades = []
    excel_tax_summary = []
    final_inventory_for_excel = {}
    
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
    except Exception as e:
        print(f"❌ Error: {e}"); return

    # 🔥 The critical change: Target ONLY 'CASH' trades
    trades_elements = [t for t in root.findall(".//Trade") if t.get('assetCategory') == 'CASH']
    trades_elements.sort(key=lambda x: x.get('tradeDate'))

    rates = ['daily', 'uniform']
    methods = ['FIFO', 'LIFO', 'GLOBAL_MAX_LOSS']
    
    print(f"\n🌍 GLOBAL CASH/FOREX TAX SIMULATION FOR YEAR: {target_year}")
    simulation_results = []

    for rate_type in rates:
        rate_key = f"{rate_type}RateCZK"
        for method in methods:
            inventory = defaultdict(list)
            sales_to_process = []
            detailed_log_entries = [] 
            total_rev, total_cost, total_fees = 0.0, 0.0, 0.0
            
            for t_el in trades_elements:
                t = t_el.attrib.copy()
                
                # Forex pairs (like USD.MYR) don't have a single "currency" in the same way stocks do,
                # but if you use your suspect_currencies filter, we keep it here just in case.
                if parameter_info.upper() in suspect_currencies:
                    if t.get('currency', '').upper() != parameter_info:
                        continue
                
                symbol = t['symbol']
                date = t['tradeDate']
                trade_year = int(date[:4])
                
                # For CASH, quantity is the base currency amount being exchanged
                qty = abs(float(t['quantity']))
                price = float(t['tradePrice'])
                rate = float(t.get(rate_key, 1.0))

                if parameter_info.upper() in suspect_currencies:
                    rate = 1.0

                # IBKR Cash commissions are typically converted using these fields
                comm_czk = abs(float(t.get(f'{rate_type}CommissionCZK', 0.0)))
                tax = abs(float(t.get('taxes', 0.0))) * rate
                trade_fees = comm_czk + tax

                if t['buySell'] == 'BUY':
                    # --- Check for previous years' usage ---
                    already_used_ratio = 0.0
                    for attr, val in t_el.attrib.items():
                        if attr.startswith("used-"):
                            try:
                                attr_year = int(attr.split("-")[1])
                                if attr_year < int(target_year):
                                    already_used_ratio += float(val)
                            except (ValueError, IndexError):
                                continue
                    
                    available_qty = qty * (1.0 - already_used_ratio)
                    
                    if already_used_ratio > 1.0000001:
                        error_msg = f"⚠️ WARNING: {symbol} CASH lot from {date} is overused! Total used: {already_used_ratio*100:.2f}%. Fixing to 0."
                        sys.exit(error_msg)
                    
                    if available_qty > 1e-9:
                        cost = (price * qty * rate)
                        inventory[symbol].append({
                            'qty': available_qty,
                            'orig_qty': qty,
                            'unit_cost': cost/qty, 
                            'date': date,
                            'raw_buy_price': price,
                            'purchase_fees': trade_fees,
                            'xml_el': t_el 
                        })

                    if trade_year == int(target_year): 
                        detailed_log_entries.append({
                            'symbol': symbol, 'type': 'BUY', 'qty': available_qty, 'price': price,
                            'price_total': price * available_qty * rate, 'pnl': 0.0, 'is_sell': False, 'date': date,
                            'trade_fees': trade_fees  
                        })
                
                elif t['buySell'] == 'SELL':
                    sell_price_czk = price * rate
                    rev = (sell_price_czk * qty) 

                    if trade_year < int(target_year):
                        continue
                    elif trade_year == int(target_year):
                        total_rev += rev
                        sales_to_process.append({
                            'symbol': symbol, 'qty': qty, 'sell_price_czk': sell_price_czk, 
                            'date': date, 'trade_fees': trade_fees
                        })
                    else:
                        continue
                else:
                    print("NEW BUY/SELL TYPE ENCOUNTERED IN CASH!")
            
            # --- Process target year sales ---
            for sale in sales_to_process:
                symbol, needed = sale['symbol'], sale['qty']
                trade_fees = sale['trade_fees']

                if method == 'FIFO': 
                    inventory[symbol].sort(key=lambda x: x['date'])
                elif method == 'LIFO': 
                    inventory[symbol].sort(key=lambda x: x['date'], reverse=True)
                elif method == 'GLOBAL_MAX_LOSS':
                    inventory[symbol].sort(key=lambda x: x['unit_cost'], reverse=True)

                while needed > 0 and inventory[symbol]:
                    lot = inventory[symbol][0]
                    current_lot_qty = lot['qty']
                    take = min(needed, current_lot_qty)
                    
                    share_fee = take / sale['qty']
                    partial_trade_fee = trade_fees * share_fee
                    #tady je chyba - needed je 0


                    lot_cost = take * lot['unit_cost']
                    lot_rev = take * sale['sell_price_czk']
                    sold_price = sale['sell_price_czk']
                    lot_realized_pnl = lot_rev - lot_cost
                    
                    purchase_fees = lot['purchase_fees'] 
                    partial_purchase_fees = purchase_fees * (take / lot['orig_qty'])
                    total_fees += (partial_purchase_fees + partial_trade_fee)
                    
                    # 🔥 ADDING THE 'USED-YEAR' ATTRIBUTE 🔥
                    if save_xml and method == preferred_method and rate_type == preferred_rate:
                        usage_ratio = take / lot['orig_qty']
                        if usage_ratio > 0:
                            current_val = float(lot['xml_el'].get(attr_name, 0))
                            new_val = round(current_val + usage_ratio, 8)
                            lot['xml_el'].set(attr_name, str(new_val))
                            
                    if lot['date'][:4] < target_year:
                        detailed_log_entries.append({
                            'symbol': symbol, 'type': 'BUY*', 'qty': take, 'price': lot['raw_buy_price'], 
                            'price_total': lot_cost, 'pnl': 0.0, 'is_sell': False, 'date': lot['date'],
                            'trade_fees': partial_purchase_fees
                        })
                        
                    detailed_log_entries.append({
                        'symbol': symbol, 'type': 'SOLD', 'qty': take, 'price': lot['raw_buy_price'], 
                        'price_total': lot_rev, 'pnl': lot_realized_pnl, 'is_sell': True,
                        'sold_price': sold_price, 'date': sale['date'],
                        'trade_fees': partial_trade_fee, 'lot_cost': lot_cost,
                        'purchase_fees': partial_purchase_fees
                    })

                    total_cost += lot_cost
                    lot['qty'] -= take
                    needed -= take
                    if lot['qty'] <= 1e-9: 
                        inventory[symbol].pop(0)

            alternative_total_rev = 0.0
            
            if rate_type == preferred_rate and method == preferred_method:
                final_inventory_for_excel = {k: [lot.copy() for lot in v] for k, v in inventory.items()}
            
                excel_tax_summary.append({
                    'Preferred Method': f"{method} / {rate_type}",
                    'Total Revenue (Příjmy) [CZK]': total_rev,
                    'Total Purchase Cost (Výdaje) [CZK]': total_cost,
                    'Total Fees [CZK]': total_fees,
                    'Net Trade PnL [CZK]': total_rev - total_cost - total_fees
                })
                
                print(f"\n--- DETAILED FOREX/CASH REPORT: {rate_type.upper()} / {method} ---")
                print(f"{'Pair':<8} | {'Type':<6} | {'Date':<10} | {'Qty':>10} | {'Rate':>10} | {'Total CZK':>14} | {'Fees CZK':>12} | {'PnL CZK':>14}")
                print("-" * 115)

                grouped = defaultdict(list)
                for entry in detailed_log_entries:
                    grouped[entry['symbol']].append(entry)
                
                for ticker in sorted(grouped.keys()):
                    t_buy_total, t_sell_total, t_pnl_total, t_fees_total, t_buy_sold_shares, t_pre_fees_pnl_total = 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
                    
                    print(f"{ticker:<8} | {'Type':<6} | {'Date':<10} | {'Qty':>10} | {'Rate':>10} | {'Total CZK':>14} | {'Fees CZK':>12} | {'PnL CZK':>14}")
                    print("-" * 115)

                    for e in grouped[ticker]:
                        excel_detailed_trades.append({
                            'Pair': ticker, 'Type': e['type'], 'Date': e['date'],
                            'Qty': e['qty'], 'Rate': e['price'], 'Total CZK': e['price_total'],
                            'Fees CZK': e.get('trade_fees', 0.0), 'PnL CZK': e['pnl']
                        })
                        
                        row_fees = e.get('trade_fees', 0.0)
                        print(f"{e['symbol']:<8} | {e['type']:<6} | {e['date']:<10} | {e['qty']:>10.2f} | {e['price']:>10.2f} | {e['price_total']:>14.2f} | {row_fees:>12.2f} | {e['pnl']:>14.2f}")

                        if e['is_sell']:
                            t_fees_total += e['purchase_fees'] + e['trade_fees']
                            t_sell_total += e['price_total']
                            t_pre_fees_pnl_total += e['pnl']
                            t_buy_sold_shares += e['lot_cost']
                        else:
                            if e['date'].startswith(target_year):
                                t_buy_total += e['price_total']
                                
                    t_pnl_total = t_pre_fees_pnl_total - t_fees_total
                    alternative_total_rev += t_sell_total
                    
                    print(f"{'---':<8} | {'SUM':<6} | {'Pair: ' + ticker:<30}")
                    print(f"{'':<8} | {'':<6} | Total Cost:     {t_buy_total + t_fees_total:>14.2f} CZK")
                    print(f"{'':<8} | {'':<6} | SOLD Purchase: {t_buy_sold_shares:>14.2f} CZK")
                    print(f"{'':<8} | {'':<6} | SOLD Fees:     {t_fees_total:>14.2f} CZK")
                    print(f"{'':<8} | {'':<6} | Total Sold:     {t_sell_total:>14.2f} CZK")
                    print(f"{'':<8} | {'':<6} | Net PnL:        {t_pnl_total:>14.2f} CZK")
                    print("-" * 65)
                
                # --- The Error Print You Requested ---
                #if abs(alternative_total_rev - total_rev) > 0.1:
                    #error_msg = f"❌ ERROR: Mismatch in CASH total revenue calculation for {method} / {rate_type}! Calculated: {total_rev:.2f} CZK, Alternative Calc: {alternative_total_rev:.2f} CZK\n(You likely have CASH sales with no matching BUY history)"
                    #sys.exit(error_msg)

            res_entry = {
                'rate': rate_type.upper(), 'method': method, 'rev': alternative_total_rev,
                'cost': total_cost, 'fees': total_fees, 'net': alternative_total_rev - total_cost - total_fees
            }
            simulation_results.append(res_entry)

    for res in simulation_results:
        print(f"{res['rate']:<10} | {res['method']:<15} | {res['rev']:>15.2f} | {res['cost']:>15.2f} | {res['fees']:>15.2f} | {res['net']:>15.2f} |")
    # --- XML Saving logic using your exact requested path ---
    if save_xml and parameter_info.upper() not in suspect_currencies:
        output_path = f"history/{target_year}+{preferred_method}.xml"
        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        print(f"✅ Saved CASH simulation XML to: {output_path}")

        # Cleanup the attributes from the 'root' object in memory
        for el in root.findall(".//Trade"):
            if attr_name in el.attrib:
                del el.attrib[attr_name]
# --- Usage ---
    
if __name__ == "__main__":
    # If there is any argument after the filename, save_enabled becomes True
    save_enabled = len(sys.argv) > 1
    if save_enabled:
        parameter_info = " ".join(sys.argv[1:])
    if save_enabled and not os.path.exists('history'):
        os.makedirs('history')
    # 3. REDIRECT PRINTING HERE
    if save_enabled and parameter_info.upper() not in suspect_currencies:
        log_file_path = f"./history/print{target_year}.txt"
        f = open(log_file_path, 'w', encoding='utf-8') # 'a' appends, 'w' overwrites
    
        original_stdout = sys.stdout
        sys.stdout = Tee(sys.stdout, f)
    try:
        # 4. Run your simulation
        run_tax_simulations(save_xml=save_enabled)
        stock_saved_file = f"history/{target_year}+{preferred_method}.xml"
        run_cash_tax_simulation(xml_file=stock_saved_file, save_xml=True)

    finally:
        # 5. Clean up: Close the file and reset stdout
        if save_enabled and parameter_info.upper() not in suspect_currencies:
            sys.stdout = original_stdout
            f.close()
