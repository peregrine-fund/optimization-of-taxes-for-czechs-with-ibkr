import pandas as pd
import lxml.etree as ET
from collections import defaultdict

def update_xml_and_calculate_from_ibkr(csv_filename='2024.csv', xml_filename='final.xml'):
    try:
        # --- 1. LOAD AND CLEAN CSV ---
        # We use header=None because the first row is often a 'Trades' metadata row
        raw_df = pd.read_csv(csv_filename, header=None)
        
        # Only keep rows where the second column is 'Data'
        df_csv = raw_df[raw_df[1] == 'Data'].copy()
        
        # Mapping columns based on the IBKR 'Header' row structure
        columns = [
            'Trades_Prefix', 'DataDiscriminator', 'AssetClass', 'AssetCategory', 
            'Currency', 'Symbol', 'DateTime', 'Quantity', 'TPrice', 'CPrice', 
            'Proceeds', 'CommFee', 'Basis', 'RealizedPL', 'MTMPL', 'Code'
        ]
        df_csv.columns = columns[:len(df_csv.columns)]

        # Clean numeric columns (remove commas, handle quotes)
        num_cols = ['Quantity', 'Proceeds', 'Basis', 'RealizedPL']
        for col in num_cols:
            df_csv[col] = pd.to_numeric(df_csv[col].astype(str).str.replace(',', '').str.replace('"', ''), errors='coerce').fillna(0)

        # Process Dates for XML matching
        df_csv['DT'] = pd.to_datetime(df_csv['DateTime'].str.replace('"', '').str.strip())
        df_csv['XMLDate'] = df_csv['DT'].dt.strftime('%Y%m%d')

        # --- 2. FINANCIAL CALCULATION (Trusting IBKR Columns) ---
        # We group by symbol to get the final tax figures
        summary = defaultdict(lambda: {'Revenue': 0.0, 'Cost': 0.0, 'Net': 0.0})
        
        # Track quantities for the Usage % logic
        total_sold_qty = defaultdict(float)

        for _, row in df_csv.iterrows():
            sym = row['Symbol']
            if row['Quantity'] < 0:  # This is a Closing Trade (SELL)
                summary[sym]['Revenue'] += row['Proceeds']
                summary[sym]['Cost'] += abs(row['Basis'])
                summary[sym]['Net'] += row['RealizedPL']
                total_sold_qty[sym] += abs(row['Quantity'])

        # --- 3. PRINT FINANCIAL SUMMARY ---
        print(f"\n{'Symbol':<10} | {'Revenue':>12} | {'Cost (Basis)':>12} | {'Realized P/L':>12}")
        print("-" * 58)
        
        t_rev, t_cost, t_net = 0, 0, 0
        for sym, v in sorted(summary.items()):
            print(f"{sym:<10} | {v['Revenue']:>12.2f} | {v['Cost']:>12.2f} | {v['Net']:>12.2f}")
            t_rev += v['Revenue']
            t_cost += v['Cost']
            t_net += v['Net']
            
        print("-" * 58)
        print(f"{'TOTAL':<10} | {t_rev:>12.2f} | {t_cost:>12.2f} | {t_net:>12.2f}")

        # --- 4. UPDATE XML USAGE PERCENTAGES ---
        # We allocate the 'total_sold_qty' against 'BUY' lots to see which are exhausted
        parser = ET.XMLParser(remove_blank_text=True)
        tree = ET.parse(xml_filename, parser)
        root = tree.getroot()

        for symbol in total_sold_qty:
            sold_rem = total_sold_qty[symbol]
            # Find all BUY trades for this symbol in the XML, sorted by date
            buy_trades = root.xpath(f".//Trade[@symbol='{symbol}' and @buySell='BUY']")
            # Sort by date (assuming FIFO matching for the usage tag)
            buy_trades.sort(key=lambda x: x.get('tradeDate'))

            for trade in buy_trades:
                buy_qty = float(trade.get('quantity', 0))
                if sold_rem <= 0:
                    trade.set("used2024", "0.0%")
                elif sold_rem >= buy_qty:
                    trade.set("used2024", "100.0%")
                    sold_rem -= buy_qty
                else:
                    pct = round((sold_rem / buy_qty) * 100, 2)
                    trade.set("used2024", f"{pct}%")
                    sold_rem = 0

        tree.write(xml_filename, encoding="utf-8", xml_declaration=True, pretty_print=True)
        print(f"\nXML '{xml_filename}' updated based on IBKR closing logic.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    update_xml_and_calculate_from_ibkr()