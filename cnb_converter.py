import pandas as pd
import requests
import io
import calendar
from datetime import datetime, timedelta

class CNBConverter:
    def __init__(self):
        self.daily_rate_cache = {}
        self.daily_multipliers_cache = {}
        self.other_rate_cache = {}

    def _get_daily_data(self, year):
        if year in self.daily_rate_cache:
            return self.daily_rate_cache[year], self.daily_multipliers_cache[year]
        
        url = f"https://www.cnb.cz/en/financial-markets/foreign-exchange-market/central-bank-exchange-rate-fixing/central-bank-exchange-rate-fixing/year.txt?year={year}"
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            df = pd.read_csv(io.StringIO(response.text), sep='|')
            
            multipliers, new_cols = {}, {}
            for col in df.columns:
                if col == 'Date': continue
                parts = col.split()
                if len(parts) == 2:
                    amt, curr = parts
                    new_cols[col] = curr
                    multipliers[curr] = float(amt)
            
            df.rename(columns=new_cols, inplace=True)
            df['Date'] = pd.to_datetime(df['Date'], dayfirst=True).dt.strftime('%Y%m%d')
            df.set_index('Date', inplace=True)
            
            self.daily_rate_cache[year] = df
            self.daily_multipliers_cache[year] = multipliers
            return df, multipliers
        except:
            return None, None

    def _get_obscure_rate(self, year, month, target_currency):
        # Shift back 1 month to get the rate declaration date
        target_dt = datetime(int(year), int(month), 1) - timedelta(days=1)
        req_year, req_month = target_dt.year, target_dt.month

        cache_key = f"{req_year}{req_month:02d}"
        if cache_key in self.other_rate_cache and target_currency in self.other_rate_cache[cache_key]:
            return self.other_rate_cache[cache_key][target_currency]

        url = f"https://www.cnb.cz/en/financial-markets/foreign-exchange-market/fx-rates-of-other-currencies/fx-rates-of-other-currencies/fx_rates.txt?month={req_month}&year={req_year}"
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            lines = response.text.strip().split('\n')
            start_index = next(i for i, line in enumerate(lines) if "Country" in line and "Code" in line)
            
            df = pd.read_csv(io.StringIO("\n".join(lines[start_index:])), sep='|')
            df.columns = [c.strip() for c in df.columns]

            if cache_key not in self.other_rate_cache:
                self.other_rate_cache[cache_key] = {}
                
            for _, row in df.iterrows():
                code = str(row['Code']).strip()
                rate = float(str(row['Rate']).replace(',', '.'))
                amt = float(str(row['Amount']).replace(',', '.'))
                # Store rate for 1 unit
                self.other_rate_cache[cache_key][code] = rate / amt

            return self.other_rate_cache[cache_key].get(target_currency)
        except:
            return None
    #PLEASE IN FUTURE CHECK WHETHER GET OBSCURE RATE IS TRUE. I MOVED month by 1 before as stated on cnb site that its for next period of month the forex list.

    def get_rate(self, currency="TWD", date_str=None, method="daily"):
        """
        Returns the exchange rate (CZK for 1 unit of foreign currency).
        date_str: 'YYYYMMDD'. method: 'daily' or 'uniform'.
        """
        if not date_str:
            date_str = datetime.now().strftime('%Y%m%d')
        year, month = date_str[:4], int(date_str[4:6])
        
        # Try Daily Fixings (USD, EUR...)
        df_daily, multipliers = self._get_daily_data(year)
        if df_daily is not None and currency in df_daily.columns:
            mult = multipliers.get(currency, 1.0)
            if method == "uniform":
                rates = []
                for m in range(1, 13):
                    last_day = calendar.monthrange(int(year), m)[1]
                    check_dt = datetime(int(year), m, last_day)
                    for _ in range(7):
                        look = check_dt.strftime('%Y%m%d')
                        if look in df_daily.index:
                            val = float(str(df_daily.loc[look, currency]).replace(',', ''))
                            rates.append(val / mult)
                            break
                        check_dt -= timedelta(days=1)
                final_rate = sum(rates) / len(rates) if rates else None
            else:
                check_dt = datetime.strptime(date_str, "%Y%m%d")
                final_rate = None
                for _ in range(7):
                    look = check_dt.strftime('%Y%m%d')
                    if look in df_daily.index:
                        val = float(str(df_daily.loc[look, currency]).replace(',', ''))
                        final_rate = val / mult
                        break
                    check_dt -= timedelta(days=1)
        
        # Try Obscure Currencies (TWD...)
        else:
            if method == "uniform":
                rates = [self._get_obscure_rate(year, m, currency) for m in range(1, 13)]
                rates = [r for r in rates if r]
                final_rate = sum(rates) / len(rates) if rates else None
            else:
                final_rate = self._get_obscure_rate(year, month, currency)

        # We return the rate for 1 unit (e.g., 23.45 CZK per 1 USD)
        return round(final_rate, 4) if final_rate else None