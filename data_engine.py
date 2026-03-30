import pandas as pd
from datetime import datetime
import io, zipfile, requests, os

BHAVCOPY_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "bhavcopy")

class DataEngine:
    def __init__(self):
        os.makedirs(BHAVCOPY_DIR, exist_ok=True)

    def fetch_atm_strike(self, index_name, data_date):
        """Fetches spot price and returns rounded ATM strike."""
        try:
            idx = index_name.strip().upper()
            data = self.fetch_bhavcopy(data_date)
            if not data or data.get('index') is None:
                return None
            
            df_idx = data['index']
            i_lbl = "NIFTY 50" if idx == "NIFTY" else ("NIFTY BANK" if idx == "BANKNIFTY" else "NIFTY FIN SERVICE")
            matches = df_idx[df_idx['Index Name'].str.upper() == i_lbl]
            
            if matches.empty:
                return None
            
            spot = float(matches.iloc[0]['Closing Index Value'])
            gap = 50 if idx in ["NIFTY", "FINNIFTY"] else 100
            atm = round(spot / gap) * gap
            return int(atm)
        except Exception as e:
            print(f"ATM Error: {e}")
            return None

    def fetch_bhavcopy(self, date_str, force_download=False):
        """Loads Bhavcopy from disk or downloads from NSE archives."""
        try: ds = datetime.strptime(date_str, "%Y-%m-%d")
        except: return None
        
        ds_ymd = ds.strftime("%Y%m%d")
        ds_dmy = ds.strftime('%d%b%Y').upper()
        ds_dmy_num = ds.strftime("%d%m%Y")
        
        fo_dest = os.path.join(BHAVCOPY_DIR, f"fo_{ds_ymd}.zip")
        idx_dest = os.path.join(BHAVCOPY_DIR, f"idx_{ds_dmy_num}.csv")
        results = {'fo': None, 'index': None}

        if not force_download:
            if os.path.exists(fo_dest):
                try:
                    with zipfile.ZipFile(fo_dest) as z:
                        results['fo'] = pd.read_csv(z.open(z.namelist()[0]))
                except: pass
            if os.path.exists(idx_dest):
                try: results['index'] = pd.read_csv(idx_dest)
                except: pass
            
            if results['fo'] is not None and results['index'] is not None:
                return results

        session = requests.Session()
        session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"})
        
        # FO Download
        if results['fo'] is None:
            fo_urls = [
                f"https://archives.nseindia.com/content/fo/BhavCopy_NSE_FO_0_0_0_{ds_ymd}_F_0000.csv.zip",
                f"https://nsearchives.nseindia.com/content/fo/fo{ds_dmy}bhav.csv.zip"
            ]
            for url in fo_urls:
                try:
                    r = session.get(url, timeout=12)
                    if r.status_code == 200:
                        with open(fo_dest, "wb") as f: f.write(r.content)
                        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                            results['fo'] = pd.read_csv(z.open(z.namelist()[0]))
                        break
                except: continue

        # Index Download
        if results['index'] is None:
            idx_urls = [
                f"https://nsearchives.nseindia.com/content/indices/ind_close_all_{ds_dmy_num}.csv",
                f"https://archives.nseindia.com/content/indices/ind_close_all_{ds_dmy_num}.csv"
            ]
            for url in idx_urls:
                try:
                    r = session.get(url, timeout=12)
                    if r.status_code == 200:
                        with open(idx_dest, "w") as f: f.write(r.text)
                        results['index'] = pd.read_csv(io.StringIO(r.text))
                        break
                except: continue

        if results['fo'] is not None and results['index'] is not None:
            return results
        
        return None
