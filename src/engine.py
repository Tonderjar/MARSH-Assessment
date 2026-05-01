import os
from datetime import datetime
import pandas as pd


REQUIRED_COLS = {"Client", "Country", "Currency", "Transaction"}

############################################################################################
#                           File reading and parsing logic                                 #
############################################################################################

def _find_header_row(raw: pd.DataFrame) -> int | None:
    """Return the integer row index whose values contain all four required column names."""
    for i, row in raw.iterrows():
        if REQUIRED_COLS.issubset({str(v).strip() for v in row}):
            return i
    return None


def _read_file(filepath: str) -> pd.DataFrame | None:
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext == ".csv":
            # sep=None + engine='python' lets pandas auto-detect comma vs semicolon
            raw = pd.read_csv(filepath, header=None, sep=None, engine="python", dtype=str)
        else:
            raw = pd.read_excel(filepath, header=None, dtype=str)
    except Exception as e:
        print(f"Skipping {filepath}: {e}")
        return None

    header_row = _find_header_row(raw)
    if header_row is None:
        print(f"Skipping {filepath}: required columns not found")
        return None

    # Use the located row as column names, then drop it from the data
    raw.columns = [str(v).strip() for v in raw.iloc[header_row]]
    df = raw.iloc[header_row + 1:].reset_index(drop=True)

    # Keep only the four required columns (ignores any extra/empty columns)
    df = df[list(REQUIRED_COLS)].copy()
    return df


def load_all_transactions(data_root: str) -> pd.DataFrame:
    """
    Walk the data folder tree and concatenate every DataTrans_*.csv / *.xlsx file
    into a single DataFrame.

    Expected layout:
        <data_root>/<year>/<NN Month Name>/<DD-MM-YY>/DataTrans_N.{csv,xlsx}

    Handles:
        - Semicolon-delimited CSVs (auto-detected)
        - Extra leading empty columns
        - Header row not on row 0
        - Any column order

    Added column:
        log_date (datetime) — parsed from the DD-MM-YY folder name.
    """
    frames = []

    for dirpath, _, filenames in os.walk(data_root):
        for filename in sorted(filenames):
            if not filename.startswith("DataTrans_"):
                continue
            if os.path.splitext(filename)[1].lower() not in {".csv", ".xlsx"}:
                continue

            date_folder = os.path.basename(dirpath)
            try:
                log_date = datetime.strptime(date_folder, "%d-%m-%y")
            except ValueError:
                log_date = pd.NaT

            df = _read_file(os.path.join(dirpath, filename))
            if df is None:
                continue

            df["log_date"] = log_date
            frames.append(df)

    if not frames:
        return pd.DataFrame(columns=["Client", "Country", "Currency", "Transaction", "log_date"])

    return pd.concat(frames, ignore_index=True)

def load_fx_rates(fx_path: str) -> dict[str, float]:
    """
    Parse the '2023 Budget FX Rates.xlsx' file and return a dict mapping
    currency CODE (e.g. 'USD', 'EUR') to its RATE (float, relative to USD).
    Uses _find_header_row so it is resilient to extra title rows at the top.
    """
    raw = pd.read_excel(fx_path, header=None, dtype=str)
    header_row = _find_header_row(raw)
    if header_row is None:
        raise ValueError(f"Could not find COUNTRY/CURRENCY/CODE/RATE header in {fx_path}")

    raw.columns = [str(v).strip() for v in raw.iloc[header_row]]
    df = raw.iloc[header_row + 1:].reset_index(drop=True)
    df = df[["CODE", "RATE"]].dropna(subset=["CODE", "RATE"])
    df["RATE"] = pd.to_numeric(df["RATE"], errors="coerce")
    return dict(zip(df["CODE"].str.strip(), df["RATE"]))

############################################################################################
#                           Report generation logic                                        #
############################################################################################

def currency_client_transactions(df: pd.DataFrame) -> pd.DataFrame:
    """
    Count the transactions by Client and Currency.
    """

    # Group by Client and Currency
    counts = (
        df.groupby(["Client", "Currency"]).size().reset_index(name="Count")
    )
    return counts
    
def client_totals_usd(df: pd.DataFrame, fx_rates: dict[str, float]) -> pd.DataFrame:
    """
    Sum the Transaction amounts by Client, converting to USD using fx_rates.

    fx_rates: dict mapping currency CODE -> rate vs USD,
              as returned by load_fx_rates().
    Transactions in currencies not present in fx_rates are excluded from the total.
    """
    df = df.copy()
    df["Transaction"] = pd.to_numeric(df["Transaction"], errors="coerce")
    df["Exchange_Rate"] = df["Currency"].map(fx_rates)
    df["Transaction_USD"] = df["Transaction"] * df["Exchange_Rate"]

    totals = (
        df.groupby("Client")["Transaction_USD"].sum().reset_index(name="Total_USD")
    )
    return totals