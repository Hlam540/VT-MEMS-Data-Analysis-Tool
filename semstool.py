import pandas as pd
import matplotlib.pyplot as plt

# Show all columns and pretty-print floats
pd.set_option('display.max_columns', None)
pd.set_option('display.float_format', '{:.6f}'.format)

class SEMSReader:
    def __init__(self,
                 path: str,
                 sheet: str = "Sheet1",
                 date_col: int = 0,
                 time_col: int = 1,
                 first_data_col: int = 2):
        """
        Reader for SEMS Excel data.

        :param path:           filepath to Excel file
        :param sheet:          sheet name
        :param date_col:       zero-based index of the column with dates (YYMMDD)
        :param time_col:       zero-based index of the column with times (HH:MM:SS or h:mm:ss AM/PM)
        :param first_data_col: zero-based index where the 30 bin counts begin (column C => 2)
        """
        self.path = path
        self.sheet = sheet
        self.date_col = date_col
        self.time_col = time_col
        self.first_data_col = first_data_col
        self.df = None
        self.bin_centers = None

    def load(self):
        raw = pd.read_excel(self.path, sheet_name=self.sheet, header=None)

        # Row 2 (index 1) has the 30 bin centers in nm
        self.bin_centers = raw.iloc[1, self.first_data_col:].astype(float)

        # Data starts on row 3 (index 2)
        data = raw.iloc[2:, :]

        # Build one string per row: "YYMMDD HH:MM:SS[ AM/PM]"
        dt_str = (
            data.iloc[:, self.date_col].astype(str).str.zfill(6) + ' ' +
            data.iloc[:, self.time_col].astype(str).str.strip()
        )

        # Try 12-hr with AM/PM first, then fall back to 24-hr
        timestamps = pd.to_datetime(dt_str,
                                    format="%y%m%d %I:%M:%S %p",
                                    errors="coerce")
        if timestamps.isna().any():
            mask = timestamps.isna()
            timestamps.loc[mask] = pd.to_datetime(
                dt_str.loc[mask],
                format="%y%m%d %H:%M:%S",
                errors="raise"
            )

        # Extract counts, index by timestamp, columns by bin center
        counts = data.iloc[:, self.first_data_col:].astype(float)
        counts.index = timestamps
        counts.columns = self.bin_centers
        counts.index.name = "Timestamp"

        self.df = counts

    def mean_counts(self, start, end) -> pd.Series:
        """
        Compute mean counts per bin between two timestamps.

        :param start: inclusive start (string or pd.Timestamp)
        :param end:   inclusive end
        :return:      Series indexed by bin center (nm), name 'Size (nm)'
        """
        s = self.df.loc[start:end].mean()
        s.index.name = "Size (nm)"
        return s

class SEMSPenetrationEfficiency:
    def __init__(self, reader: SEMSReader):
        self.reader = reader

    def compute_pe(self,
                   up_start, up_end,
                   down_start, down_end,
                   up_factor: float = 1.0,
                   down_factor: float = 1.0) -> pd.DataFrame:
        orig_up   = self.reader.mean_counts(up_start, up_end)
        orig_down = self.reader.mean_counts(down_start, down_end)
        corr_up   = orig_up * up_factor
        corr_down = orig_down * down_factor
        pe        = (corr_down / corr_up) * 100

        df = pd.DataFrame({
            "Original Upstream":    orig_up,
            "Original Downstream":  orig_down,
            "Corrected Upstream":   corr_up,
            "Corrected Downstream": corr_down,
            "Penetration Efficiency (%)":  pe
        })
        df.index.name = "Size (nm)"
        return df

    def plot_pe(self, pe_df: pd.DataFrame):
        diam = pe_df.index.astype(float)
        pe   = pe_df["Penetration Efficiency (%)"]
        fig, ax = plt.subplots()
        ax.semilogx(diam, pe, marker="o")
        ax.set_xlabel("Particle diameter (nm)")
        ax.set_ylabel("Penetration Efficiency (%)")
        ax.grid(True, which="both", linestyle="--", linewidth=0.5)
        import matplotlib.ticker as ticker
        ax.xaxis.set_major_formatter(ticker.FormatStrFormatter("%.3g"))
        ax.xaxis.set_minor_formatter(ticker.NullFormatter())
        plt.tight_layout()
        plt.show()

if __name__ == "__main__":
    # --- Configuration ---
    file_path  = "sems file updated.xlsx"
    sheet_name = "Sheet1"

    # Column C (2) has Bin_Cnts1 … Bin_Cnts30; row 2 has the bin sizes
    reader = SEMSReader(
        path=file_path,
        sheet=sheet_name,
        date_col=0,
        time_col=1,
        first_data_col=2
    )
    reader.load()

    # --- Prompt for correction factors ---
    raw_up   = input("Enter upstream factor (e.g. 1375/800): ")
    raw_down = input("Enter downstream factor (e.g. 1330/800): ")
    try:
        up_factor   = float(eval(raw_up))
        down_factor = float(eval(raw_down))
    except Exception:
        raise ValueError("Invalid factor expression. Please enter a ratio or number.")

    # --- Hard-coded time windows (modify as needed) ---
    up_start   = pd.Timestamp("2025-06-03 16:09:51")
    up_end     = pd.Timestamp("2025-06-03 16:22:09")
    down_start = pd.Timestamp("2025-06-03 16:25:51")
    down_end   = pd.Timestamp("2025-06-03 16:40:37")

    pe_calc = SEMSPenetrationEfficiency(reader)
    pe_df   = pe_calc.compute_pe(
        up_start, up_end, down_start, down_end,
        up_factor, down_factor
    )

    # Print nicely spaced table
    print(pe_df.reset_index().to_string(index=False))

    # And show the semilog PE plot
    pe_calc.plot_pe(pe_df)
