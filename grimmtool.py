import pandas as pd
import matplotlib.pyplot as plt

# Display all columns and format floats clearly
pd.set_option('display.max_columns', None)
pd.set_option('display.float_format', '{:.6f}'.format)

class GrimmReader:
    def __init__(self, path: str, sheet: str = "Sheet1", first_data_col: int = 1):
        """
        Reader for GRIMM Excel data.

        :param path: filepath to Excel file
        :param sheet: sheet name to load
        :param first_data_col: zero-based index where size bins begin (1 for B, 0 for A)

        Expects:
        - Row 1: bucket labels (ignored)
        - Row 2: mean diameters in µm (used as bin centers)
        - Rows 3+: timestamps + counts for each bin
        """
        self.path = path
        self.sheet = sheet
        self.first_data_col = first_data_col
        self.df = None
        self.bin_centers = None

    def load(self):
        # Read raw sheet without header
        raw = pd.read_excel(self.path, sheet_name=self.sheet, header=None)
        # Extract bin centers from row 2 (index 1)
        self.bin_centers = raw.iloc[1, self.first_data_col:].astype(float)
        # Data rows start at row 3 (index 2)
        data = raw.iloc[2:, :]
        # Name columns: timestamp + each bin center
        data.columns = ["timestamp"] + list(self.bin_centers)
        # Parse and set timestamp index
        data["timestamp"] = pd.to_datetime(data["timestamp"])
        data.set_index("timestamp", inplace=True)
        # Store DataFrame of counts
        self.df = data

    def mean_counts(self, start: str, end: str) -> pd.Series:
        """
        Compute mean counts per bin between two timestamps.

        :param start: inclusive start timestamp (parseable string)
        :param end: inclusive end timestamp
        :return: Series indexed by mean diameter (µm), name 'Size'
        """
        s = self.df.loc[start:end].mean()
        s.index.name = 'Size'
        return s

class PenetrationEfficiency:
    def __init__(self, reader: GrimmReader):
        self.reader = reader

    def compute_pe(self,
                   up_start: str, up_end: str,
                   down_start: str, down_end: str,
                   up_factor: float = 1.0,
                   down_factor: float = 1.0) -> pd.DataFrame:
        """
        Build a DataFrame showing original vs corrected counts and PE.

        Corrected Upstream   = orig_up * up_factor
        Corrected Downstream = orig_down * down_factor
        PE                   = (Corrected Downstream / Corrected Upstream) * 100

        Returns DataFrame indexed by 'Size' (µm) with columns:
        ['Original Upstream', 'Original Downstream',
         'Corrected Upstream', 'Corrected Downstream',
         'Penetration Efficiency']
        """
        orig_up   = self.reader.mean_counts(up_start, up_end)
        orig_down = self.reader.mean_counts(down_start, down_end)
        corr_up   = orig_up * up_factor
        corr_down = orig_down * down_factor
        pe        = (corr_down / corr_up) * 100
        df = pd.DataFrame({
            'Original Upstream':      orig_up,
            'Original Downstream':    orig_down,
            'Corrected Upstream':     corr_up,
            'Corrected Downstream':   corr_down,
            'Penetration Efficiency': pe
        })
        df.index.name = 'Size'
        return df

    def plot_pe(self, pe_df: pd.DataFrame):
        """
        Plot penetration efficiency vs particle diameter on log x-axis.
        """
        pe = pe_df['Penetration Efficiency']
        diam = pe.index.astype(float)
        fig, ax = plt.subplots()
        ax.semilogx(diam, pe.values, marker='o')
        ax.set_xlabel('Particle diameter (µm)')
        ax.set_ylabel('Penetration Efficiency (%)')
        ax.grid(True, which='both', linestyle='--', linewidth=0.5)
        import matplotlib.ticker as ticker
        ax.xaxis.set_major_formatter(ticker.FormatStrFormatter('%.3g'))
        ax.xaxis.set_minor_formatter(ticker.NullFormatter())
        plt.tight_layout()
        plt.show()

if __name__ == '__main__':
    # --- Configuration ---
    file_path = 'grim file.xlsx'
    sheet_name = 'Sheet1'
    # If mean-diameter row starts in column A, set first_data_col=0
    reader = GrimmReader(file_path, sheet=sheet_name, first_data_col=1)
    reader.load()

    # --- Prompt for correction factors ---
    raw_up   = input('Enter upstream factor (e.g. 1375/800): ') 
    raw_down = input('Enter downstream factor (e.g. 1.2): ')
    try:
        up_factor   = float(eval(raw_up))
        down_factor = float(eval(raw_down))
    except Exception:
        raise ValueError('Invalid factor expression. Please enter a ratio or number.')

    # --- Time windows ---
    up_start   = "2025-06-03 15:46:03" #6/3/2025  3:46:03 PM
    up_end     = "2025-06-03 15:55:57" #6/3/2025  3:55:57 PM
    down_start = "2025-06-03 15:58:03" #6/3/2025  3:58:03 PM
    down_end   = "2025-06-03 16:07:45" #6/3/2025  4:07:45 PM

    # --- Compute and display results ---
    pe_calc = PenetrationEfficiency(reader)
    df = pe_calc.compute_pe(
        up_start, up_end, down_start, down_end,
        up_factor, down_factor
    )
    # Reset index so 'Size' appears as its own column in the output
    df_reset = df.reset_index()
    print(df_reset.to_string(index=False))

    # --- Plot ---
    pe_calc.plot_pe(df)
