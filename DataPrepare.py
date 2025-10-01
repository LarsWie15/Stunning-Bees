import pandas as pd

# --- Step 1: Preview Excel to find actual header row ---
input_file = r"C:\Users\Lars\OneDrive\Documents\Stunning-Bees\Shanghai_Containerized_Freight_Prepared.xlsx"

# Read without header to inspect first 10 rows
df_preview = pd.read_excel(input_file, header=None)
print("Preview of first 10 rows to locate headers:")
print(df_preview.head(10))

# --- Step 2: Read Excel with correct header row ---
# Replace `header=1` with the correct row index after inspecting preview
df = pd.read_excel(input_file, header=1)

# --- Step 3: Rename columns for easier use ---
df.rename(columns={
    "the period (YYYY-MM-DD)": "date",
    "Comprehensive Index": "comprehensive_index",
    "Europe (Base port)": "europe_base_port",
    "Mediterranean (Base port)": "mediterranean_base_port",
    "Persian Gulf and Red Sea (Dubai)": "persian_gulf_red_sea_dubai"
}, inplace=True)

print("Columns after header fix:", df.columns.tolist())

# --- Step 4: Ensure date column is datetime and sort ---
df['date'] = pd.to_datetime(df['date'])
df = df.sort_values('date').reset_index(drop=True)

# --- Step 5: Create lag features ---
lag_weeks = 3  # configurable: number of past weeks to use as lag
columns_to_lag = ["comprehensive_index", "europe_base_port", 
                  "mediterranean_base_port", "persian_gulf_red_sea_dubai"]

for col in columns_to_lag:
    for lag in range(1, lag_weeks + 1):
        df[f'{col}_lag_{lag}'] = df[col].shift(lag)

# --- Step 6: Create rolling averages ---
rolling_windows = [3, 6]  # rolling windows in weeks
for col in columns_to_lag:
    for window in rolling_windows:
        df[f'{col}_rollmean_{window}'] = df[col].shift(1).rolling(window=window).mean()

# --- Step 7: Drop rows with NaN caused by lags or rolling averages ---
df = df.dropna().reset_index(drop=True)

# --- Step 8: Save the prepared file as Excel ---
output_file = r"C:\Users\Rares\Desktop\GroupProject\Shanghai_Containerized_Freight_Prepared.xlsx"
df.to_excel(output_file, index=False)

print(f"Prepared data saved to: {output_file}")
