
import pandas as pd
from pathlib import Path
from statsmodels.tsa.holtwinters import SimpleExpSmoothing

# =====================================
# 1. Load Data
# =====================================

BASE_DIR = Path(__file__).resolve().parent

df = pd.read_csv(
    BASE_DIR / "Cottonil_Monthly_Sales_2020.csv",
    parse_dates=["Month"]
).sort_values("Month").reset_index(drop=True)

sales = df["Net_Sales"].astype(float)

# =====================================
# 2. Train Final SES Model
# =====================================

model = SimpleExpSmoothing(
    sales,
    initialization_method="estimated"
).fit(optimized=True)

print("Optimized Alpha:", model.params["smoothing_level"])

# =====================================
# 3. Forecast Next 3 Months
# =====================================

forecast_values = model.forecast(3)

future_dates = pd.date_range(
    start=df["Month"].max() + pd.offsets.MonthBegin(1),
    periods=3,
    freq="MS"
)

forecast_df = pd.DataFrame({
    "Month": future_dates,
    "Forecast_Sales": forecast_values.values
})

# =====================================
# 4. Export
# =====================================

output_file = BASE_DIR / "Cottonil_Final_SES_Forecast.csv"

forecast_df.to_csv(
    output_file,
    index=False,
    encoding="utf-8-sig"
)

print("\nFINAL FORECAST")
print(forecast_df.to_string(index=False))

print("\nSaved:", output_file)
