
import pandas as pd
import numpy as np

from pathlib import Path
from sklearn.linear_model import LinearRegression
from statsmodels.tsa.holtwinters import ExponentialSmoothing


# =====================================
# 1. Load Data
# =====================================

BASE_DIR = Path(__file__).resolve().parent

df = pd.read_csv(
    BASE_DIR / "Cottonil_Monthly_Sales_2020.csv",
    parse_dates=["Month"]
).sort_values("Month").reset_index(drop=True)

y = df["Net_Sales"].astype(float)

train = y.iloc[:9]
test = y.iloc[9:]

X_train = np.arange(len(train)).reshape(-1, 1)
X_test = np.arange(len(train), len(train) + 3).reshape(-1, 1)


# =====================================
# 2. Models
# =====================================

predictions = {}

# Linear Regression
model = LinearRegression()
model.fit(X_train, train)

predictions["Linear Regression"] = model.predict(X_test)

# Naive
predictions["Naive"] = np.repeat(train.iloc[-1], 3)

# Moving Average
predictions["Moving Average"] = np.repeat(
    train.tail(3).mean(), 3
)

# Exponential Smoothing
es_model = ExponentialSmoothing(
    train,
    trend=None,
    seasonal=None,
    initialization_method="estimated"
).fit()

predictions["Exponential Smoothing"] = (
    es_model.forecast(3).values
)

# Holt Damped Trend
holt_model = ExponentialSmoothing(
    train,
    trend="add",
    damped_trend=True,
    seasonal=None,
    initialization_method="estimated"
).fit()

predictions["Holt Damped Trend"] = (
    holt_model.forecast(3).values
)


# =====================================
# 3. Evaluation
# =====================================

results = []

for name, pred in predictions.items():

    actual = test.values

    mae = np.mean(np.abs(actual - pred))

    wape = (
        np.sum(np.abs(actual - pred))
        / np.sum(np.abs(actual))
    ) * 100

    mape = np.mean(
        np.abs((actual - pred) / actual)
    ) * 100

    results.append({
        "Model": name,
        "MAE": mae,
        "MAPE": mape,
        "WAPE": wape
    })


results_df = pd.DataFrame(results)

results_df = results_df.sort_values(
    "WAPE"
).reset_index(drop=True)


# =====================================
# 4. Display Results
# =====================================

print("\nMODEL COMPARISON")
print("=" * 60)

print(
    results_df.to_string(
        index=False,
        formatters={
            "MAE": "{:,.2f}".format,
            "MAPE": "{:.2f}%".format,
            "WAPE": "{:.2f}%".format
        }
    )
)


# =====================================
# 5. Export Results
# =====================================

results_df.to_csv(
    BASE_DIR / "Cottonil_Model_Comparison.csv",
    index=False,
    encoding="utf-8-sig"
)

print("\nComparison exported successfully!")
