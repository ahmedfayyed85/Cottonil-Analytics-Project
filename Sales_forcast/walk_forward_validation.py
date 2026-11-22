
import pandas as pd
import numpy as np

from pathlib import Path
from sklearn.linear_model import LinearRegression
from statsmodels.tsa.holtwinters import (
    SimpleExpSmoothing,
    ExponentialSmoothing
)

# =====================================
# 1. Load Data
# =====================================

BASE_DIR = Path(__file__).resolve().parent

df = pd.read_csv(
    BASE_DIR / "Cottonil_Monthly_Sales_2020.csv",
    parse_dates=["Month"]
).sort_values("Month").reset_index(drop=True)

y = df["Net_Sales"].astype(float)

models_results = []

# =====================================
# 2. Walk Forward Validation
# =====================================

for i in range(6, len(y)):

    train = y.iloc[:i]
    actual = y.iloc[i]

    # Linear Regression
    lr = LinearRegression()
    lr.fit(
        np.arange(len(train)).reshape(-1, 1),
        train
    )

    lr_pred = lr.predict([[i]])[0]

    # Moving Average
    ma_pred = train.tail(3).mean()

    # Simple Exponential Smoothing
    ses = SimpleExpSmoothing(
        train,
        initialization_method="estimated"
    ).fit()

    ses_pred = ses.forecast(1).iloc[0]

    # Holt Damped Trend
    holt = ExponentialSmoothing(
        train,
        trend="add",
        damped_trend=True,
        initialization_method="estimated"
    ).fit()

    holt_pred = holt.forecast(1).iloc[0]

    predictions = {
        "Linear Regression": lr_pred,
        "Moving Average": ma_pred,
        "Simple Exponential Smoothing": ses_pred,
        "Holt Damped Trend": holt_pred
    }

    for name, predicted in predictions.items():

        error = actual - predicted

        models_results.append({
            "Month": df.loc[i, "Month"],
            "Model": name,
            "Actual": actual,
            "Predicted": predicted,
            "Absolute_Error": abs(error),
            "APE": abs(error) / abs(actual) * 100
        })

# =====================================
# 3. Compare Models
# =====================================

results = pd.DataFrame(models_results)

comparison = (
    results.groupby("Model")
    .apply(
        lambda x: pd.Series({
            "MAE": x["Absolute_Error"].mean(),
            "MAPE": x["APE"].mean(),
            "WAPE": (
                x["Absolute_Error"].sum()
                / x["Actual"].abs().sum()
            ) * 100
        }),
        include_groups=False
    )
    .sort_values("WAPE")
    .reset_index()
)

print("\nWALK FORWARD MODEL COMPARISON")
print("=" * 65)

print(comparison.to_string(index=False))

# =====================================
# 4. Export
# =====================================

results.to_csv(
    BASE_DIR / "Cottonil_All_Models_WalkForward.csv",
    index=False,
    encoding="utf-8-sig"
)

comparison.to_csv(
    BASE_DIR / "Cottonil_WalkForward_Comparison.csv",
    index=False,
    encoding="utf-8-sig"
)

print("\nFiles exported successfully!")
