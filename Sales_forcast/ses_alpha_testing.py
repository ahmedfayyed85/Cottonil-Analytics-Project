
import pandas as pd
import numpy as np
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

y = df["Net_Sales"].astype(float)

# =====================================
# 2. Walk Forward Alpha Testing
# =====================================

alphas = [0.2, 0.5, 0.8, None]

results = []

for alpha in alphas:

    for i in range(6, len(y)):

        train = y.iloc[:i]
        actual = y.iloc[i]

        if alpha is None:

            model = SimpleExpSmoothing(
                train,
                initialization_method="estimated"
            ).fit(optimized=True)

            model_name = "Optimized Alpha"
            used_alpha = model.params["smoothing_level"]

        else:

            model = SimpleExpSmoothing(
                train,
                initialization_method="estimated"
            ).fit(
                smoothing_level=alpha,
                optimized=False
            )

            model_name = f"Alpha {alpha}"
            used_alpha = alpha

        predicted = model.forecast(1).iloc[0]

        results.append({
            "Model": model_name,
            "Alpha": used_alpha,
            "Month": df.loc[i, "Month"],
            "Actual": actual,
            "Predicted": predicted,
            "Absolute_Error": abs(actual - predicted),
            "APE": abs(actual - predicted) / abs(actual) * 100
        })

# =====================================
# 3. Model Comparison
# =====================================

results_df = pd.DataFrame(results)

comparison = (
    results_df.groupby("Model")
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

print("\nSES ALPHA COMPARISON")
print(comparison.to_string(index=False))

# =====================================
# 4. Export
# =====================================

results_df.to_csv(
    BASE_DIR / "SES_Alpha_Details.csv",
    index=False,
    encoding="utf-8-sig"
)

comparison.to_csv(
    BASE_DIR / "SES_Alpha_Comparison.csv",
    index=False,
    encoding="utf-8-sig"
)

print("\nResults exported successfully!")
