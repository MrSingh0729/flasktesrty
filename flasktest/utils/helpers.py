import pandas as pd

def get_top_n_counts(df, col, n=3):
    if col not in df.columns or df.empty:
        return pd.DataFrame(columns=[col, "Count"])
    counts = df[col].value_counts().head(n).reset_index()
    counts.columns = [col, "Count"]
    return counts
