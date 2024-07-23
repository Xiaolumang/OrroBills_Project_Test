import pandas as pd
import numpy as np

# Sample DataFrame
data = {
    'colA': [np.nan, -3, 2, 0, -1, 1, 0, 3, np.nan, -2],
    'colB': [5, 2, 3, 4, 1, 6, 8, 7, 0, 9]
}

df = pd.DataFrame(data)

# Define a custom sorting key
def custom_sort_key(row):
    colA = row['colA']
    if pd.isna(colA):
        return (0, np.nan)
    elif colA > 0:
        return (1, -colA)
    elif colA < 0:
        return (2, colA)
    elif colA == 0:
        return (3, 0)

# Apply the custom sorting key to create a new column
df['sort_key'] = df.apply(custom_sort_key, axis=1)

# Sort the DataFrame by the custom key and then by colB
df_sorted = df.sort_values(by=['sort_key', 'colB'], ascending=[True, True])

# Drop the sort_key column as it's no longer needed
df_sorted = df_sorted.drop(columns=['sort_key'])

print(df_sorted)
