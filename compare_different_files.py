from pathlib import Path

import pandas as pd
import xlwings as xw

initial_version = Path.cwd() / 'initial_file.csv'
new_version = Path.cwd() / 'new_file.csv'
df_initial = pd.read_csv(initial_version)
print(df_initial.head(3))

df_new = pd.read_csv(new_version)
print(df_new.head(3))

df_initial.shape == df_updated.shape

df_new = df_new.reset_index()
print(df_new.head(3))

# Merge dataframes and add indicator column
df_diff = pd.merge(df_initial, df_new, how="outer", indicator="Exist")
print(df_diff)

# Add new column with separated value
new = df_diff['market'].str.split('[A-Z][^A-Z]*', n=1, expand=True)
# df_diff['country'] = new[0]
df_diff['market_1'] = new[1]

# Show only difference
df_diff = df_diff.query("Exist != 'both'")
print(df_diff)

# Show only data existing in updated version
df_right = df_diff.query("Exist == 'right_only'")
print(df_right)

right_rows = df_right['index'].tolist()

right_rows = [int(row) for row in right_rows]

# change case to lower
df_right['market'] = df_right['market'].str.casefold()
print(df_right)

df_diff.drop_duplicates(subset=['column_name'], keep='last')
