!pip install openpyxl xlwings

from pathlib import Path
import pandas as pd
import xlwings as xw

initial_version = Path.cwd() /'GA_America_traffic_final.csv'
updated_version = Path.cwd() /'GA_America_traffic_final_date.csv'

df_initial = pd.read_csv(initial_version)
df_initial.head(3)

df_update = pd.read_csv(updated_version)
df_update.head(3)

df_initial.shape
df_update.shape

df_update.drop(columns="Unnamed: 0.1.1", axis=1, inplace=True)

# Align the differences in columns
diff = df_update.compare(df_initial, align_axis=1)

# Stack the differences in columns
diff = df_update.compare(df_initial, align_axis=0)

# Keep all oroginal rows and columns
diff = df_update.compare(df_initial, keep_shape=True, keep_equal=False)

# Keep all original rows and columns and also all original values
diff = df_update.compare(df_initial, keep_shape=True, keep_equal=True)

# Export difference to Excel
diff = df_update.compare(df_initial, align_axis=1)
diff.to_excel(Path.cwd() / 'Difference.xlsx')

initial_version = Path.cwd() /'GA_America_traffic_final.xlsx'
updated_version = Path.cwd() /'GA_America_traffic_final_date.xlsx'

# Highlight the difference in Excel
with xw.App(visible=False) as app:
    initial_wb = app.books.open(initial_version)
    initial_ws = initial_wb.sheets(1)

    updated_wb = app.books.open(updated_version)
    updated_ws = initial_wb.sheets(1)

    for cell in updated_ws.used_range:
        old_value = initial_ws.range((cell.row, cell.column)).value
        if cell.value != old_value:
            cell.api.AddComment(f"Value from {initial_wb.name}: {old_value}")
            cell.color = (255, 71, 76) # light red
    
    updated_wb.save(Path.cwd() / "Difference_Highlighted.xlsx")
