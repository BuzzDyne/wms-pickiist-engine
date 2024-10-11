import pandas as pd

# Load the entire sheet, starting from row 6
df = pd.read_excel("files/DAFTAR PRODUK TIKTOK.xlsx", skiprows=5)

# Select only column 'D' (which is index 3, since indexing starts at 0)
column_d_data = df.iloc[:, 3]  # 3 refers to the 4th column, which is 'D'

# Drop any rows that are empty in this column
column_d_data_cleaned = column_d_data.dropna()

# Display the result
print(column_d_data_cleaned)

print('end')
