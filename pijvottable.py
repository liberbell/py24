import pandas as pd

sales_df = pd.read_excel("sales_record.xlsx")
print(sales_df.head())

print(sales_df.sort_values(by = ["Region", "Country", "Item Type"]).head(10))
print(sales_df.sort_values(by = ["Region", "Country", "Item Type"]).tail(10))
