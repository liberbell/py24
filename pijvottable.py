import pandas as pd
import numpy as np

sales_df = pd.read_excel("sales_record.xlsx")
print(sales_df.head())

print(sales_df.sort_values(by = ["Region", "Country", "Item Type"]).head(10))
print(sales_df.sort_values(by = ["Region", "Country", "Item Type"]).tail(10))

table = pd.pivot_table(sales_df,
                       index=["Region", "Country"],
                       values=["Units_Sold", "Total_Revenue", "Total_Profit"],
                       aggfunc=[np.sum])