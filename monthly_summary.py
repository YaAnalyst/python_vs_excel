
#summarizing_monthly_sales_by_categories

import openpyxl
import pandas as pd

data1 = pd.read_excel('/Users/yana/online_retail.xlsx')
data1.head(5)
data1.describe

data1 = data1.dropna(subset=['customer_id'])
data1.loc[:, 'customer_id'] = data1.loc[:, 'customer_id'].astype(int)

data1['year_month'] = data1['InvoiceDate'].dt.to_period('M')
data1['total_price'] = data1['quantity'] * data1['unit_price']
data1.head(5)

monthly_summary = data1.groupby(['customer_id', 
                                 'year_month', 
                                 'country']
                                 ).agg(
    {'total_price': 'sum',
    'quantity': 'sum'
    }).reset_index()
monthly_summary



output_path = '/Users/yana/online_retail.xlsx'
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    countries = monthly_summary['country'].unique()
    
    # Create a new sheet for each country
    for country in countries:
        country_data = monthly_summary[monthly_summary['country'] == country]
        
        country_data.to_excel(writer, sheet_name=country, index=False)

