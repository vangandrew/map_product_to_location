import pandas as pd

data = pd.read_excel('testing.xlsx')

product_dict = {
  'Product 1': '123',
  'Product 2': '456',
  'Product 3': '789'
}

for index, row in data.iterrows():
  product = row['Product Name']
  if product in product_dict:
    location = product_dict[product]
    data.at[index, 'Location Code'] = location

data.to_excel('testing.xlsx', index=False)