import pandas as pd

df = pd.read_csv('/tmp/claude-0/-home-claude/85218b8e-007c-5ba2-bfe2-a5a971ad410a/scratchpad/retail-sales-project/data_raw/superstore.csv', encoding='latin1')
before = len(df)

df = df.dropna(how='all', subset=['Order Date', 'Sales'])
after_null = len(df)
dupes = df.duplicated().sum()
df = df.drop_duplicates()
after_dedup = len(df)

df['Order Date'] = pd.to_datetime(df['Order Date'], format='mixed')
df['Ship Date'] = pd.to_datetime(df['Ship Date'], format='mixed')
df['Year'] = df['Order Date'].dt.year
df['Postal Code'] = df['Postal Code'].astype('Int64')

cols = ['Order ID','Order Date','Ship Date','Ship Mode','Customer ID','Customer Name',
        'Segment','Country','City','State','Postal Code','Region','Product ID',
        'Category','Sub-Category','Product Name','Sales','Quantity','Discount','Profit','Year']
df = df[cols].sort_values('Order Date').reset_index(drop=True)

out = '/tmp/claude-0/-home-claude/85218b8e-007c-5ba2-bfe2-a5a971ad410a/scratchpad/retail-sales-project/build/clean_data.csv'
df.to_csv(out, index=False)

print(f"Raw rows: {before}")
print(f"After dropping {before - after_null} fully-blank rows: {after_null}")
print(f"Duplicate rows removed: {dupes}")
print(f"Final clean row count: {after_dedup}")
print(f"Columns ({len(cols)}): {cols}")
print(f"Date range: {df['Order Date'].min().date()} to {df['Order Date'].max().date()}")
print("Saved:", out)
