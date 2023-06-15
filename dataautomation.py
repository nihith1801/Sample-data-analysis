import pandas as pd
import matplotlib.pyplot as plt
df = pd.read_excel('Worksheet in Hood Slam Bumper Vision System_1.xlsx')
# Create a bar chart
print(df.dtypes)
print(df.columns)
df = df[pd.to_numeric(df['In-store Revenue'], errors='coerce').notnull()]
df['In-store Revenue'] = pd.to_numeric(df['In-store Revenue'])

df.plot(kind='bar', x='In-store Channel Costs', y='In-store Revenue')
plt.show()
# Create a scatter plot
df.plot(kind='scatter', x='In-store Channel Costs', y='In-store Revenue')
plt.show()
# Create daily report
df_daily = df[df['Profit'] >= df['Ad Expenditure']]
df_daily.to_excel('Profitmore.xlsx', index=False)
# Create weekly report
df_daily = df[df['Profit'] < df['Ad Expenditure']]
df_daily.to_excel('Profitless.xlsx', index=False)
df_profitlevel = df[df['Unnamed: 13'] == 'OP']
df_profitlevel.to_excel('Moderateprofit.xlsx', index=False)

