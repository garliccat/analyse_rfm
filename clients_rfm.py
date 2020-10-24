import pandas as pd
import datetime
import glob
import os


dateparse = lambda x: datetime.datetime.strptime(x, '%d.%m.%Y %H:%M:%S')


### variables
# R (recency) boundaries (days since last purchase)
# 0 ---1 cat.--- r_1 ---2 cat.--- r_2 ---3 cat.---
# lower - the better
r_1 = 30
r_2 = 180

# F (frequency) boundaries (number of purchases in selected period)
# assuming the period equals to 12 months
# higher - the better
# 0 ---3 cat.--- f_2 ---2 cat.--- f_1 ---1 cat.---
f_1 = 12
f_2 = 6

# M (monetary) boundaries (just like ABC)
# 0 ---3 cat.--- m_2 ---2 cat.--- m_1 ---1 cat.---
# higher - the better
m_1 = 80
m_2 = 15

period_start = '2020-01'
period_stop = '2020-10'


# importing sales excel files
os.chdir('..')
files = glob.glob('datasets/sl*.xlsx')
frames = []
for f in files:
    df = pd.read_excel(f, usecols=[1, 3, 6], skiprows=7, index_col=0, parse_dates=True, date_parser=dateparse)
    df.columns = ['client', 'sum']
    df.index.name = 'date'
    frames.append(df)
sales_df = pd.concat(frames, axis=0, sort=False)
os.chdir('analyse_rfm')
# selecting the date interval
sales_df = sales_df.loc[period_start:period_stop]


# counting RECENCY
df_r = sales_df.copy()
df_r['r'] = datetime.date.today() - df_r.index.date
df_r['r'] = df_r['r'].dt.days
df_r = df_r[['client', 'r']].groupby('client').agg({'r': min})
df_r = df_r['r'].apply(lambda x: 1 if x <= r_1 else (2 if x > r_1 and x <= r_2 else 3))

print('\n ----- R values -----')
print(df_r.head(20))

# counting FREQUENCY
df_f = sales_df['client'].copy()
df_f.index = df_f.index.date
df_f = df_f.reset_index().drop_duplicates()
df_f = pd.DataFrame(df_f['client'].value_counts())
df_f = df_f['client'].apply(lambda x: 1 if x >= f_1 else (2 if x < f_1 and x >= f_2 else 3))
df_f = df_f.rename('f')

print('\n ----- F values -----')
print(df_f.head(20))

# counting MONETARY (just like ABC analysis)
df_m = sales_df.copy()
df_m = df_m.groupby('client').sum()['sum']
df_m = df_m.div(df_m.sum()).mul(100)
df_m = df_m.sort_values(ascending=False)
df_m = df_m.cumsum()
df_m = df_m.apply(lambda x: 1 if x <= 80 else (2 if x > 80 and x <= 90 else 3))
df_m = df_m.rename('m')

print('\n ----- M values -----')
print(df_m.head(20))


# joining all together and exporting to excel file (oh god)

df = pd.concat([df_r, df_f, df_m], axis=1)
df = df.sort_values(by=['r', 'f', 'm'])
df.to_excel('rfm.xlsx')

print('\n ----- RESULTING TABLE -----')
print(df)
