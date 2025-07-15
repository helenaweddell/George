import pandas as pd

# open the excel files
# update the filepath to match whatever's in your PC !

# 3 minute sampling data
df_5Hz = pd.read_excel("Timeseries1.xlsx",
                        sheet_name="Sheet1",
                        header=0)
# 5 minute sampling data
df_random = pd.read_excel("Timeseries2.xlsx",
                        sheet_name="Sheet1",
                        header=0)


df_5Hz.dropna(inplace=True) # clean up missing values
df_random.dropna(inplace=True)
# example of removing invalid values less than 100
df_5Hz = df_5Hz[df_5Hz['Values1']<100].reset_index(drop=True)

# make sure your datetime columns are in the correct format (otherwise code won't work)
df_5Hz['Datetime'] = pd.to_datetime(df_5Hz['Datetime'])
df_random['Datetime'] = pd.to_datetime(df_random['Datetime'])

# combine the two datasets using the datetimes
df_combined = df_5Hz.merge(df_random, on='Datetime', how='outer')

# interpolate the missing data
df_filled = df_combined.interpolate()

# decide how to aggregate data
agg_type = {'Values1': 'mean', 'Values2': 'mean', 'Breaths': 'mean'}

df_filled.set_index('Datetime', inplace=True)
df_resampled = df_filled.resample('1s').agg(agg_type)

print(df_resampled.head()) # this displays the top of the dataframe

df_resampled.to_excel('Data.xlsx')