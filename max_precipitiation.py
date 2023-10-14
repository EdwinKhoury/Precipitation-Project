import pandas as pd
from datetime import datetime, timedelta

input_file = 'QPCP.xlsx'

def convert_date(df):
    fulltime = []
    for index, row in df.iterrows():
        current_time = datetime.strptime(f'{row["Year"]}-{row["Month"]}-{row["Day"]} {row["Time"]}', '%Y-%m-%d %H:%M')
        fulltime.append(current_time)
    df['datetime'] = fulltime
    return df

rawdata = convert_date(pd.read_excel(input_file))

data = rawdata.groupby('Year')
final_table = pd.DataFrame(index=data.groups.keys(), columns=['15 min', '30 min', '1 hour', '6 hours', '12 hours', '24 hours'])
timedeltas = [timedelta(minutes=15), timedelta(minutes=30), timedelta(hours=1), timedelta(hours=6), timedelta(hours=12), timedelta(hours=24)]
for year in data.groups:
    print(year)
    current = data.get_group(year)
    maximums = [0] * 6
    current_reversed = current[::-1]
    for index, datapoint in current_reversed.iterrows():
        # print(datapoint['datetime'])
        for i in range(len(timedeltas)):
            # Checking ahead
            ahead_check_index = current_reversed.index.get_loc(index)
            ahead_check = datapoint
            interval_sum = 0
            while ahead_check['datetime'] > (datapoint['datetime'] - timedeltas[i]):
                # print(ahead_check['datetime'], datapoint['datetime'] + timedeltas[i], end='\r')
                interval_sum += ahead_check['QPCP (in)']
                ahead_check_index += 1
                if ahead_check_index >= current_reversed.shape[0]:
                    break
                ahead_check = current_reversed.iloc[ahead_check_index]
            if interval_sum > maximums[i]:
                maximums[i] = interval_sum
    final_table.loc[year] = maximums

print(final_table)
with pd.ExcelWriter(input_file, mode='a', if_sheet_exists='replace') as writer:
    final_table.to_excel(writer, sheet_name='Maximums')
