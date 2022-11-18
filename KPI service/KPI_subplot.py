import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

pd.options.mode.chained_assignment = None
pd.plotting.register_matplotlib_converters()

header=["starttime","endtime","interval","3","4","5","6","7","8","9","10","11","total_data","failed_data","total_voice","failed_voice","16","total_ivr","failed_ivr","total_USSD","failed_USSD","total_SMS","failed_SMS","23"]
df = pd.read_csv("Tool_KPI_QoS.csv",names=header) #add thêm header cho file .csv

df['QoS_VOICE'] = (((df['total_voice'] - df['failed_voice']))/df['total_voice'])*100
df['QoS_IVR'] = (((df['total_ivr'] - df['failed_ivr']))/df['total_ivr'])*100
df['QoS_USSD'] = (((df['total_USSD'] - df['failed_USSD']))/df['total_USSD'])*100
df['QoS_SMS'] = (((df['total_SMS'] - df['failed_SMS']))/df['total_SMS'])*100
df['QoS_DATA'] = (((df['total_data'] - df['failed_data'] + df['16']))/df['total_data'])*100
# print(df.head())

df_kpi = df[['starttime','QoS_VOICE','QoS_IVR','QoS_USSD','QoS_SMS','QoS_DATA']]

df_kpi['time'] = pd.to_datetime(df_kpi['starttime'], unit='s') + pd.Timedelta('07:00:00')
print(df_kpi.time)
excel_file="KPI_result.xlsx"
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
df_kpi.to_excel(writer, index=False,sheet_name='Sheet1')

service_name = ['VOICE', 'SMS', 'DATA', 'USSD', 'IVR']

def create_chart(service):
    fig, ax = plt.subplots(figsize=(10,4))
    ax.plot(df_kpi.time, df_kpi[service],"b-")
    ax.grid(linestyle='dotted')
    ax.set_title(service)
    # ax.set_xticklabels(df_kpi.time, rotation=30) # using set_xticklabels, there was "UserWarning: FixedFormatter should only be used together with FixedLocator"
    ax.tick_params(axis='x', labelrotation=30)
    hour = mdates.HourLocator()
    hour_fmt = mdates.DateFormatter('%H:%M')
    ax.xaxis.set_major_locator(hour)
    ax.xaxis.set_major_formatter(hour_fmt)
    return plt

for i in service_name:
    if i == 'VOICE':
        cell = 'K1'
    elif i == 'SMS':
        cell = 'K25'
    elif i == 'DATA':
        cell = 'K50'
    elif i == 'USSD':
        cell = 'K75'
    elif i == 'IVR':
        cell = 'K100'
    service = f'QoS_{i}'
    create_chart(service=service)
    plt.savefig(f'{i}.png')
    worksheet = writer.sheets['Sheet1']
    worksheet.insert_image(cell,f'{i}.png')
writer.save()
