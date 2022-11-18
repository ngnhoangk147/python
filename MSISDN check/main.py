import pandas as pd
from datetime import date

# need to pip install openpyxl

#data processing
file_name=input("Enter the file: ")
df = pd.read_excel(file_name)
OCSHN1	=['772',' 775','773','774','776','777','778','779']
OCSHN2	=['762','763','764','765','1210','766','767','768','769']
OCSHN3	=['901','902','903','904','905']
OCSHN4	=['931','932','933','934','935','702','703','704','705']
OCSHCM1	=['786','787','788','789','782','783','784','785']
OCSHCM2	=['792','793','794','795','796','797','798','799']
OCSHCM3	=['906','907','896','898','908','909','899']
OCSHCM4	=['936','937','706','707','708','938','939']
df_count=df['CALLING_NUMBER_M'].count()

NODE=[]

for i in range(df['CALLING_NUMBER_M'].count()):
    if str(df['CALLING_NUMBER_M'][i])[0:3] in OCSHN1:
        NODE.append('HNOCS1')
    elif str(df['CALLING_NUMBER_M'][i])[0:3] in OCSHN2:
        NODE.append('HNOCS2')
    elif str(df['CALLING_NUMBER_M'][i])[0:3] in OCSHN3:
        NODE.append('HNOCS3')
    elif str(df['CALLING_NUMBER_M'][i])[0:3] in OCSHN4:
        NODE.append('HNOCS4')
    elif str(df['CALLING_NUMBER_M'][i])[0:3] in OCSHCM1:
        NODE.append('HCMOCS1')
    elif str(df['CALLING_NUMBER_M'][i])[0:3] in OCSHCM2:
        NODE.append('HCMOCS2')
    elif str(df['CALLING_NUMBER_M'][i])[0:3] in OCSHCM3:
        NODE.append('HCMOCS3')
    elif str(df['CALLING_NUMBER_M'][i])[0:3] in OCSHCM4:
        NODE.append('HCMOCS4')
    else:
        NODE.append('None')

df['Node']=NODE
today = date.today()
excel_file = f'msisdn_check_{today}.xlsx'
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
# df.to_excel(writer,index=False, sheet_name='Sheet1')

print("Data has been processed !!!")

#insert column chart to excel file
workbook = writer.book

subscribers = df.groupby(by='Node', as_index=False).agg({'CALLING_NUMBER_M': pd.Series.count})
subscribers.sort_values('CALLING_NUMBER_M',ascending=False, inplace=True)
subscribers.columns = ['Node','NoOfSub'] #Changing column name
print(subscribers)
subscribers.to_excel(writer,index=False, sheet_name='Sheet2')
count_rows=subscribers.Node.count()+1
chart = workbook.add_chart({'type': 'column'})
chart.add_series({'categories': f'=Sheet2!$A$2:$A${count_rows}',
    'values':     f'=Sheet2!$B$2:$B${count_rows}',
    'gap':        2,
})
chart.set_legend({'delete_series': [0]})
chart.set_x_axis({'name': 'NODE'})
chart.set_y_axis({'name': 'Number of subscribers', 'major_gridlines': {'visible': False}})
worksheet = writer.sheets['Sheet2']
worksheet.insert_chart('G2', chart,{'x_scale': 2, 'y_scale': 1.5})
writer.save()
