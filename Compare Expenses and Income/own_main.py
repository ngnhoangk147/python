# TODO: Import module
import pandas as pd
import plotly.express as px
import seaborn as sns
import xlwings as xw
import matplotlib.pyplot as plt

pd.set_option('mode.chained_assignment', None)
# TODO: Open transactions.csv
df_transaction = pd.read_csv('transactions.csv')

# TODO: Separate expenses from income
# df_category = df_transaction.groupby(by=df_transaction.Category, as_index=False )
new_df = df_transaction[['Date','Amount', 'Simple Description', 'Category']]
# new_df['Date'] = pd.to_datetime(new_df['Date'])
new_df['Amount'] = new_df['Amount'].str.replace(',','',regex=True).astype(float)

expenses_df = new_df[new_df['Amount'] < 0]
expenses_df['Amount'] = expenses_df['Amount']*-1
sum_expenses_df = expenses_df.groupby('Category', as_index=False ).agg({'Amount': 'sum'})
sum_expenses_df.sort_values('Amount', ascending=False, inplace=True)
print(sum_expenses_df)

income_df = new_df[new_df['Amount'] > 0]
sum_income_df = income_df.groupby('Category', as_index=False ).agg({'Amount': 'sum'})
sum_income_df.sort_values('Amount', ascending=False, inplace=True)
print(sum_income_df)

deviant_df1 = expenses_df[['Date','Amount']]
deviant_df1['Type'] = 'Expenses'

deviant_df2 = income_df[['Date','Amount']]
deviant_df2['Type'] = 'Income'

deviant = pd.concat([deviant_df1,deviant_df2])

deviant.Date = pd.to_datetime(deviant.Date).dt.date.astype(str)
deviant.Date = [date[0:7] for date in deviant.Date]

deviant_df = deviant.groupby(by=['Date','Type'],as_index=False ).agg({'Amount': 'sum'})

# TODO: Load Excel workbook
excel_file = 'Expenses.xlsx'
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
income_df.to_excel(writer, index=False, sheet_name='Income')
expenses_df.to_excel(writer, index=False, sheet_name='Expenses')
deviant_df.to_excel(writer, index=False, sheet_name='Deviant')

# writer.save()
writer.close()
# TODO: Add charts
# Helper function to insert 'Headings' into Excel cells
def insert_title(rng, text):
    rng.value = text
    rng.font.bold = True
    rng.font.size = 20
    rng.font.color = (0, 0, 139)

# Create Pie Chart for Expenses
fig_exp = px.pie(sum_expenses_df, values=sum_expenses_df.Amount, names=sum_expenses_df.Category, hole=0.5)
# Update Traces
fig_exp.update_traces(textposition='outside', textinfo='percent+label')
# Export Pie chart to HTML
# plotly.offline.plot(fig_exp, filename='Piechart.html')
# Insert Pie chart to sheet
wb = xw.Book('Expenses.xlsx')
sht_expenses = wb.sheets['Expenses']
insert_title(sht_expenses.range("H1"), "Percentage of Expenses by category")
sht_expenses.pictures.add(
    fig_exp,
    name="plotly",
    update=True,
    left=sht_expenses.range("H2").left,
    top=sht_expenses.range("H2").top,
    height=400,
    width=550,
)

# Create Pie Chart for Income
fig_income = px.pie(sum_income_df, values=sum_income_df.Amount, names=sum_income_df.Category, hole=0.5)
# Update Traces
fig_income.update_traces(textposition='outside', textinfo='percent+label')
# Export Pie chart to HTML
# plotly.offline.plot(fig, filename='Piechart.html')
# Insert Pie chart to sheet
sht_income = wb.sheets['Income']
insert_title(sht_income.range("H1"), "Percentage of Income by category")
sht_income.pictures.add(
    fig_income,
    name="plotly",
    update=True,
    left=sht_income.range("H2").left,
    top=sht_income.range("H2").top,
    height=400,
    width=550,
)

# Create Bar chart
bar_deviant = plt.figure()
sns.set_style("whitegrid")
sns.barplot(data=deviant_df, x="Date", y="Amount", hue="Type")
# Insert Bar chart to sheet
sht_deviant = wb.sheets['Deviant']
insert_title(sht_deviant.range("H1"), "Deviant between Expenses and Income")
sht_deviant.pictures.add(
    bar_deviant,
    name="plotly",
    update=True,
    left=sht_deviant.range("H2").left,
    top=sht_deviant.range("H2").top,
    height=400,
    width=550,
)

# TODO: Save workbook
wb.save()
if len(wb.app.books) == 1:
    wb.app.quit()
else:
    wb.close()