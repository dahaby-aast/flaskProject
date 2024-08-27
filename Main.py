import openpyxl
import pandas as pd

wb = openpyxl.Workbook()
ws = wb.active
#wb.save('testabc.xlsx')
df = pd.read_excel('testabc.xlsx')
#selected_data = df[df['regno'] == 'some_value']
#print(selected_data)
column_a_data = df['regno']
column_a_df = pd.DataFrame(column_a_data)
sorted_column_a_data = column_a_data.sort_values()

sorted_column_a_data.to_excel('new_output_file.xlsx', index=False, header=True)
print("data successfully saved in output_file.xlsx")
#print("sorted data",sorted_column_a_data)


