import openpyxl as xl
from openpyxl.chart import BarChart,Reference
import matplotlib.pyplot as plt
import math

wb = xl.load_workbook('Activity_1_Data.xlsx')
sheet = wb['Sheet1']
#cell = sheet['a1']
#cell = sheet.cell(2, 1)
#print(cell.value)
#print(sheet.max_row)

x = []
y = []

for i in range(sheet.max_row):
    v = sheet.cell(i+1,2)
    val = v.value
    v1 = sheet.cell(i+1,3)
    val1 = v1.value
    x.append(val)
    y.append(val1)
print(f'x = {x}')
print(f'y = {y}')

sum = 0
for j in range(170):
    sum = sum + y[j]

mean = sum/170
print(f'mean = {mean}')

nby2 = y[85]
nby21 = y[86]
median = (nby2+nby21)/2
print(f'median = {median}')

mode = max(y, key=y.count)
print(f'mode = {mode}')

y2 = []
for k in range(170):
    v3 = y[i]-mean
    v4 = v3**2
    y2.append(v4)
print(f'y2 = {y2}')

sum2 = 0
for l in range(170):
    sum = sum + y2[l]
variance = sum/170
print(f'variance = {variance}')

standard_deviation = math.sqrt(variance)
print(f'standard_deviation = {standard_deviation}')


values = Reference(sheet,min_row = 1,max_row = sheet.max_row,min_col = 2,max_col = 3 )
chart = BarChart()
chart.add_data(values)
sheet.add_chart(chart,'e2')


wb.save('activity1.xlsx')
plt.scatter(x,y)
plt.show()




