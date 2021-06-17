import xlrd
import xlwt
#将Clothes_data.xlsx文件的数据提取出来，并经过一定的计算
#将部分数据打印出来
#将部分数据写入Clothes.xlsx文件
filename = r'F:\python_project3\day9\Clothes_data.xlsx'
readbook = xlrd.open_workbook(filename)

sheet = readbook.sheet_by_index(0)
nrows = sheet.nrows #行
ncols = sheet.ncols #列
#该字典存放各个种类衣服的总金额
clothing_Sale = {}
money = 0
#lng = sheet.cell(1,3).value #获取sheet表第一行三列的数据

def Clothing_sales(clothing_str,money_day):
    if clothing_str in clothing_Sale.keys():
        clothing_Sale[clothing_str] += money_day
        return 0
    clothing_Sale[clothing_str] = money_day

for i in range(nrows-1):
    date = sheet.row_values(i+1)    #获取第i+1行的数据，并返回列表
    money1 = date[2]*date[4]
    money+=money1
    clothing_str = date[1]
    Clothing_sales(clothing_str,money1)

money=int(money*100)/100
sales_volume = '当月销售为'+str(money)+'元'
avg_sale = '当月每日销售为'+str(money/(nrows-1))+'元'
print(sales_volume)
print(avg_sale)

inputbook = xlwt.Workbook(filename)
in_sheet = inputbook.add_sheet('result')
k = 1
i = 1
for clothing_name in clothing_Sale:
    m = clothing_Sale[clothing_name]
    percentage =clothing_name+'当月销量百分比为'+str((int(m/money*1000))/10)+'%'
    print(percentage)
    in_sheet.write(i,4,percentage)
    i+=1
in_sheet.write(1,0,sales_volume)
in_sheet.write(1,3,avg_sale)
inputbook.save('Clothes.xlsx')  # 一定要记得保存


