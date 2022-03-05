      import xlwings as xw 
      # https://www.kancloud.cn/gnefnuy/xlwings-docs/1127455
      # https://www.jianshu.com/p/e430539db4be #文档

    
 ==========xlwinges 可以读公示值==============
sheet = xw.sheets.active # 读取活跃工作表
### sheet = xw.Book('基金.xls').sheets['基金']

列表 = ['Fo1', 'Fo2', 'Fo3']
列表元组 = [('110025', '易方达资源行业混合', '1.5063', '0.8693', '2022-02-17 15:00:00', '1.5300', '1.5300', '2.4782', '2022-02-17'), ('006533', '易方达科融混合', '2.9038', '0.6029', '2022-02-17 15:00:00', '2.9356', '2.9356', '1.7151', '2022-02-17'), ('001513', '易方达信息产业混合', '2.6250', '0.4566', '2022-02-17 15:00:00', '2.6280', '2.6280', '0.5741', '2022-02-17')]

# print(sheet['I2'].value）
取数据3行5列 = sheet[2, 4].value #从0开始
取的5行1列数值 = sheet[:5, 0].value
取2行5列值 = sheet[:2, :5].value#从0开始
# print(取的5行1列数值)


<img src="http://static.runoob.com/images/runoob-logo.png" width="50%">


![链接](http://static.runoob.com/images/runoob-logo.png)



#### ====提取数据 元组方法============
  # print(sheet.range('D2').value)
  # print(sheet.range('A1:D3').value)

  # 要强制单个单元格作为列表到达，请使用：
  # print("CES",sheet.range('A3').options(ndim=1).value)





#### ============按行=========================
print(sheet.range(1, 1).value)#从1开始计数
print('从A1:C3',sheet.range(('A1'),('C3')).value) #A1:C3 横向读取
print('方法2从A1:C3',sheet.range((1,1),(3,3)).value) #A1:C3 横向读取

#### ============按列写============================
# 将列表[1,2,3]储存在B1:B3中
sheet.range('B1').options(transpose=True).value=[1,2,3]
# print(sheet.range('A1').expand().value) # 从A1开始扩展读取右边所有 有数据的单元格 内容

# sj = ('QQd','dsds','tex')

# # sheet.range('A8').expand().value = ['Fo1', 'Fo2', 'Fo2']#数据写入1
# sheet.range('A8').expand().value = 列表 #数据写入2

# sheet.range('A8').value = ['Fo1', 'Fo2', 'Fo2']
# sheet.range('A8').value = [['Fo1', 'Fo2', 'Fo2'], [10, 20, 30]]
# 代码说明 *从A8位置开始 A8=Fo1，B8=8=Fo2, C8=Fo3   A9=10 B9=20 C9=30

# ==================重点：=====================
# 将2x2表格，即二维数组，储存在A1:B2中，如第一行1，2，第二行3，4
sheet.range('A5').options(expand='table').value=列表元组