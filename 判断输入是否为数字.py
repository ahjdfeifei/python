
mima = input("请出入您的密码")

if mima.isdigit():
    # 判断mima 师傅是只由数字组成
    print('支付数据合法')
else:
    print('支付数据不合法')

print('-'*50)
print('支付数据合法' if mima.isdigit()else'支付密码数据不合法')