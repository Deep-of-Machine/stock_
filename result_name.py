import win32com.client
inCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")

list = ['A000320', 'A001550', 'A003925', 'A004270', 'A009415', 'A010130', 'A010640', 'A010780', 'A011090', 'A014680', 'A033270', 'A044380', 'A083420', 'A207940', 'A234310', 'A272220', 'A292190', 'A306530', 'A368190', 'A368470', 'A371460', 'A383220', 'A385720', 'A395760', 'A404260', 'A419430', 'Q530088', 'Q610013', 'Q700003']
result = []
for i in list:
    a = inCpStockCode.CodeToName(i)
    result.append(a)
    print(a)
print(result)