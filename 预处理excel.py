import openpyxl

def sort (element):
    return element[1]


workbook = openpyxl.load_workbook('./结果.xlsx')
dic2={
    'A' :'早起',
    'B' :'运动',
    'C' :'阅读',
    'D' :'学习'

}

sheet = workbook["获奖者"]

dic={}
li = []
i=0

for z in ['A','B','C','D'] :
    for dyg in sheet['{}'.format(z)]:
        name = dyg.value
        if name != None:
            dic[name] = dic.get(name,0) + 1

lis = list(dic.items())
lis.sort(key = sort,reverse = True)



for li in lis:
    i+=1
    sheet['E{}'.format(i)] = li[0]
    sheet['F{}'.format(i)] = li[1]





workbook.save('结果5.xlsx')
