import openpyxl
import pyperclip
import tkinter.filedialog
# 0. 选择文件
path = tkinter.filedialog.askopenfilename()

# 1.加载excel表
# path="C:/Users/18768/Desktop/导出数据.xlsx"
workbook = openpyxl.load_workbook(path)
sheet_names = workbook.sheetnames
sheet0 = workbook[sheet_names[0]] #加载第一个数据表

# 2.筛选数据表中【班级】含机械合01的数据
Datas = []
for row in sheet0.rows:
    lines = [cell.value for cell in row]
    Datas.append(lines)
# print(len(Datas))
me1= []
for line in Datas:
    if line[37] == '18机械合01' or line[37] == '17机械合01':
        me1.append(line[1])  #人名添加到列表中
# print(len(me1))
# print(me1)

namelist='''
黄悦岑
李佳霏
张慕辰
蔡延明
常浥尘
陈梦贤
单晨宇
方欣冉
付豪祥
郭安南
郭金华
江旻涛
康喆
孔祥瑞
刘翰桥
刘思言
陆淋康
栾梦颖
任维龙
邵琛文
宋波苇
宋文哲
田雨彤
王梦潇
王晓宇
王奕霖
吴一丹
肖丽珍
熊致弘
伊保钊
于治航
臧姝珂
张所航
张子昂
赵浩然
郑枭
钟一帆
周子帆
左海姣
'''
correct_list = namelist.split( )

a= set(correct_list).symmetric_difference(set(me1))  #查找未打卡名单
print(a)
answer = ' '.join(a)
pyperclip.copy(answer)  #得到的人名自动存在剪贴板上

# 将以上内容打包成exe文件
'''
conda create -n random_env python=3.8.8 pip=21.0.1
conda activate random_env
pip install pyinstaller pyperclip openpyxl

转到程序所在的文件夹之后
pyinstaller -F -w --icon=favicon.ico choice_stu_no_register.py --noconsole
'''
