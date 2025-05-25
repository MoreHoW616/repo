#直接处理资管通中导出的可质押券。主要是对T0可用转换为金额，并进行了排列处理
#按照产品名称进行分类，生成相应的表格

import pandas as pd


#读取excel文件
file_path = "E:/work/每日协回/0429协回.xlsx"
df = pd.read_excel(file_path, engine='openpyxl',dtype={'证券代码': str})
#print(df)
#对可用进行处理，并读取日期
df['可用金额(万)'] = df['T+0委托可用(张)'] /100


df['业务日期'] = df['业务日期'].str.replace('/', '')

time=df.iloc[0,0]


#选取特定列构成一个excel
df = df[['产品名称', '证券代码','证券名称','可用金额(万)']]

#针对与df中第4列进行排序
df=df.sort_values(df.columns[3],ascending=False)


#按照产品名称生成不同的excel
	#获取分类唯一值
categories = df['产品名称'].unique()
	# 根据类目值将数据拆分成多个DataFrame
dfs = {}
for 产品名称 in categories:
	   dfs[产品名称] = df[df['产品名称'] == 产品名称]

    # 将每个DataFrame保存到不同的Excel文件中，并按照文件+日期命名
for 产品名称, df in dfs.items():
    	df.to_excel(f'{产品名称}_{time}.xlsx', index=False)






 
