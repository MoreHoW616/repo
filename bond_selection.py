import pandas as pd
import csv
from collections import defaultdict
from pathlib import Path

#根据账户持仓情况筛选不同交易市场不同评级债券

# 配置参数
file_path =  Path(__file__).parent / "C:/Users/unis/Desktop/data.xlsx"  # 替换为你的实际文件路径
market = input("请输入要筛选的交易市场（例如：银行间/上交所/深交所）：").strip()
    
# 数据结构：{评级: [行数据]}
classified = defaultdict(list)
valid_market = False
    
# 读取文件
try:
    df = pd.read_excel(file_path,
        engine='openpyxl',     # 指定引擎
        dtype={'证券代码': str})
    df['T+0指令可用(张)'] = df['T+0指令可用(张)'].astype(str)
    name = df['产品名称'][1]
    for _, row in df.iterrows():
        try:
            #筛选交易市场
            if market not in row['交易市场'].strip():
                continue
            valid_market = True
            #排除ABS
            if '资产支持' in row['证券类别'].strip():
                continue
            #数据清洗
            t0_available = str(row['T+0指令可用(张)']).replace("，", "").strip()
            t0_available = int(float(t0_available)/100)
            if(t0_available == 0):
                continue
            #print(type(row['T+0指令可用(张)']))
            #分类存储
            rating = row['外部评级孰高'].strip()
            classified[rating].append(
                f"{row['证券名称'].strip()}\t{row['证券代码'].strip()}\t{t0_available}")
        except KeyError as e:
            print(f"数据缺失关键列：{e}")
            continue
        except Exception as e:
            print(f"处理数据时出错：{e}")
            continue
except FileNotFoundError:
    print(f"文件不存在：{file_path}")
except Exception as e:
    print(f"读取文件失败：{e}")
if not valid_market:
     print(f"未找到交易市场 '{market}' 的数据")
#定义评级排序规则（可根据需要调整顺序）
rating_order = ['AAA', 'AA+', 'AA', 'AA-','A+', 'A', 'A-', 'BBB+', 'BBB','BB+', 'BB',  'BB-', 'B+', 'B']
    
#生成输出内容
output = []
for rating in sorted(classified.keys(), 
                    key=lambda x: (rating_order.index(x) if x in rating_order else len(rating_order))): 
    output.append(f"评级：{rating}")
    output.extend(classified[rating])
    output.append("")
    
#写入文件
with open(f"C:/Users/unis/Desktop/{market}_债券分类.txt", "w", encoding="utf-8") as f:#更换为自己的文件路径
    f.write(f"{name}\n")
    f.write("\n".join(output))
    
print(f"分类完成，结果已保存到{market}_债券分类.txt")
