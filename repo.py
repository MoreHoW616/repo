import pandas as pd


file_path = '统计分析-资产查询-在途业务查询(含质押券).xlsx'#文件路径名称请自行调整
df = pd.read_excel(file_path, header=None, engine='openpyxl')

data = df.values.tolist()
#print(data)   
output = []
    
#按两行一组处理数据 (存在问题)
for i in range(1, len(data), 2):
    main_row = data[i]
    pledge_row = data[i+1] if i+1 < len(data) else []
        
    #交易信息
    产品名称 = main_row[4] if len(main_row) > 4 else ""
    组合名称 = main_row[6] if len(main_row) > 6 else ""
    委托方向 = main_row[9] if len(main_row) > 9 else ""
    发生金额 = int(main_row[15]/10000) if len(main_row) > 15 else ""
    交易对手 = main_row[25] if len(main_row) > 17 else ""
    指令执行人 = main_row[22] if len(main_row) > 18 else ""
    期限 = main_row[20] if len(main_row) > 14 else ""
    利率 = main_row[13] if len(main_row) > 13 else ""


    

    #质押券信息
    证券代码 = pledge_row[35] if len(pledge_row) > 34 else ""
    证券名称 = pledge_row[36] if len(pledge_row) > 35 else ""
    券面总额 = pledge_row[37] if len(pledge_row) > 36 else ""
    质押比例 = pledge_row[38] if len(pledge_row) > 37 else ""
    担保价值 = pledge_row[39] if len(pledge_row) > 38 else ""
    

    #上交所代码，请自行补充
    while True:
        if 产品名称 in ['臻享4号','安享1号','添鑫5号']:
            交易所代码 = '中信建投Z11810'
            break
        if 产品名称 in ['臻享1号','臻享2号','臻享3号','熙元流动管家','尊享1号','熙元91天理财D类']:
            交易所代码 = '中信证券Z06308'
            break
        if 产品名称 in ['睿享添利1号']:
            交易所代码 = '海通证券Z06514'
            break
        if 产品名称 in ['尊享2号']:
            交易所代码 = '广发证券Z08712'
            break
    #格式化输出 
    output.append(f"{产品名称}   {组合名称}     {委托方向}     {发生金额}w")
    output.append(f"{交易对手}")
    output.append(f"{产品名称}   {指令执行人}    {交易所代码}")
    output.append(f"{期限}d   {利率}%   {发生金额}w")
    output.append(f"{产品名称}   {证券代码}   {证券名称}   {券面总额}   {质押比例}   {担保价值}")
    output.append("")
    
    #输出文件,保存为同一路径
with open('output.txt', 'wt', encoding='utf-8') as f:
    f.write("\n".join(output))
    
print(f"处理完成，结果已保存至output.txt")