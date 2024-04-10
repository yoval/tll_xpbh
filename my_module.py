import pandas as pd
import os
import shutil
import glob

# 返回所有加盟店
def mendian_format(file):
    df = pd.read_excel(file, sheet_name = '门店信息表', header=1, converters={'U8C客商编码': str})
    df = df[df['大区经理'].notna()] #排除大区经理为空的门店
    df = df[~df['门店名称'].str.contains('测试|茶语|茶讯')] #排除测试门店、教学门店
    df['U8C客商编码'] = df['U8C客商编码'].astype(str)
    # 删除各级经理列中的括号及括号内的内容
    df['大区经理'] = df['大区经理'].str.replace('(\(.*?\))', '', regex=True)
    df['省经理'] = df['省经理'].str.replace('(\(.*?\))', '', regex=True)
    df['区域经理'] = df['区域经理'].str.replace('(\(.*?\))', '', regex=True)
    # 将 "区域经理" 列为空的值填充为 "省经理" 的值
    df['区域经理'] = df['区域经理'].fillna(df['省经理'])
    df = df.loc[:, ['门店编号', '门店名称', '大区经理', '省经理', '区域经理', '南北战区', '运营状态', '省', '市', '区', 'U8C客商编码']]
    df.rename(columns={'门店编号': '门店编码', '省经理': '省区经理'}, inplace=True)
    df = df.sort_values(['大区经理', '省区经理', '区域经理'], ascending=[True, True, True])
    return df

# 美团格式化
def meituan_format(file):
    df =  pd.read_excel(file,header=[2,3],skipfooter=1)
    df.columns = df.columns.map(''.join).str.replace(' ', '')
    old_header = df.columns
    new_header = [s.split('Unnamed:')[0] if 'Unnamed:' in s else s for s in old_header]
    df.columns = new_header
    caipin_name = df['菜品名称'].unique()[0]
    df['机构编码'] = df['机构编码'].str.split('-').str[0]
    df = df.groupby('机构编码')['销售数量'].sum()
    return df,caipin_name

# 移动文件
def move_file(source_file, target_folder):
    file_name = os.path.basename(source_file)
    target_file = os.path.join(target_folder, file_name)
    shutil.move(source_file, target_file)

# 获取文件夹中所有.xlxs、xls文件
def list_excel_files(folder):
    file_list = glob.glob(os.path.join(folder, '*.xlsx')) + glob.glob(os.path.join(folder, '*.xls'))
    for file in file_list:
        print(file)
    return file_list

#u8c导出文件完整性校验
def check_u8c_export(df):
    second_last_row = df.iloc[-2]
    if second_last_row.存货名称 == '--------' :
        print('数据已完整导出！')
    else: 
        print('数据导出不全！')
    df = df[:-2]
    print('当报货表中，共有%s件商品'%len(df['存货名称'].unique()))
    return df