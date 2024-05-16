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
    df = df.sort_values(['南北战区','大区经理', '省区经理', '区域经理'], ascending=True)
    return df

# 美团格式化
def meituan_caipin_format(file):
    df =  pd.read_excel(file,header=[2,3],skipfooter=1)
    df.columns = df.columns.map(''.join).str.replace(' ', '')
    old_header = df.columns
    new_header = [s.split('Unnamed:')[0] if 'Unnamed:' in s else s for s in old_header]
    df.columns = new_header
    caipin_name = df['菜品名称'].unique()[0]
    df['机构编码'] = df['机构编码'].str.split('-').str[0]
    df = df.groupby('机构编码')['销售数量'].sum()
    return df,caipin_name

# 综合营业统计
def meituan_zonghe_format(file):
    df =  pd.read_excel(file,header=[2,3,4],skipfooter=1)
    df.columns = df.columns.map(''.join).str.replace(' ', '')
    old_header = df.columns
    new_header = [s.split('Unnamed:')[0] if 'Unnamed:' in s else s for s in old_header]
    df.columns = new_header
    df = df.loc[:,['机构编码','商户号','门店','营业天数','营业额（元）','营业收入（元）','订单量','渠道营业构成收银POS营业额（元）','渠道营业构成收银POS营业收入（元）',
         '渠道营业构成美团外卖营业额（元）','渠道营业构成美团外卖营业收入（元）','渠道营业构成饿了么外卖营业额（元）','渠道营业构成饿了么外卖营业收入（元）','渠道营业构成自营外卖营业额（元）','渠道营业构成自营外卖营业收入（元）'
          ,'渠道营业构成第三方小程序营业额（元）','渠道营业构成第三方小程序营业收入（元）'
         ]]
    df_new = df.rename(columns={
        '机构编码':'门店编码',
        '商户号': '门店ID',
        '订单量':'账单数',
        '营业额（元）':'流水金额',
        '营业收入（元）':'实收金额',
        '渠道营业构成收银POS营业额（元）':'堂食流水',
        '渠道营业构成收银POS营业收入（元）':'堂食实收',
        '渠道营业构成美团外卖营业额（元）':'美团流水',
        '渠道营业构成美团外卖营业收入（元）':'美团实收',
        '渠道营业构成饿了么外卖营业额（元）':'饿了么流水',
        '渠道营业构成饿了么外卖营业收入（元）':'饿了么实收',
        '渠道营业构成自营外卖营业额（元）':'自营流水',
        '渠道营业构成自营外卖营业收入（元）':'自营实收',
        '渠道营业构成第三方小程序营业额（元）':'小程序流水',
        '渠道营业构成第三方小程序营业收入（元）':'小程序实收'
        })
    df_new['门店编码'] = df_new['门店编码'].str.split('-').str[0]
    df_new['外卖流水'] = df_new['美团流水']+df_new['饿了么流水']+df_new['自营流水']
    df_new['外卖实收'] = df_new['美团实收']+df_new['饿了么实收']+df_new['自营实收']
    return df_new
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

def format_meituan_table(file):
    columns = ['门店编码', '营业天数','流水金额', '实收金额', '订单数', '堂食流水', '堂食实收', '堂食订单数', '外卖流水', '外卖实收', '外卖订单数', '小程序流水', '小程序实收', '小程序订单数']
    df_emt = pd.DataFrame(columns=columns)
    df =  pd.read_excel(file,header=[2,3,4],skipfooter=1)
    df.columns = df.columns.map(''.join).str.replace(' ', '')
    old_header = df.columns
    new_header = [s.split('Unnamed:')[0] if 'Unnamed:' in s else s for s in old_header]
    df.columns = new_header
    df_emt['门店编码'] = df['机构编码']
    df_emt['营业天数'] = df['营业天数']
    df_emt['流水金额'] = df['营业额（元）']
    df_emt['实收金额'] = df['营业收入（元）']
    df_emt['订单数'] = df['订单量']
    df_emt['堂食流水'] = df['渠道营业构成收银POS营业额（元）']
    df_emt['堂食实收'] = df['渠道营业构成收银POS营业收入（元）']
    df_emt['堂食订单数'] = df['渠道营业构成收银POS订单量']
    df_emt['外卖流水'] = df['渠道营业构成饿了么外卖营业额（元）'] + df['渠道营业构成美团外卖营业额（元）']
    df_emt['外卖实收'] = df['渠道营业构成饿了么外卖营业收入（元）'] + df['渠道营业构成美团外卖营业收入（元）']
    df_emt['外卖订单数'] = df['渠道营业构成饿了么外卖订单量'] + df['渠道营业构成美团外卖订单量']
    df_emt['小程序流水'] = df['渠道营业构成第三方小程序营业额（元）']
    df_emt['小程序实收'] = df['渠道营业构成第三方小程序营业收入（元）']
    df_emt['小程序订单数'] = df['渠道营业构成第三方小程序订单量'] 
    df_emt['门店编码'] = df_emt['门店编码'].str.split('-').str[0]
    pivot_df = df_emt.groupby('门店编码').sum().reset_index()
    pivot_df = pivot_df[columns]
    return pivot_df


def format_hualala_table(file):
    shouyinji_path = r"C:\Users\Administrator\OneDrive\甜啦啦\代码底表\收银机管理表_2024.3.22.xlsx"
    df_syj = pd.read_excel(shouyinji_path)
    df_syj = df_syj[['门店编码','组织编码']]
    df_syj = df_syj.dropna(subset=['组织编码'])
    columns = ['门店编码', '营业天数','流水金额', '实收金额', '订单数', '堂食流水', '堂食实收', '堂食订单数', '外卖流水', '外卖实收', '外卖订单数', '小程序流水', '小程序实收', '小程序订单数']
    df_emt = pd.DataFrame(columns=columns)
    df = pd.read_excel(file,header=[2,3],skipfooter=1)
    df.columns = df.columns.map(''.join).str.replace(' ', '')
    old_header = df.columns
    new_header = [s.split('Unnamed:')[0] if 'Unnamed:' in s else s for s in old_header]
    df.columns = new_header
    df = df.dropna(subset=['店铺组织编码'])
    df= pd.merge(df, df_syj, how='left', left_on='店铺组织编码',right_on = '组织编码')
    df_emt['门店编码'] = df['门店编码']
    df_emt['流水金额'] = df['合计流水金额']
    df_emt['实收金额'] = df['合计实收金额']
    df_emt['订单数'] = df['合计账单数']
    df_emt['外卖流水'] = df['美团外卖流水金额'] + df['饿了么外卖流水金额']
    df_emt['外卖实收'] = df['美团外卖实收金额'] + df['饿了么外卖实收金额']
    df_emt['外卖订单数'] = df['美团外卖账单数'] + df['饿了么外卖账单数']
    df_emt['小程序流水'] = df['微信小程序流水金额'] + df['支付宝小程序流水金额']
    df_emt['小程序实收'] = df['微信小程序实收金额'] + df['支付宝小程序实收金额']
    df_emt['小程序订单数'] = df['微信小程序账单数'] + df['支付宝小程序账单数']
    df_emt['堂食流水'] = df_emt['流水金额'] - df_emt['外卖流水'] - df_emt['小程序流水'] 
    df_emt['堂食实收'] = df_emt['实收金额'] - df_emt['外卖实收'] - df_emt['小程序实收'] 
    df_emt['堂食订单数'] = df_emt['订单数'] - df_emt['外卖订单数'] - df_emt['小程序订单数'] 
    pivot_df = df_emt.groupby('门店编码').sum().reset_index()
    pivot_df = pivot_df[columns]
    pivot_df= pivot_df.fillna(0)
    pivot_df['营业天数'] = ''
    return pivot_df










