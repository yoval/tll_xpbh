# -*- coding: utf-8 -*-
"""
Created on Wed Apr 10 14:06:51 2024

@author: Administrator
"""
from my_module import mendian_format, meituan_format
from flask import Flask, request, url_for, render_template,render_template_string
import os
import pandas as pd
import time

#输出文件夹
folder = 'outputs'

# 判断销售报货
def sales_status(row):
    col_name_1 = f'{cunhuo_name}报货数量'
    col_name_2 = f'{caipin_name}销售数量'
    if row[col_name_1] == 0:
        baohuo = '无报货'
    else:
        baohuo = '有报货'
    if row[col_name_2] == 0:
        xiaoshou = '无销售'
    else:
        xiaoshou = '有销售'
    return baohuo+xiaoshou

# 判断报货
def check_baohuo(i):
    if i > 0:
        return '有报货'
    else:
        return '无报货'


# 创建 Flask 应用
app = Flask(__name__, template_folder='./templates',static_folder='outputs')

@app.route('/')
def index():
    return render_template('index.html')


# 定义上传文件的路由
@app.route('/upload', methods=['POST'])
def upload_files():
    global now
    now = time.strftime('%Y%m%d_%H%M', time.localtime())
    # 获取上传的文件
    file1 = request.files['file1']
    file2 = request.files['file2']
    file3 = request.files.get('file3')
    # 为文件添加当前时间并保存到服务器
    file1_path = os.path.join('uploads', f'{now}_{file1.filename}')
    file1.save(file1_path)
    file2_path = os.path.join('uploads', f'{now}_{file2.filename}')
    file2.save(file2_path)
    
    if file3:
        file3_path = os.path.join('uploads', f'{now}_{file3.filename}')
        file3.save(file3_path)
        gen_file_name = process_files(file1_path, file2_path, file3_path)
    else:
        gen_file_name = process_files(file1_path, file2_path)
        
    file_link = url_for('static', filename=gen_file_name)
    html = f'<a href="{file_link}">下载{gen_file_name}</a>'
    return render_template_string(html)
    

# 处理文件
def process_files(file1_path, file2_path, file3_path=None):
    global cunhuo_name,caipin_name
    if file3_path is None:
        mendian_format_df = mendian_format(file1_path)
        baohuo_df = pd.read_excel(file2_path, header = 5, converters={'客商编码': str})
        cunhuo_name = baohuo_df['存货名称'].unique()[0]
        col_name = f'{cunhuo_name}报货周期'  # 报货情况
        baohuo_df = baohuo_df.loc[:,['客商编码','数量']]
        baohuo_df = baohuo_df.groupby('客商编码')['数量'].sum()
        df_merge= pd.merge(mendian_format_df, baohuo_df, how='left', left_on='U8C客商编码',right_on='客商编码').fillna(0)
        df_merge.rename(columns={'数量':f'{cunhuo_name}报货数量'}, inplace=True)
        df_merge[col_name] = df_merge[f'{cunhuo_name}报货数量'].apply(check_baohuo)
        df_merge_weibaohuo = df_merge[(df_merge[col_name] == '无报货') & (df_merge['运营状态'] == '营业中')]
        with pd.ExcelWriter(f'{folder}\\{cunhuo_name}_(新品)报货信息_{now}.xlsx') as writer:
            df_merge.to_excel(writer, sheet_name=f'{cunhuo_name}底表', index=False)
            df_merge_weibaohuo.to_excel(writer, sheet_name='未报货表', index=False)
        return f'{cunhuo_name}_(新品)报货信息_{now}.xlsx'
    else:
        mendian_format_df = mendian_format(file1_path)
        baohuo_df = pd.read_excel(file2_path, header = 5, converters={'客商编码': str})
        cunhuo_name = baohuo_df['存货名称'].unique()[0]
        col_name = f'{cunhuo_name}报货周期'  # 报货情况
        baohuo_df = baohuo_df.loc[:,['客商编码','数量']]
        baohuo_df = baohuo_df.groupby('客商编码')['数量'].sum()
        df_merge= pd.merge(mendian_format_df, baohuo_df, how='left', left_on='U8C客商编码',right_on='客商编码').fillna(0)
        df_merge.rename(columns={'数量':f'{cunhuo_name}报货数量'}, inplace=True)
        xiaoshou_df,caipin_name = meituan_format(file3_path)
        df_merge= pd.merge(df_merge, xiaoshou_df, how='left', left_on='门店编码',right_on='机构编码').fillna(0)
        df_merge.rename(columns={'销售数量':f'{caipin_name}销售数量'}, inplace=True)
        df_merge[col_name] = df_merge.apply(sales_status, axis=1)
        df_merge_weibaohuo = df_merge[(df_merge[col_name] == '无报货无销售') & (df_merge['运营状态'] == '营业中')]
        with pd.ExcelWriter(f'{folder}\\{cunhuo_name}_(新品)销售报货信息_{now}.xlsx') as writer:
            df_merge.to_excel(writer, sheet_name=f'{cunhuo_name}底表', index=False)
            df_merge_weibaohuo.to_excel(writer, sheet_name='未报未销表', index=False)
        
        return f'{cunhuo_name}_(新品)销售报货信息_{now}.xlsx'
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000,debug=0)