# -*- coding: utf-8 -*-
"""
Created on Sat May 18 10:07:52 2024

@author: Administrator
"""

from my_module import mendian_format, meituan_caipin_format,format_meituan_table,format_hualala_table,format_zhongtai_table,add_summary,highlight_summary_rows
from flask import Flask, request, url_for, render_template,render_template_string,send_from_directory
import os
import pandas as pd
import numpy as np
import time
import openpyxl
from datetime import datetime, timedelta
import markdown
from dateutil.relativedelta import relativedelta
import socket,requests
import math
import re
from openpyxl.styles import PatternFill


# 创建 Flask 应用
app = Flask(__name__, template_folder='./templates',static_folder='outputs')
app.config.from_pyfile('config.py')

output_folder = app.config['OUTPUT_FOLDER']
upload_floder = app.config['UPLOAD_FOLDER']

# 主页
@app.route('/')
def index():
    file_list = []
    for file in os.listdir(output_folder):
        if file == 'ReadMe.txt':
            continue
        file_path = os.path.join(output_folder, file)
        creation_date = datetime.fromtimestamp(os.path.getctime(file_path))
        file_list.append((file, creation_date))
    sorted_file_list = sorted(file_list, key=lambda x: x[1], reverse=True)
    if len(sorted_file_list) > 13:
       sorted_file_list = sorted_file_list[:13]
    return render_template('index.html', file_list=sorted_file_list)

# 新品报货
@app.route('/xinpin')
def xinpin():
    return render_template('xinpin.html')

def xinpin_process_files(file1_path, file2_path, file3_path=None):
    """处理上传的文件并生成结果文件

    Args:
        file1_path (str): 门店管理表
        file2_path (str): 新品报货
        file3_path (str, optional): 新品销售（可选）

    Returns:
        str: 生成的结果文件名
    """
    now = time.strftime('%Y%m%d_%H%M', time.localtime())
    mendian_format_df = mendian_format(file1_path)
    baohuo_df = pd.read_excel(file2_path, header = 5, converters={'客商编码': str})
    cunhuo_name = baohuo_df['存货名称'].unique()[0]
    col_name = f'{cunhuo_name}报货周期'  # 报货情况
    baohuo_df = baohuo_df.loc[:,['客商编码','数量']]
    baohuo_df = baohuo_df.groupby('客商编码')['数量'].sum()
    df_merge= pd.merge(mendian_format_df, baohuo_df, how='left', left_on='U8C客商编码',right_on='客商编码').fillna(0)
    df_merge.rename(columns={'数量':f'{cunhuo_name}报货数量'}, inplace=True)
    if file3_path is None:
        df_merge[col_name] = df_merge[f'{cunhuo_name}报货数量'].apply(lambda x: '有报货' if x > 0 else '无报货')
        df_merge_weibaohuo = df_merge[(df_merge[col_name] == '无报货') & (df_merge['运营状态'] == '营业中')]
        with pd.ExcelWriter(f'{output_folder}\\{cunhuo_name}_(新品)报货信息_{now}.xlsx') as writer:
            df_merge.to_excel(writer, sheet_name=f'{cunhuo_name}底表', index=False)
            df_merge_weibaohuo.to_excel(writer, sheet_name='未报货表', index=False)
        return f'{cunhuo_name}_(新品)报货信息_{now}.xlsx'
    else:
        xiaoshou_df,caipin_name = meituan_caipin_format(file3_path)
        df_merge= pd.merge(df_merge, xiaoshou_df, how='left', left_on='门店编码',right_on='机构编码').fillna(0)
        df_merge.rename(columns={'销售数量':f'{caipin_name}销售数量'}, inplace=True)
        df_merge[col_name] = df_merge.apply(lambda row: '有报货有销售' if row[f'{cunhuo_name}报货数量'] > 0 and row[f'{caipin_name}销售数量'] > 0 else '有报货无销售' if row[f'{cunhuo_name}报货数量'] > 0 else '无报货有销售' if row[f'{caipin_name}销售数量'] > 0 else '无报货无销售', axis=1)
        df_merge_weibaohuo = df_merge[(df_merge[col_name] == '无报货无销售') & (df_merge['运营状态'] == '营业中')]
        with pd.ExcelWriter(f'{output_folder}\\{cunhuo_name}_(新品)销售报货信息_{now}.xlsx') as writer:
            df_merge.to_excel(writer, sheet_name=f'{cunhuo_name}底表', index=False)
            df_merge_weibaohuo.to_excel(writer, sheet_name='未报未销表', index=False)
        
        return f'{cunhuo_name}_(新品)销售报货信息_{now}.xlsx'

@app.route('/xinpin_upload', methods=['POST'])
def xinpin_upload_files():
    now = time.strftime('%Y%m%d_%H%M', time.localtime())
    # 获取上传的文件
    file1 = request.files['file1']
    file2 = request.files['file2']
    file3 = request.files.get('file3')
    file1_path = os.path.join('uploads', f'{now}_{file1.filename}')
    file1.save(file1_path)
    file2_path = os.path.join('uploads', f'{now}_{file2.filename}')
    file2.save(file2_path)

    if file3:
        file3_path = os.path.join('uploads', f'{now}_{file3.filename}')
        file3.save(file3_path)
        gen_file_name = xinpin_process_files(file1_path, file2_path, file3_path)
    else:
        gen_file_name = xinpin_process_files(file1_path, file2_path)

    file_link = url_for('static', filename=gen_file_name)
    html = f'<a href="{file_link}">下载{gen_file_name}</a>'
    return render_template_string(html)










































































































if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000,debug=1)
    
