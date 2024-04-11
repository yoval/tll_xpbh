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
import openpyxl
from datetime import datetime

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

#计算周期
def calc_zhouqi(days):
    if days == '-':
        zhouqi = '90日内无报货'
    elif days< 30:
        zhouqi = '30日内有报货'
    elif days <60:
        zhouqi = '60日内有报货'
    elif days <90:
        zhouqi = '90日内有报货'
    else:
        zhouqi = '90日内无报货'
    return zhouqi

def jiexi(customer_code):
    try:
        filtered_rows = baohuo_df[baohuo_df['客商编码'] == customer_code ]
        max_date = filtered_rows['单据日期'].max()
        result = filtered_rows[filtered_rows['单据日期'] == max_date]
        count = result['数量'].iloc[0]
    except:
        max_date = '-'
        count = '-'
    return max_date, count

# 创建 Flask 应用
app = Flask(__name__, template_folder='./templates',static_folder='outputs')

@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/xinpin')
def xinpin():
    return render_template('xinpin.html')

@app.route('/caipin')
def caipin():
    return render_template('caipin.html')

@app.route('/danpin')
def danpin():
    return render_template('danpin.html')

# 定义上传文件的路由
@app.route('/xinpin_upload', methods=['POST'])
def xinpin_upload_files():
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
        gen_file_name = xinpin_process_files(file1_path, file2_path, file3_path)
    else:
        gen_file_name = xinpin_process_files(file1_path, file2_path)
        
    file_link = url_for('static', filename=gen_file_name)
    html = f'<a href="{file_link}">下载{gen_file_name}</a>'
    return render_template_string(html)

# 处理文件
def xinpin_process_files(file1_path, file2_path, file3_path=None):
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
        
        return f'{food_name}_(新品)销售报货信息_{now}.xlsx'    
# 定义上传文件的路由
@app.route('/danpin_upload', methods=['POST'])
def danpin_upload_files():
    global now
    now = time.strftime('%Y%m%d_%H%M', time.localtime())
    # 获取上传的文件
    file1 = request.files['file1']
    file2 = request.files['file2']

    # 为文件添加当前时间并保存到服务器
    file1_path = os.path.join('uploads', f'{now}_{file1.filename}')
    file1.save(file1_path)
    file2_path = os.path.join('uploads', f'{now}_{file2.filename}')
    file2.save(file2_path)
    gen_file_name = danpin_process_files(file1_path, file2_path)
    file_link = url_for('static', filename=gen_file_name)
    html = f'<a href="{file_link}">下载{gen_file_name}</a>'
    return render_template_string(html)

    
# 处理文件
def danpin_process_files(file1_path, file2_path):
    global food_name,baohuo_df
    mendian_df = mendian_format(file1_path)
    baohuo_df = pd.read_excel(file2_path, header = 5, converters={'客商编码': str})
    try:
        baohuo_df['客商编码'] = baohuo_df['客商编码'].astype(int).astype(str)
    except:
        baohuo_df['客商编码'] = baohuo_df['客商编码'].astype(str) 
    food_name = baohuo_df['存货名称'].loc[0]
    output_filename = f'{folder}\\{food_name}_报货信息_{now}.xlsx'
    mendian_df[['日期', f'{food_name}数量']] = mendian_df['U8C客商编码'].apply(jiexi).apply(pd.Series)
    mendian_df['日期'] = pd.to_datetime(mendian_df['日期'], errors='coerce')
    today_ = datetime.now().date()
    mendian_df['上次报货距今'] = mendian_df['日期'].apply(lambda x: (today_ - x.date()).days if pd.notnull(x) else "-")
    mendian_df[f'{food_name}报货周期'] = mendian_df['上次报货距今'].apply(calc_zhouqi)
    mendian_df.rename(columns={'日期': '上次报货日期'}, inplace=True)
    mendian_df = mendian_df.sort_values(by=['大区经理', '省区经理', '区域经理'], ascending=True) #升序排序
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        sheet_name = f'{food_name}底表'
        mendian_df.to_excel(writer, sheet_name=sheet_name, index=False)
        # 格式设置
    workbook = openpyxl.load_workbook(output_filename)
    worksheet = workbook.active
    # 设置 L 列的日期格式
    for row in range(1, worksheet.max_row + 1):
        cell = worksheet.cell(row=row, column=12)  # L 列的索引为 12
        cell.number_format = 'yyyy/m/d'
    workbook.save(output_filename)
    return f'{food_name}_报货信息_{now}.xlsx'  


@app.route('/caipin_upload', methods=['POST'])
def caipin_upload_files():
    now = time.strftime('%Y%m%d_%H%M', time.localtime())
    output_filename = f'{folder}\\菜品标准名称还原_{now}.xlsx'
    # 获取上传的文件
    file1 = request.files['file1']
    file2 = request.files['file2']

    # 为文件添加当前时间并保存到服务器
    file1_path = os.path.join('uploads', f'{now}_{file1.filename}')
    file1.save(file1_path)
    file2_path = os.path.join('uploads', f'{now}_{file2.filename}')
    file2.save(file2_path)
    try:
        caipin_df = pd.read_excel(file1_path,header=2)
        # 删除前两行
        caipin_df = caipin_df.drop(caipin_df.iloc[0:2].index)
        caipin_df = caipin_df.loc[:,['机构编码','门店','订单编号','营业日期','菜品名称','销售数量']]
    except:
        caipin_df = pd.read_excel(file1_path)
    duizhao_df = pd.read_excel(file2_path,header=2)
    duizhao_df = duizhao_df.loc[:,['新增套餐名单品名','标准单品名称','数量']]
    result_df = pd.merge(caipin_df, duizhao_df, how='left', left_on='菜品名称',right_on='新增套餐名单品名')
    try:
        result_df['合计数量'] = result_df['销售数量'] *result_df['数量']
    except:
        pass
    result_df.to_excel(output_filename, index=False)
    file_link = url_for('static', filename=f'菜品标准名称还原_{now}.xlsx')
    html = f'<a href="{file_link}">下载菜品标准名称还原_{now}.xlsx</a>'
    return html


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000,debug=0)