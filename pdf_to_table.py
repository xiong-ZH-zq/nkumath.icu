# -*- coding: utf-8 -*-      
"""
This file provides the ability for this app to
extract tables in NKU lecture cookbook

本文件为程序提供能在南开选课手册中提取表格的能力.
提取表格在原 pdf 跨页部分有一定的瑕疵，需要检查
"""

import pdfplumber
import pandas as pd

# 表格抽取设置/table extract settings
table_settings = {
    "vertical_strategy": "lines_strict",
    "horizontal_strategy": "lines_strict"
}

# 最终数据表的指标/index of final table
final_table_index = ["选课序号","课程名称","课程归属模块","选课名额",
                     "跨专业名额","教师","星期","节次",
                     "起止周次","教室","备注"]

def combine_strings(series):
    return ''.join(series)

def pdf_import(pdf_path:str,page_list:list,table_path:str) -> pd.DataFrame:
    """
    :param filepath: PDF文件路径/the path of pdf file
    :param page_list: 需要提取的页码列表/the list of page index
    """
    pdf =  pdfplumber.open(pdf_path)
    all_tables = pd.DataFrame()
    for index in page_list:
        page = pdf.pages[index]
        table = page.extract_table(table_settings=table_settings)
        
        # table 提取成功/ table is not None
        if table:
            table_df = pd.DataFrame(table[0:],columns=final_table_index)
            all_tables = pd.concat([all_tables,table_df],ignore_index=True)

    all_tables.reset_index(drop=True, inplace=True)
    # 填补合并单元格造成的缺失
    all_tables.ffill(inplace=True)
    # 丢弃多余表头/Drop duplicated table title
    all_tables.drop_duplicates(keep=False,inplace=True)

    # 处理跨页部分/Deal with parts that cross pages
    # TODO 尽管能部分问题，但是无法完全消除跨页的问题.
    missing_rows = all_tables[all_tables["选课序号"].isnull()]
    for index in missing_rows.index:  
        all_tables.loc[index-1] = all_tables.loc[(index-1):(index)].astype(str).apply(combine_strings,axis=0)

    all_tables.dropna(subset=["选课序号"],inplace=True)
    all_tables.to_csv(table_path,index=False)

if __name__ == '__main__':
    
    # 文件路径
    pdf_path = r"./example/选课手册.pdf"
    table_path = r"./example/选课手册.csv"

    # 页码范围
    page_range = range(10,140)
    
    # 请根据每届不同的 pdf 选课手册进行提取页码更改
    pdf_import(pdf_path,page_list= page_range)