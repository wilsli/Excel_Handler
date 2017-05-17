#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Handler
处理从Excel文件导入的数据表，包括数据清洗、格式转换等。

handle_merged_cells(workbook): 将整个工作簿里的合并单元格进行单元格值分发。
ulersheet_to_df(worksheet,dropna=True): 将openpyxl工作表放入pandas的DataFrame。
dtype_vector(dList): 传入array-like的参数dList，返回其值类型的权值列表。

@author: wilson
Created on Thu Mar 30 23:40:06 2017

Dependencies: pandas, openpyxl

"""
import numpy as np
from pandas import DataFrame, isnull, Timestamp
import datetime, xlrd, openpyxl
import scipy.cluster


def cancel_merged_cells(worksheet):
    """
    把所有空白的合并单元格赋予所属合并区域的值。
    ----------
    参数： worksheet - openpyxl的worksheet对象
    返回值：worksheet对象
    """
    rgs = worksheet.merged_cell_ranges         # 列出所有合并单元格区域 list
    for c_area in rgs:                         # 对每个合并单元格区域执行该操作
        v = worksheet[c_area][0][0].value          # 合并区域左上角单元格的值
#        print("{0}'s value is '{1}',".format(worksheet[c_area][0][0],v))
        for c_row in worksheet[c_area]:      # 在合并区域的每行间循环 worksheet[c_area]是指向合并区域的元组，li是指向合并单元格区域每行的元组
            for cell in c_row:               # 对一行中每个单元格执行该操作
                cell.value = v            # 写入合并区域左上角单元格的值
#                print("updated {0}'s value to '{1}'.".format(cell,cell.value))
    return worksheet


def sheet_to_df(worksheet, dropna=True):
    """
    将openpyxl的worksheet工作表对象传入写入DataFrame结构并删除空列，返回DataFrame。
    若不删除空列，则参数dropna取False。
    ----------
    参数: worksheet - openpyxl的worksheet对象
         dropna - 删除空列，默认True
    返回值：工作表的DataFrame对象
    """
    df = DataFrame(worksheet.values)        # 将工作表写入DataFrame
    if dropna is True:
        df = df.dropna(axis=1, how='all')    # 删除空列
        df = df.dropna(axis=0, how='all')    # 删除空行
    return df.reset_index(drop=True)


def isFormula(string):
    """
    判断一个字符串string是否为公式
    ----------
    参数：string - 单元格的字符串值
    返回值：bool
    """
    if string.startswith('=') and len(string.split('=')) == 2:
        return True
    else:
        return False


def isNumberString(variable):
    """
    判断一个值是否为str类型的数值
    """
    if type(variable) == str:
        try:
            float(variable)                 # 尝试转换为float类型
            return True
        except ValueError:
            pass
    return False


def dtype_vector(dList):
    """
    传入list-like的参数dList，返回其值类型的数字化列表。
    ----------
    使用场景：
    读入一个数据表（例如Excel工作表）时，需要按照每行数据的类型判断该行数据是表头标题行还是数据内容行。
    由于在数据分析中主要是要对数字类型(int,float等)进行分析，而表头标题原则上都是字符串类型(str)，故数字类型和字符串应分居权值的两端。而空值在大量数据表中可能会在数据行的某些位置存在，而在标题行中一般不会存在（excel中合并单元格会造成被合并单元格为空值，但可用cancel_merged_cells(worksheet)将所有表中合并单元格的值分发到各子单元格），因此空值取权值3。时间类型和布尔类型在统计分析中也比较常见和有用，因此赋权值4。其它的数据类型pandas中较少见，但为了与“标题”中常用的str区分开来，给予稍高的权值2。
    ----------
    参数： dList - 一行数据记录的列表对象 list
    返回值：对应的数据类型数字化列表 list
    """
    dt_vec = list()
    for i in range(len(dList)):
        if isnull(dList[i]):
            dt_vec.append(3)                        # Null -> 3
        elif isinstance(dList[i], int):
            dt_vec.append(6)                        # int -> 6
        elif isinstance(dList[i], float):
            dt_vec.append(7)                        # float -> 7
        elif isinstance(dList[i],
                        (datetime.datetime,
                         datetime.date,
                         datetime.time,
                         datetime.timedelta,
                         datetime.timezone,
                         Timestamp)):
            dt_vec.append(5)                        # time -> 5
        elif isinstance(dList[i], bool):
            dt_vec.append(4)                        # Bool -> 4
        elif isinstance(dList[i], str):
            if isFormula(dList[i]):
                dt_vec.append(8)                    # 单元格公式 -> 8
            else:
                dt_vec.append(1)                    # str -> 1
        else:
            dt_vec.append(2)                        # 其它类型 -> 2
    return dt_vec


def sheet_to_typematrix(sheet_df):
    """
    处理传入由worksheet转换的DataFrame对象，返回其数字化矩阵。用于下一步用k均值算法对数据表每行进行分类识别出是表头或数据。
    ----------
    参数： sheet_df - 数据表DataFrame对象
    返回值：权值矩阵，ndarray类型。
                    1 - 字符串类型
                    2 - 其它类型
                    3 - 空值
                    4 - 布尔类型
                    5 - 时间类型
                    6 - 整数类型
                    7 - 浮点类型
                    8 - 公式类型
    """
    sheet_ary = np.zeros(sheet_df.shape)
    for line_num in range(len(sheet_df)):
        sheet_ary[line_num] = dtype_vector(list(sheet_df.iloc[line_num, :]))
    return sheet_ary


def get_label_list(sheet_df):
    """
    分析传入的工作表DataFrame对象，返回行类型归类列表label_list，列表元素为0的表示该行是标签行，为1表示数据行。
    ----------
    参数： sheet_df - 数据表的DataFrame对象
    返回值：工作表记录分类列表list对象
    """
    type_mat = sheet_to_typematrix(sheet_df)
    [centroid, label_list_ary] = scipy.cluster.vq.kmeans2(
            type_mat, np.array([np.ones((1, type_mat.shape[1]))[0], type_mat[-1, :]]))
    return list(label_list_ary)


def has_no_header(label_list):
    """
    根据行类型识别列表label_list判断是否为纯数据表（表中不含表头标题行），纯数据表返回True。
    ----------
    参数：label_list - 数据表逐行归类列表，由0和1分别代表标题行或数据行。list对象。
    返回值：bool类型。
    """
    if label_list[0] * len(label_list) == np.sum(label_list):
        return True
    else:
        return False

def get_type_str(data_record):
    """
    传入一条数据记录，返回类型名列表。
    ----------
    参数： data_record - 一行数据Series对象      pandas.Series
    返回值： 类型名的字符串列表type_str_list      list
    """
    type_str_list = list()
    for cell in data_record:
        type_str_list.append(type(cell).__name__)
    return type_str_list


def dtype_list(sheet_df, label_list):
    """
    根据sheet_df和label_list判断各列的数据类型
    ----------
    参数：数据表DataFrame对象sheet_df，数据记录属性列表label_list
    返回值：数据表各列数据类型名的列表type_str_list      list
    """
    first_dr = first_data_row(label_list)
    type_str_list = get_type_str(sheet_df.iloc[first_dr, :])
    ncol = null_col(type_str_list)
    for col in ncol:
        for row in np.arange(first_dr+1, len(label_list)):
            type_str_list[col] = type(sheet_df.iloc[row, col]).__name__
            if type_str_list[col] == 'NoneType':
                continue
            else:
                break
    return type_str_list

def null_col(type_str_list):
    """
    寻找type_str_list中的‘NoneType'类型所在的列号
    ----------
    参数： type_str_list - 记录类型名列表     list
    返回值： 'NoneType'所在的列号列表n_col     list
    """
    n_col = list()
    for col in range(len(type_str_list)):
        if type_str_list[col] == 'NoneType':
            n_col.append(col)
        else:
            pass
    return n_col

def first_data_row(label_list):
    """
    寻找第一条数据记录的行号
    ----------
    参数： 记录属性列表label_list        list
    返回值： 第一条数据记录的行号          int
    """
    for row in range(len(label_list)):
        if label_list[row] == 0:
            pass
        else:
            break
    return row


def header_rows(label_list):
    """
    根据行类型识别列表label_list识别出表头行，返回行号列表header_rows。
    ----------
    参数：label_list - 数据表逐行归类列表，由0和1分别代表标题行或数据行。list对象。
    返回值：标题行的行号列表，list对象。
    """
    header_rows = list()
    if has_no_header(label_list):                    # 纯数据没有标题行，返回空列表
        return header_rows
    else:
        for r in range(len(label_list)):
            if label_list[r] == label_list[0]:
                header_rows.append(r)               # 写入标题行行号
            else:
                break
        # print('header rows: %s' % str(header_rows))
        return header_rows


def clean_sheet(sheet_df):
    """
    将worksheet的标题行合并，返回合并后表格的DataFrame，以及各字段数据类型字典。
    ----------
    参数： sheet_df - 数据表的DataFrame对象
    返回值：合并标题行后的数据表DataFrame对象new_df    DataFrame
          {列名：数据类型}的字段数据类型字典dt_scheme  Dict
    """
    new_df = sheet_df.copy()                            # 定义目标DataFrame对象
    label_list = get_label_list(sheet_df)               # 获得标识每行记录属性的列表
    type_str_list = dtype_list(sheet_df, label_list)    # 获得字段类型列表
    h_rows = header_rows(label_list)                    # 获得标题行的行号列表
    if len(h_rows) != 0:                                # 无标题行则直接返回原DataFrame
        for r in h_rows:
            new_df.iloc[r, :] = cells_to_str(new_df.iloc[r, :])
            if r == 0:
                next
            else:
                new_df.iloc[0, :] = new_df.iloc[0, :].fillna('*') + '/' + new_df.iloc[r, :].fillna('*')   # 合并第一行和第r行
        new_df.columns = list(new_df.iloc[0, :])
        new_df.drop(h_rows, inplace=True)
        new_df = new_df.reset_index(drop=True)
    else:
        pass
    dt_scheme = dict()
    for col in range(len(type_str_list)):
        dt_scheme[new_df.columns[col]] = type_str_list[col]
    return new_df, dt_scheme


def cells_to_str(row):
    """
    将Series中非字符串和非空的单元值转换成字符串类型str
    ----------
    参数：row - 一行数据，Series对象。
    返回值：类型转换后的Series对象。
    """
    new_row = row.copy()
    for i in range(len(new_row)):
        if isinstance(row.iloc[i], (str)) or isnull(row.iloc[i]):
            next
        else:
            new_row.iloc[i] = str(row.iloc[i])
    return new_row


def xls_to_xlsx(*args, **kw):
    """
    打开并转换XLS文件为openpyxl的Workbook对象
    ----------
    @param args: args for xlrd.open_workbook 例如：文件路径是必须的参数
    @param kw: kwargs for xlrd.open_workbook 的关键字参数
    @return: openpyxl.workbook.Workbook对象
    """
    book_xls = xlrd.open_workbook(*args, formatting_info=True, ragged_rows=True, **kw)
    book_xlsx = openpyxl.workbook.Workbook()

    sheet_names = book_xls.sheet_names()
    for sheet_index in range(len(sheet_names)):
        sheet_xls = book_xls.sheet_by_name(sheet_names[sheet_index])

        if sheet_index == 0:
            sheet_xlsx = book_xlsx.active
            sheet_xlsx.title = sheet_names[sheet_index]
        else:
            sheet_xlsx = book_xlsx.create_sheet(title=sheet_names[sheet_index])

        for crange in sheet_xls.merged_cells:
            rlo, rhi, clo, chi = crange

            sheet_xlsx.merge_cells(
                start_row=rlo + 1, end_row=rhi,
                start_column=clo + 1, end_column=chi,
            )

        def _get_xlrd_cell_value(cell):
            value = cell.value
            if cell.ctype == xlrd.XL_CELL_DATE:
                datetime_tup = xlrd.xldate_as_tuple(value, 0)   # xldate类型的单元计算时间元组
                if datetime_tup[0:3] == (0, 0, 0):              # 没有日期的用time处理
                    value = datetime.time(*datetime_tup[3:])    # 将元组第四个往后的元素分别传入
                else:
                    value = datetime.datetime(*datetime_tup)    # 将元组的全部元素分别传入
            return value

        for row in range(sheet_xls.nrows):
            sheet_xlsx.append((
                _get_xlrd_cell_value(cell)
                for cell in sheet_xls.row_slice(row, end_colx=sheet_xls.row_len(row))
            ))
    return book_xlsx
