#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Apr 27 21:19:47 2017

@author: wilson
"""

import sys, os, time
# sys.path.append(os.getcwd())
sys.path.append('/home/webApp/ehApp')
import excel_handler as eh
import openpyxl
from pandas import ExcelWriter
from flask import Flask, request, jsonify, render_template
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['JSON_AS_ASCII'] = False
app.config['JSON_SORT_KEYS'] = False
ALLOWED_EXTENSIONS = set(['csv', 'xls', 'xlsx'])
FOLDER_IN = os.path.abspath('/home/webApp/ehApp/infiles')
FOLDER_OUT = os.path.abspath('/home/webApp/ehApp/outfiles')
app.config['FOLDER_IN'] = FOLDER_IN
app.config['FOLDER_OUT'] = FOLDER_OUT
ALLOWED_EXTENSIONS = set(['xls', 'xlsx'])


def allowed_file(filename):
    """
    根据文件名filename判断是否是允许的文件类型
    ----------
    参数： filename - 文件名String
    返回值：Bool
    """
    return '.' in filename and filename.rsplit('.', maxsplit=1)[1] in ALLOWED_EXTENSIONS


@app.route('/')
def normal():
    return jsonify(cwd=os.getcwd(),
        stdout_encoding=sys.stdout.encoding)


@app.route('/send')
def test():
    return render_template('upload.html')


@app.route('/api/clean_xl', methods=['POST'])
def clean_xl():
    file = request.files['filename']
    if file and allowed_file(file.filename):
        sec_name = secure_filename(file.filename)
        prx_name = sec_name.rsplit('.', 1)[0]
        tstamp = str(int(time.time()))
        if sec_name == prx_name:
            ext = prx_name
        else:
            ext = sec_name.rsplit('.', 1)[1]
        new_filename = prx_name + tstamp + '.' + ext
        file.save(os.path.join(FOLDER_IN, new_filename))     # 保存到FOLDER_IN目录

        org_file = os.path.join(FOLDER_IN, new_filename)
        out_file = os.path.join(FOLDER_OUT, new_filename)

        if ext == 'xls':                                    # 旧版本的xls文件
            wb = eh.xls_to_xlsx(org_file)                   # 使用xls_to_xls函数打开
        elif ext == 'xlsx':                                         # 新版本xlsx文件
            wb = openpyxl.load_workbook(org_file, data_only=True)   # 用openpyxl打开
        else:
            return jsonify(
                errno=1001,
                msg='文件格式错误。')
        sheet_df = dict()
        dt_scheme = dict()
        excel_writer = ExcelWriter(out_file, engine='openpyxl')

        os.remove(org_file)                               # 删除源文件

        for sheetname in wb.get_sheet_names():
            if len(wb[sheetname]._cells) == 0:              # 跳过空表
                pass
            else:
                sheet_df[sheetname], dt_scheme[sheetname] = eh.clean_sheet(eh.sheet_to_df(
                    eh.cancel_merged_cells(wb[sheetname])))
                sheet_df[sheetname].to_excel(excel_writer, sheetname, index=False)   # 向ExcelWriter对象添加worksheet
        excel_writer.save()                                     # 保存xlsx文件

        return jsonify(
            errno=0,
            msg='处理成功。',
            path2file=os.path.join(FOLDER_OUT, new_filename),
            type_scheme=dt_scheme)
    else:
        return jsonify(
            errno=1001,
            msg='文件格式错误。')


if __name__ == '__main__':
    app.run('0.0.0.0')
