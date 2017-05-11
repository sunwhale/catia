# -*- coding: utf-8 -*-
"""
Created on Fri May 05 11:06:31 2017

@author: SunJingyu
"""

import os
import numpy as np
import xlrd
import xlwt
import sys
import qrcode

def getFiles(path):
    fns=[]
    for root, dirs, files in os.walk( path ):
        for fn in files:
            fns.append( [ root , fn ] )
    return fns

def isSuffixFile(filename,suffixs=['txt']):
    b = False
    for suffix in suffixs:
        if filename.split('.')[-1] == suffix:
            b = True
    return b

def getRowsNumber(list1,key):
    for index, item in enumerate(list1):
        if item in key and item<>'':
            key_index = index
    return key_index
   
class Catia:
    """
    Catia类，定义一些函数。
    """
    
    def readExcel(self):
#         打开文件
        workbook = xlrd.open_workbook(r'F:\GitHub\catia\name_space.xlsx')

#         根据sheet索引或者名称获取sheet内容
        sheet1 = workbook.sheet_by_name(u'M01单元体')
        sheet2 = workbook.sheet_by_name(u'M02单元体')
        
        data = {}
        for sheet in [sheet1]:
#            print sheet.name,sheet.nrows,sheet.ncols
            header = sheet.row_values(0)
            for i in range(sheet.ncols):
                data[header[i]] = sheet.col_values(i)[1:]
                
        for sheet in [sheet2]:
#            print sheet.name,sheet.nrows,sheet.ncols
            header = sheet.row_values(0)
            for i in range(sheet.ncols):
                data[header[i]] += sheet.col_values(i)[1:]

        for i, item in enumerate(data[u'页数']): # 将数字转化为unicode
            if item <> '':
                data[u'页数'][i] = unicode(int(item))
        
        for i, item in enumerate(data[u'发动机类型编号']): # 将数字转化为unicode
            if item <> '':
                data[u'发动机类型编号'][i] = unicode(int(item))
                
        for i, item in enumerate(data[u'原始编号']): # 将数字转化为unicode
            if item <> '':
                data[u'原始编号'][i] = unicode(item)
                
#        print data[u'原始编号']
#        print data[u'中文名称']
#        print data[u'中文材料']
#        print data[u'英文材料']
#        print data[u'装配处']
#        print data[u'页数']
#        print data[u'图纸尺寸']
#        print data[u'任务分配']
#        print data[u'发动机类型']
#        print data[u'发动机类型编号']
#        print data[u'发动机级别']
#        print data[u'发动机级别编号']
#        print data[u'部件分组']
#        print data[u'部件分组编号']
#        print data[u'零件分组']
#        print data[u'零件分组编号']
#        print data[u'零件分组统计']
#        print data[u'零件号']
#        print data[u'三维文件名称']
#        print data[u'页号']
#        print data[u'版本号']
#        print data[u'二维文件名称']
#        print data[u'产品序号']
#        print data[u'产品编号']
#        print header
                    
        return header,data

    def createCATScript(self,catia_directory_dict,rename):
        
        debug_directory = 'F:\\GitHub\\catia\\debug\\'
                
        script_directory = debug_directory + 'CATScript\\' # 存放CATScript文件夹
        
        html_directory = debug_directory + 'html\\' # 存放html文件夹
        
        qrcode_directory = debug_directory + 'qrcode\\' # 存放html文件夹
        
        all_directory = [debug_directory,script_directory,html_directory,qrcode_directory]
        
        for directory in all_directory:
            if not os.path.isdir(directory):
                os.makedirs(directory)
                print 'Create new directory:',directory

        titleblock_script_template_file_name = 'F:\\GitHub\\catia\\GB_Titleblock.CATScript'
        
        titleblock_batch_outfile_name = debug_directory + 'create_titleblock.bat'
        titleblock_batch_outfile = open(titleblock_batch_outfile_name, 'w') # 添加标题栏批处理文件
        
        rename_batch_outfile_name = debug_directory + 'rename.bat'
        rename_batch_outfile = open(rename_batch_outfile_name, 'w') # 文件改名批处理文件
        
        html_code_type = 'utf-8' # html文件编码
        script_code_type = 'utf-8' # CATScript文件编码
        cmd_code_type = 'gbk' # cmd文件编码
        
        header,data = self.readExcel() # 读取excel文件数据
        
        for catia_directory in catia_directory_dict.keys():
            old_directory = catia_directory
            new_directory = catia_directory_dict[catia_directory]

            if not os.path.isdir(new_directory):
                os.makedirs(new_directory)
                print 'Create new directory:',new_directory
            
            filenames = getFiles(old_directory)
            
            directories = {}
            
            for filename in filenames:
                if isSuffixFile(filename[1],suffixs=['CATPart','CATDrawing']):
                    if filename[0] not in directories:
                        directories[filename[0]] = {'CATPart':[],'CATDrawing':[]}
            
            for directory in directories.keys():
                for filename in filenames:
                    if filename[0] is directory:
                        if isSuffixFile(filename[1],suffixs=['CATPart']):
                            directories[directory]['CATPart'].append(filename[1])
                        if isSuffixFile(filename[1],suffixs=['CATDrawing']):
                            directories[directory]['CATDrawing'].append(filename[1])
            
            for directory in directories.keys()[:]:
#                print directory
                directories[directory]['CATDrawing'].sort()
#                print directories[directory]['CATDrawing']
                for i, drawing_name in enumerate(directories[directory]['CATDrawing']):
                    
                    part_name_old = u'%s\%s' % (directory,directories[directory]['CATPart'][0])
#                    print part_name_old
                    
                    drawing_name_old = u'%s\%s' % (directory,drawing_name)
#                    print drawing_name_old
                    
                    if rename == u'原始编号':
                        rows_number = getRowsNumber(data[u'原始编号'],part_name_old)
                    if rename == u'三维文件名称':
                        rows_number = getRowsNumber(data[u'三维文件名称'],part_name_old)
#                    print rows_number
                    
                    part_directory_name = u'%s %s' % (data[u'三维文件名称'][rows_number],data[u'中文名称'][rows_number])
#                    print 'part_directory_name',part_directory_name
                    
                    
                    part_directory_full_name = u'%s%s' % (new_directory,part_directory_name)
#                    print 'part_directory_full_name',part_directory_full_name
                    
                    part_name_new = u'%s\%s.CATPart' % (part_directory_full_name,data[u'三维文件名称'][rows_number])
#                    print part_name_new
                    
                    drawing_numner_new = data[u'二维文件名称'][rows_number]
                    drawing_numner_new = unicode(int(drawing_numner_new) + i*100)
#                    print drawing_numner_new

                    drawing_name_new = u'%s\%s.CATDrawing' % (part_directory_full_name,drawing_numner_new)
                    print drawing_name_new.encode('gbk')
                    
                    export_suffix = u'pdf'
                    export_name_new= u'%s\%s.%s' % (part_directory_full_name,drawing_numner_new,export_suffix)
#                    print export_name_new
                    
                    drawing_chinese_material = data[u'中文材料'][rows_number]
#                    print drawing_chinese_material
                    
                    drawing_english_material = data[u'英文材料'][rows_number]
#                    print drawing_english_material
                    
                    drawing_chinese_name = data[u'中文名称'][rows_number]
#                    print drawing_chinese_name
                    
                    if not os.path.isdir(part_directory_full_name):
                        os.makedirs(part_directory_full_name)
                        print 'Create new directory:',part_directory_full_name

#==============================================================================
# 生成html文件
#==============================================================================
                    html_outfile_name = u'%s\\%s.htm' % (html_directory,drawing_numner_new)
                    html_outfile = open(html_outfile_name, 'w')

                    line = u"""<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>%s</title>
</head>
<body>

<table width="100%%">""" % part_directory_name
                    print >>html_outfile,line.encode(html_code_type)
                    
                    keys = [
#                            u'原始编号',
                            u'中文名称',
                            u'二维文件名称',
                            u'三维文件名称',
#                            u'装配处',
#                            u'页数',
#                            u'任务分配',
                            u'发动机类型',
#                            u'发动机类型编号',
                            u'发动机级别',
#                            u'发动机级别编号',
                            u'部件分组',
#                            u'部件分组编号',
                            u'零件分组',
#                            u'零件分组编号',
#                            u'零件分组统计',
                            u'零件号',
                            u'图纸尺寸',
                            u'页号',
                            u'版本号',
                            u'中文材料',
                            u'英文材料',
#                            u'产品序号',
#                            u'产品编号',
                            ]
                            
                    for key in keys:
                        line = '<tr><td width=\"40%%\"><b>%s</b></td>          <td>%s</td></tr>' % (key,data[key][rows_number])
                        print >>html_outfile,line.encode(html_code_type)
                        
                    line = u"""</table>

</body>
</html>"""
                    print >>html_outfile,line.encode(html_code_type)
                    
                    html_outfile.close()
#==============================================================================
# 生成标题栏CATScript文件
#==============================================================================
                    infile = open(titleblock_script_template_file_name, 'r')
                    lines = infile.readlines()
                    for i, line in enumerate(lines):
                        lines[i] = line.decode(script_code_type)
                    infile.close()

                    qrcode_image_url = u'http://47.93.195.1/%s.htm' % drawing_numner_new
                    qrcode_image = qrcode.make(qrcode_image_url)
                    qrcode_image_full_name = u'%s%s.png' % (qrcode_directory,drawing_numner_new)
                    qrcode_image.save(qrcode_image_full_name)
                    
                    for i, line in enumerate(lines):
                        if u'Text_18 =' in line:
#                            print line
                            lines[i] = u'  Text_18 = \"%s\" + vbLf + \"%s\"\n' % (drawing_chinese_material,drawing_english_material)
                        if u'Text_19 =' in line:
#                            print line
                            lines[i] = u'  Text_19 = \"%s\"\n' % u'清华大学'
                        if u'Text_20 =' in line:
#                            print line
                            lines[i] = u'  Text_20 = \"%s\"\n' % drawing_chinese_name
                        if u'Text_21 =' in line:
#                            print line
                            lines[i] = u'  Text_21 = \"%s\"\n' % drawing_numner_new
                        if u'drawing_document' in line:
#                            print line
                            lines[i] = u'  Set drawingDocument1 = documents1.Open(\"%s\")\n' % drawing_name_old
                        if u'save_as' in line:
#                            print line
                            lines[i] = u'  drawingDocument1.SaveAs \"%s\"\n' % drawing_name_new
                        if u'export_data' in line:
#                            print line
                            lines[i] = u'  drawingDocument1.ExportData \"%s\", \"%s\"\n' % (export_name_new,export_suffix)
                        if u'Qrcode =' in line:
#                            print line
                            lines[i] = u'  Qrcode = \"%s\"\n' % (qrcode_image_full_name)

                    titleblock_outfile_name = '%sGB_Titleblock_%s.CATScript' % (script_directory,drawing_numner_new)
                    titleblock_outfile = open(titleblock_outfile_name, 'w')
                    for line in lines:
                        titleblock_outfile.writelines(line.encode(script_code_type))
                    titleblock_outfile.close()
#==============================================================================
# 生成改名CATScript文件
#==============================================================================
                    rename_outfile_name = '%srename_%s.CATScript' % (script_directory,drawing_numner_new)
                    rename_outfile = open(rename_outfile_name, 'w')
                    print >>rename_outfile, u'Language=\"VBSCRIPT\"'.encode(html_code_type)
                    print >>rename_outfile, u'Sub CATMain()'.encode(html_code_type)
                    lines = u"""
CATIA.DisplayFileAlerts = False
Set documents1 = CATIA.Documents
Set drawingDocument1 = documents1.Open("%s")
Set partDocument1 = documents1.Open("%s")
Set product1 = partDocument1.GetItem("Part1")
product1.PartNumber = "%s"
partDocument1.SaveAs "%s"
drawingDocument1.SaveAs "%s"
drawingDocument1.ExportData "%s", "%s"
partDocument1.Close
drawingDocument1.Close
""" % (drawing_name_old,part_name_old,part_directory_name,part_name_new,drawing_name_new,export_name_new,export_suffix)
                    print >>rename_outfile, lines.encode(html_code_type)
                    print >>rename_outfile, u'End Sub'.encode(html_code_type)
                    rename_outfile.close()
#==============================================================================
# 写入批处理文件
#==============================================================================
                    print >>titleblock_batch_outfile, (r'cnext -batch -macro %s' % titleblock_outfile_name)
                    print >>rename_batch_outfile, (r'cnext -batch -macro %s' % rename_outfile_name)
                    print
                    
        titleblock_batch_outfile.close()
        rename_batch_outfile.close()
        return titleblock_batch_outfile_name,rename_batch_outfile_name

#==============================================================================
# main
#==============================================================================
catia_directory_dict = {
                        u'F:\\Temp\\catia\\3\\M01\\':u'F:\\Temp\\catia\\3\\M01\\',
                        u'F:\\Temp\\catia\\3\\M02\\':u'F:\\Temp\\catia\\3\\M02\\',
                        } # 字典 old_directory:new_directory

#catia_directory_dict = {
#                        u'F:\\Temp\\catia\\3\\M02\\20106014003 回油接管嘴\\':u'F:\\Temp\\catia\\3\\M02\\20106014003 回油接管嘴\\',
#                        } # 字典 old_directory:new_directory
                        
#catia_directory_dict = {
#                        u'F:\\Temp\\catia\\1\\M01\\':u'F:\\Temp\\catia\\3\\M01\\',
#                        u'F:\\Temp\\catia\\1\\M02\\':u'F:\\Temp\\catia\\3\\M02\\',
#                        } # 字典 old_directory:new_directory
      
catia = Catia()
#titleblock_batch_outfile_name,rename_batch_outfile_name = catia.createCATScript(catia_directory_dict,u'原始编号') # 从原始编号检索
titleblock_batch_outfile_name,rename_batch_outfile_name = catia.createCATScript(catia_directory_dict,u'三维文件名称') # 从新编号检索
#catia.readExcel()
#print rename_batch_outfile_name
#os.system(rename_batch_outfile_name)
raw_input("Enter enter key to exit...") 