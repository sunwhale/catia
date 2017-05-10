# -*- coding: utf-8 -*-
"""
Created on Fri May 05 11:06:31 2017

@author: SunJingyu
"""

import os
import numpy as np
import xlrd
import xlwt
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
    
    def createBatchFile(self):
        catscript_name = 'rename.CATScript'
        outfile = open('rename.bat', 'w')
        print >>outfile, (r'cnext -batch -macro %s' % catscript_name)
        outfile.close()
    
    def readExcel(self):
        # 打开文件
        workbook = xlrd.open_workbook(r'F:\GitHub\catia\name_space.xlsx')

        # 根据sheet索引或者名称获取sheet内容
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
        
#        print data[u'原始编号']
#        print data[u'中文名称']
#        print data[u'中文材料']
#        print data[u'英文材料']
#        print data[u'装配处']
        print data[u'页数']
#        print data[u'图纸尺寸']
#        print data[u'任务分配']
#        print data[u'发动机类型']
        print data[u'发动机类型编号']
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

    def createTitleblockCATScript(self):
        batch_outfile = open('create_titleblock.bat', 'w')
        
        script_directory = 'F:\\GitHub\\catia\\createTitleblockCATScript'
        if not os.path.isdir(script_directory):
            os.makedirs(script_directory)
            print 'Create new directory:',script_directory
                        
        input_code_type = 'gbk'
        output_code_type = 'gbk'
        html_code_type = 'utf-8'
        
        header,data = self.readExcel()
        
        catia_directory_dict = {'F:\\Temp\\catia\\2\\M01\\':'F:\\Temp\\catia\\2\\M01\\',
                                'F:\\Temp\\catia\\2\\M02\\':'F:\\Temp\\catia\\2\\M02\\'}
        
        for catia_directory in catia_directory_dict.keys():
            old_directory = catia_directory
            new_directory = catia_directory_dict[catia_directory]
            
#            print old_directory
#            print new_directory
            
            if not os.path.isdir(new_directory):
                os.makedirs(new_directory)
                print 'Create new directory:',new_directory
            
            filenames = getFiles(old_directory)
            directories = {}
            for filename in filenames:
                if isSuffixFile(filename[1],suffixs=['CATPart','CATDrawing']):
                    if filename[0] not in directories:
                        directories[filename[0]] = {'CATPart':[],'CATDrawing':[]}
            
            for directory in directories:
                for filename in filenames:
                    if filename[0] is directory:
                        if isSuffixFile(filename[1],suffixs=['CATPart']):
                            directories[directory]['CATPart'].append(filename[1])
                        if isSuffixFile(filename[1],suffixs=['CATDrawing']):
                            directories[directory]['CATDrawing'].append(filename[1])
        
            for directory in directories:
                print directory
                directories[directory]['CATDrawing'].sort()
                print directories[directory]['CATDrawing']
                for i, drawing_name in enumerate(directories[directory]['CATDrawing']):
                    
                    part_name_old = u'%s\%s' % (directory,directories[directory]['CATPart'][0])
#                    part_name_old = part_name_old.decode(input_code_type)
#                    print repr(part_name_old)
                    print part_name_old
                    
                    drawing_name_old = u'%s\%s' % (directory,drawing_name)
                    print drawing_name_old
                    
                    rows_number = getRowsNumber(data[u'三维文件名称'],part_name_old)
                    print rows_number
                    
                    part_number = u'%s %s' % (data[u'三维文件名称'][rows_number],data[u'中文名称'][rows_number])
                    print part_number
                    
                    
                    new_directory_name = u'%s%s' % (new_directory,part_number)
                    print new_directory_name
                    
                    part_name_new = u'%s\%s.CATPart' % (new_directory_name,data[u'三维文件名称'][rows_number])
                    print part_name_new
                    
                    drawing_numner_new = data[u'二维文件名称'][rows_number]
                    drawing_numner_new = unicode(int(drawing_numner_new) + i*100)
                    print drawing_numner_new

                    drawing_name_new = u'%s\%s.CATDrawing' % (new_directory_name,drawing_numner_new)
                    print drawing_name_new
                    
                    export_suffix = u'pdf'
                    export_name_new= u'%s\%s.%s' % (new_directory_name,drawing_numner_new,export_suffix)
                    print export_name_new
                    
                    drawing_chinese_material = data[u'中文材料'][rows_number]
                    print drawing_chinese_material
                    
                    drawing_english_material = data[u'英文材料'][rows_number]
                    print drawing_english_material
                    
                    drawing_chinese_name = data[u'中文名称'][rows_number]
                    print drawing_chinese_name

                    drawing_number = data[u'二维文件名称'][rows_number]
                    print drawing_number

                    
                    infile = open('GB_Titleblock.CATScript', 'r')
                    lines = infile.readlines()
                    infile.close()

                    qrcode_image_url = u'http://47.93.195.1/%s.htm' % drawing_numner_new
                    qrcode_image = qrcode.make(qrcode_image_url)
                    qrcode_image_full_name = u'%s\\%s.png' % (script_directory,drawing_numner_new)
                    qrcode_image.save(qrcode_image_full_name)
                    
                    html_directory = u'F:\\Cloud\\wwwroot\\turbine\\'
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

<table width="100%%">""" % part_number
                    print >>html_outfile,line.encode(html_code_type)
                    
                    keys = [u'原始编号',
                            u'中文名称',
                            u'中文材料',
                            u'英文材料',
                            u'装配处',
                            u'页数',
                            u'图纸尺寸',
                            u'任务分配',
                            u'发动机类型',
                            u'发动机类型编号',
                            u'发动机级别',
                            u'发动机级别编号',
                            u'部件分组',
                            u'部件分组编号',
                            u'零件分组',
                            u'零件分组编号',
                            u'零件分组统计',
                            u'零件号',
                            u'三维文件名称',
                            u'页号',
                            u'版本号',
                            u'二维文件名称',
                            u'产品序号',
                            u'产品编号',]
                            
                    for key in keys:
                        line = '<tr><td width=\"40%%\"><b>%s</b></td>          <td>%s</td></tr>' % (key,data[key][rows_number])
                        print >>html_outfile,line.encode(html_code_type)
                        
                    line = u"""</table>

</body>
</html>"""
                    print >>html_outfile,line.encode(html_code_type)
                    
                    html_outfile.close()
                    
                    for i, line in enumerate(lines):
                        if 'Text_18 =' in line:
                            print line.decode('utf-8').encode('gbk')
                            lines[i] = '  Text_18 = \"%s\" + vbLf + \"%s\"\n' % (drawing_chinese_material,drawing_english_material)
                        if 'Text_19 =' in line:
                            print line.decode('utf-8').encode('gbk')
                            lines[i] = '  Text_19 = \"%s\"\n' % '清华大学'
                        if 'Text_20 =' in line:
                            print line.decode('utf-8').encode('gbk')
                            lines[i] = '  Text_20 = \"%s\"\n' % drawing_chinese_name
                        if 'Text_21 =' in line:
                            print line.decode('utf-8').encode('gbk')
                            lines[i] = '  Text_21 = \"%s\"\n' % drawing_numner_new
                        if 'drawing_document' in line:
                            lines[i] = '  Set drawingDocument1 = documents1.Open(\"%s\")\n' % drawing_name_old
                        if 'save_as' in line:
                            lines[i] = '  drawingDocument1.SaveAs \"%s\"\n' % drawing_name_new
                        if 'export_data' in line:
                            lines[i] = '  drawingDocument1.ExportData \"%s\", \"%s\"\n' % (export_name_new,export_suffix)
                        if 'Qrcode =' in line:
                            lines[i] = '  Qrcode = "%s"\n' % (qrcode_image_full_name)
                            
                    outfile_name = '%s\\GB_Titleblock_%s.CATScript' % (script_directory,drawing_numner_new)
                    outfile = open(outfile_name, 'w')
                    for line in lines:
                        outfile.writelines(line)
                    outfile.close()
                    
                    print >>batch_outfile, (r'cnext -batch -macro %s' % outfile_name)
        batch_outfile.close()
        
        
    def createRenameCATScript(self):
        old_directory = 'F:\\Temp\\catia\\1\\M01\\'
        new_directory = 'F:\\Temp\\catia\\2\\M01\\'
        
        if not os.path.isdir(new_directory):
            os.makedirs(new_directory)
            print 'Create new directory:',new_directory
            
        name_space_file = 'name_space_m01.csv'
        data = np.genfromtxt(name_space_file, delimiter=',', skip_header=1, dtype=str)
        part_number_old_array_m01 = data[:,1]
        part_chinese_name_array_m01 = data[:,2]
        part_directory_new_array_m01 = data[:,19]
        part_number_new_array_m01 = data[:,22]
        product_number_array_m01 = data[:,24]
        
        name_space_file = 'name_space_m02.csv'
        data = np.genfromtxt(name_space_file, delimiter=',', skip_header=1, dtype=str)
        part_number_old_array_m02 = data[:,1]
        part_chinese_name_array_m02 = data[:,2]
        part_directory_new_array_m02 = data[:,19]
        part_number_new_array_m02 = data[:,22]
        product_number_array_m02 = data[:,24]
        
        part_number_old_array = list(part_number_old_array_m01) + list(part_number_old_array_m02)
        part_chinese_name_array = list(part_chinese_name_array_m01) + list(part_chinese_name_array_m02)
        part_directory_new_array = list(part_directory_new_array_m01) + list(part_directory_new_array_m02)
        part_number_new_array = list(part_number_new_array_m01) + list(part_number_new_array_m02)
        product_number_array = list(product_number_array_m01) + list(product_number_array_m02)
        
        filenames = getFiles(old_directory)
        directories = {}
        for filename in filenames:
            if isSuffixFile(filename[1],suffixs=['CATPart','CATDrawing']):
                if filename[0] not in directories:
                    directories[filename[0]] = {'CATPart':[],'CATDrawing':[]}
        
        for directory in directories:
            for filename in filenames:
                if filename[0] is directory:
                    if isSuffixFile(filename[1],suffixs=['CATPart']):
                        directories[directory]['CATPart'].append(filename[1])
                    if isSuffixFile(filename[1],suffixs=['CATDrawing']):
                        directories[directory]['CATDrawing'].append(filename[1])
        
        outfile = open('rename.CATScript', 'w')
        print >>outfile, 'Language=\"VBSCRIPT\"'
        print >>outfile, 'Sub CATMain()'

        for directory in directories:
            print directory
            directories[directory]['CATDrawing'].sort()
#            print directories[directory]['CATDrawing']
            for i, drawing_name in enumerate(directories[directory]['CATDrawing']):
#                print drawing_name
#                print i
                part_name_old = r'%s\%s' % (directory,directories[directory]['CATPart'][0])
                drawing_name_old = r'%s\%s' % (directory,drawing_name)
                rows_number = getRowsNumber(part_number_old_array,part_name_old)
                part_name_old = part_name_old.decode('gbk').encode('utf-8')
                drawing_name_old = drawing_name_old.decode('gbk').encode('utf-8')
                part_number = r'%s %s' % (part_directory_new_array[rows_number],part_chinese_name_array[rows_number])
                new_directory_name = '%s%s' % (new_directory,part_number)
                part_name_new = r'%s\%s.CATPart' % (new_directory_name,part_directory_new_array[rows_number])
                drawing_numner_new = part_number_new_array[rows_number]
                drawing_numner_new = str(int(drawing_numner_new) + i*100)
                drawing_name_new = r'%s\%s.CATDrawing' % (new_directory_name,drawing_numner_new)
                export_suffix = 'pdf'
                export_name_new= r'%s\%s.%s' % (new_directory_name,drawing_numner_new,export_suffix)
                
#                print rows_number
#                print part_name_old.decode('utf-8').encode('gbk')
#                print drawing_name_old.decode('utf-8').encode('gbk')
#                print part_number.decode('utf-8').encode('gbk')
#                print new_directory_name.decode('utf-8').encode('gbk')
#                print part_name_new.decode('utf-8').encode('gbk')
#                print drawing_name_new.decode('utf-8').encode('gbk')
#                print export_name_new.decode('utf-8').encode('gbk')
#                print
                
                new_directory_name = new_directory_name.decode('utf-8').encode('gbk')
    
                if not os.path.isdir(new_directory_name):
                    os.makedirs(new_directory_name)
#                    print 'Create new directory:',new_directory_name.decode('gbk').encode('utf-8')
                    print 'Create new directory:',new_directory_name
                
                lines = r"""
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
""" % (drawing_name_old,part_name_old,part_number,part_name_new,drawing_name_new,export_name_new,export_suffix)
            
                print >>outfile, lines
            
        print >>outfile, 'End Sub'
        outfile.close()
        
    def rename(self):
        old_name = 'F:\\Temp\\catia\\rename\\M01\\97.174 安全绳锁头\\99.pdf'
        new_name = 'F:\\Temp\\catia\\rename\\M01\\97.174 安全绳锁头\\98.pdf'
        cmd = 'move \"%s\" \"%s\"' % (old_name,new_name)
        cmd = cmd.decode('utf-8').encode('gbk')
        print cmd
        os.system(cmd)
        
    
        

catia = Catia()
#catia.createRenameCATScript()
#catia.createTitleblockCATScript()
#catia.createBatchFile()
catia.readExcel()