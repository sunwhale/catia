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
            print sheet.name,sheet.nrows,sheet.ncols
            header = sheet.row_values(0)
            for i in range(sheet.ncols):
                data[header[i]] = sheet.col_values(i)
                
        for sheet in [sheet2]:
            print sheet.name,sheet.nrows,sheet.ncols
            header = sheet.row_values(0)
            for i in range(sheet.ncols):
                data[header[i]] += sheet.col_values(i)
        
        
        print data[u'原始编号']
        print data[u'中文名称']
        print data[u'中文材料']
        print data[u'英文材料']
        print data[u'装配处']
        print data[u'页数']
        print data[u'图纸尺寸']
        print data[u'任务分配']
        print data[u'发动机类型']
        print data[u'发动机类型编号']
        print data[u'发动机级别']
        print data[u'发动机级别编号']
        print data[u'部件分组']
        print data[u'部件分组编号']
        print data[u'零件分组']
        print data[u'零件分组编号']
        print data[u'零件分组统计']
        print data[u'零件号']
        print data[u'三维文件名称']
        print data[u'页号']
        print data[u'版本号']
        print data[u'二维文件名称']
        print data[u'产品序号']
        print data[u'产品编号']



    def createTitleblockCATScript(self):
        batch_outfile = open('create_titleblock.bat', 'w')
        
        old_directory = 'F:\\Temp\\catia\\2\\M01\\'
        new_directory = 'F:\\Temp\\catia\\2\\M01\\'
        
        if not os.path.isdir(new_directory):
            os.makedirs(new_directory)
            print 'Create new directory:',new_directory
            
        name_space_file = 'name_space_m01.csv'
        data = np.genfromtxt(name_space_file, delimiter=',', skip_header=1, dtype=str)

        part_number_old_array = list(data[:,19])
        part_chinese_name_array = list(data[:,2])
        part_chinese_material_array = list(data[:,3])
        part_english_material_array = list(data[:,4])
        
        paper_size_array = list(data[:,7])
        engine_type_array = list(data[:,9])
        engine_level_array = list(data[:,11])
        engine_component_array = list(data[:,13])
        engine_element_type_array = list(data[:,15])
        engine_element_number_array = list(data[:,18])
        
        page_number_array = list(data[:,20])
        version_number_array = list(data[:,21])

        part_directory_new_array = list(data[:,19])
        part_number_new_array = list(data[:,22])
        product_number_array = list(data[:,24])
        del(data)
        
        name_space_file = 'name_space_m02.csv'
        data = np.genfromtxt(name_space_file, delimiter=',', skip_header=1, dtype=str)
        part_number_old_array += list(data[:,19])
        part_chinese_name_array += list(data[:,2])
        part_chinese_material_array += list(data[:,3])
        part_english_material_array += list(data[:,4])
        
        paper_size_array += list(data[:,7])
        engine_type_array += list(data[:,9])
        engine_level_array += list(data[:,11])
        engine_component_array += list(data[:,13])
        engine_element_type_array += list(data[:,15])
        engine_element_number_array += list(data[:,18])
        
        page_number_array += list(data[:,20])
        version_number_array += list(data[:,21])
    
        part_directory_new_array += list(data[:,19])
        part_number_new_array += list(data[:,22])
        product_number_array += list(data[:,24])
        del(data)
        
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
            print directories[directory]['CATDrawing']
            for i, drawing_name in enumerate(directories[directory]['CATDrawing']):
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
                drawing_chinese_material = part_chinese_material_array[rows_number]
                drawing_english_material = part_english_material_array[rows_number]
                drawing_chinese_name = part_chinese_name_array[rows_number]
                drawing_number = part_number_new_array[rows_number]
                
                print rows_number
                print part_name_old.decode('utf-8').encode('gbk')
                print drawing_name_old.decode('utf-8').encode('gbk')
                print part_number.decode('utf-8').encode('gbk')
                print new_directory_name.decode('utf-8').encode('gbk')
                print part_name_new.decode('utf-8').encode('gbk')
                print drawing_name_new.decode('utf-8').encode('gbk')
                print export_name_new.decode('utf-8').encode('gbk')
                print drawing_chinese_material.decode('utf-8').encode('gbk')
                print drawing_english_material.decode('utf-8').encode('gbk')
                print
                
                infile = open('GB_Titleblock.CATScript', 'r')
                lines = infile.readlines()
                infile.close()
                
                script_directory = 'F:\\GitHub\\catia\\createTitleblockCATScript'
                if not os.path.isdir(script_directory):
                    os.makedirs(script_directory)
                    print 'Create new directory:',script_directory
            
                qrcode_image_url = 'http://47.93.195.1/%s.htm' % drawing_numner_new
                qrcode_image = qrcode.make(qrcode_image_url)
                qrcode_image_full_name = '%s\\%s.png' % (script_directory,drawing_numner_new)
                qrcode_image.save(qrcode_image_full_name)
                
                html_directory = 'F:\\Cloud\\wwwroot\\turbine\\'
                html_outfile_name = '%s\\%s.htm' % (html_directory,drawing_numner_new)
                outfile = open(html_outfile_name, 'w')
                print >>outfile, """<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>%s</title>
</head>
<body>

<table width="100%%">""" % part_number
                print >>outfile,'<tr><td width=\"40%%\"><b>中文名称</b></td>         <td>%s</td></tr>' % part_chinese_name_array[rows_number]
                print >>outfile,'<tr><td><b>图纸文件编号</b></td>     <td>%s</td></tr>' % drawing_numner_new
                print >>outfile,'<tr><td><b>对应3D零件编号</b></td>   <td>%s</td></tr>' % part_directory_new_array[rows_number]
                print >>outfile,'<tr><td><b>发动机类型</b></td>       <td>%s</td></tr>' % engine_type_array[rows_number]
                print >>outfile,'<tr><td><b>发动机级别</b></td>       <td>%s</td></tr>' % engine_level_array[rows_number]
                print >>outfile,'<tr><td><b>部件</b></td>             <td>%s</td></tr>' % engine_component_array[rows_number]
                print >>outfile,'<tr><td><b>零件类型</b></td>         <td>%s</td></tr>' % engine_element_type_array[rows_number]
                print >>outfile,'<tr><td><b>零件图纸号</b></td>       <td>%s</td></tr>' % engine_element_number_array[rows_number]
                print >>outfile,'<tr><td><b>纸张大小</b></td>         <td>%s</td></tr>' % paper_size_array[rows_number]
                print >>outfile,'<tr><td><b>页号</b></td>             <td>%s</td></tr>' % page_number_array[rows_number]
                print >>outfile,'<tr><td><b>版本号</b></td>           <td>%s</td></tr>' % version_number_array[rows_number]
                print >>outfile,'<tr><td><b>中文材料</b></td>         <td>%s</td></tr>' % drawing_chinese_material
                print >>outfile,'<tr><td><b>英文材料</b></td>         <td>%s</td></tr>' % drawing_english_material
                print >>outfile,"""</table>

</body>
</html>"""
                outfile.close()
                
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