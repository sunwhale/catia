# -*- coding: utf-8 -*-
"""
Created on Fri May 05 11:06:31 2017

@author: SunJingyu
"""

import os
import numpy as np

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
        
    def createTitleblockCATScript(self):
        batch_outfile = open('create_titleblock.bat', 'w')
        
        old_directory = 'F:\\Temp\\catia\\M02_2\\'
        new_directory = 'F:\\Temp\\catia\\M02_2\\'
        
        if not os.path.isdir(new_directory):
            os.makedirs(new_directory)
            print 'Create new directory:',new_directory
            
        name_space_file = 'name_space_m02.csv'
        data = np.genfromtxt(name_space_file, delimiter=',', skip_header=1, dtype=str)
        part_number_old_array = data[:,19]
        part_chinese_name_array = data[:,2]
        part_chinese_material_array = data[:,3]
        part_english_material_array = data[:,4]
        part_directory_new_array = data[:,19]
        part_number_new_array = data[:,22]
        product_number_array = data[:,24]
        
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
#            print directory
            
            part_name_old = r'%s\%s' % (directory,directories[directory]['CATPart'][0])
            drawing_name_old = r'%s\%s' % (directory,directories[directory]['CATDrawing'][0])
            rows_number = getRowsNumber(part_number_old_array,part_name_old)
            part_name_old = part_name_old.decode('gbk').encode('utf-8')
            drawing_name_old = drawing_name_old.decode('gbk').encode('utf-8')
            part_number = r'%s %s' % (part_directory_new_array[rows_number],part_chinese_name_array[rows_number])
            new_directory_name = '%s%s' % (new_directory,part_number)
            part_name_new = r'%s\%s.CATPart' % (new_directory_name,part_directory_new_array[rows_number])
            drawing_name_new = r'%s\%s.CATDrawing' % (new_directory_name,part_number_new_array[rows_number])
            export_suffix = 'pdf'
            export_name_new= r'%s\%s.%s' % (new_directory_name,part_number_new_array[rows_number],export_suffix)
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
                    lines[i] = '  Text_21 = \"%s\"\n' % drawing_number
                if 'drawing_document' in line:
                    lines[i] = '  Set drawingDocument1 = documents1.Open(\"%s\")\n' % drawing_name_old
                if 'save_as' in line:
                    lines[i] = '  drawingDocument1.SaveAs \"%s\"\n' % drawing_name_new
                if 'export_data' in line:
                    lines[i] = '  drawingDocument1.ExportData \"%s\", \"%s\"\n' % (export_name_new,export_suffix)
                    
            outfile_name = 'GB_Titleblock_%s.CATScript' % drawing_number
            outfile = open(outfile_name, 'w')
            for line in lines:
                outfile.writelines(line)
            outfile.close()
            
            print >>batch_outfile, (r'cnext -batch -macro %s' % outfile_name)
        batch_outfile.close()
        
        
    def createRenameCATScript(self):
        old_directory = 'F:\\Temp\\catia\\M01\\'
        new_directory = 'F:\\Temp\\catia\\M01_2\\'
        
        if not os.path.isdir(new_directory):
            os.makedirs(new_directory)
            print 'Create new directory:',new_directory
            
        name_space_file = 'name_space_m01.csv'
        data = np.genfromtxt(name_space_file, delimiter=',', skip_header=1, dtype=str)
        part_number_old_array = data[:,1]
        part_chinese_name_array = data[:,2]
        part_directory_new_array = data[:,19]
        part_number_new_array = data[:,22]
        product_number_array = data[:,24]
        
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
            
            part_name_old = r'%s\%s' % (directory,directories[directory]['CATPart'][0])
            drawing_name_old = r'%s\%s' % (directory,directories[directory]['CATDrawing'][0])
            rows_number = getRowsNumber(part_number_old_array,part_name_old)
            part_name_old = part_name_old.decode('gbk').encode('utf-8')
            drawing_name_old = drawing_name_old.decode('gbk').encode('utf-8')
            part_number = r'%s %s' % (part_directory_new_array[rows_number],part_chinese_name_array[rows_number])
            new_directory_name = '%s%s' % (new_directory,part_number)
            part_name_new = r'%s\%s.CATPart' % (new_directory_name,part_directory_new_array[rows_number])
            drawing_name_new = r'%s\%s.CATDrawing' % (new_directory_name,part_number_new_array[rows_number])
            export_suffix = 'pdf'
            export_name_new= r'%s\%s.%s' % (new_directory_name,part_number_new_array[rows_number],export_suffix)
            
            print rows_number
            print part_name_old.decode('utf-8').encode('gbk')
            print drawing_name_old.decode('utf-8').encode('gbk')
            print part_number.decode('utf-8').encode('gbk')
            print new_directory_name.decode('utf-8').encode('gbk')
            print part_name_new.decode('utf-8').encode('gbk')
            print drawing_name_new.decode('utf-8').encode('gbk')
            print export_name_new.decode('utf-8').encode('gbk')
            print
            
            new_directory_name = new_directory_name.decode('utf-8').encode('gbk')

            if not os.path.isdir(new_directory_name):
                os.makedirs(new_directory_name)
                print 'Create new directory:',new_directory_name.decode('gbk').encode('utf-8')
            
#            new_name = 'new'
#            export_suffix = r'pdf'
#            drawing_name_old = r'F:\Temp\catia\rename\M01\97.175 内六角螺钉\97.175.CATDrawing'
#            part_name_old = r'F:\Temp\catia\rename\M01\97.175 内六角螺钉\97.175.CATPart'
#            part_number = r'%s' % new_name
#            drawing_name_new = r'F:\Temp\catia\rename\M01\97.175 内六角螺钉\%s.CATDrawing' % new_name
#            part_name_new = r'F:\Temp\catia\rename\M01\97.175 内六角螺钉\%s.CATPart' % new_name
#            export_name_new = r'F:\Temp\catia\rename\M01\97.175 内六角螺钉\%s.%s' % (new_name,export_suffix)
        
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
catia.createTitleblockCATScript()
#catia.createBatchFile()