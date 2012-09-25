# -*- coding: utf-8 -*-
'''
Created on 2012-3-26

@author: yunshandi
'''
import re
import subprocess
import os.path
#import codecs
import xlrd
import xlwt
from xlutils.copy import copy #http://pypi.python.org/pypi/xlutils
from xlrd import open_workbook #http://pypi.python.org/pypi/xlrd
from xlwt import easyxf #http://pypi.python.org/pypi/xlwt

# workbook object to store the optimized results
rb = open_workbook("android-permission-check\\template.xls",formatting_info=True)
r_sheet = rb.sheet_by_index(0) #read only copy to introspect the file
#r_sheet = rb.sheet_by_index(1) #read only copy to introspect the file
wb = copy(rb) #a writable copy (I can't read values out of this, only write to it)
w_sheet = wb.get_sheet(0) #the sheet to write to within the writable copy
#w_sheet = wb.get_sheet(1) #the sheet to write to within the writable copy

# initial row and col
row = 1
col = 169
x =0

# the repository files to be analysis
path = 'android-permission-check\\data\\ls'
# file for write final results
results = open('android-permission-check\\results.txt', 'a')

# search token when walk through AndroidManifest.xml
token = "uses-permission"
# regex
regex = '\<uses-permission\s[^\t\n\s]+|\<uses-permission[\t\n\s]+.*|\<permission\s[^\t\n\s]+|\<permission[\t\n\s]+.*'

# work through the svn ls results
for root, dirs, files in os.walk(path):
    for file in files:
        print "start:"
        f = os.path.join(root, file)
        print f
        reponame = file.split('.')[0]
        print reponame
        
        # find repository path
        # file with list of repositories
        reps = open('android-permission-check\\reps.txt', 'r')
        while 1:
            repsline = reps.readline()
            if not repsline:
                break
            if repsline.find(reponame) > -1:
                repopath = "".join(repsline.strip())
                print repopath
                break
        reps.close()
        
        # write repository header to txt file 
        results.write('####################################################\n')
        results.write('Repository: ' + reponame +'\n')
        results.write('####################################################\n')
        
        # write repository name to excel
        #w_sheet.write(row, 0, reponame)
        ls = open(f, 'r')
        while 1:
            l = ls.readline()
            if not l:
                break
            else:
                if l.find('AndroidManifest.xml') > -1 and l.find('proj/tags') == -1 and l.find('/bvt/') == -1 and l.find('/doc/') == -1 and l.find('proj/branches') == -1 and l.find('/test/') == -1 and l.find('/autotest') == -1 and l.find('/document/') == -1 and l.find('/tests/') == -1 and l.find('/samples/') == -1 and l.find('/sample/') == -1:
                    manifestpath = repopath + "".join(l.strip())
                    print manifestpath
                    
                    # manifile = open('android-permission-check/data/manifile/' + reponame + '.xml', 'a')
                    # print manifestpath
                    cmd = 'svn cat ' + manifestpath
                    process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
                    out = process.stdout.read()                    

                    q = re.compile(regex)
                    ms = q.findall(out)
                    print ms
                    
                    if len(ms) >= 1: 
                        prj = l.split('/')[0]
                        #results.write('Project: ' + prj +'\n')
                        print 'Project: ' + prj
                        # write project name into excel
                        #if r_sheet.cell_value(row-1, 0) == prj:
                        #    row = row - 1
                        
                        # write svn repository and project name
                        w_sheet.write(row, 0, reponame)
                        w_sheet.write(row, 1, prj)
                            
                        #results.write('Path: ' + manifestpath +'\n')
                        print 'Path: ' + manifestpath
                        #results.write('Permissions: ' +'\n') 
                        print 'Permissions: '
                        for m in ms:
                            # remove the line ends
                            ll = "".join(m.strip())
                            print ll
                            #ll = j.split('"')[1]
                            #print ll
                            
                            if ll.find('uses-permission') > -1:
                                if ll.find('android.permission') > -1:
                                    jj = ll.split('"')[1].split('.')[-1]
                                else:
                                    jj = ll.split('"')[1]
                                    # x = x + 1
                                #results.write('uses-permission: ' + jj +'\n')
                                print jj
                            else:
                                jj = ll.split('"')[1]
                                #results.write('permission: ' + jj +'\n')
                                #x = x + 1
                                print jj
                            
                            #xx = col + x
                            #print 'xx is: ' + str(xx)
                            for colx in range(2, col):
                                if r_sheet.cell_value(0,colx) == jj:
                                    w_sheet.write(row, colx, 'X')
                                else:
                                    if colx == col:
                                        x = x + 1
                                        w_sheet.write(0, col + x, jj)   # If the permission not found, write it in col header
                        #results.write('--------------------------------------------------------------------------------------------------------\n')
                        print '-------------------------------------------------------------------------------------------------------------'
                    row = row + 1
        ls.close()

wb.save('android-permission-check\\android-permission-final.xls')
#results.write('@@@@@@@@@@@@@@@\n')
#results.write('Done!\n')
print 'Done!'
#results.write('@@@@@@@@@@@@@@@\n')
#results.close()
