# coding=utf-8
__author__ = '@dmarcelinobr'

import pandas as pd
from openpyxl import load_workbook

import xlrd, xlwt
from xlutils.copy import copy as xl_copy

tot = 15
quota  = 'quota.xls'

def runApp():
    qname = input('Enter Question name:')
    templateBuild(qname)
    makeQuota(qname)

def templateBuild(n):
    designFile = 'Design.dat'
    fin = open("src/template.xml", "rt")
    fout = open(n+".xml", "wt")
    for line in fin:
        fout.write(line.replace('Q1', n))
    fin.close()
    fout.close()

def getVersions():
    print("get number of versions")

def appendDefines():
    print("append defines for each version")
    
def makeQuota(n):
    rb = xlrd.open_workbook(quota, formatting_info=True)
    wb = xl_copy(rb)
    sheetName = n+"_Maxdiff"
    # add_sheet is used to create sheet.  
    sheet1 = wb.add_sheet(sheetName)
    sheet1.write(0, 0, "#="+sheetName)
    for x in range(tot):
        sheet1.write(x+1, 0, 'ver_'+ str(x+1)) 
        sheet1.write(x+1, 1, 'inf') 
    wb.save(quota)

runApp()