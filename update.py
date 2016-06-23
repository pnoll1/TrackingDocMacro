# -*- coding: utf-8 -*-
"""
Created on Thu Jun 16 11:35:44 2016

@author: pnoll
"""
from shutil import copyfile
from xlwings import *

def update():
#update info for jobs that are already in tracking document
#project names in master schedule & tracking must match
#only searches 100 lines in master schedule & tracking document

    copyfile("S:\Pacific Tower Cranes\Sales\master schedule.xlsx", "E:\Documents\master schedule.xlsx")

    t =xw.Workbook.caller()
    ms = Workbook("E:\Documents\master schedule.xlsx")
    
    #items that need to be matched to current column number; items are split into 2 rows
    #names do not properly describe items, each column description is split across 2 rows
    col_list1 = [ 'Crane','Hook', 'Initial', 'Start','FCC#', 'Erct', 'Dis', 'Disassem']
    col_list2 = ['Customer', 'Job', 'Base','Status']
    
    #finds which column name is in and creates named variable that equals column number
    For n in range(1,43):
        For i in col_list1:
            if i is in Range((1,n).value:
                i == Range((1,n))
                
    For n in range(1,43):
        For i in col_list2:
            if i is in Range((2,n).value:
                i == Range((2,n))
    
    For J = 3 to 100:
        For K = 2 to 60:
            If ms.Worksheets("Tower Cranes").Range((J,Customer)).value = t.Worksheets("Crane").Range((K,2)).value:
                t.Worksheets("Crane").Range((K,12)) = ms.Worksheets("Tower Cranes").Range((J,Crane)).value