import datetime 
import pandas as pd
import pandas.io.sql
from tkinter import *
from tkinter import filedialog
import numpy as np
import _mssql
import pymssql
import matplotlib.pyplot as plt
import matplotlib.ticker as plticker
from matplotlib import style
style.use('bmh')

now = datetime.datetime.now()
Y = str(now.year)

master = Tk()
Label(master, text="Job Number").grid(row=0)
Label(master, text="% Complete").grid(row=1)
Label(master, text="Operation").grid(row=2)
e = Entry(master)
e.grid(row=0, column=1)
e2 = Entry(master)
e2.grid(row=1, column=1)
e3 = Entry(master)
e3.grid(row=2, column=1)
e2.insert(0, '100')
e.focus_set()
conn = pymssql.connect(
    host=r'host',
    user=r'user',
    password='password',
    database='database')

###Compares Estimated vs. Actual Labor based on timecard data
def laborCostsGraph():
    j = e.get()
    jo = "%" + j + "%"
    job = "'" + jo + "'"
    o = e3.get()
    op = "%" + o + "%"
    opr = "'" + op + "'"
    pc = float(e2.get())
    p = str(pc/100)
    q = "SELECT SUM(ACT_RUN_HRS) AS ACTUAL, ("+ p + " * SUM(RUN_HRS)) AS ESTIMATE, RESOURCE_ID AS OPERATION FROM \
               OPERATION WHERE WORKORDER_TYPE = 'W' AND RESOURCE_ID NOT LIKE\
               'FIELD INSTALL' AND RESOURCE_ID NOT LIKE 'IMP WAREHOUSE'\
               AND RESOURCE_ID NOT LIKE 'INSTALLATION' AND RESOURCE_ID NOT LIKE\
               'JOBSITE' AND RESOURCE_ID NOT LIKE 'OUTSIDE SERVICE'\
               AND RESOURCE_ID NOT LIKE 'CA WAREHOUSE'\
               AND RESOURCE_ID NOT LIKE 'WAREHOUSE'\
               AND RESOURCE_ID NOT LIKE 'SHIPPING' AND RESOURCE_ID LIKE" + opr + "AND WORKORDER_BASE_ID LIKE" + job
    sql = q + 'GROUP BY RESOURCE_ID'
    laborCosts = pd.io.sql.read_sql(sql, conn)
    try:
        df = pd.DataFrame(laborCosts)
        pt = pd.pivot_table(df, index = ['OPERATION'])
        pt.plot.bar()
        plt.title('Estimate vs. Actual Hours')
        plt.xlabel('Department')
        plt.ylabel('Hours')
        plt.xticks(rotation='40')
        plt.legend()
        plt.tight_layout()
        plt.show()
    except:
        top = Toplevel()
        msg = Message(top, text="No labor found for provided job number or operation", width = 2000)
        msg.pack() 

def laborCostsGraphYTD():
    j = e.get()
    jo = "%" + j + "%"
    job = "'" + jo + "'"
    o = e3.get()
    op = "%" + o + "%"
    opr = "'" + op + "'"
    pc = float(e2.get())
    p = str(pc/100)
    q = "SELECT SUM(ACT_RUN_HRS) AS ACTUAL, ("+ p + " * SUM(RUN_HRS)) AS ESTIMATE, RESOURCE_ID AS OPERATION FROM \
               OPERATION WHERE CLOSE_DATE LIKE '%" + Y + "%' AND WORKORDER_TYPE = 'W' AND RESOURCE_ID NOT LIKE\
               'FIELD INSTALL' AND RESOURCE_ID NOT LIKE 'IMP WAREHOUSE'\
               AND RESOURCE_ID NOT LIKE 'INSTALLATION' AND RESOURCE_ID NOT LIKE\
               'JOBSITE' AND RESOURCE_ID NOT LIKE 'OUTSIDE SERVICE'\
               AND RESOURCE_ID NOT LIKE 'CA WAREHOUSE'\
               AND RESOURCE_ID NOT LIKE 'WAREHOUSE'\
               AND RESOURCE_ID NOT LIKE 'SHIPPING' AND RESOURCE_ID LIKE" + opr + "AND WORKORDER_BASE_ID LIKE" + job
    sql = q + 'GROUP BY RESOURCE_ID'
    laborCosts = pd.io.sql.read_sql(sql, conn)
    try:
        df = pd.DataFrame(laborCosts)
        pt = pd.pivot_table(df, index = ['OPERATION'])
        pt.plot.bar()
        plt.title('Estimate vs. Actual Hours')
        plt.xlabel('Department')
        plt.ylabel('Hours')
        plt.xticks(rotation='40')
        plt.legend()
        plt.tight_layout()
        plt.show()
    except:
        top = Toplevel()
        msg = Message(top, text="No labor found for provided job number or operation", width = 2000)
        msg.pack()         

###Shows costs of reworks using work orders with 05-, 06-, 07-, 08- part header cards
def COPQGraph():
    j = e.get()
    jo = "%" + j + "%"
    job = "'" + jo + "'"
    q = "SELECT PART_ID, SUM(ACT_MATERIAL_COST) AS MATERIAL, SUM(ACT_LABOR_COST) AS LABOR,  SUM(ACT_BURDEN_COST) AS BURDEN, SUM(ACT_SERVICE_COST) AS SERVICE FROM WORK_ORDER WHERE PART_ID LIKE '0%' AND BASE_ID LIKE" + job
    sq = q + 'GROUP BY PART_ID'
    sq2 = "SELECT PART_ID, SUM(ACT_MATERIAL_COST) AS MATERIAL, SUM(ACT_LABOR_COST) AS LABOR,  SUM(ACT_BURDEN_COST) AS BURDEN, SUM(ACT_SERVICE_COST) AS SERVICE FROM WORK_ORDER WHERE BASE_ID LIKE" + job
    sql = sq2 + "AND PART_ID LIKE '0%' GROUP BY PART_ID"
    reworkCosts = pd.io.sql.read_sql(sql, conn)
    try:
        df = pd.DataFrame(reworkCosts)
        pt = pd.pivot_table(df, index = ['PART_ID'])
        reworkCosts2 = pd.io.sql.read_sql(sq, conn)
        df2 = pd.DataFrame(reworkCosts2)
        pt2 = pd.pivot_table(df2, columns=['MATERIAL','LABOR','BURDEN','SERVICE'], aggfunc='sum')
        top = Toplevel()
        msg = Message(top, text=pt, width = 2000)
        msg.pack()
        chart = pt.plot.bar()
        plt.title('Rework Costs')
        plt.xlabel('Cause')
        plt.ylabel('Cost')
        plt.xticks(rotation='horizontal')
        plt.legend()
        plt.show()
    except:
        top = Toplevel()
        msg = Message(top, text="No reworks found for provided job number or operation", width = 2000)
        msg.pack() 

def COPQGraphYTD():
    j = "'%" + e.get() + "%'"
    q = "SELECT PART_ID, SUM(ACT_MATERIAL_COST) AS MATERIAL, SUM(ACT_LABOR_COST) AS LABOR, SUM(ACT_BURDEN_COST) AS BURDEN, SUM(ACT_SERVICE_COST) AS SERVICE FROM WORK_ORDER WHERE PART_ID LIKE '0%' AND BASE_ID LIKE" + j + " AND CLOSE_DATE LIKE '%" + Y + "%' GROUP BY PART_ID"
    sq2 = "SELECT PART_ID, SUM(ACT_MATERIAL_COST) AS MATERIAL, SUM(ACT_LABOR_COST) AS LABOR,  SUM(ACT_BURDEN_COST) AS BURDEN, SUM(ACT_SERVICE_COST) AS SERVICE FROM WORK_ORDER WHERE BASE_ID LIKE" + j + "AND PART_ID LIKE '0%' AND CLOSE_DATE LIKE '%" + Y + "%' GROUP BY PART_ID"
    reworkCosts = pd.io.sql.read_sql(sq2, conn)
    try:
        df = pd.DataFrame(reworkCosts)
        pt = pd.pivot_table(df, index = ['PART_ID'])
        reworkCosts2 = pd.io.sql.read_sql(q, conn)
        df2 = pd.DataFrame(reworkCosts2)
        pt2 = pd.pivot_table(df2, columns=['MATERIAL','LABOR','BURDEN','SERVICE'], aggfunc='sum')
        top = Toplevel()
        msg = Message(top, text=pt, width = 2000)
        msg.pack()
        chart = pt.plot.bar()
        plt.title('Rework Costs')
        plt.xlabel('Cause')
        plt.ylabel('Cost')
        plt.xticks(rotation='horizontal')
        plt.legend()
        plt.show()
    except:
        top = Toplevel()
        msg = Message(top, text="No reworks found for provided job number or operation", width = 2000)
        msg.pack() 

def clear():
    e.delete(0, 'end')
    e2.delete(0, 'end')
    e3.delete(0, 'end')
    e2.insert(0, '100')
    e.focus_set()

menubar = Menu(master)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Estimate vs. Actual Labor", command=laborCostsGraph)
filemenu.add_command(label="Estimate vs. Actual Labor YTD", command=laborCostsGraphYTD)
filemenu.add_command(label="Total Rework Costs", command=COPQGraph)
filemenu.add_command(label="Rework Costs YTD", command=COPQGraphYTD)
menubar.add_cascade(label="Reports", menu=filemenu)
menubar.add_command(label="Reset", command=clear)
menubar.add_command(label="Exit", command=master.quit)
master.config(menu=menubar)
master.title("Costs")
master.minsize(width=200, height=70)
mainloop()
