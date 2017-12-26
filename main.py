# -*- coding: utf-8 -*-
from __future__ import division
import wx
import os
import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats
import xlrd
import gc
import traceback
import datetime


class Fund():
    def __init__(self, name = "NA", share = 0, valueOrig = 1, buyInDate = 0, expireDate = 0, customerInterest = 0, feeRate = 0, fundInterest = 0):
        self.name = name
        self.share = share
        self.valueOrig = valueOrig
        self.buyInDate = buyInDate
        self.expireDate = expireDate
        self.customerInterest = customerInterest
        self.fundInterest = fundInterest
        self.feeRate = feeRate
        self.relativeDays = 0 # days of expireDate - referenceDate
        pass
    
class Capital():
    def __ini__(self):
        pass

class FundOfFund():
    def __init__(self, netValue = 1, refDate = 0, cash = 0, share = 0):
        self.netValue = netValue
        self.refDate = refDate
        self.cash = cash
        self.share = share
        self.pendingAsset = 0

    def clone(self):
        copy = FundOfFund(netValue = self.netValue, refDate = self.refDate, cash = self.cash, share = self.share)
        copy.pendingAsset = 0
        return copy

class FileLoader():
    def __init__(self, parent, axis=""):
        self.loadButton = wx.Button(parent, label = "Open " + axis)
        self.fileText = wx.TextCtrl(parent)
        self.fileText.SetEditable(False)
        self.nameList = wx.ListBox(parent)
#        self.filterText = wx.TextCtrl(parent)
#        self.filterButton = wx.Button(parent, label = "Filter")
        self.search = wx.SearchCtrl(parent, size=(100,-1),
                                    style=wx.TE_PROCESS_ENTER)
        self.search.ShowCancelButton(True)
#        filterBox = wx.BoxSizer()
#        filterBox.Add(self.filterText)
#        filterBox.Add(self.filterButton)

        self.vbox = wx.BoxSizer(wx.VERTICAL)
        self.vbox.Add(self.loadButton, proportion = 0,flag = wx.WEST |  wx.EAST |wx.NORTH, border =5)
        self.vbox.Add(self.fileText, proportion = 0, flag = wx.WEST | wx.NORTH | wx.EAST | wx.EXPAND, border = 5)
        self.vbox.Add(self.search, proportion = 0, flag = wx.WEST | wx.NORTH |  wx.EAST | wx.EXPAND, border = 5)
#        self.vbox.Add(filterBox, proportion = 0,flag = wx.WEST | wx.NORTH, border =5)
        self.vbox.Add(self.nameList, proportion = 1, flag = wx.WEST |  wx.EAST |wx.NORTH | wx.EXPAND, border =5)
        self.nameList.Clear()

        self.loadButton.Bind(wx.EVT_BUTTON, self.OnLoadButton)
        self.search.Bind(wx.EVT_TEXT_ENTER, self.OnSearch)
        self.search.Bind(wx.EVT_SEARCHCTRL_CANCEL_BTN, self.OnSearch)
        
        self.parent = parent

        self.names_orig = np.array([])
        self.values_orig = np.array([])

        
    def OnLoadButton(self, evt):
        dlg = wx.FileDialog(
            self.parent, message="Choose a file",
            defaultDir=os.getcwd(), 
            defaultFile="",
            style=wx.OPEN | wx.CHANGE_DIR
            )

        # Show the dialog and retrieve the user response. If it is the OK response, 
        # process the data.
        if dlg.ShowModal() == wx.ID_OK:
            # This returns a Python list of files that were selected.
            path = dlg.GetPath()
            #print paths
            self.fileText.SetValue(os.path.basename(path))
            self.LoadFileContent(path)
            self.RefreshList(self.search.GetValue())

        # Destroy the dialog. Don't do this until you are done with it!
        # BAD things can happen otherwise!
        dlg.Destroy()
    
    def LoadFileContent(self, filename):
        wx.BeginBusyCursor()
        contentArr = None
        try :
            f = open(filename, "r")
            contents = f.readlines()[1:]
            #print contents
            contentArr = np.array([l.split() for l in contents])
            f.close()

            name_orig_temp = contentArr[:, 0];
            sort_arg = np.argsort(name_orig_temp)
            
            self.names_orig = name_orig_temp[sort_arg]
            self.values_orig = contentArr[:,1:][sort_arg].astype(np.float)

        except BaseException:
            showError(traceback.format_exc())
        finally :
            if f != None:
                f.close()
        del contentArr
        gc.collect()
        wx.EndBusyCursor()

        
        #print values
        #self.nameList.Clear()
        #for item in items:
        #    self.nameList.Append(item)
        #pass
    def RefreshList(self, searchString, sort=0):
        wx.BeginBusyCursor()
        try :
            self._RefreshList(searchString, sort)
        except BaseException:
            showError(traceback.format_exc())
        gc.collect()
        wx.EndBusyCursor()

    def _RefreshList(self, searchString, sort=0):
        searchString = searchString.upper().strip()
        self.nameList.Clear()
        if searchString == "" or self.names_orig.size <= 0:
            self.names_disp = self.names_orig
            self.values_disp = self.values_orig
        else :
        #sort_arg = np.argsort(self.names_orig)
        #names_sort = self.names_orig[sort_arg]
            search_array = np.array([ l.upper().find(searchString) >= 0 for l in self.names_orig])
        #print search_array
            self.names_disp = self.names_orig[search_array]
            self.values_disp = self.values_orig[search_array]
        
        #for name in self.names_disp :
        #    self.nameList.Append(name)
        self.nameList.SetItems(self.names_disp)

def showError(msg):
    dlg = wx.MessageDialog(None, msg,
                           'Message',
                           wx.OK | wx.ICON_INFORMATION
                           )
    dlg.ShowModal()
    dlg.Destroy()


def onDumpResult(evt):
    pass

def onAbout(evt):
    msg = '''All rights reserved by BoC

Author: Xiang Wang (wangxfdu@163.com)
    '''
    dlg = wx.MessageDialog(None, msg,
                           'About',
                           wx.OK | wx.ICON_INFORMATION
                           )
    dlg.ShowModal()
    dlg.Destroy()

def logPrint(log, level = 's'):
    '''
    s: standard outoupt
    v: verbose output
    '''
    doPrint = False
    if (level == 'v' and verboseChk.IsChecked()) or level == 's' :
        doPrint = True
    else:
        doPrint = False
    if doPrint:
        logText.AppendText(log + "\n")
    

def cashFormat(num):
    return format(num, ',f')

#https://www.tuicool.com/articles/AbAFJfe
def wxdate2pydate(date):
	assert isinstance(date, wx.DateTime)
	if date.IsValid():
		ymd = map(int, date.FormatISODate().split('-'))
		return datetime.date(*ymd)
	else:
		return None
    
def doCalculate(evt):
    global fundList, FOF
    netValueCalculate(FOF = FOF.clone(), fundList = fundList, refDate = datetime.date(2017, 10, 12),
                      refNetValue = 1, targetDate = wxdate2pydate(dpc.GetValue()))

def doClear(evt):
    logText.Clear()
    
def netValueCalculate(FOF, fundList, refDate, refNetValue, targetDate):
    for i in range(len(fundList)):
        fundList[i].relativeDays = (fundList[i].expireDate - refDate).days
        logPrint("%d" % fundList[i].relativeDays)

    fundList.sort(key = lambda d: d.relativeDays)
    logPrint("After sort")
    for i in range(len(fundList)):
        logPrint("%d" % fundList[i].relativeDays)


    targetRelDays = (targetDate - refDate).days
    if targetRelDays <= 0 :
        logPrint(u"错误: 日期选择不在有效范围内!")
        return
    tmpDate = refDate

    while targetRelDays > 0 :
        tmpDays = targetRelDays
        for i in range(len(fundList)):
            if fundList[i].relativeDays <= 0 :
                continue
            elif fundList[i].relativeDays == 1:
                tmpDays = 1
                break
            elif fundList[i].relativeDays <= targetRelDays:
                tmpDays = fundList[i].relativeDays - 1
                break
            else:
                tmpDays = targetRelDays
                break
        tmpDate = tmpDate + datetime.timedelta(days = tmpDays)
        logPrint(u"----- 计算点: 日期 %s -----" % tmpDate.strftime("%Y-%m-%d"))
        for i in range(len(fundList)):
            FOFValueCalculate(FOF, childFund = fundList[i], days = tmpDays)
        FOF.netValue = (FOF.netValue * FOF.share + FOF.pendingAsset)/FOF.share
        FOF.pendingAsset = 0
        logPrint(u"母基金净值:%s 现金 %s" % (cashFormat(FOF.netValue), cashFormat(FOF.cash) ))

        targetRelDays = targetRelDays - tmpDays
        pass

def FOFValueCalculate(FOF, childFund, days):
    # we can only update the copy of original FOF
    if childFund.relativeDays <= 0 :
        return
    elif childFund.relativeDays == 1:
        assert days == 1
        logPrint(u"提取 " + childFund.name)
        period = (childFund.expireDate - childFund.buyInDate).days
        cashInc = childFund.share * FOF.netValue - \
            childFund.share * childFund.valueOrig * childFund.customerInterest /100 * period / 365
        FOF.pendingAsset = FOF.pendingAsset + cashInc
        FOF.cash = FOF.cash + cashInc
        logPrint(u"%s 赎回，天数 %d，增加现金 %s 至母基金" % (childFund.name, period, cashFormat(cashInc)))
        childFund.relativeDays = 0
    else:
        logPrint(u"净值计算 " + childFund.name)
        assetInc = childFund.share * childFund.valueOrig * (childFund.fundInterest - childFund.feeRate)/100 * days /365
        FOF.pendingAsset = FOF.pendingAsset + assetInc
        childFund.relativeDays = childFund.relativeDays - days   
    #update FOF.cash
    #update FOF.value
    pass

def loadValueTable(file):
    global cost

    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    #sanity check
    tableName = sheet.cell_value(0,0)
    logPrint(tableName)
    if tableName.find(u"估值表") == -1:
        return False
    dateCell = sheet.cell(2,0)
    if dateCell.ctype != xlrd.XL_CELL_DATE:
        return False
    tableDate = xlrd.xldate_as_datetime(dateCell.value, wb.datemode).date()
    logPrint(u"日期: " + tableDate.strftime("%Y-%m-%d"))

    nrows = sheet.nrows
    return True

def loadLayerTable(file):
    '''
    global cost

    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    #sanity check
    tableName = sheet.cell_value(0,0)
    logPrint(tableName)
    if tableName.find(u"估值表") == -1:
        return False
    dateCell = sheet.cell(2,0)
    if dateCell.ctype != xlrd.XL_CELL_DATE:
        return False
    tableDate = xlrd.xldate_as_datetime(dateCell.value, wb.datemode).date()
    logPrint(u"日期: " + tableDate.strftime("%Y-%m-%d"))

    nrows = sheet.nrows
    '''
    return True

def selectValueTable(evt):
    dlg = wx.FileDialog(
        bkg, message=u"请选择 资产估值表",
        defaultDir=os.getcwd(), 
        defaultFile="",
        wildcard="Excel files (*.xlsx)|*.xlsx|All files(*.*)|*.*",
        style=wx.OPEN | wx.CHANGE_DIR
        )

    # Show the dialog and retrieve the user response. If it is the OK response, 
    # process the data.
    if dlg.ShowModal() == wx.ID_OK:
        # This returns a Python list of files that were selected.
        path = dlg.GetPath()
        #print paths
        valueTableText.SetValue(os.path.basename(path))
        logPrint('open ' + path)
        if loadValueTable(path) == True :
            logPrint(u"读取 资产估值表成功!")
        else:
            logPrint(u"读取 资产估值表失败!")
        #self.LoadFileContent(path)
        #self.RefreshList(self.search.GetValue())

    # Destroy the dialog. Don't do this until you are done with it!
    # BAD things can happen otherwise!
    dlg.Destroy()
    
def selectLayerTable(evt):
    dlg = wx.FileDialog(
        bkg, message=u"请选择 分层表",
        defaultDir=os.getcwd(), 
        defaultFile="",
        wildcard="Excel files (*.xlsx)|*.xlsx|All files(*.*)|*.*",
        style=wx.OPEN | wx.CHANGE_DIR
        )

    # Show the dialog and retrieve the user response. If it is the OK response, 
    # process the data.
    if dlg.ShowModal() == wx.ID_OK:
        # This returns a Python list of files that were selected.
        path = dlg.GetPath()
        #print paths
        valueTableText.SetValue(os.path.basename(path))
        logPrint('open ' + path)
        if loadLayerTable(path) == True :
            logPrint(u"读取 分层表成功!")
        else:
            logPrint(u"读取 分层表失败!")
        #self.LoadFileContent(path)
        #self.RefreshList(self.search.GetValue())

    # Destroy the dialog. Don't do this until you are done with it!
    # BAD things can happen otherwise!
    dlg.Destroy()

app = wx.App(False)
frame = wx.Frame(None, title = u"估值小程序(v0.1)", size = (500, 600))

bkg = wx.Panel(frame)
"""
load1 = FileLoader(bkg, "[X]")
load2 = FileLoader(bkg, "[Y]")
"""

hbox1 = wx.BoxSizer()
valueTableText = wx.TextCtrl(bkg, style = wx.TE_READONLY)
valueTableBtn = wx.Button(bkg, label = u"估值表")
valueTableBtn.Bind(wx.EVT_BUTTON, selectValueTable)
hbox1.Add(valueTableText, proportion = 1, flag = wx.ALL, border = 3)
hbox1.Add(valueTableBtn, flag = wx.ALL, border = 3)

hbox2 = wx.BoxSizer()
layerTableText = wx.TextCtrl(bkg, style = wx.TE_READONLY)
layerTableBtn = wx.Button(bkg, label = u"分层表")
layerTableBtn.Bind(wx.EVT_BUTTON, selectLayerTable)
hbox2.Add(layerTableText, proportion = 1, flag = wx.ALL, border = 3)
hbox2.Add(layerTableBtn, flag = wx.ALL, border = 3)

hbox3 = wx.BoxSizer()
dpc = wx.DatePickerCtrl(bkg, style = wx.DP_DROPDOWN)
caclBtn = wx.Button(bkg, label = u"计算")
caclBtn.Bind(wx.EVT_BUTTON, doCalculate)
clearBtn = wx.Button(bkg, label = u"清空输出")
clearBtn.Bind(wx.EVT_BUTTON, doClear)
verboseChk = wx.CheckBox(bkg, -1, u"详细输出")

hbox3.Add(dpc, flag = wx.ALL, border = 3)
hbox3.Add(caclBtn, flag = wx.ALL, border = 3)
hbox3.Add(clearBtn, flag = wx.ALL, border = 3)
hbox3.Add(verboseChk, flag = wx.ALL | wx.EXPAND, border = 3)

logText = wx.TextCtrl(bkg, style = wx.TE_AUTO_SCROLL | wx.TE_MULTILINE | wx.TE_READONLY)
#logText.Enabled = False
logPrint(u"初始化完成")

vbox = wx.BoxSizer(wx.VERTICAL)
vbox.Add(hbox1, flag = wx.EXPAND)
vbox.Add(hbox2, flag = wx.EXPAND)
vbox.Add(hbox3)
vbox.Add(logText, proportion = 1, flag = wx.EXPAND | wx.ALL, border = 3)

bkg.SetSizer(vbox)
"""
hbox = wx.BoxSizer()

hbox.Add(load1.vbox, flag = wx.EXPAND, proportion = 1)
hbox.Add(load2.vbox, flag = wx.EXPAND, proportion = 1)

runButton = wx.Button(bkg, label = "RUN")
bkgBox = wx.BoxSizer(wx.VERTICAL)
bkgBox.Add(hbox, proportion = 1, flag = wx.EXPAND)
bkgBox.Add(runButton, proportion = 0, flag = wx.ALL, border = 10)

runButton.Bind(wx.EVT_BUTTON, doAnalyze)

bkg.SetSizer(bkgBox)
"""
#Menu
filemenu = wx.Menu()

ID_DUMP1 = 1
ID_DUMP2 = 2
# wx.ID_ABOUT和wx.ID_EXIT是wxWidgets提供的标准ID
menuDump1 = filemenu.Append(ID_DUMP1, "TODO")
#filemenu.AppendSeparator()
menuDump2 = filemenu.Append(ID_DUMP2, "TODO")

# Help Menu
helpMenu = wx.Menu()
menuAbout = helpMenu.Append(wx.ID_ABOUT, "About", "About")

# 创建菜单栏
menuBar = wx.MenuBar()
menuBar.Append(filemenu, "&File")    # 在菜单栏中添加filemenu菜单
menuBar.Append(helpMenu, "&Help")
frame.SetMenuBar(menuBar)    # 在frame中添加菜单栏

# 设置events
frame.Bind(wx.EVT_MENU, onDumpResult, menuDump1)
frame.Bind(wx.EVT_MENU, onDumpResult, menuDump2)
frame.Bind(wx.EVT_MENU, onAbout, menuAbout)

#-----------------------
#test data
fundList = []
fundList.append(Fund(name = "SF001", share = 10000, buyInDate = datetime.date(2016, 11, 12), expireDate = datetime.date(2017, 11, 12), 
                     customerInterest = 5.5, feeRate = 1, fundInterest = 7))
fundList.append(Fund(name = "SF002", share = 20000, buyInDate = datetime.date(2016, 11, 2), expireDate = datetime.date(2017, 11, 2), 
                     customerInterest = 3, feeRate = 0.5, fundInterest = 7))
fundList.append(Fund(name = "SF003", share = 30000, buyInDate = datetime.date(2017, 3, 4), expireDate = datetime.date(2018, 3, 4), 
                     customerInterest = 4, feeRate = 1.5, fundInterest = 7))

FOF = FundOfFund(netValue = 1, refDate = datetime.date(2017, 10, 12), cash = 10000, share = 100000)

frame.Show()

app.MainLoop()
