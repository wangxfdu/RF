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
    def __init__(self, name = "NA", share = 0, valueOrig = 1, buyInDate = 0, expireDate = 0, customerInterest = 0, feeRate = 0):
        self.name = name
        self.share = share
        self.valueOrig = valueOrig
        self.buyInDate = buyInDate
        self.expireDate = expireDate
        self.customerInterest = customerInterest
        self.feeRate = feeRate #not used now
        self.relativeDays = 0 # days of expireDate - referenceDate
        pass
    
    def dump(self, level = 'v'):
        logPrint(u"[产品] {}: 份额 {}, 初始净值 {}, 交易日 {}, 到期日 {}, 业绩比较基准 {}, 费率 {}".format(
                 self.name,
                 self.share,
                 self.valueOrig,
                 self.buyInDate,
                 self.expireDate,
                 self.customerInterest,
                 self.feeRate), level)
    
    def isValid(self):
        if self.share == None or self.valueOrig == None or self.buyInDate == None or self.expireDate == None \
                or self.customerInterest == None or self.feeRate == None :
            return False
        return True
    
class Capital():
    def __init__(self, name = u"底层资产", money = 0, buyInDate = 0, valueDate = None, \
                 expireDate = 0, feeRate = 0, fundInterest = 0):
        self.name = name
        self.money = money
        self.buyInDate = buyInDate
        #起息日
        self.valueDate = valueDate
        self.expireDate = expireDate
        self.feeRate = feeRate
        self.fundInterest = fundInterest
        if valueDate == None:
            valueDate = buyInDate + datetime.timedelta(days = 0)
        self.relativeDays = 0 # days of expireDate - referenceDate

    def isValid(self):
        if self.money == None or self.buyInDate == None or self.valueDate == None \
                or self.expireDate == None or self.feeRate == None or self.fundInterest == None :
            return False
        return True
            
    def dump(self, level = 'v'):
        logPrint(u"[底层资产] {}: 金额 {}, 交易日 {}, 起息日 {}, 到期日 {}, 收益率 {}, 费率 {}".format(
                 self.name,
                 self.money,
                 self.buyInDate,
                 self.valueDate,
                 self.expireDate,
                 self.fundInterest,
                 self.feeRate), level)


class FundOfFund():
    def __init__(self, netValue = 1, refDate = 0, cash = 0, share = 0):
        self.netValue = netValue
        self.refDate = refDate
        self.cash = cash
        self.share = share
        self.pendingAsset = 0
        self.pendingShareDraw = 0

    def clone(self):
        copy = FundOfFund(netValue = self.netValue, refDate = self.refDate, cash = self.cash, share = self.share)
        copy.pendingAsset = 0
        self.pendingShareDraw = 0
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

def logPrint(log, level = 'i'):
    '''
    e: error
    w: warning
    i: info
    v: verbose output
    '''

    if level == 'v' and not verboseChk.IsChecked():
        return

    if level == 'e' :
        logText.SetDefaultStyle(wx.TextAttr(wx.RED))
    elif level == "w":
        logText.SetDefaultStyle(wx.TextAttr(wx.RED))
    elif level == "v":
        logText.SetDefaultStyle(wx.TextAttr(wx.BLUE))
    else:
        logText.SetDefaultStyle(wx.TextAttr(wx.BLACK))

    logText.AppendText(log + "\n")

def cellGetDate(cell, datemode):
    if cell.ctype == xlrd.XL_CELL_DATE:
        return xlrd.xldate_as_datetime(cell.value, datemode).date()
    else:
        return None
    
def cellGetNumber(cell):
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        return cell.value
    elif cell.ctype == xlrd.XL_CELL_TEXT:
        return float(cell.value.replace(',',''))
    else:
        return None

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
    netValueCalculate(FOF = FOF.clone(), fundList = fundList, refDate = FOF.refDate,
                      targetDate = wxdate2pydate(dpc.GetValue()))

def doClear(evt):
    logText.Clear()
    
def netValueCalculate(FOF, fundList, refDate, targetDate):
    for i in range(len(fundList)):
        fundList[i].relativeDays = (fundList[i].expireDate - refDate).days
        logPrint(u'{} 剩余天数 {}'.format(fundList[i].name, fundList[i].relativeDays), 'v')
        if (fundList[i].buyInDate - refDate).days > 0:
            logPrint(u'{} 买入日期晚于估值表日期'.format(fundList[i].name), 'w')

    fundList.sort(key = lambda d: d.relativeDays)
    logPrint(u"基金按到期日期排序", 'v')
    for i in range(len(fundList)):
        logPrint(u'{} 剩余天数 {}'.format(fundList[i].name, fundList[i].relativeDays), 'v')

    for i in range(len(capitalList)):
        capitalList[i].relativeDays = (capitalList[i].expireDate - refDate).days
        logPrint(u'{} 剩余天数 {}'.format(capitalList[i].name, capitalList[i].relativeDays), 'v')
        if (capitalList[i].buyInDate - refDate).days > 0:
            logPrint(u"{} 买入日期晚于估值表日期".format(capitalList[i].name), 'w')

    targetRelDays = (targetDate - refDate).days
    if targetRelDays <= 0 :
        logPrint(u"错误: 日期选择不在有效范围内!", 'w')
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

        for i in range(len(capitalList)):
            HandleCapital(FOF, capital = capitalList[i], days = tmpDays)  

        for i in range(len(fundList)):
            HandleChildFund(FOF, childFund = fundList[i], days = tmpDays)

        print "pending  {} ".format(FOF.pendingAsset)
        FOF.netValue = (FOF.netValue * FOF.share + FOF.pendingAsset)/ (FOF.share - FOF.pendingShareDraw)
        FOF.pendingAsset = 0
        FOF.pendingShareDraw = 0
        logPrint(u"母基金净值:%s 现金 %s" % (cashFormat(FOF.netValue), cashFormat(FOF.cash) ))

        targetRelDays = targetRelDays - tmpDays
        pass

def HandleCapital(FOF, capital, days):
    # we can only update the copy of original FOF
    if capital.relativeDays <= 0 :
        return
    else:
        logPrint(u'{} 净值增量计算，天数 {}'.format(capital.name, days), 'v')
        if days > capital.relativeDays :
            validDays = capital.relativeDays
        else :
            validDays = days
        valueInc = capital.money * (capital.fundInterest - capital.feeRate)/100 * validDays /365
        FOF.pendingAsset = FOF.pendingAsset + valueInc
        capital.relativeDays = capital.relativeDays - days
        print valueInc
        if capital.relativeDays <= 0:
            #资产到期，变现后增加现金资产
            period = (capital.expireDate - capital.buyInDate).days
            cashInc = capital.money + \
                capital.money * (capital.fundInterest - capital.feeRate)/100 * period / 365
            FOF.cash = FOF.cash + cashInc
            logPrint(u'产品 {} 到期，现金资产增加 {:,f}，当前总现金资产 {:,f}'.format(
                     capital.name,
                     cashInc,
                     FOF.cash
                     ))

def HandleChildFund(FOF, childFund, days):
    # we can only update the copy of original FOF
    if childFund.relativeDays <= 0 :
        return
    elif childFund.relativeDays == 1:
        assert days == 1
        cashDesc = childFund.share * FOF.netValue
        logPrint(u"提取 {}, 金额 {:,f}".format(childFund.name, cashDesc))
        '''
        period = (childFund.expireDate - childFund.buyInDate).days
        cashInc = childFund.share * FOF.netValue - \
            childFund.share * childFund.valueOrig * childFund.customerInterest /100 * period / 365
        FOF.pendingAsset = FOF.pendingAsset cashDesc
        '''
        FOF.cash = FOF.cash - cashDesc
        FOF.share = FOF.share - childFund.share
        #logPrint(u"%s 赎回，天数 %d，增加现金 %s 至母基金" % (childFund.name, period, cashFormat(cashInc)))
        childFund.relativeDays = 0
    else:
        #计算通道费用
        channelFee = childFund.share * childFund.valueOrig * childFund.feeRate / 100 * days / 365
        FOF.pendingAsset = FOF.pendingAsset - channelFee
        childFund.relativeDays = childFund.relativeDays - days   
    #update FOF.cash
    #update FOF.value
    pass

def loadValueTable(file):
    global cost, FOF

    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    #sanity check
    tableName = sheet.cell_value(0,0)
    logPrint(tableName)
    if tableName.find(u"估值表") == -1:
        return False
    nameCol = 2
    moneyCol = 5
    deposit = -1
    fundInvest = -1
    share = -1
    value = None
    for i in range(sheet.nrows):
        if value == None and sheet.cell_value(i, 1).find(u"今日单位净值") >=0 :
            value = cellGetNumber(sheet.cell(i, 2))
            continue
        elif value < 0 and sheet.cell_value(i, 1).find(u"实收资本份额折算") >=0 :
            share = sheet.cell_value(i, 3)
            continue
        elif deposit < 0 and sheet.cell_value(i, nameCol).find(u"银行存款") >= 0:
            deposit = sheet.cell_value(i, moneyCol)
            continue
        elif fundInvest < 0 and sheet.cell_value(i, nameCol) == u"基金投资" :
            fundInvest = sheet.cell_value(i, moneyCol)
            continue

    if share < 0 or value == None or deposit < 0 or fundInvest < 0 :
        return False
    dateCell = sheet.cell(2,0)
    if dateCell.ctype != xlrd.XL_CELL_DATE:
        return False

    tableDate = xlrd.xldate_as_datetime(dateCell.value, wb.datemode).date()
    logPrint(u"日期: {}，净值 {}, 份额 {}, 银行存款 {}，基金投资 {}".format(tableDate.strftime("%Y-%m-%d"), value, share, deposit, fundInvest))

    FOF = FundOfFund(netValue = value,
                     refDate = tableDate,
                     cash = deposit + fundInvest,
                     share = share)

    return True

def loadLayerTable(file):
    global fundList, capitalList
    fundList = []
    capitalList = []
    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    #sanity check
    fofCol = -1
    directFundCol = -1
    capitalCol = -1
    otherCol = -1
    for i in range(sheet.ncols):
        if fofCol < 0 and sheet.cell_value(0, i).find(u"母基金层面") >=0 :
            fofCol = i
        elif directFundCol < 0 and sheet.cell_value(0, i).find(u"理财直投") >=0 :
            directFundCol = i
        elif capitalCol < 0 and sheet.cell_value(0, i).find(u"底层资产") >=0 :
            capitalCol = i
        elif otherCol < 0 and sheet.cell_value(0, i).find(u"其他测算") >=0 :      
            otherCol = i
    if fofCol < 0 or directFundCol < 0 or capitalCol < 0 or otherCol < 0:
        return False

    logPrint(u"定位col: 母基金 {}, 直投 {}, 底层资产 {}, 其他信息 {}".format(fofCol, directFundCol, capitalCol, otherCol), 'v')

    contentRow = 2
    for i in range(contentRow, sheet.nrows):

        if sheet.cell(i, 0).ctype != xlrd.XL_CELL_EMPTY :
            f = Fund(
                    name = sheet.cell_value(i, 0).strip(),
                    share = cellGetNumber(sheet.cell(i, fofCol + 4)),
                    valueOrig = cellGetNumber(sheet.cell(i, fofCol + 3)),
                    buyInDate = cellGetDate(sheet.cell(i, fofCol), wb.datemode),
                    expireDate = cellGetDate(sheet.cell(i, fofCol + 5), wb.datemode),
                    customerInterest =  cellGetNumber(sheet.cell(i, otherCol)),
                    feeRate = cellGetNumber(sheet.cell(i, otherCol + 1))
                    )
            f.dump()
            if f.isValid() :
                #f.dump()
                fundList.append(f)
            else:
                logPrint(u"{} 数据错误, 不纳入计算".format(f.name), 'w')

        if sheet.cell(i, capitalCol).ctype != xlrd.XL_CELL_EMPTY and \
                sheet.cell_value(i, otherCol + 4) == u"是" :
            c = Capital(
                    name = sheet.cell_value(i, capitalCol + 2),
                    money = cellGetNumber(sheet.cell(i, capitalCol + 3)),
                    buyInDate = cellGetDate(sheet.cell(i, capitalCol + 0), wb.datemode),
                    valueDate = cellGetDate(sheet.cell(i, capitalCol + 5), wb.datemode),
                    expireDate = cellGetDate(sheet.cell(i, capitalCol + 6), wb.datemode),
                    feeRate = cellGetNumber(sheet.cell(i, otherCol + 2)),
                    fundInterest = cellGetNumber(sheet.cell(i, capitalCol + 4))
                    )

            if c.isValid() :
                c.dump()
                capitalList.append(c)
            else:
                logPrint(u"{} 数据错误, 不纳入计算".format(c.name), 'w')

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
        layerTableText.SetValue(os.path.basename(path))
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

logText = wx.TextCtrl(bkg, style = wx.TE_AUTO_SCROLL | wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2)
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
'''
fundList.append(Fund(name = "RF001", share = 10000, buyInDate = datetime.date(2016, 11, 12), expireDate = datetime.date(2017, 11, 12), 
                     customerInterest = 5.5, feeRate = 0))
fundList.append(Fund(name = "RF002", share = 20000, buyInDate = datetime.date(2016, 11, 2), expireDate = datetime.date(2017, 11, 2), 
                     customerInterest = 3, feeRate = 0))
'''
capitalList = []
'''
capitalList.append(Capital(name = "HS18-1", money = 15000, buyInDate = datetime.date(2017, 1, 4), expireDate = datetime.date(2018, 3, 4), 
                     feeRate = 1.5, fundInterest = 7))
capitalList.append(Capital(name = "HS18-2", money = 12000, buyInDate = datetime.date(2017, 1, 1), expireDate = datetime.date(2018, 3, 4), 
                     feeRate = 1.4, fundInterest = 6))
'''
FOF = None

frame.Show()

app.MainLoop()
