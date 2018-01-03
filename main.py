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
    def __init__(self, name = "NA", share = 0, valueOrig = 1,
                 buyInDate = None, expireDate = None,
                 valueDate = None, drawDate = None,
                 customerInterest = 0, feeRate = 0):
        self.name = name
        self.share = share
        self.valueOrig = valueOrig
        self.buyInDate = buyInDate
        self.expireDate = expireDate
        if drawDate == None and self.expireDate != None :
            self.drawDate = self.expireDate + datetime.timedelta(days = 0)
        else:
            self.drawDate = drawDate
        if valueDate == None and self.drawDate != None :
            self.valueDate = self.drawDate + datetime.timedelta(days = -1)
        else:
            self.valueDate = valueDate

        self.customerInterest = customerInterest
        self.drawValue = 0 #init value = FOF ref value
        self.feeRate = feeRate #not used now
        self.relativeDays = 0 # days of expireDate - referenceDate
        self.datePoint = None #计算点
        pass
    
    def dump(self, level = 'v'):
        logPrint(u"[产品] {}: 份额 {}, 初始净值 {}, 交易日 {}, 到期日 {}, 业绩比较基准 {}, 费率 {}，提取日 {}, 净值日 {}".format(
                 self.name,
                 self.share,
                 self.valueOrig,
                 self.buyInDate,
                 self.expireDate,
                 self.customerInterest,
                 self.feeRate,
                 self.drawDate,
                 self.valueDate), level)
    
    def isValid(self):
        if self.share == None or self.valueOrig == None or self.buyInDate == None or self.expireDate == None \
                or self.drawDate == None or self.valueDate == None \
                or self.customerInterest == None or self.feeRate == None :
            return False
        #TODO: other sanity checks for date
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
        self.datePoint = None #计算点

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

def showErrorDialog(msg):
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
    s: success
    '''

    if level == 'v' and not verboseChk.IsChecked():
        return

    if level == 'e' :
        logText.SetDefaultStyle(wx.TextAttr(wx.RED))
    elif level == "w":
        logText.SetDefaultStyle(wx.TextAttr(wx.RED))
    elif level == "v":
        logText.SetDefaultStyle(wx.TextAttr("#808080"))
    elif level == "s":
        logText.SetDefaultStyle(wx.TextAttr("#32CD32"))
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
    if FOF == None :
        logPrint(u"估值表未读取", 'e')
        return
    if len(fundList) == 0 :
        logPrint(u"分层表未读取", 'e')
        return
        
    netValueCalculate(FOF = FOF.clone(), fundList = fundList, refDate = FOF.refDate,
                      targetDate = wxdate2pydate(dpc.GetValue()))

def doClear(evt):
    logText.Clear()
    
def netValueCalculate(FOF, fundList, refDate, targetDate):
    dayPoints = []

    for i in range(len(fundList)):
        dayPoints.append((fundList[i].drawDate - refDate).days)
        dayPoints.append((fundList[i].valueDate - refDate).days)
        #init fund information
        fundList[i].datePoint = refDate + datetime.timedelta(days = 0)
        fundList[i].drawValue = FOF.netValue
        #TODO: add buy in date
        if (fundList[i].buyInDate - refDate).days > 0:
            logPrint(u'{} 买入日期晚于估值表日期，暂不支持'.format(fundList[i].name), 'w')

    for i in range(len(capitalList)):
        dayPoints.append((capitalList[i].expireDate - refDate).days)
        #TODO: add buy in date
        #capitalList[i].relativeDays = (capitalList[i].expireDate - refDate).days
        #logPrint(u'{} 剩余天数 {}'.format(capitalList[i].name, capitalList[i].relativeDays), 'v')
        capitalList[i].datePoint = refDate + datetime.timedelta(days = 0)
        if (capitalList[i].buyInDate - refDate).days > 0:
            logPrint(u"{} 买入日期晚于估值表日期，暂不支持".format(capitalList[i].name), 'w')

    #所有计算点排序，可能有重复
    dayPoints.sort()

    targetRelDays = (targetDate - refDate).days
    if targetRelDays <= 0 :
        logPrint(u"错误: 日期选择不在有效范围内!", 'w')
        return

    tmpDate = refDate

    while targetRelDays > 0 :
        for i in range(len(dayPoints)):
            if dayPoints[i] <= 0 :
                continue
            else:
                break
        #找到最小天数或未找到
        if dayPoints[i] > 0 and dayPoints[i] < targetRelDays:
            tmpDays = dayPoints[i]
        else:
            tmpDays = targetRelDays

        tmpDate = tmpDate + datetime.timedelta(days = tmpDays)
        logPrint(u"----- 计算点: 日期 {} -----".format(tmpDate))
        print dayPoints

        #处理资产 收益/费用/到期
        for i in range(len(capitalList)):
            HandleCapital(FOF, capital = capitalList[i], nDays = tmpDays)  

        #处理基金买入
        #处理产品到期        
        for i in range(len(fundList)):
            HandleChildFund(FOF, childFund = fundList[i], nDays = tmpDays)

        #处理资产买入

        print "pending  {} ".format(FOF.pendingAsset)
        FOF.netValue = (FOF.netValue * FOF.share + FOF.pendingAsset)/ (FOF.share - FOF.pendingShareDraw)
        FOF.pendingAsset = 0
        FOF.pendingShareDraw = 0
        logPrint(u"母基金净值:%s 现金 %s" % (cashFormat(FOF.netValue), cashFormat(FOF.cash) ))

        #更新净值     
        for i in range(len(fundList)):
            UpdateDrawValue(FOF, childFund = fundList[i], nDays = tmpDays)

        for i in range(len(fundList)):
            #update datePoint
            fundList[i].datePoint = fundList[i].datePoint + datetime.timedelta(days = tmpDays)
        targetRelDays = targetRelDays - tmpDays
        for i in range(len(dayPoints)) :
            dayPoints[i] = dayPoints[i] - tmpDays

def HandleCapital(FOF, capital, nDays):
    # we can only update the copy of original FOF
    if (capital.datePoint - capital.expireDate).days >= 0 :
        return
    assert (capital.expireDate - capital.datePoint).days >= nDays

    logPrint(u'{} 净值增量计算，天数 {}'.format(capital.name, nDays), 'v')

    valueInc = capital.money * (capital.fundInterest - capital.feeRate)/100 * nDays /365
    FOF.pendingAsset = FOF.pendingAsset + valueInc
    

    print valueInc
    if (capital.expireDate - capital.datePoint).days == nDays:
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

    #update datePoint
    capital.datePoint = capital.datePoint + datetime.timedelta(days = nDays)

def UpdateDrawValue(FOF, childFund, nDays):
    if (childFund.datePoint - childFund.valueDate).days >= 0 :
        return

    if (childFund.valueDate - childFund.datePoint).days == nDays :
        childFund.drawValue = FOF.netValue
        print "update draw value {}".format(childFund.datePoint)

def HandleChildFund(FOF, childFund, nDays):
    if (childFund.datePoint - childFund.drawDate).days >= 0 :
        return

    #计算托管费用
    channelFee = childFund.share * childFund.valueOrig * childFund.feeRate / 100 * nDays / 365
    FOF.pendingAsset = FOF.pendingAsset - channelFee

        
    if (childFund.drawDate - childFund.datePoint).days == nDays :
        cashDesc = childFund.share * childFund.drawValue
        logPrint(u"提取 {}, 金额 {:,f}，提取净值 {:,f}".format(
                childFund.name, cashDesc, childFund.drawValue))

        FOF.cash = FOF.cash - cashDesc
        FOF.pendingAsset = FOF.pendingAsset - cashDesc
        FOF.pendingShareDraw = FOF.pendingShareDraw + childFund.share
        #logPrint(u"%s 赎回，天数 %d，增加现金 %s 至母基金" % (childFund.name, period, cashFormat(cashInc)))
    pass

def loadValueTable(file):
    global cost, FOF

    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    #sanity check
    tableName = unicode(sheet.cell_value(0,0))
    if tableName.find(u"估值表") == -1:
        return False

    codeCol = None
    for i in range(min(4, sheet.nrows)):
        for j in range(min(4, sheet.ncols)):
            if unicode(sheet.cell_value(i, j)).find(u"科目代码") >= 0 :
                codeCol = j
    if codeCol == None:
        logPrint(u"找不到科目代码栏",'e')
        return False
                
    nameCol = codeCol + 1
    moneyCol = codeCol + 4
    deposit = None
    fundInvest = None
    share = None
    value = None
    for i in range(sheet.nrows):
        if value == None and unicode(sheet.cell_value(i, codeCol)).find(u"今日单位净值") >=0 :
            value = cellGetNumber(sheet.cell(i, codeCol + 1))
            continue
        elif share == None and unicode(sheet.cell_value(i, codeCol)).find(u"实收资本") >=0 :
            share = cellGetNumber(sheet.cell(i, codeCol + 2))
            continue
        elif deposit == None and unicode(sheet.cell_value(i, nameCol)).find(u"银行存款") >= 0:
            deposit = cellGetNumber(sheet.cell(i, moneyCol))
            continue
        elif fundInvest == None and unicode(sheet.cell_value(i, nameCol)) == u"基金投资" :
            fundInvest = cellGetNumber(sheet.cell(i, moneyCol))
            continue

    tableDate = cellGetDate(sheet.cell(2,0), wb.datemode)

    if share == None or value == None or deposit == None or fundInvest == None or \
        tableDate == None:
        return False

    logPrint(u"日期: {}，净值 {:,f}, 份额 {:,f}, 银行存款 {:,f}，基金投资 {:,f}".format(tableDate.strftime("%Y-%m-%d"), value, share, deposit, fundInvest))

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
    fofOtherCol = -1
    capitalOtherCol = -1
    #otherCol = -1
    for i in range(sheet.ncols):
        if fofCol < 0 and unicode(sheet.cell_value(0, i)).find(u"母基金层面交易") >=0 :
            fofCol = i
        elif directFundCol < 0 and unicode(sheet.cell_value(0, i)).find(u"理财直投") >=0 :
            directFundCol = i
        elif capitalCol < 0 and unicode(sheet.cell_value(0, i)).find(u"底层资产") >=0 :
            capitalCol = i
        elif fofOtherCol < 0 and unicode(sheet.cell_value(0, i)).find(u"母基金层面其他") >=0 :
            fofOtherCol = i
        elif capitalOtherCol < 0 and unicode(sheet.cell_value(0, i)).find(u"底层资产其他") >=0 :
            capitalOtherCol = i
    if fofCol < 0 or directFundCol < 0 or capitalCol < 0 or fofOtherCol < 0 or capitalOtherCol < 0:
        return False

    logPrint(u"定位col: 母基金 {}, 直投 {}, 底层资产 {}, 母基金其他 {}, 底层资产其他 {}".format(
            fofCol, directFundCol, capitalCol, fofOtherCol, capitalOtherCol),
            'v')

    contentRow = 2
    for i in range(contentRow, sheet.nrows):

        if sheet.cell(i, 0).ctype != xlrd.XL_CELL_EMPTY :
            f = Fund(
                    name = sheet.cell_value(i, 0).strip(),
                    share = cellGetNumber(sheet.cell(i, fofCol + 4)),
                    valueOrig = cellGetNumber(sheet.cell(i, fofCol + 3)),
                    buyInDate = cellGetDate(sheet.cell(i, fofCol), wb.datemode),
                    expireDate = cellGetDate(sheet.cell(i, fofCol + 5), wb.datemode),
                    drawDate = cellGetDate(sheet.cell(i, fofOtherCol + 1), wb.datemode),
                    valueDate = cellGetDate(sheet.cell(i, fofOtherCol + 2), wb.datemode),
                    feeRate = cellGetNumber(sheet.cell(i, fofOtherCol + 0))
                    )

            if f.isValid() :
                f.dump()
                fundList.append(f)
            else:
                logPrint(u"{} 数据错误, 不纳入计算".format(f.name), 'w')
                f.dump()

        if sheet.cell(i, capitalCol).ctype != xlrd.XL_CELL_EMPTY and \
                sheet.cell_value(i, capitalOtherCol + 1) == u"是" :
            c = Capital(
                    name = sheet.cell_value(i, capitalCol + 2),
                    money = cellGetNumber(sheet.cell(i, capitalCol + 3)),
                    buyInDate = cellGetDate(sheet.cell(i, capitalCol + 0), wb.datemode),
                    valueDate = cellGetDate(sheet.cell(i, capitalCol + 5), wb.datemode),
                    expireDate = cellGetDate(sheet.cell(i, capitalCol + 6), wb.datemode),
                    feeRate = cellGetNumber(sheet.cell(i, capitalOtherCol + 0)),
                    fundInterest = cellGetNumber(sheet.cell(i, capitalCol + 4))
                    )

            if c.isValid() :
                capitalList.append(c)
            else:
                logPrint(u"{} 数据错误, 不纳入计算".format(c.name), 'w')
                c.dump()

    return True

def selectValueTable(evt):
    dlg = wx.FileDialog(
        bkg, message=u"请选择 资产估值表",
        defaultDir=os.getcwd(), 
        defaultFile="",
        wildcard="Excel files (*.xls*)|*.xls*|All files(*.*)|*.*",
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
            logPrint(u"读取 资产估值表成功!", 's')
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
        wildcard="Excel files (*.xls*)|*.xls*|All files(*.*)|*.*",
        style=wx.OPEN | wx.CHANGE_DIR
        )

    # Show the dialog and retrieve the user response. If it is the OK response, 
    # process the data.
    if dlg.ShowModal() == wx.ID_OK:
        # This returns a Python list of files that were selected.
        path = dlg.GetPath()
        #print paths
        layerTableText.SetValue(os.path.basename(path))
        logPrint('open ' + path, 'v')
        if loadLayerTable(path) == True :
            logPrint(u"读取 分层表成功!", 's')
        else:
            logPrint(u"读取 分层表失败!")
        #self.LoadFileContent(path)
        #self.RefreshList(self.search.GetValue())

    # Destroy the dialog. Don't do this until you are done with it!
    # BAD things can happen otherwise!
    dlg.Destroy()

app = wx.App(False)
frame = wx.Frame(None, title = u"估值小程序(v0.2)", size = (500, 600))

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
logPrint(u"=====================================")
logPrint(u"注意事项：")
logPrint(u"1. 估值表A3单元格必须是Date格式的日期")
logPrint(u"2. 分层表中所有的日期单元格必须是Date格式")
logPrint(u"=====================================")

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
