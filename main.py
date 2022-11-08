from PySide2 import QtCore
from PySide2.QtWidgets import QApplication, QMainWindow, QPushButton, QPlainTextEdit, QMessageBox, QFileDialog, QRadioButton
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, QDate,   QDateTime , QTime, QUrl
from PySide2.QtGui import QDesktopServices

import os
from datetime import datetime, date, timedelta
import xlsxwriter
import io
import urllib
import time

import requests

import hmac
import math

import xlrd
import xlwt
import PySide2

import _thread

import Common.ImageHandler as ImageHandler

dirname = os.path.dirname(PySide2.__file__)
plugin_path = os.path.join(dirname, 'plugins', 'platforms')
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path

AppKey = {'联球制衣厂':'9468443','朝雄制衣厂':'4048570', '朝逸饰品厂':'7464845'}
AppSecret =  {"联球制衣厂":b'SuL1xQz2I88',"朝雄制衣厂":b'vxMiqUGDZ7p',"朝逸饰品厂":b'Oo0hRGjaaPb'}
access_token = {'联球制衣厂':'4c6a563e-a96c-41f5-a691-6d7a6f92023b','朝雄制衣厂':'f7a85f0e-fb90-40b1-9bc6-b66ad3cf120b', '朝逸饰品厂':'37454060-a51a-497d-9866-0d722a1bd5cc'}

en_code = ['s','S','m','M', 'l','L','x','X']

base_url = 'https://gw.open.1688.com/openapi/'


path = os.path.realpath(os.curdir)

price_path = path + '/price.xlsx'

request_type = {
    'trade':"param2/1/com.alibaba.trade/",
    'delivery':'param2/1/com.alibaba.logistics/'
}

def GetPriceGrid():
    workbook = xlrd.open_workbook(price_path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    return worksheet

worksheet = GetPriceGrid()

def NumFormate4Print(numStr):
    res = ""
    if numStr[0] in en_code:
        for item in numStr:
            if item in en_code:
                res += item
            else:
                break
        fomateSpaceNum = 5 - len(res)
        for k in range(fomateSpaceNum):
            res = " " + res

    else:
        for item in numStr:
            if item >= '0' and item <= '9':
                res += item
            else:
                if res[0] == '9':
                    res += " "
                break
        res += "cm"
    return res

def CalSize(sizeDescription):
    _size = 0
    for i in range(len(sizeDescription)):
        if sizeDescription[i] >= '0' and sizeDescription[i] <= '9':
            _size = _size*10 + int(sizeDescription[i])
        else:
            break
    return _size

def CalPriceLocationENCode(_size):
    if _size == "":
        return -1

    if _size[0] == 'S' or _size[0] == 's':
        return 28

    if _size[0] == 'M' or _size[0] == 'm':
        return 29

    if _size[0] == 'L' or _size[0] == 'l':
        return 30

    if _size[0] == 'x' or _size[0] == 'X':
        if _size[1] == 'L' or _size[1] == 'l':
            return 31
        elif _size[1] == 'x' or _size[1] == 'X':
            if _size[2] == 'L' or _size[2] == 'l':
                return 32
            elif _size[2] == 'x' or _size[2] == 'X':
                return 33

    return -1

def CalPriceLocation(_size):
    if _size[0] in en_code:
        return CalPriceLocationENCode(_size)
    # print(CalSize(_size))
    theta = (CalSize(_size) - 90)/5
    _col = 9 + theta
    # print(_col)
    return _col

def GetCost(cargoNumber,skuInfosValue, colNum=0):
    rowIndex = -1
    for t in range(1, worksheet.nrows):
        if str(cargoNumber) == str(worksheet.cell(t, colNum).value):
            rowIndex = t
            break
    if rowIndex == -1:
        print(cargoNumber + "1==========" + worksheet.cell(t, colNum).value)
        return -1
    colIndex = CalPriceLocation(skuInfosValue)
    if colIndex != None:
        _price = worksheet.cell(rowIndex, int(colIndex)).value
    else:
        _price = ''
    if _price == '':
        print(rowIndex, colIndex)
        print(cargoNumber + "2==========" + worksheet.cell(t, colNum).value)
        _price = -1
    return float(_price)

def GetAdressAndShopName(cargoNumber):
    rowIndex = -1
    for t in range(1, worksheet.nrows):
        if cargoNumber == worksheet.cell(t, 0).value:
            rowIndex = t
            break
    if rowIndex == -1:
        return ["","",""]
    productName = worksheet.cell(rowIndex, 1).value
    adress = worksheet.cell(rowIndex, 2).value
    shopName = worksheet.cell(rowIndex, 3).value
    return [productName,adress,shopName]

def CalPageNum(totalRecord):
    return math.ceil(totalRecord/20)

def CalculateSignature(urlPath,data, shopName):
    # 构造签名因子：urlPath

    # 构造签名因子：拼装参数
    params = list()
    for key in data.keys():
        params.append(key+str(data[key]))
    params.sort()
    sortedParams = params
    assembedParams = str()
    for param in sortedParams:
        assembedParams = assembedParams+param

    # 合并签名因子
    mergedParams = urlPath + assembedParams
    mergedParams = bytes(mergedParams, 'utf8')

    # 执行hmac_sha1算法 && 转为16进制
    hex_res1 = hmac.new(AppSecret[shopName], mergedParams, digestmod="sha1").hexdigest()

    # 转为全大写字符
    hex_res1 = hex_res1.upper()

    return hex_res1

# 交易数据获取  默认1页
def GetTradeData(data, shopName):
    data['access_token'] = access_token[shopName]
    _aop_signature = CalculateSignature(request_type['trade'] + "alibaba.trade.getSellerOrderList/" + AppKey[shopName], data, shopName)
    data['_aop_signature'] = _aop_signature
    url = base_url + request_type['trade'] + "alibaba.trade.getSellerOrderList/" + AppKey[shopName]
    response = requests.post(url, data=data)

    return response.json()

# 已发货物流信息
def GetDeliveryData(data, shopName):
    data['access_token'] = access_token[shopName]
    _aop_signature = CalculateSignature(request_type['delivery'] + "alibaba.trade.getLogisticsTraceInfo.sellerView/" + AppKey[shopName], data, shopName)
    data['_aop_signature'] = _aop_signature
    url = base_url + request_type['delivery'] + "alibaba.trade.getLogisticsTraceInfo.sellerView/" + AppKey[shopName]
    response = requests.post(url, data=data)

    return response.json()


def GetSingleTradeData(data, shopName):
    data['access_token'] = access_token[shopName]
    _aop_signature = CalculateSignature(request_type['trade'] + "alibaba.trade.get.sellerView/" + AppKey[shopName], data, shopName)
    data['_aop_signature'] = _aop_signature
    url = base_url + request_type['trade'] + "alibaba.trade.get.sellerView/" + AppKey[shopName]
    response = requests.post(url, data=data)

    return response.json()

# 获取订单号列表
# 筛除刷单
def GetOrderBill2(createStartTime, createEndTime, orderstatusStr, shopName, mode = 0, limitDeliveredTime = {}):
    orderList = []

    orderstatusList = orderstatusStr.split(',')

    for orderstatus in orderstatusList:
        data = {'createStartTime': createStartTime, 'createEndTime': createEndTime, 'orderStatus': orderstatus,
                'needMemoInfo': 'true'}
        response = GetTradeData(data, shopName)
        if orderstatus == 'waitsellersend' :
            orderstatusStr = '待发货'
        if orderstatus == 'waitbuyerreceive' :
            orderstatusStr = '已发货'
        # self.LogOut('# ' + orderstatusStr + ' : ' + str(response['totalRecord']) + '条记录')
        pageNum = CalPageNum(response['totalRecord'])

        # 规格化数据
        for pageId in range(pageNum):
            data = {'page': str(pageId + 1), 'createStartTime': createStartTime, 'createEndTime': createEndTime,
                    'orderStatus': orderstatus, 'needMemoInfo': 'true'}
            response = GetTradeData(data, shopName)

            if orderstatus == 'waitsellersend' or orderstatus == 'waitbuyerreceive':
                for order in response['result']:
                    if ('sellerRemarkIcon' in order['baseInfo']) and ( order['baseInfo']['sellerRemarkIcon'] == '2' or order['baseInfo']['sellerRemarkIcon'] == '3'):
                        continue
                    elif mode != 0 and 'sellerRemarkIcon' not in order['baseInfo']:
                        if mode == 1:
                            order['baseInfo']['sellerRemarkIcon'] = '1'
                        elif mode == 4:
                            order['baseInfo']['sellerRemarkIcon'] = '4'

            orderList += response['result']

    return orderList

def CheckDelivery2(createStartTime, createEndTime, orderstatusStr, shopName, mode = 0, limitDeliveredTime = {}):
    # 1. 获得待查询订单列表
    orderIdListRaw = GetOrderBill2(createStartTime, createEndTime, orderstatusStr, shopName, mode, limitDeliveredTime)

    orderIdList = []

    deliveryErrorList = []

    for each in orderIdListRaw:
        if 'allDeliveredTime' in each['baseInfo'] and len(limitDeliveredTime) > 0:  # 根据发货时间判断是否要输出
            allDeliveredTime = int(each['baseInfo']['allDeliveredTime'][:-8])
            if allDeliveredTime < limitDeliveredTime['deleveredStartTime'] or allDeliveredTime > limitDeliveredTime[
                'deleveredEndTime']:
                continue
        orderIdList.append(each['baseInfo']['idOfStr'])

    # 2. 查询物流信息
    for orderId in orderIdList:
        data = {'orderId': int(orderId), 'webSite': '1688'}
        response = GetDeliveryData(data, shopName)
        if 'errorMessage' in response and response['errorMessage'] == '该订单没有物流跟踪信息。':
            deliveryErrorList.append(orderId)

    return deliveryErrorList

class Window:

    def __init__(self):
        super(Window, self).__init__()

        # 从文件中加载UI定义
        qfile = QFile("QtUi.ui")
        qfile.open(QFile.ReadOnly)
        qfile.close()
        self.ui = QUiLoader().load(qfile)

        # 订单时间
        today = datetime.strptime(str(date.today()), '%Y-%m-%d')
        startDateTime = today + timedelta(days = -10)
        endDateTime   = today + timedelta(days = 1)
        self.ui.startTime.setDateTime(QDateTime(QDate(startDateTime.year, startDateTime.month, startDateTime.day), QTime(0, 0, 0)))
        self.ui.endTime.setDateTime(QDateTime(QDate(endDateTime.year, endDateTime.month, endDateTime.day), QTime(0, 0, 0)))

        # 发货时间
        self.ui.deleveredStartTime.setDateTime(QDateTime(QDate(today.year, today.month, today.day), QTime(0, 0, 0)))
        self.ui.deleveredEndTime.setDateTime(QDateTime(QDate(today.year, today.month, today.day + 1), QTime(0, 0, 0)))

        self.ui.priceTablePath.setText("price.xlsx")
        self.ui.priceTablePathButton.clicked.connect(self.CheckPriceTablePath)

        self.ui.shopName.addItem("联球制衣厂")
        self.ui.shopName.addItem("朝雄制衣厂")

        self.ui.Tag.addItem("无")
        self.ui.Tag.addItem("红")
        self.ui.Tag.addItem("蓝")
        self.ui.Tag.addItem("绿")
        self.ui.Tag.addItem("黄")
        self.ui.Tag.addItem("按单号")

        self.ui.orderStatus.addItem("已发货")
        self.ui.orderStatus.addItem("待发货 + 已发货")
        self.ui.orderStatus.addItem("待发货")
        self.ui.orderStatus.addItem("待付款")

        self.ui.commit.clicked.connect(self.CheckAllParams)
        # self.ui.openUrl.clicked.connect(self.click_window2)
        # self.ui.timeType.addItem("按发货时间")
        # self.ui.timeType.addItem("按付款时间")
        # self.ui.timeType.addItem("按下单时间")

        self.errorUrl = ""

        self.calStartTime = datetime.now()

        self.ui.saveFilePath.setText("BHtmp")
        self.ui.saveFilePathButton.clicked.connect(self.CheckSaveFilePath)

        # 物流检查按钮
        self.ui.checkDelivery.clicked.connect(self.CheckDelivery)

    def CheckDelivery(self):
        self.LogOut("# 启动计算时间 : " + self.calStartTime.strftime('%Y-%m-%d %H:%M:%S'))
        shopId = self.ui.shopName.currentIndex() + 1
        shopName = self.ui.shopName.currentText()
        mode = self.ui.Tag.currentIndex()

        orderStatus = self.ui.orderStatus.currentIndex()

        startYear = self.ui.startTime.date().toString("yyyy")
        startMonth = self.ui.startTime.date().toString("MM")
        startDay = self.ui.startTime.date().toString("dd")

        createStartTime = datetime(int(startYear), int(startMonth), int(startDay)).strftime('%Y%m%d') + '000000000+0800'

        endYear = self.ui.endTime.date().toString("yyyy")
        endMonth = self.ui.endTime.date().toString("MM")
        endDay = self.ui.endTime.date().toString("dd")
        deleveredStartTime = int(self.ui.deleveredStartTime.dateTime().toString('yyyyMMddHHmmss'))
        deleveredEndTime = int(self.ui.deleveredEndTime.dateTime().toString('yyyyMMddHHmmss'))

        createEndTime = datetime(int(endYear), int(endMonth), int(endDay)).strftime('%Y%m%d') + '000000000+0800'

        isPrintOwn = self.ui.IsPrintOwn.isChecked()

        self.LogOut("# 店铺名 ：" + shopName)
        self.LogOut("# 色标 ：" + self.ui.Tag.currentText())
        self.LogOut("# 订单开始时间 ：" + self.ui.startTime.date().toString("yyyy-MM-dd"))
        self.LogOut("# 订单截止时间 ：" + self.ui.endTime.date().toString("yyyy-MM-dd"))

        isLimitDeleveredTime = self.ui.isLimitDeleveredTime.isChecked()

        if isLimitDeleveredTime:
            limitDeliveredTime = {
                'deleveredStartTime': deleveredStartTime,
                'deleveredEndTime': deleveredEndTime
            }
            self.LogOut("# 订单开始时间 ：" + self.ui.startTime.date().toString("yyyy-MM-dd"))
            self.LogOut("# 订单截止时间 ：" + self.ui.endTime.date().toString("yyyy-MM-dd"))
        else:
            limitDeliveredTime = {}

        try:
            _thread.start_new_thread(CheckDelivery2, (createStartTime, createEndTime, 'waitbuyerreceive', shopName, mode, limitDeliveredTime))
        except:
            self.LogOut("Error: 无法计算启动线程")

    def click_window2(self):
        QDesktopServices.openUrl(QUrl(self.errorUrl))


    def CheckPriceTablePath(self):
        filePath = QFileDialog.getOpenFileName(QMainWindow(), "选择文件")  # 选择目录，返回选中的路径
        self.ui.priceTablePath.setText(filePath[0])


    def CheckSaveFilePath(self):
        FileDirectory = QFileDialog.getExistingDirectory(QMainWindow(), "选择文件夹")  # 选择目录，返回选中的路径
        self.ui.saveFilePath.setText(FileDirectory)

    def CheckAllParams(self):
        self.LogOut("# 启动计算时间 : " + self.calStartTime.strftime('%Y-%m-%d %H:%M:%S'))
        shopId = self.ui.shopName.currentIndex() + 1
        shopName = self.ui.shopName.currentText()
        mode = self.ui.Tag.currentIndex()

        orderStatus = self.ui.orderStatus.currentIndex()

        # timeType = self.ui.timeType.currentIndex()

        tmp = self.ui.startTime.date();

        startYear = self.ui.startTime.date().toString("yyyy")
        startMonth = self.ui.startTime.date().toString("MM")
        startDay = self.ui.startTime.date().toString("dd")

        createStartTime = datetime(int(startYear), int(startMonth), int(startDay)).strftime('%Y%m%d') + '000000000+0800'

        endYear = self.ui.endTime.date().toString("yyyy")
        endMonth = self.ui.endTime.date().toString("MM")
        endDay = self.ui.endTime.date().toString("dd")
        deleveredStartTime = int(self.ui.deleveredStartTime.dateTime().toString('yyyyMMddHHmmss'))
        deleveredEndTime   = int(self.ui.deleveredEndTime.dateTime().toString('yyyyMMddHHmmss'))

        createEndTime = datetime(int(endYear), int(endMonth), int(endDay)).strftime('%Y%m%d') + '000000000+0800'

        isPrintOwn = self.ui.IsPrintOwn.isChecked()

        self.LogOut("# 店铺名 ：" + shopName)
        self.LogOut("# 色标 ：" + self.ui.Tag.currentText())
        self.LogOut("# 订单开始时间 ：" + self.ui.startTime.date().toString("yyyy-MM-dd"))
        self.LogOut("# 订单截止时间 ：" + self.ui.endTime.date().toString("yyyy-MM-dd"))

        isLimitDeleveredTime = self.ui.isLimitDeleveredTime.isChecked()

        if isLimitDeleveredTime:
            limitDeliveredTime = {
                'deleveredStartTime':deleveredStartTime,
                'deleveredEndTime':deleveredEndTime
            }
        else:
            limitDeliveredTime = {}

        try:
            _thread.start_new_thread(self.OrderList, (shopName, int(mode), createStartTime, createEndTime, orderStatus, limitDeliveredTime))
        except:
            self.LogOut("Error: 无法计算启动线程")

    def LogOut(self, text):
        self.ui.output.append(text)


    def OrderList(self, shopName, mode, createStartTime, createEndTime, orderStatus, isPrintOwn, limitDeliveredTime = {}):
        if mode == 5:
            orderId = int(self.ui.orderId.toPlainText())
            self.GetSingleOrder(shopName, orderId, isPrintOwn)
        elif orderStatus == 0:
            self.GetOrderBill(createStartTime, createEndTime, 'waitsellersend,waitbuyerreceive', shopName, isPrintOwn, mode, limitDeliveredTime)
        elif orderStatus == 1:
            self.GetOrderBill(createStartTime, createEndTime, 'waitsellersend', shopName, isPrintOwn, mode, limitDeliveredTime)
        elif orderStatus == 2:
            self.GetOrderBill(createStartTime, createEndTime, 'waitbuyerreceive', shopName, isPrintOwn, mode, limitDeliveredTime)
        elif orderStatus == 3:
            self.GetOrderBill(createStartTime, createEndTime, 'waitbuyerpay', shopName, isPrintOwn, mode, limitDeliveredTime)

        self.LogOut("# 统计完成 \n")
        self.LogOut("###############################################################")

    def GetOrderBill(self, createStartTime, createEndTime, orderstatusStr, shopName, isPrintOwn, mode = 0, limitDeliveredTime = {}):
        orderList = []

        orderstatusList = orderstatusStr.split(',')

        for orderstatus in orderstatusList:
            data = {'createStartTime': createStartTime, 'createEndTime': createEndTime, 'orderStatus': orderstatus,
                    'needMemoInfo': 'true'}
            response = GetTradeData(data, shopName)
            if orderstatus == 'waitsellersend' :
                orderstatusStr = '待发货'
            if orderstatus == 'waitbuyerreceive' :
                orderstatusStr = '已发货'
            self.LogOut('# ' + orderstatusStr + ' : ' + str(response['totalRecord']) + '条记录')
            pageNum = CalPageNum(response['totalRecord'])

            # 规格化数据
            for pageId in range(pageNum):
                data = {'page': str(pageId + 1), 'createStartTime': createStartTime, 'createEndTime': createEndTime,
                        'orderStatus': orderstatus, 'needMemoInfo': 'true'}
                response = GetTradeData(data, shopName)

                if orderstatus == 'waitsellersend' or orderstatus == 'waitbuyerreceive':
                    for order in response['result']:
                        if ('sellerRemarkIcon' in order['baseInfo']) and ( order['baseInfo']['sellerRemarkIcon'] == '2' or order['baseInfo']['sellerRemarkIcon'] == '3'):
                            continue
                        elif mode != 0 and 'sellerRemarkIcon' not in order['baseInfo']:
                            if mode == 1:
                                order['baseInfo']['sellerRemarkIcon'] = '1'
                            elif mode == 4:
                                order['baseInfo']['sellerRemarkIcon'] = '4'

                orderList += response['result']

        self.GetBeihuoJson(orderList, isPrintOwn, mode, limitDeliveredTime)

    def GetSingleOrder(self, shopName, orderId, isPrintOwn):
        orderList = []
        data = {}
        data['orderId'] = orderId
        tmp = GetSingleTradeData(data, shopName)
        orderList.append(tmp['result'])
        self.GetBeihuoJson(orderList, isPrintOwn, 0)


    def GetBeihuoJson(self, orderList, isPrintOwn, mode=0, limitDeliveredTime = {}):
        BeihuoJson = {}
        for order in orderList:
            if ('sellerRemarkIcon' in order['baseInfo']) and (order['baseInfo']['sellerRemarkIcon'] == '2' or order['baseInfo']['sellerRemarkIcon'] == '3'):
                continue
            else:
                if mode != 0 and ('sellerRemarkIcon' not in order['baseInfo'] or order['baseInfo']['sellerRemarkIcon'] != str(mode)):
                    continue

                if 'allDeliveredTime' in order['baseInfo'] and len(limitDeliveredTime) > 0:  # 根据发货时间判断是否要输出
                    allDeliveredTime = int(order['baseInfo']['allDeliveredTime'][:-8])
                    if allDeliveredTime < limitDeliveredTime['deleveredStartTime'] or allDeliveredTime > limitDeliveredTime['deleveredEndTime']:
                        continue

                for productItem in order['productItems']:
                    # 货号
                    if 'cargoNumber' in productItem:
                        cargoNumberTag = 'cargoNumber'
                        cargoNumber = productItem['cargoNumber']
                    else:
                        cargoNumberTag = 'productCargoNumber'
                        cargoNumber = productItem['productCargoNumber']
                    if cargoNumber not in BeihuoJson:
                        BeihuoJson[cargoNumber] = {}

                    # 颜色
                    color = productItem['skuInfos'][0]['value']
                    if color not in BeihuoJson[productItem[cargoNumberTag]]:
                        BeihuoJson[productItem[cargoNumberTag]][color] = {
                            'products':{},
                            'productImgUrl':productItem['productImgUrl'][1],
                        }

                    # 身高
                    height = productItem['skuInfos'][1]['value']
                    if height not in BeihuoJson[cargoNumber][color]['products']:
                        BeihuoJson[cargoNumber][color]['products'][height] = {
                            # 数量
                            'quantity': productItem['quantity'],
                        }
                    else:
                        # 数量
                        BeihuoJson[cargoNumber][color]['products'][height]['quantity'] = BeihuoJson[cargoNumber][color]['products'][height]['quantity'] + productItem['quantity']

                    # 单价
                    BeihuoJson[cargoNumber][color]['products'][height]['price'] = GetCost(cargoNumber, height)

                    # 总价
                    BeihuoJson[cargoNumber][color]['products'][height]['cost'] = BeihuoJson[cargoNumber][color]['products'][height]['price']*BeihuoJson[cargoNumber][color]['products'][height]['quantity']
        self.GetTable(BeihuoJson, isPrintOwn)

    def GetTable(self, BeihuoJson, isPrintOwn):
        # 制表
        productsCountByShopName = {}
        BeihuoList = []
        BeihuoTable = []
        for item in BeihuoJson:
            BeihuoList.append(item)
            adressAndShopName = GetAdressAndShopName(item)
            BeihuoList.append(adressAndShopName[2])
            BeihuoList.append(adressAndShopName[1])
            BeihuoList.append(adressAndShopName[0])
            if adressAndShopName[2] not in productsCountByShopName:
                productsCountByShopName[adressAndShopName[2]] = [0,0]
            for color in BeihuoJson[item]:
                BeihuoList.append(color)
                BeihuoList.append(BeihuoJson[item][color]['productImgUrl'])
                heightTable = []
                heightList  = []
                for height in BeihuoJson[item][color]['products']:
                    heightList.append(height)
                    heightList.append(BeihuoJson[item][color]['products'][height]['quantity'])
                    productsCountByShopName[adressAndShopName[2]][0] += BeihuoJson[item][color]['products'][height]['quantity'] # 叠加总数
                    productsCountByShopName[adressAndShopName[2]][1] += (BeihuoJson[item][color]['products'][height]['quantity'] * BeihuoJson[item][color]['products'][height]['price'])
                    heightList.append(BeihuoJson[item][color]['products'][height]['price'])
                    heightTable.append(heightList.copy())
                    heightList.clear()
                heightTable.sort(key=lambda x:x[0])
                BeihuoList.append(heightTable)
                BeihuoTable.append(BeihuoList.copy())
                BeihuoList.pop()
                BeihuoList.pop()
                BeihuoList.pop()
            BeihuoList.pop()
            BeihuoList.pop()
            BeihuoList.pop()
            BeihuoList.pop()

        # 排序 拿货地规整
        BeihuoTable.sort(key=lambda x:[x[1],x[2]])

        # 写表
        savePath = self.ui.saveFilePath.toPlainText().split('.')[0]
        BH_wb = xlsxwriter.Workbook(savePath + self.calStartTime.strftime("%m_%d_%H_%M_%S") + ".xlsx")
        BH_sheet = BH_wb.add_worksheet('BH')
        BH_x = 0
        BH_y = 0

        shopNameTmp = ''

        sumCountX = 0
        for _list in BeihuoTable:
            if (not isPrintOwn) and (_list[1] == "朝新" or _list[2] == "本厂"):
                continue
            BH_sheet.write(BH_x,BH_y, _list[0]) # 货号
            BH_y += 1
            BH_sheet.write(BH_x, BH_y, _list[1])
            BH_y += 1
            # BH_sheet.write(BH_x, BH_y, _list[2])
            # BH_y += 1
            BH_sheet.write(BH_x, BH_y, _list[3])
            BH_y += 1
            BH_sheet.write(BH_x, BH_y, _list[4])
            BH_y += 1

            amount = ""
            for height in _list[6]:
                if amount != "":
                    amount += '\n'
                amount = amount + NumFormate4Print(height[0]) + ' ' + str(height[1]) + '件'

            BH_sheet.write(BH_x, BH_y, amount)
            BH_y += 1

            # 多尺码标识
            if len(amount.split('\n')) >= 2:
                BH_sheet.write(BH_x, BH_y + 3, '多尺码')

            # 插图
            imageName = _list[5].split('.jpg')[0].split('/')[-1]
            if ImageHandler.IsImageExist(imageName) :
                # 本地存有图片，读出
                imageData = ImageHandler.ReadImageFromDir(imageName)

                self.LogOut("读取本地图片")
            else:
                self.LogOut("下载图片")
                rt  = self.RequestPic(_list[5])


                if rt == 420:
                    # 本地存有图片，读出
                    imageData = ImageHandler.ReadImageFromDir(imageName)

                    self.LogOut("读取到手动下载图片")
                else:

                    imageDataRaw =rt.read()

                    imageData = io.BytesIO(imageDataRaw)

                    # BH_sheet.insert_image(BH_x, BH_y, _list[4],
                    #                       {'image_data': imageData, 'x_offset': 5, 'x_scale': 0.1, 'y_scale': 0.1})
                    # 保存图片
                    ImageHandler.SaveImage(imageData.getvalue(), imageName)

            BH_sheet.insert_image(BH_x, BH_y, _list[4],
                                  {'image_data': imageData, 'x_offset': 3, 'x_scale': 0.14, 'y_scale': 0.14})


            if _list[1] != shopNameTmp or _list == BeihuoTable[-1]:
                if _list == BeihuoTable[-1]:
                    sumCountX += 5
                if shopNameTmp != '':
                    self.LogOut(shopNameTmp + ' 拿货总件数 ： ' + str(productsCountByShopName[shopNameTmp][0]) + "  ||  总货款： " + str(round(productsCountByShopName[shopNameTmp][1],3)))
                    # BH_sheet.write(sumCountX, 1, shopNameTmp + ' 拿货总件数 ： ' + str(productsCountByShopName[shopNameTmp][0]) + "  ||  总货款： " + str(round(productsCountByShopName[shopNameTmp][1],3)))
                shopNameTmp = _list[1]

            sumCountX = BH_x + 1

            if len(_list[6]) >= 5:
                BH_x += 2
            else:
                theta = 5 - len(_list[6])
                BH_x = BH_x + theta + 2

            BH_y =  0

        # BH_sheet.write(sumCountX, 1, shopNameTmp + ' 拿货总件数 ： ' + str(productsCountByShopName[shopNameTmp]))
        BH_wb.close()

    def RequestPic(self, url):
        flag = True
        while flag:
            try:
                if self.errorUrl != url:
                    self.LogOut("<a href='" + url + "'>" + url + "</a>")
                picData = urllib.request.urlopen(url)
                flag = False
            except:
                self.LogOut("# 获取图片异常 重试")
                if self.errorUrl != url:
                    self.errorUrl = url
                QDesktopServices.openUrl(QUrl(self.errorUrl))
                time.sleep(3)

                imageName = url.split('.jpg')[0].split('/')[-1]
                if ImageHandler.IsImageExist(imageName):
                    return 420

            # picData = urllib.request.urlopen(url)

        return picData

if __name__ == '__main__':
    # 高分辨率适配
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication([])
    w = Window()
    w.ui.show()

    app.exec_()