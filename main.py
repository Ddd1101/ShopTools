import locale

from PySide2 import QtCore
from PySide2.QtWidgets import QApplication, QMainWindow, QFileDialog
from PySide2.QtCore import QDate, QDateTime, QTime, QUrl
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
import PySide2
import _thread

import logging

from xlsxwriter import workbook

import Common.ImageHandler as ImageHandler

# Views
from View.MenuSettingView import *

import platform

os.environ["QT_MAC_WANTS_LAYER"] = "1"

locale.setlocale(locale.LC_CTYPE, "Chinese")

dirname = os.path.dirname(PySide2.__file__)
plugin_path = os.path.join(dirname, "plugins", "platforms")
os.environ["QT_QPA_PLATFORM_PLUGIN_PATH"] = plugin_path

AppKey = {"联球制衣厂": "3527689", "朝雄制衣厂": "4834884", "万盈饰品厂": "2888236"}
AppSecret = {
    "联球制衣厂": b"Zw5KiCjSnL",
    "朝雄制衣厂": b"JeV4khKJshr",
    "万盈饰品厂": b"Zy7QvG0bQJI",
}
access_token = {
    "联球制衣厂": "999d182a-3576-4aee-97c5-8eeebce5e085",
    "朝雄制衣厂": "ef65f345-7060-4031-98ad-57d7d857f4d9",
    "万盈饰品厂": "f2f18480-0067-462f-9fac-952311ad4349",
}

# /ShopBackData/Common/GlobleDate - EnglishCode
en_code = ["s", "S", "m", "M", "l", "L", "x", "X"]
# /ShopBackData/Common/GlobleDate - AlbbBaseUrl
base_url = "https://gw.open.1688.com/openapi/"

path = os.path.realpath(os.curdir)

price_path = path + "/price.xlsx"
price_accessor_path = path + "/price_accessor.xlsx"

factory_path = path + "/factory.xlsx"

logging_path = path + "\\Log\\"

request_type = {
    "trade": "param2/1/com.alibaba.trade/",
    "delivery": "param2/1/com.alibaba.logistics/",
}

sumReporter = {}

worksheet = None

## shop type
SHOPTYPE_ALI_CHILD_CLOTH = 1
SHOPTYPE_ALI_ACCESSOR = 2

global_SHOPTYPE = SHOPTYPE_ALI_CHILD_CLOTH


# /Common/Utils/ExcelUtil - getSheet()
def GetPriceGrid():
    if global_SHOPTYPE == SHOPTYPE_ALI_CHILD_CLOTH:
        workbook = xlrd.open_workbook(price_path)  # 打开工作簿
    else:
        workbook = xlrd.open_workbook(price_accessor_path)
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(
        sheets[0]
    )  # 获取工作簿中所有表格中的的第一个表格
    return worksheet


def GetFactoryGrid():
    workbook = xlrd.open_workbook(factory_path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(
        sheets[0]
    )  # 获取工作簿中所有表格中的的第一个表格

    for t in range(1, worksheet.nrows):
        factoryName = worksheet.cell(t, 0).value
        sumReporter[factoryName] = {}
    return sumReporter


def SplitChineseAndPinyin(shopNameRaw):
    ret = -1

    for i in range(len(shopNameRaw)):
        if (shopNameRaw[i] >= "a" and shopNameRaw[i] <= "z") or (
            shopNameRaw[i] >= "A" and shopNameRaw[i] <= "Z"
        ):
            break
        ret = i
    return ret


# /ShopBackData/Common/Utils/Utils - NumFormate4Print
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
            if item >= "0" and item <= "9":
                res += item
            else:
                if len(res) and res[0] == "9":
                    res += " "
                break
        res += "cm"
    return res


# ShopBackData/Common/Utils/PriceUtils - CalSize
def CalSize(sizeDescription):
    _size = 0
    for i in range(len(sizeDescription)):
        if sizeDescription[i] >= "0" and sizeDescription[i] <= "9":
            _size = _size * 10 + int(sizeDescription[i])
        else:
            break
    return _size


# ShopBackData/Common/Utils/PriceUtils
def CalPriceLocationENCode(_size):
    if _size == "":
        return -1

    if _size[0] == "S" or _size[0] == "s":
        return 24

    if _size[0] == "M" or _size[0] == "m":
        return 25

    if _size[0] == "L" or _size[0] == "l":
        return 26

    if _size[0] == "x" or _size[0] == "X":
        if _size[1] == "L" or _size[1] == "l":
            return 27
        elif _size[1] == "x" or _size[1] == "X":
            if _size[2] == "L" or _size[2] == "l":
                return 28
            elif _size[2] == "x" or _size[2] == "X":
                return 29

    return -1


# ShopBackData/Common/Utils/PriceUtil/CalPriceLocation
def CalPriceLocation(_size):
    if _size[0] in en_code:
        return CalPriceLocationENCode(_size)
    # print(CalSize(_size))
    theta = (CalSize(_size) - 90) / 5
    _col = 5 + theta
    # print(_col)
    return _col


def CalPriceColByName(_size):
    if _size[0] in en_code:
        return CalPriceLocationENCode(_size)

    for col in range(worksheet.ncols):
        if worksheet.cell_value(0, col) == CalSize(_size):
            return col
    return -1


def GetCost(cargoNumber, skuInfosValue, colNum=0):
    global worksheet
    if worksheet == None:
        worksheet = GetPriceGrid()
    if global_SHOPTYPE == SHOPTYPE_ALI_CHILD_CLOTH:
        rowIndex = -1
        for t in range(1, worksheet.nrows):
            value = worksheet.cell(t, colNum).value
            if isinstance(value, int) or isinstance(value, float):
                value = int(value)
            # print(str(cargoNumber), str(value))
            if str(cargoNumber) == str(value):
                rowIndex = t
                break
        if rowIndex == -1:
            print(cargoNumber + " ： 未找到对应货号1")
            return 0
        colIndex = CalPriceColByName(skuInfosValue)
        if colIndex != None:
            _price = worksheet.cell(rowIndex, int(colIndex)).value
        else:
            _price = ""
        if _price == "":
            print(
                "货号："
                + cargoNumber
                + " 规格： "
                + skuInfosValue
                + " ： 未找到对应价格"
            )
            _price = 0
        return float(_price)
    elif global_SHOPTYPE == SHOPTYPE_ALI_ACCESSOR:
        rowIndex = -1
        for t in range(1, worksheet.nrows):
            if str(cargoNumber) == str(worksheet.cell(t, colNum).value):
                rowIndex = t
                break
        if rowIndex == -1:
            print(cargoNumber + " ： 未找到对应货号2")
            return 0
        colIndex = 2
        if colIndex != None:
            _price = worksheet.cell(rowIndex, int(colIndex)).value
        else:
            _price = ""
        return float(_price)


# 由货号得到产品名 - 厂家地址 - 厂家名
def GetAdressAndShopName(cargoNumber):
    rowIndex = -1
    for t in range(1, worksheet.nrows):
        value = worksheet.cell(t, 0).value
        if isinstance(value, int) or isinstance(value, float):
            value = int(value)
        print(str(cargoNumber), str(value))
        if str(cargoNumber) == str(value):
            rowIndex = t
            break
    if rowIndex == -1:
        print(str(cargoNumber) + " 没找到厂家")
        return ["", "", ""]
    productName = worksheet.cell(rowIndex, 1).value
    adress = worksheet.cell(rowIndex, 2).value
    shopName = worksheet.cell(rowIndex, 2).value
    return [productName, adress, shopName]


def CalPageNum(totalRecord):
    return math.ceil(totalRecord / 20)


def CalculateSignature(urlPath, data, shopName):
    # 构造签名因子：urlPath

    # 构造签名因子：拼装参数
    params = list()
    for key in data.keys():
        params.append(key + str(data[key]))
    params.sort()
    sortedParams = params
    assembedParams = str()
    for param in sortedParams:
        assembedParams = assembedParams + param

    # 合并签名因子
    mergedParams = urlPath + assembedParams
    mergedParams = bytes(mergedParams, "utf8")

    # 执行hmac_sha1算法 && 转为16进制
    hex_res1 = hmac.new(AppSecret[shopName], mergedParams, digestmod="sha1").hexdigest()

    # 转为全大写字符
    hex_res1 = hex_res1.upper()

    return hex_res1


# 交易数据获取  默认1页
def GetTradeData(data, shopName):
    data["access_token"] = access_token[shopName]
    _aop_signature = CalculateSignature(
        request_type["trade"] + "alibaba.trade.getSellerOrderList/" + AppKey[shopName],
        data,
        shopName,
    )
    data["_aop_signature"] = _aop_signature
    url = (
        base_url
        + request_type["trade"]
        + "alibaba.trade.getSellerOrderList/"
        + AppKey[shopName]
    )
    try:
        time.sleep(0.2)
        response = requests.post(url, data=data)
    except Exception as e:
        print("post error ", e)

    return response.json()


# 已发货物流信息
def GetDeliveryData(data, shopName):
    data["access_token"] = access_token[shopName]
    _aop_signature = CalculateSignature(
        request_type["delivery"]
        + "alibaba.trade.getLogisticsInfos.sellerView/"
        + AppKey[shopName],
        data,
        shopName,
    )
    data["_aop_signature"] = _aop_signature

    url = (
        base_url
        + request_type["delivery"]
        + "alibaba.trade.getLogisticsInfos.sellerView/"
        + AppKey[shopName]
    )

    response = requests.post(url, data=data)

    return response.json()


# 已发货物流信息+单号信息
def GetDeliveryTraceData(data, shopName):
    data["access_token"] = access_token[shopName]
    tmp = CalculateSignature(
        "param2/1/com.alibaba.logistics/alibaba.trade.getLogisticsInfos.sellerView/"
        + AppKey[shopName],
        data,
        shopName,
    )
    _aop_signature = CalculateSignature(
        request_type["delivery"]
        + "alibaba.trade.getLogisticsTraceInfo.sellerView/"
        + AppKey[shopName],
        data,
        shopName,
    )
    data["_aop_signature"] = _aop_signature
    url = (
        base_url
        + request_type["delivery"]
        + "alibaba.trade.getLogisticsTraceInfo.sellerView/"
        + AppKey[shopName]
    )

    response = requests.post(url, data=data)

    return response.json()


def GetSingleTradeData(data, shopName):
    data["access_token"] = access_token[shopName]
    _aop_signature = CalculateSignature(
        request_type["trade"] + "alibaba.trade.get.sellerView/" + AppKey[shopName],
        data,
        shopName,
    )
    data["_aop_signature"] = _aop_signature
    url = (
        base_url
        + request_type["trade"]
        + "alibaba.trade.get.sellerView/"
        + AppKey[shopName]
    )
    response = requests.post(url, data=data)

    return response.json()


# 获取订单号列表
# 筛除刷单
def GetOrderBill2(
    createStartTime,
    createEndTime,
    orderstatusStr,
    shopName,
    mode=0,
    limitDeliveredTime={},
):
    orderList = []

    orderstatusList = orderstatusStr.split(",")

    for orderstatus in orderstatusList:
        data = {
            "createStartTime": createStartTime,
            "createEndTime": createEndTime,
            "orderStatus": orderstatus,
            "needMemoInfo": "true",
        }
        response = GetTradeData(data, shopName)
        if orderstatus == "waitsellersend":
            orderstatusStr = "待发货"
        if orderstatus == "waitbuyerreceive":
            orderstatusStr = "已发货"
        # self.Logout('# ' + orderstatusStr + ' : ' + str(response['totalRecord']) + '条记录')
        pageNum = CalPageNum(response["totalRecord"])

        # 规格化数据
        for pageId in range(pageNum):
            data = {
                "page": str(pageId + 1),
                "createStartTime": createStartTime,
                "createEndTime": createEndTime,
                "orderStatus": orderstatus,
                "needMemoInfo": "true",
            }
            response = GetTradeData(data, shopName)

            if orderstatus == "waitsellersend" or orderstatus == "waitbuyerreceive":
                for order in response["result"]:
                    if ("sellerRemarkIcon" in order["baseInfo"]) and (
                        order["baseInfo"]["sellerRemarkIcon"] == "2"
                        or order["baseInfo"]["sellerRemarkIcon"] == "3"
                    ):
                        continue
                    elif mode != 0 and "sellerRemarkIcon" not in order["baseInfo"]:
                        if mode == 1:
                            order["baseInfo"]["sellerRemarkIcon"] = "1"
                        elif mode == 4:
                            order["baseInfo"]["sellerRemarkIcon"] = "4"

            orderList += response["result"]

    return orderList


# def self.Logout(text, type='info'):
#     if type == 'info':
#         logging.debug(text)
#     elif type == 'debug':
#         logging.debug(text)
# pass


class Window:
    def __init__(self):
        super(Window, self).__init__()

        # 从文件中加载UI定义
        qfile = QFile("Main.ui")
        qfile.open(QFile.ReadOnly)
        qfile.close()
        self.ui = QUiLoader().load(qfile)

        ## 菜单
        menuBar = self.ui.menuBar()
        menuSetting = menuBar.addMenu("设置")
        menuSettingKey = menuSetting.addAction("Key设置")

        menuSettingKey.triggered.connect(self.OpenMenuSetting)

        self.AliChildClothInit()

        self.AliAccessorInit()

    def AliChildClothInit(self):
        # 订单时间
        todayTmp = datetime.strptime(str(date.today()), "%Y-%m-%d")
        startDateTimeTmp = todayTmp + timedelta(days=-1)
        endDateTimeTmp = todayTmp + timedelta(days=1)
        self.ui.startTime.setDateTime(
            QDateTime(
                QDate(
                    startDateTimeTmp.year, startDateTimeTmp.month, startDateTimeTmp.day
                ),
                QTime(0, 0, 0),
            )
        )
        self.ui.endTime.setDateTime(
            QDateTime(
                QDate(endDateTimeTmp.year, endDateTimeTmp.month, endDateTimeTmp.day),
                QTime(0, 0, 0),
            )
        )

        # 发货时间
        deliverStartDateTimeTmp = todayTmp + timedelta(days=0)
        deliverEndDateTimeTmp = todayTmp + timedelta(days=1)
        self.ui.deleveredStartTime.setDateTime(
            QDateTime(
                QDate(
                    deliverStartDateTimeTmp.year,
                    deliverStartDateTimeTmp.month,
                    deliverStartDateTimeTmp.day,
                ),
                QTime(0, 0, 0),
            )
        )
        self.ui.deleveredEndTime.setDateTime(
            QDateTime(
                QDate(
                    deliverEndDateTimeTmp.year,
                    deliverEndDateTimeTmp.month,
                    deliverEndDateTimeTmp.day,
                ),
                QTime(0, 0, 0),
            )
        )

        self.ui.priceTablePath.setText("price.xlsx")
        self.ui.priceTablePathButton.clicked.connect(self.CheckPriceTablePath)

        # self.ui.shopName.addItem("万盈饰品厂")
        self.ui.shopName.addItem("联球制衣厂+朝雄制衣厂")
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
        self.ui.orderStatus.addItem("历史订单")

        self.ui.commit.clicked.connect(self.CalculateBeiHuoTable)
        # self.ui.openUrl.clicked.connect(self.click_window2)
        # self.ui.timeType.addItem("按发货时间")
        # self.ui.timeType.addItem("按付款时间")
        # self.ui.timeType.addItem("按下单时间")

        # MenuA = MenuBar.addMenu("MenuA")
        # MenuA1 = MenuA.addMenu("A1")  # 菜单嵌套子菜单
        # MenuA1.addAction("A1a")
        # MenuA1.addAction("A1b")
        # ActionA2 = MenuA.addAction("A2")  # 菜单添加Action
        # MenuB = MenuBar.addMenu("MenuB")
        # MenuB1 = MenuB.addMenu("B1")
        # MenuB1.addAction("B1a")
        # MenuB1.addAction("B1b")
        # MenuB.addAction("B2")

        self.errorUrl = ""

        self.calStartTime = datetime.now()

        self.ui.saveFilePath.setText("备货单")
        self.ui.saveFilePathButton.clicked.connect(self.CheckSaveFilePath)

        # 物流检查按钮
        self.ui.checkDelivery.clicked.connect(self.CheckDelivery)

        #
        self.shopId = self.ui.shopName.currentIndex() + 1
        self.shopName = self.ui.shopName.currentText()
        self.mode = self.ui.Tag.currentIndex()

        self.orderStatus = self.ui.orderStatus.currentIndex()

        self.startYear = self.ui.startTime.date().toString("yyyy")
        self.startMonth = self.ui.startTime.date().toString("MM")
        self.startDay = self.ui.startTime.date().toString("dd")

        self.createStartTime = (
            datetime(
                int(self.startYear), int(self.startMonth), int(self.startDay)
            ).strftime("%Y%m%d")
            + "000000000+0800"
        )

        self.endYear = self.ui.endTime.date().toString("yyyy")
        self.endMonth = self.ui.endTime.date().toString("MM")
        self.endDay = self.ui.endTime.date().toString("dd")
        self.deleveredStartTime = int(
            self.ui.deleveredStartTime.dateTime().toString("yyyyMMddHHmmss")
        )
        self.deleveredEndTime = int(
            self.ui.deleveredEndTime.dateTime().toString("yyyyMMddHHmmss")
        )

        self.createEndTime = (
            datetime(int(self.endYear), int(self.endMonth), int(self.endDay)).strftime(
                "%Y%m%d"
            )
            + "000000000+0800"
        )

        self.isPrintOwn = self.ui.IsPrintOwn.isChecked()

        self.isPrintUnitPrice = self.ui.IsPrintUnitPrice.isChecked()

        # 配置日志
        logPath = logging_path + datetime.now().strftime("%Y_%m_%d_%H_%M_%S") + ".log"
        logging.basicConfig(
            filename=logPath,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
        )
        self.Logout("end AliChildClothInit", "debug")

    def OpenMenuSetting(self):
        childWindow = MenuSettingView()
        childWindow.exec_()

    # 异常物流检测
    def CheckDelivery(self):
        self.CheckAllParams()

        try:
            _thread.start_new_thread(
                self.DoCheckDelivery,
                (
                    self.createStartTime,
                    self.createEndTime,
                    "waitbuyerreceive",
                    self.shopName,
                    self.mode,
                    self.limitDeliveredTime,
                ),
            )

            self.Logout2("\n # 启动计算 请稍后 \n")
        except:
            self.Logout2("Error: 无法计算启动线程")

    def DoCheckDelivery(
        self,
        createStartTime,
        createEndTime,
        orderstatusStr,
        shopName,
        mode=0,
        limitDeliveredTime={},
    ):
        self.Logout("start DoCheckDelivery", "debug")
        # 1. 获得待查询订单列表
        self.Logout2("1. 获得待查询订单列表")
        orderIdListRaw = GetOrderBill2(
            createStartTime,
            createEndTime,
            orderstatusStr,
            shopName,
            mode,
            limitDeliveredTime,
        )

        orderIdList = []

        deliveryErrorList = []

        for each in orderIdListRaw:
            if (
                "allDeliveredTime" in each["baseInfo"] and len(limitDeliveredTime) > 0
            ):  # 根据发货时间判断是否要输出
                allDeliveredTime = int(each["baseInfo"]["allDeliveredTime"][:-8])
                if (
                    allDeliveredTime < limitDeliveredTime["deleveredStartTime"]
                    or allDeliveredTime > limitDeliveredTime["deleveredEndTime"]
                ):
                    continue
            orderIdList.append(each["baseInfo"]["idOfStr"])

        # 2. 查询物流跟踪信息
        self.Logout2("2. 查询物流跟踪信息")
        for orderId in orderIdList:
            data = {"orderId": int(orderId), "webSite": "1688"}
            response = GetDeliveryTraceData(data, shopName)
            if "errorMessage" in response:
                deliveryErrorList.append([orderId])

        # 3.获取物流运单号
        self.Logout2("3.获取物流运单号")
        for each in deliveryErrorList:
            data = {"orderId": int(each[0]), "webSite": "1688"}
            response = GetDeliveryData(data, shopName)
            # each.append(response['result'][0]['logisticsBillNo'])
            self.Logout2("异常订单号：" + each[0])
            self.Logout2("异常运单号：" + response["result"][0]["logisticsBillNo"])
            self.Logout2("-----------------------------------")

        # 4. 打印
        # self.Logout('4. 打印')
        # for each in deliveryErrorList:
        #     self.Logout('异常订单号：' + each[0])
        #     self.Logout('异常运单号：' + each[1])
        #     self.Logout('===================================')

        self.Logout2("# 超时订单检测完成 ")
        self.Logout("end DoCheckDelivery", "debug")

        return deliveryErrorList

    def click_window2(self):
        QDesktopServices.openUrl(QUrl(self.errorUrl))

    def CheckPriceTablePath(self):
        filePath = QFileDialog.getOpenFileName(
            QMainWindow(), "选择文件"
        )  # 选择目录，返回选中的路径
        self.ui.priceTablePath.setText(filePath[0])

    def CheckSaveFilePath(self):
        FileDirectory = QFileDialog.getExistingDirectory(
            QMainWindow(), "选择文件夹"
        )  # 选择目录，返回选中的路径
        self.ui.saveFilePath.setText(FileDirectory)

    def CheckAllParams(self):
        self.Logout("start CheckAllParams", "debug")
        self.Logout(
            "# 启动计算时间 : " + self.calStartTime.strftime("%Y-%m-%d %H:%M:%S")
        )
        self.shopId = self.ui.shopName.currentIndex() + 1
        self.shopName = self.ui.shopName.currentText()
        self.mode = self.ui.Tag.currentIndex()

        self.orderStatus = self.ui.orderStatus.currentIndex()

        self.startYear = self.ui.startTime.date().toString("yyyy")
        self.startMonth = self.ui.startTime.date().toString("MM")
        self.startDay = self.ui.startTime.date().toString("dd")

        self.createStartTime = (
            datetime(
                int(self.startYear), int(self.startMonth), int(self.startDay)
            ).strftime("%Y%m%d")
            + "000000000+0800"
        )

        self.endYear = self.ui.endTime.date().toString("yyyy")
        self.endMonth = self.ui.endTime.date().toString("MM")
        self.endDay = self.ui.endTime.date().toString("dd")
        self.deleveredStartTime = int(
            self.ui.deleveredStartTime.dateTime().toString("yyyyMMddHHmmss")
        )
        self.deleveredEndTime = int(
            self.ui.deleveredEndTime.dateTime().toString("yyyyMMddHHmmss")
        )

        self.createEndTime = (
            datetime(int(self.endYear), int(self.endMonth), int(self.endDay)).strftime(
                "%Y%m%d"
            )
            + "000000000+0800"
        )

        self.isPrintOwn = self.ui.IsPrintOwn.isChecked()
        self.isPrintUnitPrice = self.ui.IsPrintUnitPrice.isChecked()

        self.Logout2("# 店铺名 ：" + self.shopName)
        self.Logout2("# 色标 ：" + self.ui.Tag.currentText())
        self.Logout2(
            "# 订单开始时间 ：" + self.ui.startTime.date().toString("yyyy-MM-dd")
        )
        self.Logout2(
            "# 订单截止时间 ：" + self.ui.endTime.date().toString("yyyy-MM-dd")
        )

        self.isLimitDeleveredTime = self.ui.isLimitDeleveredTime.isChecked()

        if self.isLimitDeleveredTime:
            self.limitDeliveredTime = {
                "deleveredStartTime": self.deleveredStartTime,
                "deleveredEndTime": self.deleveredEndTime,
            }
        else:
            self.limitDeliveredTime = {}

        self.Logout("end CheckAllParams", "debug")

    def CalculateBeiHuoTable(self):
        self.CheckAllParams()
        try:
            _thread.start_new_thread(
                self.OrderList,
                (
                    self.shopName,
                    int(self.mode),
                    self.createStartTime,
                    self.createEndTime,
                    self.orderStatus,
                    self.isPrintOwn,
                    self.limitDeliveredTime,
                    self.isPrintUnitPrice,
                    SHOPTYPE_ALI_CHILD_CLOTH,
                ),
            )
            self.Logout("\n # 启动计算 请稍后 \n")

        except:
            self.Logout("Error: 无法计算启动线程")

    def Logout(self, text, type="info"):
        # if type == 'info':
        #     # logging.debug(text)
        #     pass
        # elif type == 'debug':
        #     logging.debug(text)

        logging.info(text)
        # print(text)
        # self.ui.costomLogging.append(text)

    def Logout2(self, text, type="info"):
        # if type == 'info':
        #     # logging.debug(text)
        #     pass
        # elif type == 'debug':
        #     logging.debug(text)

        # logging.debug(text)
        # print(text)
        self.ui.costomLogging.append(text)
        # pass

    def OrderList(
        self,
        shopName,
        mode,
        createStartTime,
        createEndTime,
        orderStatus,
        isPrintOwn,
        limitDeliveredTime,
        isPrintUnitPrice,
        shopType=SHOPTYPE_ALI_CHILD_CLOTH,
    ):
        self.Logout2(
            "start OrderList shopname" + shopName + " mode:" + str(mode), "debug"
        )
        global global_SHOPTYPE
        global_SHOPTYPE = shopType
        worksheet = GetPriceGrid()
        if mode == 5:
            orderId = int(self.ui.orderId.toPlainText())
            self.order = self.GetSingleOrder(
                shopName, orderId, isPrintOwn, isPrintUnitPrice
            )
        elif orderStatus == 0:
            self.GetOrderBill(
                createStartTime,
                createEndTime,
                "waitbuyerreceive",
                shopName,
                isPrintOwn,
                mode,
                limitDeliveredTime,
                isPrintUnitPrice,
                shopType,
            )
        elif orderStatus == 1:
            self.GetOrderBill(
                createStartTime,
                createEndTime,
                "waitsellersend,waitbuyerreceive",
                shopName,
                isPrintOwn,
                mode,
                limitDeliveredTime,
                isPrintUnitPrice,
                shopType,
            )
        elif orderStatus == 2:
            self.GetOrderBill(
                createStartTime,
                createEndTime,
                "waitsellersend",
                shopName,
                isPrintOwn,
                mode,
                limitDeliveredTime,
                isPrintUnitPrice,
                shopType,
            )
        elif orderStatus == 3:
            self.GetOrderBill(
                createStartTime,
                createEndTime,
                "waitbuyerpay",
                shopName,
                isPrintOwn,
                mode,
                limitDeliveredTime,
                isPrintUnitPrice,
                shopType,
            )
        # elif orderStatus == 4:
        #     self.GetOrderHistory(createStartTime, createEndTime,
        #                          'waitbuyerreceive,waitbuyersign,signinsuccess,confirm_goods,success', shopName,
        #                          isPrintOwn, mode,
        #                          limitDeliveredTime, isPrintUnitPrice, shopType)

        elif orderStatus == 4:
            self.GetOrderHistory(
                createStartTime,
                createEndTime,
                "waitbuyerreceive,success",
                shopName,
                isPrintOwn,
                mode,
                limitDeliveredTime,
                isPrintUnitPrice,
                shopType,
            )

        # elif orderStatus == 4:
        #     self.GetOrderHistory(createStartTime, createEndTime,
        #                          '未枚举', shopName,
        #                          isPrintOwn, mode,
        #                          limitDeliveredTime, isPrintUnitPrice, shopType)

        # self.ui.output.append("# 统计完成 \n")
        self.Logout2("# 统计完成 \n")
        self.Logout2("#################################################")

    def GetOrderBill(
        self,
        createStartTime,
        createEndTime,
        orderstatusStr,
        shopNameStr,
        isPrintOwn,
        mode=0,
        limitDeliveredTime={},
        isPrintUnitPrice=False,
        shopType=SHOPTYPE_ALI_CHILD_CLOTH,
    ):
        self.Logout("start GetOrderBill", "debug")

        global_SHOPTYPE = shopType

        shopNameList = shopNameStr.split("+")

        orderListRaw = []

        orderstatusList = orderstatusStr.split(",")

        orderList = []

        for shopName in shopNameList:
            orderListRaw.clear()
            for orderstatus in orderstatusList:
                data = {
                    "createStartTime": createStartTime.strip(),
                    "createEndTime": createEndTime.strip(),
                    "orderStatus": orderstatus.strip(),
                    "needMemoInfo": "true",
                }

                response = GetTradeData(data, shopName)
                if orderstatus == "waitsellersend":
                    orderstatusStr = "待发货"
                if orderstatus == "waitbuyerreceive":
                    orderstatusStr = "已发货"

                pageNum = CalPageNum(response["totalRecord"])

                # 规格化数据
                for pageId in range(pageNum):
                    data = {
                        "page": str(pageId + 1),
                        "createStartTime": createStartTime,
                        "createEndTime": createEndTime,
                        "orderStatus": orderstatus,
                        "needMemoInfo": "true",
                    }
                    response = GetTradeData(data, shopName)

                    if (
                        orderstatus == "waitsellersend"
                        or orderstatus == "waitbuyerreceive"
                    ):
                        for order in response["result"]:
                            if ("sellerRemarkIcon" in order["baseInfo"]) and (
                                order["baseInfo"]["sellerRemarkIcon"] == "2"
                                or order["baseInfo"]["sellerRemarkIcon"] == "3"
                            ):
                                continue
                            elif (
                                mode != 0
                                and "sellerRemarkIcon" not in order["baseInfo"]
                            ):
                                if mode == 1:
                                    order["baseInfo"]["sellerRemarkIcon"] = "1"
                                elif mode == 4:
                                    order["baseInfo"]["sellerRemarkIcon"] = "4"

                    orderListRaw += response["result"]

                if len(limitDeliveredTime) >= 2:
                    for order in orderListRaw:
                        if (
                            "allDeliveredTime" in order["baseInfo"]
                            and len(limitDeliveredTime) > 0
                        ):  # 根据发货时间判断是否要输出
                            allDeliveredTime = int(
                                order["baseInfo"]["allDeliveredTime"][:-8]
                            )
                            if (
                                allDeliveredTime
                                < limitDeliveredTime["deleveredStartTime"]
                                or allDeliveredTime
                                > limitDeliveredTime["deleveredEndTime"]
                            ):
                                continue
                            else:
                                orderList.append(order)
                else:
                    orderList += orderListRaw

                orderListRaw.clear()

        self.Logout2("# " + orderstatusStr + " : " + str(len(orderList)) + "条记录")

        if shopType == SHOPTYPE_ALI_CHILD_CLOTH:
            self.GetBeihuoJson(
                orderList, isPrintOwn, mode, limitDeliveredTime, isPrintUnitPrice
            )
        elif shopType == SHOPTYPE_ALI_ACCESSOR:
            self.AliAccessorGetBeihuoJson(
                orderList, isPrintOwn, mode, limitDeliveredTime, isPrintUnitPrice
            )
        self.Logout("end GetOrderBill", "debug")

    def GetOrderHistory(
        self,
        createStartTime,
        createEndTime,
        statusStr,
        shopNameStr,
        isPrintOwn,
        mode=0,
        limitDeliveredTime={},
        isPrintUnitPrice=False,
        shopType=SHOPTYPE_ALI_CHILD_CLOTH,
    ):
        self.Logout("start GetOrderBill", "debug")

        shopNameList = shopNameStr.split("+")

        orderListRaw = []

        orderstatusList = statusStr.split(",")

        orderList = []

        statusList = statusStr.split(",")

        for shopName in shopNameList:
            orderListRaw.clear()
            for orderstatus in orderstatusList:
                data = {
                    "createStartTime": createStartTime.strip(),
                    "createEndTime": createEndTime.strip(),
                    "needMemoInfo": "true",
                }

                response = GetTradeData(data, shopName)

                pageNum = CalPageNum(response["totalRecord"])

                # 规格化数据
                for pageId in range(pageNum):
                    print(orderstatus, pageNum, pageId)
                    data = {
                        "page": str(pageId + 1),
                        "createStartTime": createStartTime,
                        "createEndTime": createEndTime,
                        "orderStatus": orderstatus,
                        "needMemoInfo": "true",
                    }
                    response = GetTradeData(data, shopName)

                    print("response 2")

                    while "result" not in response:
                        time.sleep(2)
                        print("while 'result' not in response:")
                        response = GetTradeData(data, shopName)
                        data["page"] = str(pageId - 1)
                        response = GetTradeData(data, shopName)
                        data["page"] = str(pageId + 1)
                        response = GetTradeData(data, shopName)

                    for order in response["result"]:
                        if ("sellerRemarkIcon" in order["baseInfo"]) and (
                            order["baseInfo"]["sellerRemarkIcon"] == "2"
                            or order["baseInfo"]["sellerRemarkIcon"] == "3"
                        ):
                            continue
                        elif order["baseInfo"]["status"] not in statusList:
                            continue
                        elif mode != 0 and "sellerRemarkIcon" not in order["baseInfo"]:
                            if mode == 1:
                                order["baseInfo"]["sellerRemarkIcon"] = "1"
                            elif mode == 4:
                                order["baseInfo"]["sellerRemarkIcon"] = "4"

                        orderListRaw.append(order)

                if len(limitDeliveredTime) >= 2:
                    for order in orderListRaw:
                        if (
                            "allDeliveredTime" in order["baseInfo"]
                            and len(limitDeliveredTime) > 0
                        ):  # 根据发货时间判断是否要输出
                            allDeliveredTime = int(
                                order["baseInfo"]["allDeliveredTime"][:-8]
                            )
                            if (
                                allDeliveredTime
                                < limitDeliveredTime["deleveredStartTime"]
                                or allDeliveredTime
                                > limitDeliveredTime["deleveredEndTime"]
                            ):
                                continue
                            else:
                                orderList.append(order)
                else:
                    orderList += orderListRaw

                orderListRaw.clear()

        self.Logout2("# " + statusStr + " : " + str(len(orderList)) + "条记录")

        if shopType == SHOPTYPE_ALI_CHILD_CLOTH:
            self.GetBeihuoJson(
                orderList, isPrintOwn, mode, limitDeliveredTime, isPrintUnitPrice
            )
        elif shopType == SHOPTYPE_ALI_ACCESSOR:
            self.AliAccessorGetBeihuoJson(
                orderList, isPrintOwn, mode, limitDeliveredTime, isPrintUnitPrice
            )
        self.Logout("end GetOrderBill", "debug")

    def GetOrderBillBac(
        self,
        createStartTime,
        createEndTime,
        orderstatusStr,
        shopName,
        isPrintOwn,
        mode=0,
        limitDeliveredTime={},
    ):
        orderListRaw = []

        orderstatusList = orderstatusStr.split(",")

        orderList = []

        for orderstatus in orderstatusList:
            data = {
                "createStartTime": createStartTime,
                "createEndTime": createEndTime,
                "orderStatus": orderstatus,
                "needMemoInfo": "true",
            }
            response = GetTradeData(data, shopName)
            if orderstatus == "waitsellersend":
                orderstatusStr = "待发货"
            if orderstatus == "waitbuyerreceive":
                orderstatusStr = "已发货"

            pageNum = CalPageNum(response["totalRecord"])

            # 规格化数据
            for pageId in range(pageNum):
                data = {
                    "page": str(pageId + 1),
                    "createStartTime": createStartTime,
                    "createEndTime": createEndTime,
                    "orderStatus": orderstatus,
                    "needMemoInfo": "true",
                }
                response = GetTradeData(data, shopName)

                if orderstatus == "waitsellersend" or orderstatus == "waitbuyerreceive":
                    for order in response["result"]:
                        if ("sellerRemarkIcon" in order["baseInfo"]) and (
                            order["baseInfo"]["sellerRemarkIcon"] == "2"
                            or order["baseInfo"]["sellerRemarkIcon"] == "3"
                        ):
                            continue
                        elif mode != 0 and "sellerRemarkIcon" not in order["baseInfo"]:
                            if mode == 1:
                                order["baseInfo"]["sellerRemarkIcon"] = "1"
                            elif mode == 4:
                                order["baseInfo"]["sellerRemarkIcon"] = "4"

                orderListRaw += response["result"]

            if len(limitDeliveredTime) >= 2:
                for order in orderListRaw:
                    if (
                        "allDeliveredTime" in order["baseInfo"]
                        and len(limitDeliveredTime) > 0
                    ):  # 根据发货时间判断是否要输出
                        allDeliveredTime = int(
                            order["baseInfo"]["allDeliveredTime"][:-8]
                        )
                        if (
                            allDeliveredTime < limitDeliveredTime["deleveredStartTime"]
                            or allDeliveredTime > limitDeliveredTime["deleveredEndTime"]
                        ):
                            continue
                        else:
                            orderList.append(order)
            else:
                orderList = orderListRaw

        self.Logout2("# " + orderstatusStr + " : " + str(len(orderList)) + "条记录")

        self.GetBeihuoJson(orderList, isPrintOwn, mode, limitDeliveredTime)

    def GetSingleOrder(self, shopName, orderId, isPrintOwn, isPrintUnitPrice=False):
        orderList = []
        data = {}
        data["orderId"] = orderId
        tmp = GetSingleTradeData(data, shopName)
        orderList.append(tmp["result"])
        self.GetBeihuoJson(orderList, isPrintOwn, 0)

    def GetBeihuoJson(
        self,
        orders,
        is_print_own,
        mode=0,
        limit_delivered_time={},
        isPrintUnitPrice=False,
    ):
        self.Logout("start GetBeihuoJson", "debug")
        beihuo_json = {}

        for order in orders:
            if mode != 0 and (
                "sellerRemarkIcon" not in order["baseInfo"]
                or order["baseInfo"]["sellerRemarkIcon"] != str(mode)
            ):
                continue
            if "sellerRemarkIcon" in order["baseInfo"] and (
                order["baseInfo"]["sellerRemarkIcon"] == "2"
                or order["baseInfo"]["sellerRemarkIcon"] == "3"
            ):
                continue

            for product_item in order["productItems"]:
                cargo_number_tag = (
                    "cargoNumber"
                    if "cargoNumber" in product_item
                    else "productCargoNumber"
                )
                cargo_number = product_item[cargo_number_tag]
                color = product_item["skuInfos"][0]["value"]
                height = product_item["skuInfos"][1]["value"]

                cargo_number = cargo_number.split("【")[-1]

                product_dict = beihuo_json.setdefault(cargo_number, {}).setdefault(
                    color,
                    {"products": {}, "productImgUrl": product_item["productImgUrl"][1]},
                )
                product_dict["products"].setdefault(height, {"quantity": 0})
                product_dict["products"][height]["quantity"] += product_item["quantity"]
                product_dict["products"][height]["price"] = GetCost(
                    cargo_number, height
                )
                product_dict["products"][height]["cost"] = (
                    product_dict["products"][height]["price"]
                    * product_dict["products"][height]["quantity"]
                )

        self.GetTable(beihuo_json, is_print_own, isPrintUnitPrice)
        self.Logout("end GetBeihuoJson", "debug")

    def GetTable(self, BeihuoJson, isPrintOwn, isPrintUnitPrice):
        self.Logout("start GetTable", "debug")
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
                productsCountByShopName[adressAndShopName[2]] = [0, 0]
            for color in BeihuoJson[item]:
                BeihuoList.append(color)
                BeihuoList.append(BeihuoJson[item][color]["productImgUrl"])
                heightTable = []
                heightList = []
                for height in BeihuoJson[item][color]["products"]:
                    heightList.append(height)
                    heightList.append(
                        BeihuoJson[item][color]["products"][height]["quantity"]
                    )
                    productsCountByShopName[adressAndShopName[2]][0] += BeihuoJson[
                        item
                    ][color]["products"][height][
                        "quantity"
                    ]  # 叠加总数
                    productsCountByShopName[adressAndShopName[2]][1] += (
                        BeihuoJson[item][color]["products"][height]["quantity"]
                        * BeihuoJson[item][color]["products"][height]["price"]
                    )
                    heightList.append(
                        BeihuoJson[item][color]["products"][height]["price"]
                    )
                    heightTable.append(heightList.copy())
                    heightList.clear()
                heightTable.sort(key=lambda x: x[0])
                BeihuoList.append(heightTable)
                BeihuoTable.append(BeihuoList.copy())
                BeihuoList.pop()
                BeihuoList.pop()
                BeihuoList.pop()
            BeihuoList.pop()
            BeihuoList.pop()
            BeihuoList.pop()
            BeihuoList.pop()

        self.Logout("制表结束", "info")

        # 排序 拿货地规整
        BeihuoTable.sort(key=lambda x: [x[1], x[2]])

        # 写表
        savePath = self.ui.saveFilePath.toPlainText().split(".")[0]

        # 保存备货单
        BH_wb = xlsxwriter.Workbook(
            savePath + "/" + datetime.now().strftime("%m月%d日%H时%M分%S秒") + ".xlsx"
        )
        BH_sheet = BH_wb.add_worksheet("BH")
        BH_pay_sheet = BH_wb.add_worksheet("pay")
        BH_x = 0
        BH_y = 0

        shopNameTmp = ""

        sumCountX = 0

        piecesCount = 0  # 统计总价格
        sumReporter = GetFactoryGrid()
        for _list in BeihuoTable:
            if (not isPrintOwn) and (_list[1] == "朝新" or _list[2] == "朝新"):
                continue
            BH_sheet.write(BH_x, BH_y, _list[0])  # 货号
            BH_y += 1
            BH_sheet.write(BH_x, BH_y, _list[1])
            BH_y += 1
            # BH_sheet.write(BH_x, BH_y, _list[2])
            # BH_y += 1
            BH_sheet.write(BH_x, BH_y, _list[3])
            BH_y += 1
            BH_sheet.write(BH_x, BH_y, _list[4])
            BH_y += 1

            # 多尺码序列化
            amount = ""
            for height in _list[6]:
                if amount != "":
                    amount += "\n"
                if isPrintUnitPrice:
                    amount = (
                        amount
                        + NumFormate4Print(height[0])
                        + " "
                        + str(height[1])
                        + "件"
                        + " "
                        + str(height[2])
                    )
                else:
                    amount = (
                        amount
                        + NumFormate4Print(height[0])
                        + " "
                        + str(height[1])
                        + "件"
                    )

            BH_sheet.write(BH_x, BH_y, amount)
            BH_y += 1

            # 多尺码标识
            if len(amount.split("\n")) >= 2:
                BH_sheet.write(BH_x, BH_y + 3, "多尺码")

            # 插图
            imageName = _list[5].split(".jpg")[0].split("/")[-1]

            if ImageHandler.IsImageExist(imageName):
                time.sleep(0.2)
                # 本地存有图片，读出
                self.Logout("before ReadImageFromDir" + imageName, "debug")
                imageData = ImageHandler.ReadImageFromDir(imageName)
                self.Logout("after ReadImageFromDir" + imageName, "debug")

                self.Logout("读取本地图片")

            else:
                self.Logout("下载图片")
                # time.sleep(0.5)
                rt = self.RequestPic(_list[5])

                if rt == 420:
                    # 本地存有图片，读出
                    self.Logout("before ReadImageFromDir" + imageName, "debug")
                    imageData = ImageHandler.ReadImageFromDir(imageName)
                    self.Logout("after ReadImageFromDir" + imageName, "debug")

                    self.Logout("读取到手动下载图片")
                elif rt == 400:
                    imageData = ImageHandler.ReadImageFromDir(imageName)
                else:
                    self.Logout("before 1 ReadImageFromD net " + imageName, "debug")
                    imageDataRaw = rt.read()
                    self.Logout("2  ReadImageFromD net " + imageName, "debug")

                    imageData = io.BytesIO(imageDataRaw)
                    self.Logout("after ReadImageFromD net " + imageName, "debug")

                    # BH_sheet.insert_image(BH_x, BH_y, _list[4],
                    #                       {'image_data': imageData, 'x_offset': 5, 'x_scale': 0.1, 'y_scale': 0.1})
                    # 保存图片
                    ImageHandler.SaveImage(imageData.getvalue(), imageName)
                    self.Logout("save ReadImageFromD net " + imageName)

            self.Logout("before insert_image", "debug")
            BH_sheet.insert_image(
                BH_x,
                BH_y,
                _list[4],
                {
                    "image_data": imageData,
                    "x_offset": 3,
                    "x_scale": 0.14,
                    "y_scale": 0.14,
                },
            )
            self.Logout("after insert_image", "debug")

            # 只能处理到倒数第二个厂商
            if _list[1] != shopNameTmp:
                if shopNameTmp != "":
                    piecesCount += productsCountByShopName[shopNameTmp][1]
                    self.Logout2(
                        shopNameTmp
                        + " 总件数："
                        + str(productsCountByShopName[shopNameTmp][0])
                        + " | 货款："
                        + str(round(productsCountByShopName[shopNameTmp][1], 3))
                    )
                    # 输出字体
                    priceStyle = BH_wb.add_format(
                        {
                            # "fg_color": "yellow",  # 单元格的背景颜色
                            "bold": 1,  # 字体加粗
                            "align": "left",  # 对齐方式
                            "valign": "vcenter",  # 字体对齐方式
                            "font_color": "red",  # 字体颜色
                        }
                    )
                    writeStr = (
                        shopNameTmp
                        + " 总件数："
                        + str(productsCountByShopName[shopNameTmp][0])
                    )
                    BH_sheet.merge_range(
                        "A" + str(sumCountX + 1) + ":D" + str(sumCountX + 1),
                        writeStr,
                        priceStyle,
                    )

                    # 商家 & 货款  信息收集
                    shopNameSplited = shopNameTmp[
                        0 : SplitChineseAndPinyin(shopNameTmp) + 1
                    ]

                    if shopNameSplited not in sumReporter:
                        sumReporter[shopNameSplited] = {}
                    sumReporter[shopNameSplited]["num"] = productsCountByShopName[
                        shopNameTmp
                    ][0]
                    sumReporter[shopNameSplited]["payment"] = round(
                        productsCountByShopName[shopNameTmp][1], 3
                    )

                shopNameTmp = _list[1]

            # 最后一个单独处理
            if _list == BeihuoTable[-1]:
                sumCountX += 6
                if shopNameTmp != "":
                    piecesCount += productsCountByShopName[shopNameTmp][1]
                    self.Logout2(
                        shopNameTmp
                        + " 总件数："
                        + str(productsCountByShopName[shopNameTmp][0])
                        + " | 货款："
                        + str(round(productsCountByShopName[shopNameTmp][1], 3))
                    )
                    # 输出字体
                    priceStyle = BH_wb.add_format(
                        {
                            # "fg_color": "yellow",  # 单元格的背景颜色
                            "bold": 1,  # 字体加粗
                            "align": "left",  # 对齐方式
                            "valign": "vcenter",  # 字体对齐方式
                            "font_color": "red",  # 字体颜色
                        }
                    )
                    writeStr = (
                        shopNameTmp
                        + " 总件数："
                        + str(productsCountByShopName[shopNameTmp][0])
                    )
                    BH_sheet.merge_range(
                        "A" + str(sumCountX + 1) + ":D" + str(sumCountX + 1),
                        writeStr,
                        priceStyle,
                    )

                    # 商家 & 货款  信息收集
                    shopNameSplited = shopNameTmp[
                        0 : SplitChineseAndPinyin(shopNameTmp) + 1
                    ]

                    if shopNameSplited not in sumReporter:
                        sumReporter[shopNameSplited] = {}
                    sumReporter[shopNameSplited]["num"] = productsCountByShopName[
                        shopNameTmp
                    ][0]
                    sumReporter[shopNameSplited]["payment"] = round(
                        productsCountByShopName[shopNameTmp][1], 3
                    )

                shopNameTmp = _list[1]

            sumCountX = BH_x + 1

            if len(_list[6]) >= 5:
                BH_x += 2
            else:
                theta = 5 - len(_list[6])
                BH_x = BH_x + theta + 2

            BH_y = 0

        sumCountX += 6

        # 输出统计信息
        self.PrintSumReporter(BH_wb, BH_pay_sheet, sumReporter, piecesCount)

        try:
            BH_wb.close()
        except Exception as e:
            print("程序出现异常:", e)

        self.Logout("end GetTable", "debug")

    def PrintSumReporter(self, BH_wb, BH_pay_sheet, sumReporter, piecesCount):
        header_style = BH_wb.add_format(
            {
                # "fg_color": "yellow",  # 单元格的背景颜色
                "bold": 1,  # 字体加粗
                "align": "right",  # 对齐方式
                "valign": "vcenter",  # 字体对齐方式
                "font_color": "black",  # 字体颜色
            }
        )
        BH_pay_sheet.write(0, 0, "订单发货日期", header_style)
        BH_pay_sheet.write(0, 1, "厂名", header_style)
        BH_pay_sheet.write(0, 2, "发货件数", header_style)
        BH_pay_sheet.write(0, 3, "总货款", header_style)

        priceStyle = BH_wb.add_format(
            {
                # "fg_color": "yellow",  # 单元格的背景颜色
                "align": "right",  # 对齐方式
                "valign": "vcenter",  # 字体对齐方式
                "font_color": "black",  # 字体颜色
            }
        )

        sumCountX = 1
        for shopName in sumReporter:
            if "num" in sumReporter[shopName]:
                BH_pay_sheet.write(sumCountX, 0, self.deleveredStartTimeStr, priceStyle)
                BH_pay_sheet.write(sumCountX, 1, shopName, priceStyle)
                BH_pay_sheet.write(
                    sumCountX, 2, str(sumReporter[shopName]["num"]), priceStyle
                )
                BH_pay_sheet.write(
                    sumCountX,
                    3,
                    str(round(sumReporter[shopName]["payment"], 3)),
                    priceStyle,
                )
                sumCountX += 1

        piecesCountStyle = BH_wb.add_format(
            {
                # "fg_color": "yellow",  # 单元格的背景颜色
                "bold": 1,
                "align": "right",  # 对齐方式
                "valign": "vcenter",  # 字体对齐方式
                "font_color": "red",  # 字体颜色
            }
        )

        writeStr = "总货款 ： " + str(round(piecesCount, 3))
        BH_pay_sheet.merge_range(
            "A" + str(sumCountX + 2) + ":C" + str(sumCountX + 2),
            writeStr,
            piecesCountStyle,
        )

    def RequestPic(self, url):
        self.Logout("start RequestPic " + url, "debug")
        flag = True
        while flag:
            # 图片下载尝试 2
            time.sleep(0.3)
            print("# 图片下载方法1")
            self.Logout2("# 图片下载方法1")
            try:
                response = requests.get(url)
                if response.status_code == 200:
                    imageName = url.split(".jpg")[0].split("/")[-1]
                    with open("images/" + imageName + ".jpg", "wb") as f:
                        f.write(response.content)
                    self.Logout2("# 图片下载方法1 成功")

                    return 400
                else:
                    print("无法下载图片1:", response.status_code)
            except:
                print("无法下载图片1 - 1")

            # 图片下载尝试 1
            try:
                self.Logout2("# 图片下载方法2")
                if self.errorUrl != url:
                    self.Logout("<a href='" + url + "'>" + url + "</a>")
                picData = urllib.request.urlopen(url)
                return picData
            except:
                self.Logout2("# 获取图片异常2 重试")
                time.sleep(0.1)

                imageName = url.split(".jpg")[0].split("/")[-1]
                self.Logout("images/" + imageName + ".jpg")
                if ImageHandler.IsImageExist(imageName):
                    return 420

            if self.errorUrl != url:
                self.errorUrl = url
                QDesktopServices.openUrl(QUrl(self.errorUrl))

            # picData = urllib.request.urlopen(url)

        self.Logout("end RequestPic ", "debug")
        return picData

    # 饰品 AliAccessor
    def AliAccessorInit(self):
        # 订单时间
        todayTmp = datetime.strptime(str(date.today()), "%Y-%m-%d")
        startDateTimeTmp = todayTmp + timedelta(days=-2)
        endDateTimeTmp = todayTmp + timedelta(days=1)
        self.ui.AliAccessorStartTime.setDateTime(
            QDateTime(
                QDate(
                    startDateTimeTmp.year, startDateTimeTmp.month, startDateTimeTmp.day
                ),
                QTime(0, 0, 0),
            )
        )
        self.ui.AliAccessorEndTime.setDateTime(
            QDateTime(
                QDate(endDateTimeTmp.year, endDateTimeTmp.month, endDateTimeTmp.day),
                QTime(0, 0, 0),
            )
        )

        # 发货时间
        self.ui.AliAccessorDeleveredStartTime.setDateTime(
            QDateTime(
                QDate(todayTmp.year, todayTmp.month, todayTmp.day + 0), QTime(0, 0, 0)
            )
        )
        self.ui.AliAccessorDeleveredEndTime.setDateTime(
            QDateTime(
                QDate(todayTmp.year, todayTmp.month, todayTmp.day + 1), QTime(0, 0, 0)
            )
        )

        self.ui.AliAccessorPriceTablePath.setText("price.xlsx")
        self.ui.priceTablePathButton.clicked.connect(self.CheckPriceTablePath)

        self.ui.AliAccessorShopName.addItem("万盈饰品厂")
        self.ui.AliAccessorShopName.addItem("义乌睿得")

        self.ui.AliAccessorTag.addItem("无")
        self.ui.AliAccessorTag.addItem("红")
        self.ui.AliAccessorTag.addItem("蓝")
        self.ui.AliAccessorTag.addItem("绿")
        self.ui.AliAccessorTag.addItem("黄")
        self.ui.AliAccessorTag.addItem("按单号")

        self.ui.AliAccessorOrderStatus.addItem("已发货")
        self.ui.AliAccessorOrderStatus.addItem("待发货 + 已发货")
        self.ui.AliAccessorOrderStatus.addItem("待发货")
        self.ui.AliAccessorOrderStatus.addItem("待付款")
        self.ui.AliAccessorOrderStatus.addItem("历史订单")

        self.ui.AliAccessorCommit.clicked.connect(self.AliAccessorCalculateBeiHuoTable)

        # MenuA = MenuBar.addMenu("MenuA")
        # MenuA1 = MenuA.addMenu("A1")  # 菜单嵌套子菜单
        # MenuA1.addAction("A1a")
        # MenuA1.addAction("A1b")
        # ActionA2 = MenuA.addAction("A2")  # 菜单添加Action
        # MenuB = MenuBar.addMenu("MenuB")
        # MenuB1 = MenuB.addMenu("B1")
        # MenuB1.addAction("B1a")
        # MenuB1.addAction("B1b")
        # MenuB.addAction("B2")

        self.errorUrl = ""

        self.calStartTime = datetime.now()

        self.ui.AliAccessorSaveFilePath.setText("备货单")
        self.ui.AliAccessorSaveFilePathButton.clicked.connect(self.CheckSaveFilePath)

        # 物流检查按钮
        self.ui.AliAccessorCheckDelivery.clicked.connect(self.CheckDelivery)

        #
        self.shopId = self.ui.AliAccessorShopName.currentIndex() + 1
        self.shopName = self.ui.AliAccessorShopName.currentText()
        self.mode = self.ui.AliAccessorTag.currentIndex()

        self.orderStatus = self.ui.AliAccessorOrderStatus.currentIndex()

        self.startYear = self.ui.AliAccessorStartTime.date().toString("yyyy")
        self.startMonth = self.ui.AliAccessorStartTime.date().toString("MM")
        self.startDay = self.ui.AliAccessorStartTime.date().toString("dd")

        self.createStartTime = (
            datetime(
                int(self.startYear), int(self.startMonth), int(self.startDay)
            ).strftime("%Y%m%d")
            + "000000000+0800"
        )

        self.endYear = self.ui.AliAccessorEndTime.date().toString("yyyy")
        self.endMonth = self.ui.AliAccessorEndTime.date().toString("MM")
        self.endDay = self.ui.AliAccessorEndTime.date().toString("dd")
        self.deleveredStartTime = int(
            self.ui.AliAccessorDeleveredStartTime.dateTime().toString("yyyyMMddHHmmss")
        )
        self.deleveredEndTime = int(
            self.ui.AliAccessorDeleveredEndTime.dateTime().toString("yyyyMMddHHmmss")
        )

        self.deleveredStartTimeStr = (
            self.ui.AliAccessorDeleveredStartTime.dateTime().toString("yyyy-MM-dd")
        )

        self.createEndTime = (
            datetime(int(self.endYear), int(self.endMonth), int(self.endDay)).strftime(
                "%Y%m%d"
            )
            + "000000000+0800"
        )

        self.isPrintOwn = self.ui.IsPrintOwn.isChecked()

        # 配置日志
        logPath = logging_path + datetime.now().strftime("%m_%d_%H_%M_%S") + ".log"
        logging.basicConfig(
            filename=logPath,
            level=logging.DEBUG,
            format="%(asctime)s - %(levelname)s - %(message)s",
        )
        logging.info("test")

        self.Logout("end AliAccessorInit", "debug")

    def AliAccessorCheckAllParams(self):
        self.Logout2(
            "# 启动计算时间 : " + self.calStartTime.strftime("%Y-%m-%d %H:%M:%S")
        )
        self.shopId = self.ui.AliAccessorShopName.currentIndex() + 1
        self.shopName = self.ui.AliAccessorShopName.currentText()
        self.mode = self.ui.AliAccessorTag.currentIndex()

        self.orderStatus = self.ui.AliAccessorOrderStatus.currentIndex()

        self.startYear = self.ui.AliAccessorStartTime.date().toString("yyyy")
        self.startMonth = self.ui.AliAccessorStartTime.date().toString("MM")
        self.startDay = self.ui.AliAccessorStartTime.date().toString("dd")

        self.createStartTime = (
            datetime(
                int(self.startYear), int(self.startMonth), int(self.startDay)
            ).strftime("%Y%m%d")
            + "000000000+0800"
        )

        self.endYear = self.ui.AliAccessorEndTime.date().toString("yyyy")
        self.endMonth = self.ui.AliAccessorEndTime.date().toString("MM")
        self.endDay = self.ui.AliAccessorEndTime.date().toString("dd")
        self.deleveredStartTime = int(
            self.ui.AliAccessorDeleveredStartTime.dateTime().toString("yyyyMMddHHmmss")
        )
        self.deleveredEndTime = int(
            self.ui.AliAccessorDeleveredEndTime.dateTime().toString("yyyyMMddHHmmss")
        )

        self.createEndTime = (
            datetime(int(self.endYear), int(self.endMonth), int(self.endDay)).strftime(
                "%Y%m%d"
            )
            + "000000000+0800"
        )

        self.isPrintOwn = self.ui.AliAccessorIsPrintOwn.isChecked()
        self.isPrintUnitPrice = self.ui.AliAccessorIsPrintUnitPrice.isChecked()

        self.Logout2("# 店铺名 ：" + self.shopName)
        self.Logout2("# 色标 ：" + self.ui.AliAccessorTag.currentText())
        self.Logout2(
            "# 订单开始时间 ："
            + self.ui.AliAccessorStartTime.date().toString("yyyy-MM-dd")
        )
        self.Logout2(
            "# 订单截止时间 ："
            + self.ui.AliAccessorEndTime.date().toString("yyyy-MM-dd")
        )

        self.isLimitDeleveredTime = self.ui.AliAccessorIsLimitDeleveredTime.isChecked()

        if self.isLimitDeleveredTime:
            self.limitDeliveredTime = {
                "deleveredStartTime": self.deleveredStartTime,
                "deleveredEndTime": self.deleveredEndTime,
            }
        else:
            self.limitDeliveredTime = {}

    def AliAccessorCalculateBeiHuoTable(self):
        self.Logout("AliAccessorCalculateBeiHuoTable", "debug")
        self.AliAccessorCheckAllParams()
        try:
            _thread.start_new_thread(
                self.OrderList,
                (
                    self.shopName,
                    int(self.mode),
                    self.createStartTime,
                    self.createEndTime,
                    self.orderStatus,
                    self.isPrintOwn,
                    self.limitDeliveredTime,
                    self.isPrintUnitPrice,
                    SHOPTYPE_ALI_ACCESSOR,
                ),
            )
            self.Logout2("\n # 启动计算 请稍后 \n")

        except:
            self.Logout2("Error: 无法计算启动线程")

    def AliAccessorGetBeihuoJson(
        self,
        orders,
        is_print_own,
        mode=0,
        limit_delivered_time={},
        isPrintUnitPrice=False,
    ):
        self.Logout("AliAccessorGetBeihuoJson", "debug")
        beihuo_json = {}

        for order in orders:
            if mode != 0 and (
                "sellerRemarkIcon" not in order["baseInfo"]
                or order["baseInfo"]["sellerRemarkIcon"] != str(mode)
            ):
                continue
            if "sellerRemarkIcon" in order["baseInfo"] and (
                order["baseInfo"]["sellerRemarkIcon"] == "2"
                or order["baseInfo"]["sellerRemarkIcon"] == "3"
            ):
                print("刷单")
                continue
            for product_item in order["productItems"]:
                cargo_number_tag = (
                    "cargoNumber"
                    if "cargoNumber" in product_item
                    else "productCargoNumber"
                )
                cargo_number = product_item[cargo_number_tag]
                color = product_item["skuInfos"][0]["value"]

                product_dict = beihuo_json.setdefault(
                    cargo_number,
                    {
                        "color": color,
                        "quantity": 0,
                        "price": GetCost(cargo_number, color),
                        "productImgUrl": product_item["productImgUrl"][1],
                    },
                )
                product_dict["quantity"] += product_item["quantity"]
                product_dict["cost"] = product_dict["price"] * product_dict["quantity"]

        self.AliAccessorGetTable(beihuo_json, is_print_own, isPrintUnitPrice)

    def AliAccessorGetTable(self, BeihuoJson, isPrintOwn, isPrintUnitPrice):
        # 制表
        BeihuoList = []

        for item in BeihuoJson:
            BeihuoList.append(item)

        # 写表
        savePath = self.ui.saveFilePath.toPlainText().split(".")[0]

        # 保存备货单
        BH_wb = xlsxwriter.Workbook(
            savePath + "/" + self.calStartTime.strftime("%m_%d_%H_%M_%S") + "_饰品.xlsx"
        )
        BH_sheet = BH_wb.add_worksheet("BH")
        BH_x = 0
        BH_y = 0

        sumCountX = 0

        amount = 0

        for cargoNumber in BeihuoJson:
            BH_sheet.write(BH_x, BH_y, cargoNumber)  # 货号
            BH_y += 1
            BH_sheet.write(BH_x, BH_y, BeihuoJson[cargoNumber]["color"])  # 品名
            BH_y += 1
            BH_sheet.write(
                BH_x, BH_y, str(BeihuoJson[cargoNumber]["quantity"]) + " 件"
            )  # 数量
            BH_y += 1
            amount += BeihuoJson[cargoNumber]["cost"]
            BH_sheet.write(
                BH_x, BH_y, str(BeihuoJson[cargoNumber]["cost"]) + " 元"
            )  # 总价
            BH_y += 1

            # 插图
            imgUrl = BeihuoJson[cargoNumber]["productImgUrl"]
            imageName = imgUrl.split(".jpg")[0].split("/")[-1]

            if ImageHandler.IsImageExist(imageName):
                time.sleep(0.2)
                # 本地存有图片，读出
                self.Logout("before read img " + imageName, "debug")
                imageData = ImageHandler.ReadImageFromDir(imageName)
                self.Logout("after read img " + imageName, "debug")

                self.Logout("读取本地图片")

            else:
                self.Logout("下载图片")
                # time.sleep(0.5)
                rt = self.RequestPic(imgUrl)

                if rt == 420:
                    # 本地存有图片，读出
                    imageData = ImageHandler.ReadImageFromDir(imageName)

                    self.Logout("读取到手动下载图片")
                elif rt == 400:
                    imageData = ImageHandler.ReadImageFromDir(imageName)
                else:

                    imageDataRaw = rt.read()

                    imageData = io.BytesIO(imageDataRaw)

                    # 保存图片
                    ImageHandler.SaveImage(imageData.getvalue(), imageName)

            logging.debug("before insert img " + imageName)
            BH_sheet.insert_image(
                BH_x,
                BH_y,
                imageName,
                {
                    "image_data": imageData,
                    "x_offset": 3,
                    "x_scale": 0.14,
                    "y_scale": 0.14,
                },
            )
            logging.debug("end insert img " + imageName)

            sumCountX = BH_x + 1

            BH_x += 6

            BH_y = 0

        sumCountX += 6
        # 输出字体
        priceStyle = BH_wb.add_format(
            {
                # "fg_color": "yellow",  # 单元格的背景颜色
                "bold": 1,  # 字体加粗
                "align": "left",  # 对齐方式
                "valign": "vcenter",  # 字体对齐方式
                "font_color": "red",  # 字体颜色
            }
        )
        writeStr = "拿货成本总计：" + str(amount) + " 元"
        BH_sheet.merge_range(
            "A" + str(sumCountX + 1) + ":D" + str(sumCountX + 1),
            writeStr,
            priceStyle,
        )

        BH_wb.close()


if __name__ == "__main__":
    # 高分辨率适配
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_ShareOpenGLContexts)
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication([])
    w = Window()
    w.ui.show()

    app.exec_()
