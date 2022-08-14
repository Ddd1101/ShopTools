from PySide2 import QtCore
from PySide2.QtWidgets import QApplication, QMainWindow, QPushButton, QPlainTextEdit, QMessageBox, QFileDialog
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, QDate,   QDateTime , QTime

import os
from datetime import datetime, timedelta
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


dirname = os.path.dirname(PySide2.__file__)
plugin_path = os.path.join(dirname, 'plugins', 'platforms')
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = plugin_path

AppKey = {'联球制衣厂':'2460951','朝雄制衣厂':'4048570', '朝逸饰品厂':'7464845'}
AppSecret =  {"联球制衣厂":b'AnBm2v0I80x0',"朝雄制衣厂":b'vxMiqUGDZ7p',"朝逸饰品厂":b'Oo0hRGjaaPb'}
access_token = {'联球制衣厂':'b7635cd6-18b8-4b96-b6a5-d2ab04f203ae','朝雄制衣厂':'f7a85f0e-fb90-40b1-9bc6-b66ad3cf120b', '朝逸饰品厂':'37454060-a51a-497d-9866-0d722a1bd5cc'}

en_code = ['s','S','m','M', 'l','L','x','X']

base_url = 'https://gw.open.1688.com/openapi/'


path = os.path.realpath(os.curdir)

price_path = path + '/price.xlsx'

request_type = "param2/1/com.alibaba.trade/"
url = base_url + request_type

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
        print(1)
        print(cargoNumber + "==========" + worksheet.cell(t, colNum).value)
        return -1
    colIndex = CalPriceLocation(skuInfosValue)
    if colIndex != None:
        _price = worksheet.cell(rowIndex, int(colIndex)).value
    else:
        _price = ''
    if _price == '':
        print(rowIndex, colIndex)
        print(cargoNumber + "==========" + worksheet.cell(t, colNum).value)
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
        params.append(key+data[key])
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
    _aop_signature = CalculateSignature(request_type + "alibaba.trade.getSellerOrderList/" + AppKey[shopName], data, shopName)
    data['_aop_signature'] = _aop_signature
    url = base_url + request_type + "alibaba.trade.getSellerOrderList/" + AppKey[shopName]
    response = requests.post(url, data=data)

    return response.json()

class Window:

    def __init__(self):
        super(Window, self).__init__()

        # 从文件中加载UI定义
        qfile = QFile("QtUi.ui")
        qfile.open(QFile.ReadOnly)
        qfile.close()
        self.ui = QUiLoader().load(qfile)

        # 时间
        today = datetime.now()
        self.ui.startTime.setDateTime(QDateTime(QDate(today.year, today.month, today.day - 1), QTime(0, 0, 0)))
        self.ui.endTime.setDateTime(QDateTime(QDate(today.year, today.month, today.day + 1), QTime(0, 0, 0)))

        self.ui.priceTablePath.setText("price.xlsx")
        self.ui.priceTablePathButton.clicked.connect(self.CheckPriceTablePath)

        self.ui.saveFilePath.setText("BHtmp.xlsx")
        self.ui.saveFilePathButton.clicked.connect(self.CheckSaveFilePath)

        self.ui.shopName.addItem("联球制衣厂")
        self.ui.shopName.addItem("朝雄制衣厂")

        self.ui.Tag.addItem("无")
        self.ui.Tag.addItem("红")
        self.ui.Tag.addItem("蓝")
        self.ui.Tag.addItem("绿")
        self.ui.Tag.addItem("黄")

        self.ui.commit.clicked.connect(self.CheckAllParams)


    def CheckPriceTablePath(self):
        filePath = QFileDialog.getOpenFileName(QMainWindow(), "选择文件")  # 选择目录，返回选中的路径
        self.ui.priceTablePath.setText(filePath[0])


    def CheckSaveFilePath(self):
        FileDirectory = QFileDialog.getExistingDirectory(QMainWindow(), "选择文件夹")  # 选择目录，返回选中的路径
        self.ui.saveFilePath.setText(FileDirectory)

    def CheckAllParams(self):
        shopId = self.ui.shopName.currentIndex() + 1
        shopName = self.ui.shopName.currentText()
        mode = self.ui.Tag.currentIndex()

        startYear = self.ui.startTime.date().toString("yyyy")
        startMonth = self.ui.startTime.date().toString("MM")
        startDay = self.ui.startTime.date().toString("dd")

        createStartTime = datetime(int(startYear),int(startMonth),int(startDay)).strftime('%Y%m%d') + '000000000+0800'

        endYear = self.ui.endTime.date().toString("yyyy")
        endMonth = self.ui.endTime.date().toString("MM")
        endDay = self.ui.endTime.date().toString("dd")

        createEndTime = datetime(int(endYear), int(endMonth), int(endDay)).strftime('%Y%m%d') + '000000000+0800'

        self.LogOut("店铺名 ：" + shopName)
        self.LogOut("色标 ：" + self.ui.Tag.currentText())
        self.LogOut("订单开始时间 ：" + self.ui.startTime.date().toString("yyyy-MM-dd"))
        self.LogOut("订单截止时间 ：" + self.ui.endTime.date().toString("yyyy-MM-dd"))

        try:
            _thread.start_new_thread(self.OrderList, (shopId, shopName, int(mode), createStartTime, createEndTime,))
        except:
            print("Error: 无法启动线程")

    def LogOut(self, text):
        self.ui.output.append(text)


    def OrderList(self, shopId, shopName, mode, createStartTime, createEndTime):
        if mode == 0:
            self.GetOrderBill(createStartTime, createEndTime, 'waitsellersend', shopName=shopName)
        else:
            self.GetOrderBill(createStartTime, createEndTime, 'waitsellersend,waitbuyerreceive', shopName, mode)

        self.LogOut("# 统计完成 \n")
        self.LogOut("###############################################################")

    def GetOrderBill(self, createStartTime, createEndTime, orderstatusStr, shopName, mode = 0):
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

            BeihuoJson = {}

            # 规格化数据
            for pageId in range(pageNum):
                data = {'page': str(pageId + 1), 'createStartTime': createStartTime, 'createEndTime': createEndTime,
                        'orderStatus': orderstatus, 'needMemoInfo': 'true'}
                response = GetTradeData(data, shopName)

                if orderstatus == 'waitsellersend':
                    for order in response['result']:
                        if ('sellerRemarkIcon' in order['baseInfo']) and ( order['baseInfo']['sellerRemarkIcon'] == '2' or order['baseInfo']['sellerRemarkIcon'] == '3'):
                            continue
                        elif mode != 0 and 'sellerRemarkIcon' not in order['baseInfo']:
                            if mode == 1:
                                order['baseInfo']['sellerRemarkIcon'] = '1'
                            elif mode == 4:
                                order['baseInfo']['sellerRemarkIcon'] = '4'

                orderList += response['result']

        for order in orderList:
            if ('sellerRemarkIcon' in order['baseInfo']) and (order['baseInfo']['sellerRemarkIcon'] == '2' or order['baseInfo']['sellerRemarkIcon'] == '3'):
                continue
            else:
                if mode != 0 and ('sellerRemarkIcon' not in order['baseInfo'] or order['baseInfo']['sellerRemarkIcon'] != str(mode)):
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

        #

        # 写表
        BH_wb = xlsxwriter.Workbook('BHtmp.xlsx')
        BH_sheet = BH_wb.add_worksheet('BH')
        BH_x = 0
        BH_y = 0

        shopNameTmp = ''

        sumCountX = 0
        for _list in BeihuoTable:
            BH_sheet.write(BH_x,BH_y, _list[0])
            BH_y += 1
            BH_sheet.write(BH_x, BH_y, _list[1])
            BH_y += 1
            BH_sheet.write(BH_x, BH_y, _list[2])
            BH_y += 1
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
            # 插图
            picData = self.RequestPic(_list[5])

            image_data = io.BytesIO(picData)
            BH_sheet.insert_image(BH_x, BH_y, _list[4], {'image_data': image_data, 'x_offset': 5, 'x_scale': 0.1, 'y_scale': 0.1})

            if _list[1] != shopNameTmp or _list == BeihuoTable[-1]:
                if _list == BeihuoTable[-1]:
                    sumCountX += 5
                if shopNameTmp != '':
                    self.LogOut(shopNameTmp + ' 拿货总件数 ： ' + str(productsCountByShopName[shopNameTmp][0]) + "  ||  总货款： " + str(round(productsCountByShopName[shopNameTmp][1],3)))
                    BH_sheet.write(sumCountX, 1, shopNameTmp + ' 拿货总件数 ： ' + str(productsCountByShopName[shopNameTmp][0]) + "  ||  总货款： " + str(round(productsCountByShopName[shopNameTmp][1],3)))
                shopNameTmp = _list[1]

            sumCountX = BH_x + 1

            if len(_list[6]) >= 4:
                BH_x += 2
            else:
                theta = 4 - len(_list[6])
                BH_x = BH_x + theta + 2

            BH_y =  0

        # BH_sheet.write(sumCountX, 1, shopNameTmp + ' 拿货总件数 ： ' + str(productsCountByShopName[shopNameTmp]))
        BH_wb.close()

    def RequestPic(self, url):
        flag = True
        while flag:
            try:
                self.LogOut(url)
                picData = urllib.request.urlopen(url).read()
                flag = False
                time.sleep(3)
            except:
                self.LogOut("# 获取图片异常 重试")
                time.sleep(3)

        return picData

if __name__ == '__main__':
    app = QApplication([])
    w = Window()
    w.ui.show()

    app.exec_()