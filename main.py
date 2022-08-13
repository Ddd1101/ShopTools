from PySide2 import QtCore
from PySide2.QtWidgets import QApplication, QMainWindow, QPushButton, QPlainTextEdit, QMessageBox, QFileDialog
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, QDate,   QDateTime , QTime

import os
from datetime import datetime, timedelta
import Trade as TransactionData
import Common.utils as utils
import xlsxwriter
import io
import urllib
import time

path = os.path.realpath(os.curdir)

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

        self.LogOut(str(shopId))
        self.LogOut(shopName)
        self.LogOut(str(mode))
        self.LogOut(createStartTime)
        self.LogOut(createEndTime)

        self.OrderList(shopId, shopName, int(mode), createStartTime, createEndTime)

        self.LogOut("# 统计完成")

    def LogOut(self, text):
        self.ui.output.append(text)


    def OrderList(self, shopId, shopName, mode, createStartTime, createEndTime):
        if mode == 0:
            self.GetOrderBill(createStartTime, createEndTime, 'waitsellersend', shopName=shopName)
        else:
            # GetOrderBill(createStartTime,createEndTime,  'waitsellersend', shopName, mode)
            self.GetOrderBill(createStartTime, createEndTime, 'waitsellersend,waitbuyerreceive', shopName, mode)

    def GetOrderBill(self, createStartTime, createEndTime, orderstatusStr, shopName, mode = 0):
        orderList = []

        orderstatusList = orderstatusStr.split(',')

        for orderstatus in orderstatusList:
            data = {'createStartTime': createStartTime, 'createEndTime': createEndTime, 'orderStatus': orderstatus,
                    'needMemoInfo': 'true'}
            response = TransactionData.GetTradeData(data, shopName)
            self.LogOut('# ' + orderstatus + ' : ' + str(response['totalRecord']) + '条记录')
            pageNum = utils.CalPageNum(response['totalRecord'])

            BeihuoJson = {}

            # 规格化数据
            for pageId in range(pageNum):
                data = {'page': str(pageId + 1), 'createStartTime': createStartTime, 'createEndTime': createEndTime,
                        'orderStatus': orderstatus, 'needMemoInfo': 'true'}
                response = TransactionData.GetTradeData(data, shopName)

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
                    BeihuoJson[cargoNumber][color]['products'][height]['price'] = utils.GetCost(cargoNumber, height)

                    # 总价
                    BeihuoJson[cargoNumber][color]['products'][height]['cost'] = BeihuoJson[cargoNumber][color]['products'][height]['price']*BeihuoJson[cargoNumber][color]['products'][height]['quantity']

        # 制表
        productsCountByShopName = {}
        BeihuoList = []
        BeihuoTable = []
        for item in BeihuoJson:
            BeihuoList.append(item)
            adressAndShopName = utils.GetAdressAndShopName(item)
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
                amount = amount + utils.NumFormate4Print(height[0]) + ' ' + str(height[1]) + '件'

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