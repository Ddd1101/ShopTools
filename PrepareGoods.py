from PySide2.QtWidgets import QMainWindow, QFileDialog
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, QDate, QDateTime, QTime, QUrl
from PySide2.QtGui import QDesktopServices

from datetime import datetime
import xlsxwriter
import io
import urllib
import time

import _thread

import common.ImageHandler as ImageHandler

import common.Utils as Utils

class PrepareGoods:
    def __init__(self):
        super(PrepareGoods, self).__init__()

        # 从文件中加载UI定义
        qfile = QFile("QtUi.ui")
        qfile.open(QFile.ReadOnly)
        qfile.close()
        self.ui = QUiLoader().load(qfile)

        # 订单时间
        today = datetime.now()
        self.ui.startTime.setDateTime(QDateTime(QDate(today.year, today.month, today.day - 1), QTime(0, 0, 0)))
        self.ui.endTime.setDateTime(QDateTime(QDate(today.year, today.month, today.day + 1), QTime(0, 0, 0)))

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

        self.ui.orderStatus.addItem("待发货 + 已发货")
        self.ui.orderStatus.addItem("待发货")
        self.ui.orderStatus.addItem("已发货")
        self.ui.orderStatus.addItem("待付款")

        self.ui.commit.clicked.connect(self.CheckAllParams)

        self.errorUrl = ""

        self.calStartTime = datetime.now()

        self.ui.saveFilePath.setText("BHtmp")
        self.ui.saveFilePathButton.clicked.connect(self.CheckSaveFilePath)

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
        else:
            limitDeliveredTime = {}

        try:
            _thread.start_new_thread(self.OrderList, (
            shopName, int(mode), createStartTime, createEndTime, orderStatus, isPrintOwn, limitDeliveredTime))
        except:
            self.LogOut("Error: 无法启动线程")

    def LogOut(self, text):
        self.ui.output.append(text)

    def OrderList(self, shopName, mode, createStartTime, createEndTime, orderStatus, isPrintOwn, limitDeliveredTime={}):
        if mode == 5:
            orderId = int(self.ui.orderId.toPlainText())
            self.GetSingleOrder(shopName, orderId, isPrintOwn)
        elif orderStatus == 0:
            self.GetOrderBill(createStartTime, createEndTime, 'waitsellersend,waitbuyerreceive', shopName, isPrintOwn,
                              mode, limitDeliveredTime)
        elif orderStatus == 1:
            self.GetOrderBill(createStartTime, createEndTime, 'waitsellersend', shopName, isPrintOwn, mode,
                              limitDeliveredTime)
        elif orderStatus == 2:
            self.GetOrderBill(createStartTime, createEndTime, 'waitbuyerreceive', shopName, isPrintOwn, mode,
                              limitDeliveredTime)
        elif orderStatus == 3:
            self.GetOrderBill(createStartTime, createEndTime, 'waitbuyerpay', shopName, isPrintOwn, mode,
                              limitDeliveredTime)

        self.LogOut("# 统计完成 \n")
        self.LogOut("###############################################################")

    def GetOrderBill(self, createStartTime, createEndTime, orderstatusStr, shopName, isPrintOwn, mode=0,
                     limitDeliveredTime={}):
        orderList = []

        orderstatusList = orderstatusStr.split(',')

        for orderstatus in orderstatusList:
            data = {'createStartTime': createStartTime, 'createEndTime': createEndTime, 'orderStatus': orderstatus,
                    'needMemoInfo': 'true'}
            response = Utils.GetTradeData(data, shopName)
            if orderstatus == 'waitsellersend':
                orderstatusStr = '待发货'
            if orderstatus == 'waitbuyerreceive':
                orderstatusStr = '已发货'
            self.LogOut('# ' + orderstatusStr + ' : ' + str(response['totalRecord']) + '条记录')
            pageNum = Utils.CalPageNum(response['totalRecord'])

            # 规格化数据
            for pageId in range(pageNum):
                data = {'page': str(pageId + 1), 'createStartTime': createStartTime, 'createEndTime': createEndTime,
                        'orderStatus': orderstatus, 'needMemoInfo': 'true'}
                response = Utils.GetTradeData(data, shopName)

                if orderstatus == 'waitsellersend' or orderstatus == 'waitbuyerreceive':
                    for order in response['result']:
                        if ('sellerRemarkIcon' in order['baseInfo']) and (
                                order['baseInfo']['sellerRemarkIcon'] == '2' or order['baseInfo'][
                            'sellerRemarkIcon'] == '3'):
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
        tmp = Utils.GetSingleTradeData(data, shopName)
        orderList.append(tmp['result'])
        self.GetBeihuoJson(orderList, isPrintOwn, 0)

    def GetBeihuoJson(self, orderList, isPrintOwn, mode=0, limitDeliveredTime={}):
        BeihuoJson = {}
        for order in orderList:
            if ('sellerRemarkIcon' in order['baseInfo']) and (
                    order['baseInfo']['sellerRemarkIcon'] == '2' or order['baseInfo']['sellerRemarkIcon'] == '3'):
                continue
            else:
                if mode != 0 and (
                        'sellerRemarkIcon' not in order['baseInfo'] or order['baseInfo']['sellerRemarkIcon'] != str(
                        mode)):
                    continue

                if 'allDeliveredTime' in order['baseInfo'] and len(limitDeliveredTime) > 0:  # 根据发货时间判断是否要输出
                    allDeliveredTime = int(order['baseInfo']['allDeliveredTime'][:-8])
                    if allDeliveredTime < limitDeliveredTime['deleveredStartTime'] or allDeliveredTime > \
                            limitDeliveredTime['deleveredEndTime']:
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
                            'products': {},
                            'productImgUrl': productItem['productImgUrl'][1],
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
                        BeihuoJson[cargoNumber][color]['products'][height]['quantity'] = \
                        BeihuoJson[cargoNumber][color]['products'][height]['quantity'] + productItem['quantity']

                    # 单价
                    BeihuoJson[cargoNumber][color]['products'][height]['price'] = Utils.GetCost(cargoNumber, height)

                    # 总价
                    BeihuoJson[cargoNumber][color]['products'][height]['cost'] = \
                    BeihuoJson[cargoNumber][color]['products'][height]['price'] * \
                    BeihuoJson[cargoNumber][color]['products'][height]['quantity']
        self.GetTable(BeihuoJson, isPrintOwn)

    def GetTable(self, BeihuoJson, isPrintOwn):
        # 制表
        productsCountByShopName = {}
        BeihuoList = []
        BeihuoTable = []
        for item in BeihuoJson:
            BeihuoList.append(item)
            adressAndShopName = Utils.GetAdressAndShopName(item)
            BeihuoList.append(adressAndShopName[2])
            BeihuoList.append(adressAndShopName[1])
            BeihuoList.append(adressAndShopName[0])
            if adressAndShopName[2] not in productsCountByShopName:
                productsCountByShopName[adressAndShopName[2]] = [0, 0]
            for color in BeihuoJson[item]:
                BeihuoList.append(color)
                BeihuoList.append(BeihuoJson[item][color]['productImgUrl'])
                heightTable = []
                heightList = []
                for height in BeihuoJson[item][color]['products']:
                    heightList.append(height)
                    heightList.append(BeihuoJson[item][color]['products'][height]['quantity'])
                    productsCountByShopName[adressAndShopName[2]][0] += BeihuoJson[item][color]['products'][height][
                        'quantity']  # 叠加总数
                    productsCountByShopName[adressAndShopName[2]][1] += (
                                BeihuoJson[item][color]['products'][height]['quantity'] *
                                BeihuoJson[item][color]['products'][height]['price'])
                    heightList.append(BeihuoJson[item][color]['products'][height]['price'])
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

        # 排序 拿货地规整
        BeihuoTable.sort(key=lambda x: [x[1], x[2]])

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

            amount = ""
            for height in _list[6]:
                if amount != "":
                    amount += '\n'
                amount = amount + Utils.NumFormate4Print(height[0]) + ' ' + str(height[1]) + '件'

            BH_sheet.write(BH_x, BH_y, amount)
            BH_y += 1

            # 多尺码标识
            if len(amount.split('\n')) >= 2:
                BH_sheet.write(BH_x, BH_y + 3, '多尺码')

            # 插图
            imageName = _list[5].split('.jpg')[0].split('/')[-1]
            if ImageHandler.IsImageExist(imageName):
                # 本地存有图片，读出
                imageData = ImageHandler.ReadImageFromDir(imageName)

                self.LogOut("读取本地图片")
            else:
                self.LogOut("下载图片")
                rt = self.RequestPic(_list[5])

                if rt == 420:
                    # 本地存有图片，读出
                    imageData = ImageHandler.ReadImageFromDir(imageName)

                    self.LogOut("读取到手动下载图片")
                else:

                    imageDataRaw = rt.read()

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
                    self.LogOut(
                        shopNameTmp + ' 拿货总件数 ： ' + str(productsCountByShopName[shopNameTmp][0]) + "  ||  总货款： " + str(
                            round(productsCountByShopName[shopNameTmp][1], 3)))
                    # BH_sheet.write(sumCountX, 1, shopNameTmp + ' 拿货总件数 ： ' + str(productsCountByShopName[shopNameTmp][0]) + "  ||  总货款： " + str(round(productsCountByShopName[shopNameTmp][1],3)))
                shopNameTmp = _list[1]

            sumCountX = BH_x + 1

            if len(_list[6]) >= 5:
                BH_x += 2
            else:
                theta = 5 - len(_list[6])
                BH_x = BH_x + theta + 2

            BH_y = 0

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
