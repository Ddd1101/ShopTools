# -*- coding: utf-8 -*-

import hmac
import math
from skimage import io

import urllib

import time

AppKey = {'联球制衣厂':'2460951','朝雄制衣厂':'4048570', '朝逸饰品厂':'7464845'}
AppSecret =  {"联球制衣厂":b'AnBm2v0I80x0',"朝雄制衣厂":b'vxMiqUGDZ7p',"朝逸饰品厂":b'Oo0hRGjaaPb'}
access_token = {'联球制衣厂':'b7635cd6-18b8-4b96-b6a5-d2ab04f203ae','朝雄制衣厂':'f7a85f0e-fb90-40b1-9bc6-b66ad3cf120b', '朝逸饰品厂':'37454060-a51a-497d-9866-0d722a1bd5cc'}

base_url = 'https://gw.open.1688.com/openapi/'

import Common.ExcelProcess as excelp

worksheet = excelp.GetPriceGrid()

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

def CalPageNum(totalRecord):
    return math.ceil(totalRecord/20)

def CalDayNum(day):
    day = int(day) + 1

def SavePicByUrl(url):
    io.imsave('./tmp.jpg', io.imread(url))

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

def PDDid2Commonid(cargoNumber, colNum=0):
    rowIndex = -1
    for t in range(1, worksheet.nrows):
        if str(cargoNumber) == str(worksheet.cell(t, colNum).value):
            rowIndex = t
            break
    if rowIndex == -1:
        print(1)
        print(cargoNumber + "==========" + worksheet.cell(t, colNum).value)
        return cargoNumber
    # print(skuInfosValue)
    # colIndex = CalPriceLocation(skuInfosValue)
    commonId = worksheet.cell(rowIndex, 0).value
    if commonId == '':
        print(cargoNumber + "==========" + worksheet.cell(t, colNum).value)
        commonId = cargoNumber
    return commonId

en_code = ['s','S','m','M', 'l','L','x','X']

def CalPriceLocation(_size):
    if _size[0] in en_code:
        return CalPriceLocationENCode(_size)
    # print(CalSize(_size))
    theta = (CalSize(_size) - 90)/5
    _col = 9 + theta
    # print(_col)
    return _col

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

def CalSize(sizeDescription):
    _size = 0
    for i in range(len(sizeDescription)):
        if sizeDescription[i] >= '0' and sizeDescription[i] <= '9':
            _size = _size*10 + int(sizeDescription[i])
        else:
            break
    return _size

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


# if __name__ == '__main__':
#     data = {access_token: 'ed0be2fc-75d0-4606-a3cb-bc129e4f312f', 'orderStatus': 'success'}
#     urlPath = 'param2/1/com.alibaba.trade/' + 'alibaba.trade.getSellerOrderList/' + AppKey
#     data = CalculateSignature(urlPath,data)
#     print(data)

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


def NumFormate(numStr):
    res = ""
    if numStr[0] in en_code:
        for item in numStr:
            if item in en_code:
                res += item
                break

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


# def RequestPic(url):
#     flag = True
#     while flag:
#         try:
#             print(url)
#             picData = urllib.request.urlopen(url).read()
#             flag = False
#         except:
#             print("# 获取图片异常 重试")
#             time.sleep(3)

    return picData