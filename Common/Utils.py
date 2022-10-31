import os

import requests

import hmac
import math

import xlrd
import PySide2

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
    _aop_signature = CalculateSignature(request_type + "alibaba.trade.getSellerOrderList/" + AppKey[shopName], data, shopName)
    data['_aop_signature'] = _aop_signature
    url = base_url + request_type + "alibaba.trade.getSellerOrderList/" + AppKey[shopName]
    response = requests.post(url, data=data)

    return response.json()


def GetSingleTradeData(data, shopName):
    data['access_token'] = access_token[shopName]
    _aop_signature = CalculateSignature(request_type + "alibaba.trade.get.sellerView/" + AppKey[shopName], data, shopName)
    data['_aop_signature'] = _aop_signature
    url = base_url + request_type + "alibaba.trade.get.sellerView/" + AppKey[shopName]
    response = requests.post(url, data=data)

    return response.json()