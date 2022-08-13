from Common import utils
import requests
import math

request_type = "param2/1/com.alibaba.trade/"
url = utils.base_url + request_type

# 交易数据获取  默认1页
def GetTradeData(data, shopName):
    data['access_token'] = utils.access_token[shopName]
    _aop_signature = utils.CalculateSignature(request_type + "alibaba.trade.getSellerOrderList/" + utils.AppKey[shopName], data, shopName)
    data['_aop_signature'] = _aop_signature
    url = utils.base_url + request_type + "alibaba.trade.getSellerOrderList/" + utils.AppKey[shopName]
    response = requests.post(url, data=data)

    return response.json()

# 退款数据获取  默认1页
def GetRefundData(data, shopName):
    data['access_token'] = utils.access_token[shopName]
    _aop_signature = utils.CalculateSignature(request_type + "alibaba.trade.refund.queryOrderRefundList/" + utils.AppKey[shopName], data, shopName)
    data['_aop_signature'] = _aop_signature
    url = utils.base_url + request_type + "alibaba.trade.refund.queryOrderRefundList/" + utils.AppKey[shopName]
    response = requests.post(url, data=data)

    return response.json()

# 获取总页数
def GetPageNum(shopName):
    data = dict()
    data['access_token'] = utils.access_token[shopName]
    _aop_signature = utils.CalculateSignature(request_type + "alibaba.trade.getSellerOrderList/" + utils.AppKey[shopName], data, shopName)
    data['_aop_signature'] = _aop_signature
    url = utils.base_url + request_type + "alibaba.trade.getSellerOrderList/" + utils.AppKey[shopName]
    response = requests.post(url, data=data)
    responseData = response.json()

    return math.ceil(responseData['totalRecord']/20)

# if __name__ == '__main__':
#     data = {'needMemoInfo':'true'}
#     response = GetTradeData(data)
#
#     print(len(response['result']))
#
#     pageNum = math.ceil(response['totalRecord']/20)
#
#     print(GetPageNum())

    # print(len(response['result']))
    # print(response)
    # print(response['totalRecord'])