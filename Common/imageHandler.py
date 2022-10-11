import urllib
import time

import os
import io

def SaveImage(imageData, imageName):
    # 保存图片
    with open('images/' + imageName + '.jpg', 'wb') as file:
        file.write(imageData)  # 保存到本地

def IsImageExist(filename):
    if os.access("images/" + filename + '.jpg', os.R_OK):
        return True
    else:
        return False

def ReadImageFromDir(filename):
    file = open("images/" + filename + '.jpg', 'br')  # 使用二进制

    imageData = io.BytesIO(file.read())  # 使用BytesIO读取

    return imageData