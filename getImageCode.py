from PIL import Image
import tesserocr
import requests
import time
from PIL import ImageGrab
from selenium import webdriver
from builtins import print, range, open
import pytesseract

#图片下载链接
image_url = 'http://gysfwpt.yndzyf.com/png.aspx?rnd=0.6564621833042734'
# # 图片保存路径
image_path = 'image/yzm.png'





def image_download(str):
    """
    str:验证码指定位置 截图下载
    """
    # driver = webdriver.Chrome()
    # driver.get('http://222.92.194.131:8080/hrczyy/web/login.login')
    # driver.maximize_window()
    # time.sleep(2)
    # print(str)
    im = ImageGrab.grab(str)
    im.save('image/yzm.png')

def get_image():
    """
    用Image获取图片文件
    :return: 图片文件
    """
    image = Image.open(image_path)
    return image

def image_grayscale_deal(image):
    """
    图片转灰度处理
    :param image:图片文件
    :return: 转灰度处理后的图片文件
    """
    image = image.convert('L')
    #取消注释后可以看到处理后的图片效果
    # image.show()
    return image

def image_thresholding_method(image):
    """
    图片二值化处理
    :param image:转灰度处理后的图片文件
    :return: 二值化处理后的图片文件
    """
    # 阈值
    threshold = 150
    table = []
    for i in range(256):
        if i < threshold:
            table.append(0)
        else:
            table.append(1)
    # 图片二值化，此处第二个参数为数字一
    image = image.point(table, '1')
    #取消注释后可以看到处理后的图片效果
    image.show()
    return image


def captcha_tesserocr_crack(image):
    """
    图像识别
    :param image: 二值化处理后的图片文件
    :return: 识别结果
    """
    result = tesserocr.image_to_text(image)
    # result = pytesseract.image_to_string(image)
    print('识别结果：'+result)
    return result

# str = (1310, 561, 1406, 598)
def getCode(str):
    # image_download(str)
    image = get_image()
    img1 = image_grayscale_deal(image)
    img2 = image_thresholding_method(img1)
    resultCode = captcha_tesserocr_crack(img2)
    # print('验证码：'+resultCode)
    if resultCode == '':
        while 1:
            # print('循环内if前判断')
            image_download(str)
            image = get_image()
            img1 = image_grayscale_deal(image)
            img2 = image_thresholding_method(img1)
            resultCode = captcha_tesserocr_crack(img2)
            # print('循环里面：'+resultCode)
            if resultCode != '':
                # print('if里面循环的验证码： '+resultCode)
                break
    return resultCode
    # print(captcha_tesserocr_crack(img2))

# if __name__ == '__main__':
#     # str = (1310, 561, 1406, 598)
#     getCode(str)

