# _*_ coding:utf-8 _*_
from PIL import Image
from selenium import webdriver
from PIL import ImageGrab

import pytesseract
import time

url = 'http://114.255.42.204:83/'
driver = webdriver.Chrome()
driver.maximize_window()  # 将浏览器最大化
driver.get(url)
# 截取当前网页并放到E盘下命名为printscreen，该网页有我们需要的验证码

bbox = (1149, 521, 1297, 600)
im = ImageGrab.grab(bbox)
im.save('image/zy.png')

# driver.save_screenshot('image/test.png')
# imgelement = driver.find_element_by_xpath('//*[@id="UpdatePanel1"]/ul/li[4]/img')  # 定位验证码
# location = imgelement.location  # 获取验证码x,y轴坐标
#
# size = imgelement.size  # 获取验证码的长宽
# left = location['x']
# top = location['y']
# right = left + size['width']
# bottom = top + size['height']
#   # 写成我们需要截取的位置坐标
# print((left, top, right, bottom))
# i = Image.open("image/test.png")  # 打开截图
# frame4 = i.crop((left, top, right, bottom))  # 使用Image的crop函数，从截图中再次截取我们需要的区域
# frame4.save('image/test2.png') # 保存我们接下来的验证码图片 进行打码time
# #这样获取到的图片就是验证码了。。不过是最简单的那种还需要进行调整
