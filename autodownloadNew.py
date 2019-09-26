#标注编码方式 
# -*- coding: utf-8 -*-
#创建时间、作者、文件名等其他信息 
# @Time : 2019/9/3 17:56 
# @Author : 一木Cy 
# @Site : www.guanjiu.xyz
# @File : autodownloadNew.py 
# @Software: IntelliJ IDEA

import datetime
import math
from PIL import ImageGrab
from builtins import print, sorted, len, Exception, range, str, open, globals


from selenium import webdriver
import os
# import path
from pathlib import Path
from selenium.common.exceptions import UnexpectedAlertPresentException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from PIL import Image
import time
import getImageCode


import csv
import random

# 当天日期、昨天日期
from autodownload import get_snap
from yundama import YDMHttp

today = datetime.datetime.now()
# today = '2019-05-01'
yesterday = (today - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
#获取上月
lastMonth = datetime.datetime(today.year, today.month - 1, 1).strftime("%Y-%m")
# yesterday = '2019-06-01'
# 获取上月第一天日期
first_day = datetime.datetime(today.year, today.month - 1, 1).strftime("%Y-%m-%d")
#获取上月最后一天
lastDay = datetime.date(datetime.date.today().year,datetime.date.today().month,1)-datetime.timedelta(1)
# 获取上月的当天日期
last_month_day = datetime.datetime(today.year, today.month - 1, today.day).strftime("%Y-%m-%d")
# first_day = '2019-05-01'
#获取本月第一天
thisMonth = datetime.datetime(today.year, today.month, 1).strftime("%Y-%m-%d")
imagePath = 'cache.png'

# ------------------------yundama -------------------------------------
ydmUsername = 'seaxw'  # 用户名
ydmPassword = 't46900_'  # 密码
appid = 1  # 开发者相关 功能使用和用户无关
appkey = '22cc5376925e9387a23cf797cb9ba745'  # 开发者相关 功能使用和用户无关
filename = 'cache.png'  # 验证码截图
# 验证码类型，# 例：1004表示4位字母数字，不同类型收费不同。请准确填写，否则影响识别率。在此查询所有类型 http://www.yundama.com/price.html
codetype = 1004
# 超时时间，秒
timeout = 60

# 检查
if (ydmUsername == ''):
    print('请设置好相关参数再测试')
else:
    # 初始化
    yundama = YDMHttp(ydmUsername, ydmPassword, appid, appkey)
    # 登陆云打码
    uid = yundama.login()
    print('云打码登录成功 uid: %s' % uid)
    # # 查询余额
    # balance = yundama.balance()
    # print('balance: %s' % balance)

    # # 开始识别，图片路径，验证码类型ID，超时时间（秒），识别结果
    # cid, result = yundama.decode(filename, codetype, timeout)
    # print('cid: %s, result: %s' % (cid, result))

def get_image(driver, xp, a, b, c, d):  # 对验证码所在位置进行定位，然后截取验证码图片
    # driver.refresh()  # 刷新页面
    # driver.maximize_window()  # 浏览器最大化
    # driver.set_window_size(1920,1080)
    # width_driver = driver.get_window_size()['width']
    # xp = "//*[@id='UpdatePanel1']/ul/li[4]/img"
    img = driver.find_element_by_xpath(xp)
    time.sleep(2)
    location = img.location
    size = img.size
    left = location['x'] + a
    top = location['y'] + b
    right = left + size['width'] + c
    bottom = top + size['height'] + d
    print((left, top, right, bottom))
    page_snap_obj = get_snap(driver)
    # driver.save_screenshot('full_snap.png')
    # page_snap_obj = Image.open('full_snap.png')
    image_obj = page_snap_obj.crop((left, top, right, bottom))
    # image_obj.show()
    image_obj.save("cache.png")
    return image_obj  # 得到的就是验证码

#判断是否下载成功
def is_download_finished(temp_folder):
    firefox_temp_file = sorted(Path(temp_folder).glob('*.part'))
    chrome_temp_file = sorted(Path(temp_folder).glob('*.crdownload'))
    downloaded_files = sorted(Path(temp_folder).glob('*.*'))
    if (len(firefox_temp_file) == 0) and \
            (len(chrome_temp_file) == 0) and \
            (len(downloaded_files) >= 1):
        return True
    else:
        return False


#保存图片
def saveImg(bbox):
    im = ImageGrab.grab(bbox)
    im.save('cache.png')


#  ---------------------------------------------主要方法-------------------------------------
# 获取driver
def getDriver(folder):
    PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))  # 项目根目录 -- 定义web driver时使用该目录，由于chromedriver.exe存放于该项目下
    DRIVER_BIN = os.path.join(PROJECT_ROOT, "chromedriver.exe")
    options = webdriver.ChromeOptions()
    dlto = os.path.join(PROJECT_ROOT, folder)
    prefs = {'download.default_directory': dlto}
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(executable_path=DRIVER_BIN, chrome_options=options)
    return driver



# 获取IEdriver
def getIEDriver(folder):
    PROJECT_ROOT = os.path.abspath(os.path.dirname(__file__))  # 项目根目录 -- 定义web driver时使用该目录，由于chromedriver.exe存放于该项目下
    DRIVER_BIN = os.path.join(PROJECT_ROOT, "IEDriverServer.exe")
    options = webdriver.IeOptions()
    dlto = os.path.join(PROJECT_ROOT, folder)
    prefs = {'download.default_directory': dlto}
    options.add_additional_option('Aggrs', prefs)
    driver = webdriver.Ie(executable_path=DRIVER_BIN, ie_options=options)
    return driver


# 防止验证码错误，导致下载异常
def autoDownload(workerID, url, pageUsername, pagePassword, folder):
    global download_reason
    every_start_time = datetime.datetime.now()
    format_every_start_time = datetime.datetime(every_start_time.year, every_start_time.month, every_start_time.day,
                                                every_start_time.hour,
                                                every_start_time.minute).strftime(
        "%Y%m%d-%H%M")
    try:
        globals()[workerID](url, pageUsername, pagePassword, folder)
    except UnexpectedAlertPresentException as alertException:
        print(alertException, ">>验证码错误", folder, "--", workerID, ">>>正在重新尝试...")
        autoDownload(workerID, url, pageUsername, pagePassword, folder)
    except Exception as e:
        print(e, ">>异常", folder, "--", workerID, ">>>下载失败")
        download_reason = e
        pass
    with open(file_name, mode='a', newline='', encoding='utf-8') as csv_file:
        finished = is_download_finished(folder)
        if finished is True:
            status = '下载成功'
        else:
            status = '下载失败'
        write_csv = csv.writer(csv_file)
        write_csv.writerow([folder, format_every_start_time, status, str(download_reason)])
        download_reason = ''

# 验证下载是否超时
def checkDownload(folder):
    limit = 10
    finished = False
    while limit > 0:
        limit = limit - 1
        time.sleep(1)
        if is_download_finished(folder):
            limit = 0
            finished = True
        elif limit == 0 and finished is False:
            print('download failed or take too long!')
            raise Exception

#---------------------------------------------------网载自动化代码---------------------------------------------------


#国药控股镇江有限公司 (已完成)
def hr01(url, pageUsername, pagePassword, folder):
    # 获取驱动
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(1)
    #定位账号密码框
    userName = driver.find_element_by_xpath('//*[@id="txtUser"]')
    userPassWord = driver.find_element_by_xpath('//*[@id="txtPass"]')
    actions = ActionChains(driver)
    try:
        #自动填值登录
        userName.clear()
        userName.send_keys(pageUsername)
        time.sleep(1)
        userPassWord.clear()
        userPassWord.send_keys(pagePassword)
        #定位登录按钮
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="Ibtn_ok"]').click()
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(1)
    #切换iframe
    iframe = driver.find_element_by_xpath('//*[@id="BoardTitle"]')
    driver.switch_to_frame(iframe)
    time.sleep(3)
    #进入到流向查询页面
    driver.find_element_by_xpath('//*[@id="leftNb_I2i1_T"]').click()
    time.sleep(3)
    #退出之前的iframe后在新打开iframe
    driver.switch_to_default_content()
    time.sleep(1)
    iframe2 = driver.find_element_by_xpath('//*[@id="frmright"]')
    driver.switch_to_frame(iframe2)
    time.sleep(3)
    #填写时间条件
    driver.find_element_by_xpath('//*[@id="staDate_I"]').clear()
    driver.find_element_by_xpath('//*[@id="staDate_I"]').send_keys(first_day)
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="endDate_I"]').clear()
    driver.find_element_by_xpath('//*[@id="endDate_I"]').send_keys(thisMonth)
    time.sleep(2)
    #定位点击查询
    driver.find_element_by_xpath('//*[@id="btnQuery"]').click()
    time.sleep(5)
    #点击导出按钮
    driver.find_element_by_xpath('//*[@id="lbtnExport2"]').click()
    time.sleep(2)
    checkDownload(folder)
    driver.quit()

#华润医药商业集团的 销售流向通用方法
def getSalesTo(driver,pageUsername,pagePassword,companyName):
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("txtname")
    password = driver.find_element_by_id("txtpwd")
    vCode = driver.find_element_by_id("txtcheckcode")
    bbox = (1166, 530, 1285, 591)
    saveImg(bbox)
    print(bbox)
    # resultCode = getImageCode.getCode(bbox)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 用户名、密码、验证码赋值
    username.send_keys(pageUsername)
    time.sleep(1)
    password.send_keys(pagePassword)
    time.sleep(1)
    vCode.send_keys(resultCode)
    time.sleep(2)
    # 登录系统
    driver.find_element_by_name("ImgSubmit").click()
    time.sleep(3)
    driver.find_element_by_xpath('//*[@id="nav"]/div/div[2]/ul/li/div/a').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="nav"]/div/div[2]/ul/li/ul/li[3]/div/a').click()
    time.sleep(2)
    #切换ifrmae
    ifrmae = driver.find_element_by_xpath('//*[@id="tabs"]/div[2]/div[2]/div/iframe')
    driver.switch_to_frame(ifrmae)
    time.sleep(1)
    #开始填写条件(开始日期--->结束日期--->选择品种)
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]').clear()
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]').send_keys(first_day)
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/span/span').click()
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[2]/input[1]').clear()
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[2]/input[1]').send_keys(thisMonth)
    time.sleep(2)
    driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[2]/span/span').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="selectGoods"]').click()
    time.sleep(4)
    iframe2 = driver.find_element_by_xpath('//*[@id="openIframe"]')
    driver.switch_to_frame(iframe2)
    time.sleep(5)
    if companyName == '葫芦岛':
        driver.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div/div[2]/div[2]/table/tbody/tr[2]/td[2]/div').click()
        time.sleep(2)
        #选择产品
        text = driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[1]/span/span/input')
        text.clear()
        text.send_keys('酒石酸美托洛尔片')
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[1]/span/span/span').click()
        time.sleep(1)
        xb = '//*[@id="goodsPanle"]/div/div/div/div[2]/div[2]/table/tbody/tr[2]'
        db = '//*[@id="goodsPanle"]/div/div/div/div[2]/div[2]/table/tbody/tr[1]'
        xbs = '//*[@id="goodsPanle"]/div/div/div/div[2]/div[2]/table/tbody/tr[4]'
        dbs = '//*[@id="goodsPanle"]/div/div/div/div[2]/div[2]/table/tbody/tr[3]'
        dmk = '//*[@id="goodsPanle"]/div/div/div/div[2]/div[2]/table/tbody/tr'
        atr = [db,xb,dbs,xbs]
        for index in atr:
            driver.find_element_by_xpath(index).click()
            time.sleep(1)
        time.sleep(1)
        text.clear()
        text.send_keys('格列齐特片(Ⅱ)')
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/div[4]/div[1]/div[1]/span/span/span').click()
        time.sleep(1)
        driver.find_element_by_xpath(dmk).click()
        time.sleep(1)
        text.clear()
        time.sleep(2)
    elif companyName == '本溪':
        driver.find_element_by_xpath('//*[@id="goodsPanle"]/div/div/div/div[2]/div[2]/table/tbody/tr[1]').click()
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="goodsPanle"]/div/div/div/div[2]/div[2]/table/tbody/tr[2]').click()
        time.sleep(1)
    driver.find_element_by_xpath('//*[@id="btnSelected"]').click()
    time.sleep(1)
    #退出之前的iframe
    driver.switch_to_default_content()
    time.sleep(3)
    driver.switch_to_frame(ifrmae)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="search"]/span/span').click()
    time.sleep(6)
    driver.find_element_by_xpath('//*[@id="mainPanle"]/div/div/div[1]/a[2]/span/span').click()
    time.sleep(2)
    checkDownload(folder)
    driver.quit()

#华润辽宁葫芦岛医药有限公司  （已完成）
def hr02(url, pageUsername, pagePassword, folder):
    # 获取驱动
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(2)
    companyName = '葫芦岛'
    #pageUsername:账号，pagePassword：密码，companyName：公司名称
    getSalesTo(driver,pageUsername,pagePassword,companyName)

#华润辽宁本溪医药有限公司   （已完成）
def hr03(url, pageUsername, pagePassword, folder):
    # 获取驱动
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(2)
    companyName = '本溪'
    #pageUsername:账号，pagePassword：密码，companyName：公司名称
    getSalesTo(driver,pageUsername,pagePassword,companyName)


#营销 辽宁成大方圆医药连锁有限公司 (已完成)
def hr04(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    # 设置验证码方式
    codetype = 6300
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("opcode")
    password = driver.find_element_by_id("password")
    vCode = driver.find_element_by_id("validateCode")
    xp = '//*[@id="validateimg"]'
    # 打码 参数，驱动器、路径、截图位置（left，top，right，bottom）
    bbox = (190, 107, 10, 7)
    #调用截屏保存图片
    saveImg(bbox)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 输入用户名、密码、验证码
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    vCode.send_keys(resultCode)
    # 登录系统
    driver.find_element_by_xpath('//*[@id="frmLogin"]/div[2]/div/div[2]/div[2]/div[5]/a[1]').click()
    time.sleep(3)
    # 转换菜单栏的iframe
    left_frame = driver.find_element_by_xpath('//*[@id="if_left"]')
    driver.switch_to.frame(left_frame)
    # 定位销售流向位置 //*[@id="subContent"]/dd[1]/ul/li[1]/a
    driver.find_element_by_xpath('//*[@id="subContent"]/dd[1]/ul/li[1]').click()
    time.sleep(3)
    # 转换iframe  cdfyb2b/XslxcxServlet?  //iframe[contains(@src,"DataFlowDetails")]
    driver.switch_to.default_content()
    sale_frame = driver.find_element_by_xpath('//*[@id="contentR"]/iframe')
    driver.switch_to.frame(sale_frame)
    # 设定起始时间 //*[@id="query_date1"]
    driver.find_element_by_xpath('//*[@id="query_date1"]').clear()
    driver.find_element_by_xpath('//*[@id="query_date1"]').send_keys(first_day)
    # 设定结束时间 //*[@id="query_date2"]
    driver.find_element_by_xpath('//*[@id="query_date2"]').clear()
    driver.find_element_by_xpath('//*[@id="query_date2"]').send_keys(yesterday)
    # 定位查询按钮 /html/body/div/div/div[3]/input[1]
    driver.find_element_by_xpath('/html/body/div/div/div[3]/input[1]').click()
    time.sleep(20)
    # 输入回车
    actions = ActionChains(driver)
    try:
        actions.send_keys(Keys.ENTER).perform()
    except:
        print("页面弹框，属正常现象，程序继续运行。")
    # 定位导出按钮 /html/body/div/div/div[3]/input[2]
    driver.find_element_by_xpath('/html/body/div/div/div[3]/input[2]').click()
    time.sleep(1)
    checkDownload(folder)
    driver.quit()

#江西汇仁堂药品连锁有限公司   (已完成)
def hr05(url, pageUsername, pagePassword, folder):
    # 获取驱动
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(1)
    #定位账号密码框
    userName = driver.find_element_by_xpath('//*[@id="l6d220xx"]')
    userPassWord = driver.find_element_by_xpath('//*[@id="6d220xxff"]')
    actions = ActionChains(driver)
    try:
        #自动填值登录
        userName.clear()
        userName.send_keys(pageUsername)
        time.sleep(1)
        userPassWord.clear()
        userPassWord.send_keys(pagePassword)
        #定位登录按钮
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="IbtnEnter"]').click()
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(2)
    #切换iframe
    iframe = driver.find_element_by_xpath('//*[@id="leftFrame"]')
    driver.switch_to_frame(iframe)
    #进入采购配送流向界面
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="submenu0"]/div[1]/table/tbody/tr[2]/td/a').click()
    time.sleep(2)
    #退出iframe并重新进入
    driver.switch_to_default_content()
    time.sleep(2)
    iframe2 = driver.find_element_by_xpath('//*[@id="mainFrame"]')
    driver.switch_to_frame(iframe2)
    time.sleep(3)
    #通过调用js脚本始input变为可编辑状态
    js = "document.getElementById('orderDataBegin').removeAttribute('readonly')"
    driver.execute_script(js)
    time.sleep(2)
    driver.find_element_by_id('orderDataBegin').clear()
    driver.find_element_by_id('orderDataBegin').send_keys(first_day)
    time.sleep(1)
    js2 = "document.getElementById('orderDataEnd').removeAttribute('readonly')"
    driver.execute_script(js2)
    time.sleep(2)
    driver.find_element_by_id('orderDataEnd').clear()
    driver.find_element_by_id('orderDataEnd').send_keys(thisMonth)
    time.sleep(2)
    #定位查询按钮
    driver.find_element_by_xpath('//*[@id="Submit"]').click()
    time.sleep(10)
    #定位导出按钮
    driver.find_element_by_xpath('/html/body/form/table/tbody/tr[2]/th/table/tbody/tr/td/input').click()
    time.sleep(3)
    checkDownload(folder)
    driver.quit()


#福建国大药房连锁有限公司   (已完成)
def hr06(url, pageUsername, pagePassword, folder):
    # 获取驱动
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(1)
    #定位账号密码框
    userName = driver.find_element_by_xpath('//*[@id="loginId"]')
    userPassWord = driver.find_element_by_xpath('//*[@id="password"]')
    vCode = driver.find_element_by_xpath('//*[@id="imagecode"]')
    actions = ActionChains(driver)
    bbox = (1037, 576, 1130, 620)
    saveImg(bbox)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    try:
        #自动填值登录
        userName.clear()
        userName.send_keys(pageUsername)
        time.sleep(1)
        userPassWord.clear()
        userPassWord.send_keys(pagePassword)
        time.sleep(1)
        vCode.clear()
        vCode.send_keys(resultCode)
        #定位登录按钮
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="userDiv"]/p[5]/a').click()
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(2)
    #点击进入流向查询界面
    driver.find_element_by_xpath('//*[@id="treeLeft"]/li/ul/li[3]/div/span[4]').click()
    time.sleep(2)
    #切换到iframe
    iframe = driver.find_element_by_xpath('//*[@id="tabs"]/div[2]/div[2]/div/iframe')
    driver.switch_to_frame(iframe)
    #通过调用js脚本始input变为可编辑状态
    js = "document.getElementsByClassName('combo-text validatebox-text')[0].removeAttribute('readonly')"
    driver.execute_script(js)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="searchArea"]/tbody/tr[1]/td[1]/span/input[1]').clear()
    driver.find_element_by_xpath('//*[@id="searchArea"]/tbody/tr[1]/td[1]/span/input[1]').send_keys(first_day)
    time.sleep(1)
    #js1为  改变hidden隐藏域中传给后台的日期参数
    js1 = '$("input[name=\'FilterModel.BEGINDATE\']").val("'+first_day+'")'
    driver.execute_script(js1)
    js2 = "document.getElementsByClassName('combo-text validatebox-text')[1].removeAttribute('readonly')"
    driver.execute_script(js2)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="searchArea"]/tbody/tr[1]/td[2]/span/input[1]').clear()
    driver.find_element_by_xpath('//*[@id="searchArea"]/tbody/tr[1]/td[2]/span/input[1]').send_keys(thisMonth)
    #js1为  改变hidden隐藏域中传给后台的日期参数
    js3 = '$("input[name=\'FilterModel.ENDDATE\']").val("'+thisMonth+'")'
    driver.execute_script(js3)
    #定位点击查询
    time.sleep(3)
    driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/table/tbody/tr/td[1]/a/span/span').click()
    time.sleep(8)
    #定位导出按钮
    driver.find_element_by_xpath('/html/body/div[1]/div/div[1]/table/tbody/tr/td[3]/a/span/span').click()
    time.sleep(3)
    checkDownload(folder)
    driver.quit()

#华润常州医药有限公司   (已完成，无销量)
def hr07(url, pageUsername, pagePassword, folder):
    # 获取驱动
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(1)
    #定位账号密码框
    userName = driver.find_element_by_xpath('/html/body/form/div/table/tbody/tr[1]/td/input')
    userPassWord = driver.find_element_by_xpath('/html/body/form/div/table/tbody/tr[2]/td/input')
    vCode = driver.find_element_by_xpath('/html/body/form/div/table/tbody/tr[3]/td/input')
    actions = ActionChains(driver)
    bbox = (1292, 551, 1442, 590)
    saveImg(bbox)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    try:
        #自动填值登录
        userName.clear()
        userName.send_keys(pageUsername)
        time.sleep(1)
        userPassWord.clear()
        userPassWord.send_keys(pagePassword)
        time.sleep(1)
        vCode.clear()
        vCode.send_keys(resultCode)
        #定位登录按钮
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/form/div/table/tbody/tr[5]/td/input[1]').click()
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(3)
    #点击进入销售流向查询界面
    driver.find_element_by_xpath('//*[@id="id/hrczyy/web/listview.xsmx"]').click()
    time.sleep(2)
    #切换到销售流向 iframe层
    iframe = driver.find_element_by_xpath('//*[@id="/hrczyy/web/listview.xsmx"]')
    driver.switch_to_frame(iframe)
    time.sleep(1)
    #点击查询
    driver.find_element_by_xpath('//*[@id="ext-gen30"]').click()
    time.sleep(8)
    #填写日期时间
    driver.find_element_by_xpath('//*[@id="ext-comp-1001"]').clear()
    driver.find_element_by_xpath('//*[@id="ext-comp-1001"]').send_keys(first_day)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="ext-comp-1002"]').clear()
    driver.find_element_by_xpath('//*[@id="ext-comp-1002"]').send_keys(thisMonth)
    time.sleep(2)
    #点击弹窗中的查询按钮
    driver.find_element_by_xpath('//*[@id="ext-gen79"]').click()
    time.sleep(random.randrange(8,10))
    #点击导出按钮
    driver.find_element_by_xpath('//*[@id="ext-gen34"]').click()
    time.sleep(2)
    checkDownload(folder)
    driver.quit()


#广东康美新澳医药有限公司    （已完成）
def hr08(url, pageUsername, pagePassword, folder):
    # 获取驱动
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(1)
    #定位账号密码框
    userName = driver.find_element_by_xpath('//*[@id="txtUserName"]')
    userPassWord = driver.find_element_by_xpath('//*[@id="txtPwd"]')
    vCode = driver.find_element_by_xpath('//*[@id="txtCode"]')
    actions = ActionChains(driver)
    bbox = (784, 487, 879, 525)
    saveImg(bbox)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    try:
        #自动填值登录
        userName.clear()
        userName.send_keys(pageUsername)
        time.sleep(1)
        userPassWord.clear()
        userPassWord.send_keys(pagePassword)
        time.sleep(1)
        vCode.clear()
        vCode.send_keys(resultCode)
        #定位登录按钮
        time.sleep(1)
        driver.find_element_by_xpath('//*[@id="ImageButton1"]').click()
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(3)
    #点击选择进入流向查询界面   查询报表----->销售流向查询
    driver.find_element_by_xpath('//*[@id="left_cxbb"]').click()
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="left_li_xslxcx"]').click()
    time.sleep(1)
    #填写日期条件
    driver.find_element_by_xpath('//*[@id="txtBeginTime"]').clear()
    driver.find_element_by_xpath('//*[@id="txtBeginTime"]').send_keys(first_day)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="txtEndTime"]').clear()
    driver.find_element_by_xpath('//*[@id="txtEndTime"]').send_keys(thisMonth)
    time.sleep(1)
    #点击查询
    driver.find_element_by_xpath('//*[@id="Button2"]').click()
    time.sleep(2)
    #点击导出按钮，点击后定位alert  并确定，再继续执行代码
    driver.find_element_by_xpath('//*[@id="btnSearch"]').click()
    alert = driver.switch_to_alert()
    time.sleep(2)
    print(alert.text)
    alert.accept()
    time.sleep(3)
    checkDownload(folder)
    driver.quit()

#杭州九洲大药房连锁有限公司 已完成
def hr09(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    #定位账号  密码框的位置
    userName = driver.find_element_by_xpath('//*[@id="login_id"]')
    userPassWord = driver.find_element_by_xpath('//*[@id="password"]')
    #防止页面报错
    actions = ActionChains(driver)
    try:
        #填写账号密码
        userName.clear()
        userName.send_keys(pageUsername)
        time.sleep(1)
        userPassWord.clear()
        userPassWord.send_keys(pagePassword)
        time.sleep(1)
        #定位登录按钮并模拟点击
        driver.find_element_by_xpath('//*[@id="root"]/div/div/div[2]/div[1]/div[2]/form/div[3]/div/div/button').click()
        time.sleep(1)
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(2)
    #登录成功后进入流向查询页面
    driver.find_element_by_xpath('//*[@id="side-menu"]/li[2]/ul/li[3]').click()
    time.sleep(random.randrange(3, 5))
    #切换iframe层
    iframe = driver.find_element_by_xpath('//*[@id="content-main"]/iframe[2]')
    driver.switch_to.frame(iframe)
    #定位日期填写的文本框且通过js脚本去掉只读属性
    #注意数据库查询条件是为年月 （2019-07）
    js = "document.getElementById('_easyui_textbox_input3').removeAttribute('readonly')"
    driver.execute_script(js)
    driver.find_element_by_id('_easyui_textbox_input3').clear()
    driver.find_element_by_id('_easyui_textbox_input3').send_keys(lastMonth)
    driver.find_element_by_id('_easyui_textbox_input3').click()
    time.sleep(1)
    js2 = "document.getElementById('_easyui_textbox_input4').removeAttribute('readonly')"
    driver.execute_script(js2)
    driver.find_element_by_id('_easyui_textbox_input4').clear()
    driver.find_element_by_id('_easyui_textbox_input4').send_keys(lastMonth)
    driver.find_element_by_id('_easyui_textbox_input4').click()
    time.sleep(10)
    #定位查询按钮并点击
    driver.find_element_by_xpath('//*[@id="northPanle"]/div/a[1]/span').click()
    time.sleep(random.randrange(8, 15))
    # 查询后等待8-15秒后点击导出按钮
    driver.find_element_by_xpath('//*[@id="northPanle"]/div/a[3]/span').click()
    time.sleep(2)
    checkDownload(folder)
    driver.quit()


# 营销浦江英特药业 -- 已完成
def hr10(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片 //*[@id="loginSuccess"]/div/input[1] //*[@id="loginSuccess"]/div/input[2]
    username = driver.find_element_by_xpath('//*[@id="loginSuccess"]/div/input[1]')
    password = driver.find_element_by_xpath('//*[@id="loginSuccess"]/div/input[2]')
    vCode = driver.find_element_by_xpath('//*[@id="loginSuccess"]/div/input[3]')
    vCode.click()
    xp = '//*[@id="verifyCodeId"]'
    driver.maximize_window()
    # 打码 参数，驱动器、路径、截图位置（left，top，right，bottom）
    # get_image(driver, xp, 185, 2, 10, 0)
    bbox = (185, 2, 10, 0)
    saveImg(bbox)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 防止页面错误
    actions = ActionChains(driver)
    try:
        # 输入用户名、密码、验证码
        vCode.clear()
        vCode.send_keys(resultCode)
        username.clear()
        username.send_keys(pageUsername)
        password.clear()
        password.send_keys(pagePassword)
        time.sleep(random.randrange(1, 3))
        # 登录系统 //*[@id="loginSuccess"]/div/input[4]
        driver.find_element_by_xpath('//*[@id="loginSuccess"]/div/input[4]').click()
        time.sleep(3)
        if pageUsername == 'pjsjc':
            driver.switch_to.alert.accept()
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(3)
    # 点击流向查询
    driver.find_element_by_xpath('/html/body/div[6]/div[1]/form/div[2]/a[2]').click()
    # 页面转换
    driver.switch_to.window(driver.window_handles[-1])
    # 跳转到菜单frame agentman/medicineGoto/medicinegoto.jspx
    main_frame = driver.find_element_by_id('cwin')
    driver.switch_to.frame(main_frame)
    if pageUsername == 'gykgtj':
        driver.find_element_by_xpath('//*[@id="entryId"]').click()
        time.sleep(random.randrange(1, 3))
        driver.find_element_by_xpath('//*[@id="entryId"]/option[3]').click()
        time.sleep(random.randrange(1, 3))
    # 定位输入日期 //*[@id="startTime"] //*[@id="endTime"]
    driver.find_element_by_xpath('//*[@id="startTime"]').clear()
    driver.find_element_by_xpath('//*[@id="startTime"]').send_keys(last_month_day)
    driver.find_element_by_xpath('//*[@id="endTime"]').clear()
    driver.find_element_by_xpath('//*[@id="endTime"]').send_keys(yesterday)
    # 定位销售明细
    driver.find_element_by_xpath('//*[@id="button_4"]').click()
    time.sleep(random.randrange(1, 3))
    # 定位查询按钮 //*[@id="medicinegoto_jspx"]/table[2]/tbody/tr[7]/td[2]/span/input
    driver.find_element_by_xpath('//*[@id="medicinegoto_jspx"]/table[2]/tbody/tr[7]/td[2]/span/input').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮 //*[@id="exportBtn"]
    driver.find_element_by_xpath('//*[@id="exportBtn"]').click()
    time.sleep(1)
    checkDownload(folder)
    driver.quit()

# 营销徐州恩华统一医药 -- 已完成
def hr11(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("username")
    password = driver.find_element_by_id("password")
    driver.maximize_window()
    # 防止页面错误
    try:
        # 输入用户名、密码、验证码
        username.clear()
        username.send_keys(pageUsername)
        password.clear()
        password.send_keys(pagePassword)
        time.sleep(random.randrange(1, 3))
        # 登录系统 /html/body/form/table/tbody/tr/td/div/div/div[3]/button
        driver.find_element_by_xpath('/html/body/form/table/tbody/tr/td/div/div/div[3]/button').click()
        time.sleep(3)
    except:
        print('页面错误，正在重新输入')
        driver.quit()
    time.sleep(3)
    # 点击供应商网络服务 //*[@id="navigation"]/ul/li[3]/a
    driver.find_element_by_link_text('配送明细').click()
    time.sleep(random.randrange(1, 3))
    # 定位输入日期
    driver.find_element_by_xpath('//*[@id="n1"]').clear()
    driver.find_element_by_xpath('//*[@id="n1"]').send_keys(first_day)
    driver.find_element_by_xpath('//*[@id="n2"]').clear()
    driver.find_element_by_xpath('//*[@id="n2"]').send_keys(yesterday)
    driver.find_element_by_xpath('//*[@id="n1"]').click()
    # 定位提交按钮 //*[@id="form2"]/table/tbody/tr[6]/td/input[1]
    driver.find_element_by_xpath('//*[@id="form2"]/table/tbody/tr[6]/td/input[1]').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮
    driver.find_element_by_xpath('//*[@id="form2"]/table/tbody/tr[6]/td/input[2]').click()
    time.sleep(1)
    checkDownload(folder)
    driver.quit()


# 营销国药集团临汾 -- 已完成
def hr12(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("UserName")
    password = driver.find_element_by_id("Password")
    actions = ActionChains(driver)
    # driver.maximize_window()
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    password.clear()
    password.send_keys(pagePassword)
    time.sleep(random.randrange(1, 3))
    # 登录系统 //*[@id="loginmessage"]/div[2]/button
    driver.find_element_by_xpath('//*[@id="loginmessage"]/div[2]/button').click()
    time.sleep(3)
    # 防止页面错误
    driver.switch_to.default_content()
    #  解决弹窗 //*[@id="loginbody"]/div[4]/div[3]/div/button
    driver.find_element_by_xpath('//*[@id="loginbody"]/div[4]/div[3]/div/button').click()
    # driver.switch_to.alert.accept()
    time.sleep(random.randrange(1, 3))
    # 点击流向管理 //*[@id="gnmenu"]/li[3]/a/h //*[@id="gnmenu"]/li[3]/ul/li/a/h //*[@id="gnmenu"]/li[3]/ul/li/a/h
    # actions.move_to_element('//*[@id="gnmenu"]/li[3]/a').click().perform()
    driver.find_element_by_xpath('//*[@id="gnmenu"]/li[3]/a/h').click()
    time.sleep(random.randrange(1, 3))
    driver.find_element_by_xpath('//*[@id="gnmenu"]/li[3]/ul/li/a/h').click()
    time.sleep(random.randrange(1, 3))
    # 切换frame /home/Design_index?gnbh=1001&gnmch=流向查询
    main_frame = driver.find_element_by_xpath('//*[@id="1001-1001"]')
    driver.switch_to.frame(main_frame)
    time.sleep(3)
    # 定位输入日期 //*[@id="date_1"]  //*[@id="date_2"]
    driver.find_element_by_xpath('//*[@id="date_1"]').clear()
    driver.find_element_by_xpath('//*[@id="date_1"]').send_keys(first_day)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="date_2"]').clear()
    driver.find_element_by_xpath('//*[@id="date_2"]').send_keys(thisMonth)
    time.sleep(1)
    # driver.find_element_by_xpath('//*[@id="n1"]').click()
    # 定位查询按钮 //*[@id="stageContainer"]/div[8]/ul/li[1]/a
    driver.find_element_by_xpath('//*[@id="stageContainer"]/div[8]/ul/li[1]/a').click(),
    time.sleep(10)
    # 定位导出按钮 //*[@id="div_12"]/div[1]/div[1]/div/div/button
    driver.find_element_by_xpath('//*[@id="div_12"]/div[1]/div[1]/div/div/button').click()
    # 选择csv方式导出 //*[@id="div_12"]/div[1]/div[1]/div/div/ul/li[3]/a
    driver.find_element_by_xpath('//*[@id="div_12"]/div[1]/div[1]/div/div/ul/li[3]/a').click()
    time.sleep(1)
    checkDownload(folder)
    driver.quit()


# 云南省医药嘉源有限公司  倍他乐克，达美康已完成
def hr13(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    #定位账户框与密码框,验证码位置
    userId = driver.find_element_by_xpath('//*[@id="tlxkh"]')
    userName = driver.find_element_by_xpath('//*[@id="tusername"]')
    userPassword = driver.find_element_by_xpath('//*[@id="tpass"]')
    vCode = driver.find_element_by_xpath('//*[@id="tyzm"]')
    xp = '//*[@id="flogin"]/div[7]/img'
    driver.maximize_window()
    # 打码 参数，驱动器、路径、截图位置（left，top，right，bottom）
    # get_image(driver, xp, 280, 105, 10, 5)
    bbox = (280, 105, 10, 5)
    codetype = 1005
    saveImg(bbox)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 防止页面错误
    actions = ActionChains(driver)
    try:
        userId.clear()
        userId.send_keys(pageUsername)
        time.sleep(1)
        userName.clear()
        userName.send_keys(pageUsername)
        time.sleep(1)
        userPassword.clear()
        userPassword.send_keys(pagePassword)
        time.sleep(1)
        vCode.clear()
        vCode.send_keys(resultCode)
        #定位登录按钮
        driver.find_element_by_xpath('//*[@id="btnlogin"]').click()
        time.sleep(2)
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(3)
    #模拟点击定位到查询流向页面
    driver.find_element_by_xpath('//*[@id="form_main"]/div[3]/aside/section/ul/li[2]/a').click()
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="form_main"]/div[3]/aside/section/ul/li[2]/ul/li[1]/a').click()
    time.sleep(2)
    #切换iframe
    iframe = driver.find_element_by_xpath('//*[@id="iframepage"]')
    driver.switch_to.frame(iframe)
    #日期因为被设定为不能直接对文本框修改值，所以通过driver.execute_script去执行js脚本
    js = "document.getElementById('edit2').removeAttribute('readonly')"
    driver.execute_script(js)
    driver.find_element_by_id('edit2').clear()
    driver.find_element_by_id('edit2').send_keys(first_day)
    time.sleep(2)
    js2 = "document.getElementById('edit3').removeAttribute('readonly')"
    driver.execute_script(js2)
    driver.find_element_by_id('edit3').clear()
    driver.find_element_by_id('edit3').send_keys(yesterday)
    time.sleep(2)
    #定位查询按钮
    driver.find_element_by_xpath('/html/body/div/section[2]/div[1]/div[2]/button[1]').click()
    time.sleep(random.randrange(3, 8))
    #定位导出按钮并模拟点击导出
    driver.find_element_by_xpath('//*[@id="btnExportExcel"]').click()
    checkDownload(folder)
    driver.quit()


#国药控股衡阳有限公司 (已完成)
def hr14(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(1)
    #定位账户密码元素位置
    userName = driver.find_element_by_xpath('//*[@id="txtaccount"]')
    userPassWord = driver.find_element_by_xpath('//*[@id="txtpassword"]')
    actions = ActionChains(driver)
    try:
        #自动登录
        userName.clear()
        userName.send_keys(pageUsername)
        time.sleep(1)
        userPassWord.clear()
        userPassWord.send_keys(pagePassword)
        time.sleep(2)
        #定位登录按钮并点击
        driver.find_element_by_xpath('//*[@id="btlogin"]').click()
        time.sleep(15)
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(3)
    #进入流向查询页面
    driver.find_element_by_xpath('//*[@id="accordion"]/li/ul/li/a').click()
    time.sleep(3)
    driver.find_element_by_xpath('//*[@id="accordion"]/li/ul/li/ul/li[1]').click()
    time.sleep(random.randrange(8,10))
    #切换ifrmae层
    iframe = driver.find_element_by_xpath('//*[@id="tabs_iframe_36658"]')
    driver.switch_to.frame(iframe)
    time.sleep(2)
    #定位一些查询条件的元素位置
    #商品代码，分公司名称，业务类型，开始-结束日期
    driver.find_element_by_id('goodscode').click()
    time.sleep(1)
    #小倍 编码
    driver.find_element_by_id('goodscode').send_keys('00000582')
    time.sleep(1)
    #大倍 编码
    driver.find_element_by_id('goodscode').send_keys(',00000560')
    time.sleep(1)
    #达美康 编码
    driver.find_element_by_id('goodscode').send_keys(',00001724')
    time.sleep(2)
    #点击选择分公司
    driver.find_element_by_xpath('//*[@id="QueryArea"]/ul/li[2]/button').click()
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div[2]/ul/li[5]/label').click()
    driver.find_element_by_xpath('//*[@id="QueryArea"]/ul/li[2]/button').click()
    time.sleep(2)
    #点击选择业务类型
    # driver.find_element_by_xpath('//*[@id="QueryArea"]/ul/li[3]/button').click()
    # driver.find_element_by_xpath('/html/body/div[3]/ul/li[1]/label').click()
    # driver.find_element_by_xpath('//*[@id="QueryArea"]/ul/li[3]/button').click()
    # time.sleep(3)
    #填写开始日期
    driver.find_element_by_xpath('//*[@id="makeStartTime"]').clear()
    driver.find_element_by_xpath('//*[@id="makeStartTime"]').send_keys(first_day)
    time.sleep(2)
    #结束日期
    driver.find_element_by_xpath('//*[@id="makeEndTime"]').clear()
    driver.find_element_by_xpath('//*[@id="makeEndTime"]').send_keys(thisMonth)
    time.sleep(2)
    #定位查询按钮并点击
    driver.find_element_by_xpath('//*[@id="btnSearch"]').click()
    time.sleep(random.randrange(8, 12))
    #定位导出按钮并点击
    driver.find_element_by_xpath('//*[@id="lr-derive"]').click()
    time.sleep(2)
    #退出之前的iframe
    driver.switch_to_default_content()
    # iframe2 = driver.find_element_by_xpath('//*[@id="DeriveDialog"]')
    # driver.switch_to_frame(iframe2)
    # time.sleep(2)
    # js3 = "document.getElementById('AccessView').getElementsByTagName('li')[0].removeAttribute('class')"
    # driver.execute_script(js3)
    #定位确定按钮并点击
    time.sleep(5)
    driver.find_element_by_xpath('/html/body/div[1]/table/tbody/tr[2]/td[2]/div/table/tbody/tr[3]/td/div/input[1]').click()
    time.sleep(2)
    checkDownload(folder)
    driver.quit()


# 营销云南鸿翔一心堂药业(集团)股份有限公司 -- 无法定位到导出按钮
def hr15(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    driver.maximize_window()
    time.sleep(1)
    # 点击登录按钮
    driver.find_element_by_xpath('/html/body/div/div/div[1]/div[2]/button').click()
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片 UserName PassWord
    username = driver.find_element_by_xpath('//*[@id="fullscreen"]/div[1]/div/div[1]/div/div[1]/input')
    password = driver.find_element_by_xpath('//*[@id="fullscreen"]/div[1]/div/div[1]/div/div[2]/input')
    # 防止页面错误
    actions = ActionChains(driver)
    try:
        # 输入用户名、密码、验证码
        username.clear()
        username.send_keys(pageUsername)
        password.clear()
        password.send_keys(pagePassword)
        time.sleep(random.randrange(1, 3))
        # 登录系统 //*[@id="fullscreen"]/div[1]/div/div[1]/button
        driver.find_element_by_xpath('//*[@id="fullscreen"]/div[1]/div/div[1]/button').click()
        time.sleep(3)
    except:
        print('页面弹框，正在修正...')
        driver.switch_to.alert.accept()
        driver.quit()
    time.sleep(3)
    # 点击商品流向
    driver.find_element_by_xpath('/html/body/div/aside/div[2]/ul/li[3]/ul/li[4]/div').click()
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/div/aside/div[2]/ul/li[3]/ul/li[4]/ul/li').click()
    time.sleep(random.randrange(3, 5))
    # 切换 right_frame
    right_frame = driver.find_element_by_xpath('/html/body/div/div[2]/div/div[1]/div/section/div[2]/div[2]/div/iframe')
    driver.switch_to.frame(right_frame)
    # 定位开始搜索按钮 //*[@id="fr-btn-FORMSUBMIT0"]/tbody/tr[2]/td[2]/em/button
    driver.find_element_by_xpath('//*[@id="fr-btn-FORMSUBMIT0"]/tbody/tr[2]/td[2]/em/button').click()
    time.sleep(random.randrange(5, 10))
    # 定位导出按钮  //*[@id="fr-btn-"]/tbody/tr[2]/td[2]/em/button
    driver.find_element_by_xpath('//*[@id="fr-btn-"]/tbody/tr[2]/td[2]/em/button').click()
    # 无法定位到输出 - excel 按钮
    # driver.find_element_by_xpath()
    checkDownload(folder)
    driver.quit()


# 营销福建鹭燕中宏医药有限公司-- 已完成，使用网页默认起始日期
def hr16(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("customerNum")
    password = driver.find_element_by_id("password")
    actions = ActionChains(driver)
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    time.sleep(random.randrange(1, 3))
    password.clear()
    password.send_keys(pagePassword)
    time.sleep(random.randrange(1, 3))
    # 选择区域 -- 福建
    driver.find_element_by_xpath('//*[@id="area"]').click()
    driver.find_element_by_xpath('//*[@id="area"]/option[2]').click()
    time.sleep(random.randrange(1, 3))
    # 选择公司 -- 福建鹭燕中宏医药有限公司
    driver.find_element_by_xpath('//*[@id="companys"]').click()
    driver.find_element_by_xpath('//*[@id="companys"]/option[2]').click()
    time.sleep(random.randrange(1, 3))
    # 登录系统 //*[@id="submitbtn"]
    driver.find_element_by_xpath('//*[@id="submitbtn"]').click()
    time.sleep(random.randrange(1, 3))
    # 切换到菜单frame /ws/lyzh.php/Index/menu
    menu_frame = driver.find_element_by_xpath('//frame[contains(@src,"Index/menu")]')
    driver.switch_to.frame(menu_frame)
    # 点击销售流向查询 //*[@id="items1_1"]/ul/li[3]/a
    driver.find_element_by_xpath('//*[@id="items1_1"]/ul/li[3]/a').click()
    time.sleep(random.randrange(2, 5))
    # 切换回默认frame
    driver.switch_to.default_content()
    # 切换到查询frame /ws/lyzh.php/Index/main
    main_frame = driver.find_element_by_xpath('//frame[contains(@src,"Index/main")]')
    driver.switch_to.frame(main_frame)
    time.sleep(random.randrange(2, 5))
    # 定位输入日期  -- 该网页不支持日期输入操作，仅提供点选操作
    # driver.find_element_by_id('Grid1_Toolbar2_BeginDate-inputEl').clear()
    # driver.find_element_by_id('Grid1_Toolbar2_BeginDate-inputEl').send_keys(first_day)
    # driver.find_element_by_id('Grid1_Toolbar2_EndDate-inputEl').clear()
    # driver.find_element_by_id('Grid1_Toolbar2_EndDate-inputEl').send_keys(yesterday)
    # 定位查询按钮 //*[@id="mainForm"]/div/div[4]/input[1] //*[@id="mainForm"]/div/div[4]/input[1]
    driver.find_element_by_xpath('//*[@id="mainForm"]/div/div[4]/input[1]').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮 //*[@id="doc3"]/div[2]/input[1]
    driver.find_element_by_xpath('//*[@id="doc3"]/div[2]/input[1]').click()
    time.sleep(1)
    checkDownload(folder)
    driver.quit()


# 营销福建泉南医药 -- 已完成
def hr17(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("UserName")
    password = driver.find_element_by_id("Password")
    vCode = driver.find_element_by_id("LoginSign")
    xp = '/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td[2]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[5]/td/font'
    driver.maximize_window()
    # 打码 参数，驱动器、路径、截图位置（left，top，right，bottom）
    # get_image(driver, xp, 190, 70, 10, 0)
    bbox = (190, 70, 10, 0)
    saveImg(bbox)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 防止页面错误
    try:
        # 输入用户名、密码、验证码
        username.clear()
        username.send_keys(pageUsername)
        password.clear()
        password.send_keys(pagePassword)
        vCode.clear()
        vCode.send_keys(resultCode)
        time.sleep(random.randrange(1, 3))
        # 登录系统 /html/body/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td[2]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[7]/td/input[1]
        driver.find_element_by_xpath(
            '/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table[2]/tbody/tr[1]/td[2]/table[2]/tbody/tr[2]/td[2]/table/tbody/tr[7]/td/input[1]').click()
        time.sleep(3)
    except:
        print('页面错误，正在重新输入')
        driver.quit()
    time.sleep(3)
    # 跳转到菜单frame
    menu_frame = driver.find_element_by_xpath('//frame[contains(@src,"index.asp?Do=GysMainMenu")]')
    driver.switch_to.frame(menu_frame)
    # 点击供应商网络服务 //*[@id="Flow"]
    driver.find_element_by_xpath('//*[@id="Flow"]').click()
    time.sleep(random.randrange(1, 3))
    driver.switch_to.default_content()
    # 定位到主frame./index.asp?do=blankcontenthead
    main_frame = driver.find_element_by_xpath('//frame[contains(@src,"index.asp?do=blankcontenthead")]')
    driver.switch_to.frame(main_frame)
    # 定位输入日期 /html/body/table/tbody/tr/td/table[2]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/input[1]
    path = '/html/body/table/tbody/tr/td/table[2]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/'
    driver.find_element_by_xpath(path + 'input[1]').clear()
    driver.find_element_by_xpath(path + 'input[1]').send_keys(last_month_day)
    driver.find_element_by_xpath(path + 'input[2]').clear()
    driver.find_element_by_xpath(path + 'input[2]').send_keys(yesterday)
    # 定位查询按钮 /html/body/table/tbody/tr/td/table[2]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td/input[3]
    driver.find_element_by_xpath(path + 'input[3]').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮 /html/body/table/tbody/tr/td/table[2]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/input[1]
    driver.find_element_by_xpath(
        '/html/body/table/tbody/tr/td/table[2]/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/input[1]').click()
    time.sleep(1)
    checkDownload(folder)
    driver.quit()

# 营销云南新世纪药业 -- 无导出按钮，暂未完成
def hr18(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("userName")
    vCode = driver.find_element_by_id("inputCode")
    xp = '//*[@id="checkCode"]'
    driver.maximize_window()
    # 打码 参数，驱动器、路径、截图位置（left，top，right，bottom）
    # get_image(driver, xp, 290, 94, 15, 7)
    bbox = (290, 94, 15, 7)
    saveImg(bbox)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 防止验证码错误
    try:
        # 输入用户名、密码、验证码
        username.send_keys(pageUsername)
        driver.find_element_by_id('role').click()
        vCode.send_keys(resultCode)
        time.sleep(random.randrange(1, 3))
        # 登录系统
        driver.find_element_by_id('login_link').click()
        time.sleep(3)
    except:
        print('验证码输入错误，正在重新输入')
        driver.quit()
    time.sleep(3)
    # 点击供应商网络服务 //*[@id="tree1"]/li[1]/div/div[1]
    driver.find_element_by_xpath('//*[@id="tree1"]/li[1]/div/div[1]').click()
    time.sleep(random.randrange(1, 3))
    # 点击子供应商流向查询 //*[@id="tree1"]/li[1]/ul/li[3]/div/span
    driver.find_element_by_xpath('//*[@id="tree1"]/li[1]/ul/li[3]/div/span').click()
    # 转换菜单栏的iframe flowQuery.action
    main_frame = driver.find_element_by_xpath('//iframe[contains(@src,"flowQuery")]')
    driver.switch_to.frame(main_frame)
    time.sleep(3)
    # 定位输入日期 -- 无日期选项

    # 定位查询按钮 //*[@id="submit_s"]
    driver.find_element_by_xpath('//*[@id="submit_s"]').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮 -- 暂无导出按钮
    # driver.find_element_by_xpath('//*[@id="btnExportExcel"]').click()
    time.sleep(1)
    checkDownload(folder)
    driver.quit()


#鞍山市天鸿医药有限公司 无销量   （已完成）
def hr19(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    #定位账号  密码框的位置
    userName = driver.find_element_by_xpath('//*[@id="body_TextBox1"]')
    userPassword = driver.find_element_by_xpath('//*[@id="body_TextBox2"]')
    #防止页面报错
    actions = ActionChains(driver)
    userName.clear()
    userName.send_keys(pageUsername)
    time.sleep(1)
    userPassword.clear()
    userPassword.send_keys(pagePassword)
    time.sleep(random.randrange(1, 3))
    #定位登录按钮
    driver.find_element_by_xpath('//*[@id="body_Button1"]').click()
    #休眠
    time.sleep(2)
    #定位开始日期,并重新填写值
    driver.find_element_by_xpath('//*[@id="body_TextBoxstart"]').clear()
    driver.find_element_by_xpath('//*[@id="body_TextBoxstart"]').send_keys(first_day)
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="body_TextBoxend"]').clear()
    driver.find_element_by_xpath('//*[@id="body_TextBoxend"]').send_keys(yesterday)
    time.sleep(2)
    #选择类型
    driver.find_element_by_xpath('//*[@id="body_DropDownList1"]').click()
    time.sleep(1)
    driver.find_element_by_xpath('//*[@id="body_DropDownList1"]/option[4]').click()
    #定位查询按钮（查询按钮为导出）
    time.sleep(2)
    driver.find_element_by_xpath('//*[@id="body_Button1"]').click()
    time.sleep(2)
    checkDownload(folder)
    driver.quit()
#--------------------------------------------------------------------------------------------------------------------

import pandas as pd

success_num = 0
fail_num = 0
total_num = 0
success_list = []
fail_list = []
fail_reason = []
columns = ['下载成功', '下载失败', '失败原因']
download_columns = ['客户编码', '下载时间', '下载结果', '原因']
now_time = datetime.datetime.now()
format_time = datetime.datetime(now_time.year, now_time.month, now_time.day, now_time.hour, now_time.minute).strftime(
    "%Y%m%d-%H%M")
store_path = format_time + '-'
file_name = store_path + 'download_file.csv'
result_DF = pd.DataFrame(columns=download_columns)
result_DF.to_csv(file_name, index=False, sep=',', encoding='utf-8')
download_reason = ''

# 遍历基础信息、调用各网载函数
with open('wdt.csv', newline='', encoding='utf-8') as csvfile:
    total_num += 1
    start_time = time.time()
    reader = csv.DictReader(csvfile)
    for row in reader:
        folder = row['code']
        workerID = row['cat']
        url = row['url']
        pageUsername = row['useracc']
        pagePassword = row['passwd']

        # 检测目标文件夹是否存在
        if not os.path.exists(folder):
            os.mkdir(folder)

        # 删除目标文件夹中的firefox和chrome下载专用临时文件
        apath = os.path.abspath(os.path.dirname(__file__))
        bpath = os.path.join(apath, folder)
        fl = os.listdir(bpath)
        for item in fl:
            if item.endswith('.part' or '.crdownload'):
                os.remove(os.path.join(bpath, item))
        # 防止下载中遇到问题，会循环尝试多次
        autoDownload(workerID, url, pageUsername, pagePassword, folder)
        if is_download_finished(folder) is False:
            fail_num += 1
        # globals()[workerID](url, pageUsername, pagePassword, folder)
        # try:
        #     success_num += 1
        #     globals()[workerID](url, pageUsername, pagePassword, folder)
        # # except UnexpectedAlertPresentException as alertException:
        # #     print(alertException, ">>验证码错误", folder, "--", workerID, ">>>正在重新尝试...")
        # #     globals()[workerID](url, pageUsername, pagePassword, folder)
        # except Exception as e:
        #     print(e, ">>异常", folder, "--", workerID, ">>>下载失败")
        #     fail_num += 1
        #     fail_list.append(folder)
        #     fail_reason.append(e)
        #     pass
        # success_list.append(folder)
    # data_frame = {'下载失败': fail_list, '失败原因': fail_reason}
    # finall_df = pd.DataFrame(data_frame)
    # result_path = store_path + str(total_num - fail_num) + '成功' + str(fail_num) + '失败.csv'
    # finall_df.to_csv(result_path, index=False, sep=',', encoding='utf-8')
print('总计：', total_num, '\n失败数量：', fail_num, '\n成功数量：', (total_num - fail_num))


