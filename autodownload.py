import datetime
from selenium import webdriver
import os
# import path
from pathlib import Path
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC, ui
from selenium.webdriver.support.wait import WebDriverWait
from PIL import Image
import time
from yundama import YDMHttp
import csv
import random

# 当天日期、昨天日期
today = datetime.datetime.now()
# today = '2019-05-01'
yesterday = (today - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
# yesterday = '2019-06-01'
# 获取当月第一天日期
first_day = datetime.datetime(today.year, today.month - 1, 1).strftime("%Y-%m-%d")
# 获取上月的当天日期
last_month_day = datetime.datetime(today.year, today.month - 1, today.day).strftime("%Y-%m-%d")
# first_day = '2019-05-01'
print(last_month_day)
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


### yundama -------------------------------------

def get_snap(driver):  # 对目标网页进行截屏。这里截的是全屏
    driver.save_screenshot('full_snap.png')
    page_snap_obj = Image.open(r'full_snap.png')
    return page_snap_obj


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


# 营销华润 -- 已完成（库存、销售）
def hr01(url, pageUsername, pagePassword, folder):
    # 获取驱动
    driver = getDriver(folder)
    driver.get(url)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("txtname")
    password = driver.find_element_by_id("txtpwd")
    vCode = driver.find_element_by_id("txtcheckcode")
    xp = "//*[@id='UpdatePanel1']/ul/li[4]/img"
    # 打码
    get_image(driver, xp, 179, 80, 10, 7)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 用户名、密码、验证码赋值
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    vCode.send_keys(resultCode)
    # 登录系统
    driver.find_element_by_name("ImgSubmit").click()
    time.sleep(3)
    driver.find_element_by_link_text("流向查询").click()
    time.sleep(3)
    # 库存流向查询、销售流向查询
    getSelect(driver, "库存查询", 'storelist', first_day, yesterday)
    driver.switch_to.default_content()
    getSelect(driver, "销售查询", 'salelist', first_day, yesterday)

    input('Press ENTER to close the automated browser')
    driver.quit()


# 流向查询（库存、销售）
def getSelect(driver, select_name, src_name, start_time, end_time):
    driver.find_element_by_link_text(select_name).click()
    time.sleep(3)
    kcFrame = ''
    if src_name == 'storelist':
        kcFrame = driver.find_element_by_xpath("//iframe[contains(@src,'storelist')]")
    #     DFS/work/salelist.html
    if src_name == 'salelist':
        kcFrame = driver.find_element_by_xpath("//iframe[contains(@src,'salelist')]")
    driver.switch_to.frame(kcFrame)
    # /html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span/input[1]
    if src_name == 'storelist':
        driver.find_element_by_xpath("/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]").clear()
        time.sleep(1)
        driver.find_element_by_xpath("/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]").send_keys(
            end_time)
        time.sleep(1)
    if src_name == 'salelist':
        driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]').clear()
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[1]/input[1]').send_keys(
            start_time)
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[2]/input[1]').clear()
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/div[1]/div[2]/table/tbody/tr[1]/td[2]/span[2]/input[1]').send_keys(
            end_time)
        time.sleep(1)
        # driver.find_element_by_xpath('//*[@id="batchNo"]').click()
    time.sleep(1)
    # 防止查询按钮点击不到
    driver.find_element_by_id('batchNo').click()
    # //*[@id="search"]/span/span
    driver.find_element_by_link_text("查询").click()
    time.sleep(1)
    driver.find_element_by_link_text("导出").click()
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


# 老百姓贝林网载获取 -- 异常，部分按钮点击失败
def hr02(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    # 找到用户名、密码
    username = driver.find_element_by_name("userId")
    password = driver.find_element_by_name("password")
    time.sleep(2)
    # 给用户名、密码赋值
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    # 登录系统
    driver.find_element_by_xpath("//input[@value='登 录']").click()
    time.sleep(3)
    # 找到销售流向
    driver.find_element_by_link_text("销售管理").click()
    time.sleep(3)
    driver.find_element_by_link_text("配送/物流中心出库/返仓").click()
    time.sleep(3)

    # 跳转iframe
    mainframe = driver.find_element_by_id("mainframe")  # //*[@id="mainframe"]
    driver.switch_to.frame(mainframe)
    time.sleep(1)

    # 定位起始时间 input
    driver.find_element_by_xpath('//*[@id="form"]/span[1]/span[1]/input').send_keys(first_day)
    time.sleep(1)

    # 定位结束时间 input
    driver.find_element_by_xpath("//*[@id='form']/span[2]/span[1]/input").send_keys(yesterday)
    time.sleep(3)

    # 定位商品窗口
    actions = ActionChains(driver)
    goods_span = driver.find_element_by_xpath('//*[@id="input_goods_id"]/span/span/span[2]/span')
    actions.move_to_element(to_element=goods_span).click().perform()
    # document.querySelector("#input_goods_id\\$value")  document.querySelector("#input_goods_id\\$text")
    # js = 'document.querySelector("#input_goods_id$value") = "1074091";'

    # js = """function getQueryJson(){
    #         data=form.getData(true,true);
    #         data.param.BEGINDATE="2019-06-18";
    #         data.param.ENDDATE="2019-06-19";
    #         data.param.GOODSID="1074091";
    #         data.param.DEPTID="611";
    #         return data
    #     }"""
    # driver.execute_script(js)
    time.sleep(1)

    # 切换iframe //*[@id="mini-17"]/div/div[2]/div[2]/iframe /srm/sale/MerchandiseQuery.jsp?entity=org.gocom.components.coframe.org.dataset.OrgOrganization&_winid=w2419&_t=343997
    # wait = ui.WebDriverWait(driver, 10)
    # check_box_iframe = wait.until(lambda driver: driver.find_element_by_xpath("//iframe[contains(@src,'/srm/sale/MerchandiseQuery')]"))
    check_box_iframe = driver.find_element_by_xpath('//div[@id="mini-17"]/div/div[2]/div[2]/iframe')
    driver.switch_to.frame(check_box_iframe)
    time.sleep(2)
    # 选定“全选”框，点击   '//*[@id="mini-17$headerCell2$2"]/div/div[1]/input'
    box_element = driver.find_element_by_xpath("//*[@id='mini-17checkall']")
    actions.move_to_element(box_element).click().perform()
    time.sleep(1)

    # 选定“确认”按钮，点击
    submit_button_element = driver.find_element_by_xpath('/html/body/a[1]/span')
    actions.move_to_element(submit_button_element).click().perform()
    # driver.find_element_by_link_text("确认").click()
    time.sleep(1)
    # 选定部门
    dept_element = driver.find_element_by_xpath('//*[@id="input_dept$text"]')
    actions.move_to_element(dept_element).click().perform()
    time.sleep(1)
    dept_element_select = driver.find_element_by_xpath('//*[@id="mini-10$1"]/td[2]')
    actions.move_to_element(dept_element_select).click().perform()
    time.sleep(1)
    # 选定查询按钮并点击
    select_button = driver.find_element_by_xpath('//*[@id="form"]/a[1]/span')
    actions.move_to_element(select_button).click().perform()
    # driver.find_element_by_link_text("查询").click()
    time.sleep(1)
    # 选定导出按钮并点击
    export_button = driver.find_element_by_xpath('//*[@id="form"]/a[3]/span')
    actions.move_to_element(export_button).click().perform()
    # driver.find_element_by_link_text("导出").click()
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

    input('Press ENTER to close the automated browser')

    driver.quit()


# 重庆 -- 页面跳转失败，点击操作异常
def hr03(url, pageUsername, pagePassword, folder):
    driver = getDriver(folder)
    driver.get(url)
    # 找到登录按钮，点击跳转到登录页面
    driver.find_element_by_id("login_logout").click()
    time.sleep(2)

    # 找到用户名、密码的输入框//*[@id="myform"]/div[1]/span
    username = driver.find_element_by_xpath('//*[@id="sitename"]')
    password = driver.find_element_by_xpath('//*[@id="myform"]/div[2]/span/input')
    vCode = driver.find_element_by_id("txtcheckCode")
    xp = '//*[@id="checkCode"]'
    time.sleep(2)
    # 打码
    get_image(driver, xp, 75, 120, 7, 7)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)

    # 配置用户名、密码
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    vCode.send_keys(resultCode)
    # 找到submit按钮 //*[@id="myform"]/div[5]/input[3]
    driver.find_element_by_xpath('//*[@id="myform"]/div[5]/input[3]').click()
    time.sleep(3)
    # 找到“上游客户服务”按钮 并点击 //*[@id="layoutContainers"]/div/div[2]/div/div[5]/section/div/div/div/ul/a[6]/li/h4
    # element = driver.find_element_by_xpath(
    #     '//*[@id="layoutContainers"]/div/div[2]/div/div[5]/section/div/div/div/ul/a[6]/li/span')
    # actions = ActionChains(driver)
    # # ------------------------------问题位置---------------------------
    # actions.move_to_element(to_element=element).click(
    #     '//*[@id="layoutContainers"]/div/div[2]/div/div[5]/section/div/div/div/ul/a[6]').perform()
    # driver.find_element_by_link_text('上游客户服务').click()
    # time.sleep(3)
    # //*[@id="ctl00_ContentPlaceHolder1_db__cbfgs_I"]  //*[@id="layoutContainers"]/div/div[2]/div/div[5]/section/div/div/div/ul/a[6]/li
    jump_element = driver.find_element_by_xpath(
        '//*[@id="layoutContainers"]/div/div[2]/div/div[5]/section/div/div/div/ul/a[6]/li')
    actions = ActionChains(driver)
    actions.move_to_element(jump_element).click().perform()
    # 页面跳转
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(3)

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

    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销广东康美 -- 已完成
def hr04(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    # 找到登录按钮
    driver.find_element_by_xpath('//*[@id="Map"]/area[5]').click()
    time.sleep(2)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("txtUserName")
    password = driver.find_element_by_id("txtPwd")
    vCode = driver.find_element_by_id("txtCode")
    xp = '//*[@id="imgValidate"]'
    try:
        setInfo(driver, xp, username, password, vCode, pageUsername, pagePassword)
    except:
        print('验证码输入错误，正在重新输入')
        setInfo(driver, xp, username, password, vCode, pageUsername, pagePassword)

    # 登录系统
    driver.find_element_by_name("ImageButton1").click()
    time.sleep(3)
    # 找到查询报表位置
    driver.find_element_by_xpath('//*[@id="left_cxbb"]').click()
    time.sleep(3)
    # 定位销售流向位置
    driver.find_element_by_xpath('//*[@id="left_li_xslxcx"]').click()
    time.sleep(3)
    # 设定起始时间
    driver.find_element_by_xpath('//*[@id="txtBeginTime"]').clear()
    driver.find_element_by_xpath('//*[@id="txtBeginTime"]').send_keys(first_day)
    # 设定结束时间
    driver.find_element_by_xpath('//*[@id="txtEndTime"]').clear()
    driver.find_element_by_xpath('//*[@id="txtEndTime"]').send_keys(yesterday)
    # 定位查询按钮
    driver.find_element_by_xpath('//*[@id="Button2"]').click()
    time.sleep(1)
    # 定位导出按钮
    driver.find_element_by_xpath('//*[@id="btnSearch"]').click()
    time.sleep(1)
    actions = ActionChains(driver)
    try:
        actions.send_keys(Keys.ENTER).perform()
    except:
        print("页面弹框，属正常现象，程序继续运行。")
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


def setInfo(driver, xp, username, password, vCode, pageUsername, pagePassword):
    # 打码 参数，驱动器、路径、截图位置（left，top，right，bottom）
    get_image(driver, xp, 103, 73, 10, 7)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 输入用户名、密码、验证码
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    vCode.send_keys(resultCode)


# 营销河南东森医药-- 已完成
def hr05(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("txtUserName")
    password = driver.find_element_by_id("txtPassword")
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    # 登录系统
    driver.find_element_by_link_text('登录').click()
    time.sleep(3)
    # 找到流向查询位置
    driver.find_element_by_link_text('流向查询').click()
    time.sleep(3)
    salelist = driver.find_element_by_xpath('//iframe[contains(@src,"DataFlowDetails")]')
    driver.switch_to.frame(salelist)
    # 设定起始时间
    driver.find_element_by_xpath('//*[@id="txtBeginDate"]').clear()
    driver.find_element_by_xpath('//*[@id="txtBeginDate"]').send_keys(first_day)
    # 设定结束时间
    driver.find_element_by_xpath('//*[@id="txtEndDate"]').clear()
    driver.find_element_by_xpath('//*[@id="txtEndDate"]').send_keys(yesterday)
    # 定位查询按钮
    driver.find_element_by_link_text('查询').click()
    time.sleep(1)
    # 定位导出按钮
    driver.find_element_by_link_text('导出').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销华润 南阳、周口 --- 修改日期无法实现
def hr06(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    # 定位登录窗口
    driver.find_element_by_link_text('登录').click()
    # 定位用户名、密码、验证码输入框、验证码图片 //*[@id="userName"] //*[@id="password"]
    username = driver.find_element_by_id("userName")
    password = driver.find_element_by_id("password")
    # 输入用户名、密码
    username.clear()
    username.send_keys(pageUsername)
    password.clear()
    password.send_keys(pagePassword)
    # 登录系统 //*[@id="btn_sub"]
    driver.find_element_by_id('btn_sub').click()
    time.sleep(3)
    # 点击进入办公室 //*[@id="userStateInfo"]/li[6]/a
    driver.find_element_by_link_text('进入办公室').click()
    time.sleep(3)
    # 页面跳转
    # switch_window = driver.current_window_handle
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(1)
    # 找到报表查询位置 //*[@id="sideMenu"]/li[1]
    actions = ActionChains(driver)
    report_element = driver.find_element_by_xpath('//*[@id="sideMenu"]/li[1]')
    actions.move_to_element(report_element).perform()
    time.sleep(1)
    # 点击销售流向查询 //*[@id="sideMenu"]/li[1]/div[2]/div/div[4]/a
    driver.find_element_by_xpath('//*[@id="sideMenu"]/li[1]/div[2]/div/div[4]/a').click()
    time.sleep(3)
    # 设定起始时间 //*[@id="salesFlowQuery_beginDate"]
    # actions.move_to_element('//*[@id="ext-gen123"]').click().perform()
    # actions.move_to_element('//*[@id="ext-gen124"]').click().perform()
    # actions.move_to_element('//*[@id="ext-gen451"]').click().perform()
    # actions.move_to_element('//*[@id="ext-gen450"]/tbody/tr[2]/td/table/tbody/tr[1]/td[4]/a').click().perform()
    # driver.find_element_by_xpath('//*[@id="salesFlowQuery_beginDate"]').clear()
    # driver.find_element_by_xpath('//*[@id="salesFlowQuery_beginDate"]').send_keys(first_day)
    # 设定结束时间 //*[@id="salesFlowQuery_endDate"]
    # driver.find_element_by_xpath('//*[@id="salesFlowQuery_endDate"]').clear()
    # driver.find_element_by_xpath('//*[@id="salesFlowQuery_endDate"]').send_keys(yesterday)
    # 选定经销商 //*[@id="SalesFlowQuery_spmch"] //*[@id="ext-gen578"]
    # driver.find_element_by_xpath('//*[@id="SalesFlowQuery_spmch"]').click()
    # dept_element = driver.find_element_by_xpath('//*[@id="ext-gen578"]/div[1]/span[1]')
    # actions.move_to_element(dept_element).click().perform()
    # driver.find_element_by_link_text('华润南阳医药有限公司(1)').click()
    # actions.move_to_element('//*[@id="ext-gen578"]/div[9]').perform()
    # driver.find_element_by_link_text('华润周口医药有限公司(7)').click()

    # 定位查询按钮
    # driver.find_element_by_link_text('查询').click()
    select_element = driver.find_element_by_xpath('//*[@id="ext-comp-1019"]/tbody/tr/td[2]')
    actions.move_to_element(select_element).click().perform()
    # driver.find_element_by_xpath('//*[@id="ext-gen218"]').click()
    # 由于系统本身查询时间较长，故此方法的睡眠时间定位30s
    time.sleep(30)

    # 定位导出按钮
    # driver.find_element_by_link_text('导出Excel').click()
    export_element = driver.find_element_by_xpath('//*[@id="ext-comp-1020"]/tbody/tr/td[2]')
    actions.move_to_element(export_element).click().perform()
    # driver.find_element_by_xpath('//*[@id="ext-gen227"]').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销鞍山市天鸿医药 -- 已完成，但存在无销量的情况，需确认
def hr07(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    # 定位用户名、密码、验证码输入框、验证码图片 //*[@id="body_TextBox1"]
    username = driver.find_element_by_id("body_TextBox1")
    password = driver.find_element_by_id("body_TextBox2")
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    password.clear()
    password.send_keys(pagePassword)
    # 登录系统
    driver.find_element_by_id('body_Button1').click()
    time.sleep(3)
    # 设定起始时间 //*[@id="body_TextBoxstart"]
    driver.find_element_by_xpath('//*[@id="body_TextBoxstart"]').clear()
    driver.find_element_by_xpath('//*[@id="body_TextBoxstart"]').send_keys(first_day)
    # 设定结束时间 //*[@id="body_TextBoxend"]
    driver.find_element_by_xpath('//*[@id="body_TextBoxend"]').clear()
    driver.find_element_by_xpath('//*[@id="body_TextBoxend"]').send_keys(yesterday)
    # 定位导出类型
    driver.find_element_by_xpath('//*[@id="body_DropDownList1"]').click()
    time.sleep(1)
    # 选定导出销售
    driver.find_element_by_xpath('//*[@id="body_DropDownList1"]/option[4]').click()
    time.sleep(1)
    # 点击查询、自动导出
    driver.find_element_by_xpath('//*[@id="body_Button1"]').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销浙江嘉兴百仁医药 -- 需要IE浏览器，目前无法打开
def hr08(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getIEDriver(folder)
    driver.get(url)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("tbUserName")
    password = driver.find_element_by_id("tbPassword")
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    password.clear()
    password.send_keys(pagePassword)
    # 登录系统
    driver.find_element_by_link_text('登录').click()
    time.sleep(3)
    # 找到流向查询位置
    driver.find_element_by_link_text('流向查询').click()
    time.sleep(3)
    salelist = driver.find_element_by_xpath('//iframe[contains(@src,"DataFlowDetails")]')
    driver.switch_to.frame(salelist)
    # 设定起始时间
    driver.find_element_by_xpath('//*[@id="txtBeginDate"]').clear()
    driver.find_element_by_xpath('//*[@id="txtBeginDate"]').send_keys(first_day)
    # 设定结束时间
    driver.find_element_by_xpath('//*[@id="txtEndDate"]').clear()
    driver.find_element_by_xpath('//*[@id="txtEndDate"]').send_keys(yesterday)
    # 定位查询按钮
    driver.find_element_by_link_text('查询').click()
    time.sleep(1)
    # 定位导出按钮
    driver.find_element_by_link_text('导出').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销辽宁成大方圆医药 -- 已完成
def hr09(url, pageUsername, pagePassword, folder):
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
    get_image(driver, xp, 190, 107, 10, 7)
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
    time.sleep(30)
    # 输入回车
    actions = ActionChains(driver)
    try:
        actions.send_keys(Keys.ENTER).perform()
    except:
        print("页面弹框，属正常现象，程序继续运行。")
    # 定位导出按钮 /html/body/div/div/div[3]/input[2]
    driver.find_element_by_xpath('/html/body/div/div/div[3]/input[2]').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销云南省医药保山药品 -- 已完成
def hr10(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 点击登录，跳转到登录页面 //*[@id="container"]/div/div/a
    driver.find_element_by_xpath('//*[@id="container"]/div/div/a').click()
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    code_id = driver.find_element_by_xpath('//*[@id="tlxkh"]')
    username = driver.find_element_by_id("tusername")
    password = driver.find_element_by_id("tpass")
    vCode = driver.find_element_by_id("tyzm")
    xp = '//*[@id="flogin"]/div[7]/img'
    driver.maximize_window()
    # 打码 参数，驱动器、路径、截图位置（left，top，right，bottom）
    get_image(driver, xp, 300, 107, 15, 7)
    cid, resultCode = yundama.decode(filename, codetype, timeout)
    print(resultCode)
    # 输入用户名、密码、验证码
    code_id.send_keys(pageUsername)
    username.send_keys(pageUsername)
    password.send_keys(pagePassword)
    vCode.send_keys(resultCode)
    try:
        # 登录系统
        driver.find_element_by_id('btnlogin').click()
    except:
        print('验证码输入错误，正在重新输入')
        driver.close()
        hr10(url, pageUsername, pagePassword, folder)
    time.sleep(3)
    # 点击子公司管理 //*[@id="form_main"]/div[3]/aside/section/ul/li[4]/a/span[1]
    driver.find_element_by_xpath('//*[@id="form_main"]/div[3]/aside/section/ul/li[4]/a/span[1]').click()
    time.sleep(random.randrange(1, 3))
    # 点击子公司销售流向查询
    driver.find_element_by_xpath('//*[@id="form_main"]/div[3]/aside/section/ul/li[4]/ul/li[2]').click()
    # 转换菜单栏的iframe //*[@id="iframepage"]
    main_frame = driver.find_element_by_xpath('//*[@id="iframepage"]')
    driver.switch_to.frame(main_frame)
    time.sleep(3)
    # 定位输入日期 -- 暂未实现
    # driver.find_element_by_id('edit3').click()
    # time.sleep(1)
    # driver.find_element_by_xpath('/html/body/div[2]/div[1]/table/thead/tr[2]/th[3]').click()
    # time.sleep(1)

    # 定位查询按钮 /html/body/div/section[2]/div[1]/div[2]/button[1]
    driver.find_element_by_xpath('/html/body/div/section[2]/div[1]/div[2]/button[1]').click()
    time.sleep(30)
    # 定位导出按钮 /html/body/div/div/div[3]/input[2]
    driver.find_element_by_xpath('//*[@id="btnExportExcel"]').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销云南新世纪药业 -- 无导出按钮，暂未完成
def hr11(url, pageUsername, pagePassword, folder):
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
    get_image(driver, xp, 290, 94, 15, 7)
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
        hr11(url, pageUsername, pagePassword, folder)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销徐州恩华统一医药 -- 已完成
def hr12(url, pageUsername, pagePassword, folder):
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
        # driver.quit()
        # hr12(url, pageUsername, pagePassword, folder)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销福建泉南医药 -- 已完成
def hr13(url, pageUsername, pagePassword, folder):
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
    get_image(driver, xp, 190, 70, 10, 0)
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
        # driver.quit()
        # hr12(url, pageUsername, pagePassword, folder)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销国药集团临汾 -- 已完成
def hr14(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
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
    main_frame = driver.find_element_by_xpath('//iframe[contains(@src,"Design_index?gnbh=1001&gnmch")]')
    driver.switch_to.frame(main_frame)
    # 定位输入日期 //*[@id="date_1"]  //*[@id="date_2"]
    driver.find_element_by_xpath('//*[@id="date_1"]').clear()
    driver.find_element_by_xpath('//*[@id="date_1"]').send_keys(first_day)
    driver.find_element_by_xpath('//*[@id="date_2"]').clear()
    driver.find_element_by_xpath('//*[@id="date_2"]').send_keys(yesterday)
    # driver.find_element_by_xpath('//*[@id="n1"]').click()
    # 定位查询按钮 //*[@id="stageContainer"]/div[8]/ul/li[1]/a
    driver.find_element_by_xpath('//*[@id="stageContainer"]/div[8]/ul/li[1]/a').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮 //*[@id="div_12"]/div[1]/div[1]/div/div/button
    driver.find_element_by_xpath('//*[@id="div_12"]/div[1]/div[1]/div/div/button').click()
    # 选择csv方式导出 //*[@id="div_12"]/div[1]/div[1]/div/div/ul/li[3]/a
    driver.find_element_by_xpath('//*[@id="div_12"]/div[1]/div[1]/div/div/ul/li[3]/a').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销湖南省瑞格医药 -- 已完成
def hr15(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("txtname")
    password = driver.find_element_by_id("txtpwd")
    actions = ActionChains(driver)
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    password.clear()
    password.send_keys(pagePassword)
    time.sleep(random.randrange(1, 3))
    # 登录系统 //*[@id="ImgSubmit"]
    driver.find_element_by_xpath('//*[@id="ImgSubmit"]').click()
    time.sleep(random.randrange(1, 3))
    # 点击流向管理 //*[@id="nav"]/div/div[2]/ul/li/ul/li[1]/div/a
    driver.find_element_by_xpath('//*[@id="nav"]/div/div[2]/ul/li/ul/li[1]/div/a').click()
    time.sleep(random.randrange(1, 3))
    # 切换frame DFS/work/salelist.html
    main_frame = driver.find_element_by_xpath('//iframe[contains(@src,"salelist")]')
    driver.switch_to.frame(main_frame)
    try:
        driver.find_element_by_xpath('//*[@id="jpwClose"]').click()
    except:
        print("虚拟窗口已关闭，任务继续执行...")
    # 定位输入日期 //*[@id="_easyui_textbox_input1"]  //*[@id="_easyui_textbox_input2"]
    driver.find_element_by_xpath('//*[@id="_easyui_textbox_input1"]').clear()
    driver.find_element_by_xpath('//*[@id="_easyui_textbox_input1"]').send_keys(first_day)
    driver.find_element_by_xpath('//*[@id="_easyui_textbox_input2"]').clear()
    driver.find_element_by_xpath('//*[@id="_easyui_textbox_input2"]').send_keys(yesterday)
    # 防止遮挡查询按钮 //*[@id="tools-menu-a"]/span/span[1]
    batch_element = driver.find_element_by_xpath('//*[@id="batchNo"]')
    actions.move_to_element(batch_element).click().perform()
    # 定位查询按钮 //*[@id="search"]/span
    driver.find_element_by_xpath('//*[@id="search"]/span').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮 //*[@id="mainPanle"]/div/div/div[1]/table/tbody/tr/td[3]/a/span
    driver.find_element_by_xpath('//*[@id="mainPanle"]/div/div/div[1]/table/tbody/tr/td[3]/a/span').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销无锡汇华强盛医药 -- 已完成
def hr16(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("userNameTextBox")
    password = driver.find_element_by_id("userPassTextBox")
    actions = ActionChains(driver)
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    password.clear()
    password.send_keys(pagePassword)
    time.sleep(random.randrange(1, 3))
    # 登录系统 //*[@id="loginImageButton"]
    driver.find_element_by_xpath('//*[@id="loginImageButton"]').click()
    time.sleep(random.randrange(1, 3))
    # 点击销售网络 //*[@id="headMap"]/area[5]
    driver.find_element_by_xpath('//*[@id="headMap"]/area[5]').click()
    time.sleep(random.randrange(1, 3))
    driver.find_element_by_xpath('//*[@id="subMenuTable_20"]/tbody/tr[2]/td').click()
    time.sleep(random.randrange(1, 3))
    # 定位输入日期 //*[@id="searchDateStartTextBox"]  //*[@id="searchDateEndTextBox"]
    driver.find_element_by_xpath('//*[@id="searchDateStartTextBox"]').clear()
    driver.find_element_by_xpath('//*[@id="searchDateStartTextBox"]').send_keys(first_day)
    driver.find_element_by_xpath('//*[@id="searchDateEndTextBox"]').clear()
    driver.find_element_by_xpath('//*[@id="searchDateEndTextBox"]').send_keys(yesterday)
    # 定位查询按钮 //*[@id="searchButton"]
    driver.find_element_by_xpath('//*[@id="searchButton"]').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮 //*[@id="exprotImageButton"]
    driver.find_element_by_xpath('//*[@id="exprotImageButton"]').click()
    time.sleep(1)
    # 页面跳转
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(random.randrange(8, 15))
    # 定位导出按钮 //*[@id="exportLinkButton"]
    driver.find_element_by_xpath('//*[@id="exportLinkButton"]').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销无锡汇华强盛医药 -- 已完成
def hr17(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("LoginID")
    password = driver.find_element_by_id("Pwd")
    actions = ActionChains(driver)
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    password.clear()
    password.send_keys(pagePassword)
    time.sleep(random.randrange(1, 3))
    # 登录系统 //*[@id="btnSubmit"]
    driver.find_element_by_xpath('//*[@id="btnSubmit"]').click()
    time.sleep(random.randrange(1, 3))
    # 点击流向查询 //*[@id="MainMenu_3"]
    driver.find_element_by_xpath('//*[@id="MainMenu_3"]').click()
    time.sleep(random.randrange(2, 5))
    main_frame = driver.find_element_by_xpath('//*[@id="Main"]')
    driver.switch_to.frame(main_frame)
    # 定位输入日期 //*[@id="StartDate"]  //*[@id="EndDate"]
    driver.find_element_by_xpath('//*[@id="StartDate"]').clear()
    driver.find_element_by_xpath('//*[@id="StartDate"]').send_keys(first_day)
    driver.find_element_by_xpath('//*[@id="EndDate"]').clear()
    driver.find_element_by_xpath('//*[@id="EndDate"]').send_keys(yesterday)
    # 定位查询按钮 //*[@id="btnSearch"]
    driver.find_element_by_xpath('//*[@id="btnSearch"]').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮 //*[@id="btnExcel"]
    driver.find_element_by_xpath('//*[@id="btnExcel"]').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销上海上药雷允上医药 -- 已完成部分功能，起止日期无法写入，按照默认设置，产品选择可以实现
def hr18(url, pageUsername, pagePassword, folder):
    # 获取驱动器
    driver = getDriver(folder)
    driver.get(url)
    time.sleep(1)
    # 定位用户名、密码、验证码输入框、验证码图片
    username = driver.find_element_by_id("Login1_UserName")
    password = driver.find_element_by_id("Login1_Password")
    actions = ActionChains(driver)
    # 输入用户名、密码、验证码
    username.clear()
    username.send_keys(pageUsername)
    password.clear()
    password.send_keys(pagePassword)
    time.sleep(random.randrange(1, 3))
    # 登录系统 //*[@id="Login1_LoginButton"]
    driver.find_element_by_xpath('//*[@id="Login1_LoginButton"]').click()
    time.sleep(random.randrange(1, 3))
    # 点击流向查询 //*[@id="ctl00_Menu1"]/span[2]/a
    driver.find_element_by_xpath('//*[@id="ctl00_Menu1"]/span[2]/a').click()
    time.sleep(random.randrange(2, 5))
    # 点击查询历史销售 //*[@id="ctl00_Menu1"]/span[4]/a
    driver.find_element_by_xpath('//*[@id="ctl00_Menu1"]/span[4]/a').click()
    time.sleep(random.randrange(2, 5))
    # 点击查询历史销售 //*[@id="ctl00_ContentPlaceHolder1_DropDownListShangPin"]
    driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_DropDownListShangPin"]').click()
    driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_DropDownListShangPin"]/option[2]').click()
    time.sleep(random.randrange(2, 5))
    # 定位输入日期 //*[@id="StartDate"]  //*[@id="EndDate"]
    # driver.find_element_by_xpath('//*[@id="StartDate"]').clear()
    # driver.find_element_by_xpath('//*[@id="StartDate"]').send_keys(first_day)
    # driver.find_element_by_xpath('//*[@id="EndDate"]').clear()
    # driver.find_element_by_xpath('//*[@id="EndDate"]').send_keys(yesterday)
    # 定位查询按钮 //*[@id="ctl00_ContentPlaceHolder1_ButtonSearch"]
    driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ButtonSearch"]').click()
    time.sleep(random.randrange(1, 5))
    # 定位导出按钮 //*[@id="ctl00_ContentPlaceHolder1_ButtonAllToExcel"]
    driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ButtonAllToExcel"]').click()
    time.sleep(1)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 营销浦江英特药业 --
def hr19(url, pageUsername, pagePassword, folder):
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
    get_image(driver, xp, 185, 2, 10, 0)
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
        actions.send_keys(Keys.ENTER).perform()
        driver.quit()
        # hr12(url, pageUsername, pagePassword, folder)
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
    input('Press ENTER to close the automated browser')
    driver.quit()


# 遍历基础信息、调用各网载函数
with open('wdt.csv', newline='', encoding='utf-8') as csvfile:
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

        globals()[workerID](url, pageUsername, pagePassword, folder)
