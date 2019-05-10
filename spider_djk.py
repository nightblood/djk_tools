# enconding=utf8

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import time


class SpiderDjk(object):

    def __init__(self):
        self.__drv = ''

        self.init_spider()

    def init_spider(self):
        self.openIe()
        self.__drv.get('http://154.233.7.158/ZXT/security/loc.to?page=LOGON')
        # 前台开启浏览器模式

    def openChrome(self):
        # 加启动配置
        option = webdriver.ChromeOptions()
        option.add_argument('disable-infobars')
        # 打开chrome浏览器
        self.__drv = webdriver.Chrome(chrome_options=option)

    def openIe(self):
        # iedriver = 'D:\\app\\anaconda\\IEDriverServer.exe'  # iedriver路径
        # os.environ["webdriver.ie.driver"] = iedriver  # 设置环境变量
        # self.__drv = webdriver.Ie(iedriver)
        self.__drv = webdriver.Ie()

    def process(self, account):
        """爬取客户信息"""
        time.sleep(3)
        res = []
        try:
            self.__drv.find_element_by_id('menu_code_30').click()
            time.sleep(1)
            self.__drv.find_element_by_id('sub_page_ACBAL').click()
            time.sleep(2)
            frames = self.__drv.find_elements_by_class_name('por-PopIFrame')
            target_frame = ''
            for frame in frames:
                if frame.get_attribute('src') == '/ZXT/financequery/loc.to?page=ACBAL':
                    target_frame = frame
                    break
            if target_frame == '':
                print('未能定位iframe!!!')
                return

            self.__drv.switch_to.frame(target_frame)
            self.__drv.find_element_by_name('CARD_NBR').send_keys(account)
            self.__drv.find_element_by_class_name('ui-button-text-only').click()
            time.sleep(2)


        except Exception as e:
            print(e)
