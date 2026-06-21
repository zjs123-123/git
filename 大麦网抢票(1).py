"""
知识点：
    面向对象编程
    selenium 操作浏览器
    pickle 保存和读取Cookie实现免登陆
    time 做延时操作
    os 创建文件，判断文件是否存在

第三方库：
    selenium >>> pip install selenium

开发环境：
    版 本：anaconda（python3.8.8）
    编辑器：pycharm

"""

import os  # 创建文件夹, 文件是否存在
import time  # time 计时
import pickle  # 保存和读取cookie实现免登陆的一个工具
from time import sleep
from selenium import webdriver  # 操作浏览器的工具
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.microsoft import EdgeChromiumDriverManager

"""
一. 实现免登陆
二. 抢票并且下单
"""
# 大麦网主页
damai_url = 'https://www.damai.cn/'
# 登录
login_url = 'https://passport.damai.cn/login?ru=https%3A%2F%2Fwww.damai.cn%2F'
# 抢票目标页
target_url = 'https://detail.damai.cn/item.htm?spm=a2oeg.search_category.0.0.2c473ec7RCTS1x&id=1059104921191&clicktitle=%E5%8D%95%E4%BE%9D%E7%BA%AF%E3%80%8C%E7%BA%AF%E5%A6%B9%E5%A6%B92.0%E3%80%8D2026%E5%B7%A1%E5%9B%9E%E6%BC%94%E5%94%B1%E4%BC%9A-%E4%B8%8A%E6%B5%B7%E7%AB%99'

# 超时配置（秒）
LOGIN_TIMEOUT = 120
RETRY_TIMEOUT = 300
REFRESH_INTERVAL = 2

# class Concert:
class Concert:
    # 初始化加载
    def __init__(self):
        self.status = 0  # 状态, 表示当前操作执行到了哪个步骤
        self.login_method = 1  # {0:模拟登录, 1:cookie登录}自行选择登录的方式
        self.driver = webdriver.Edge(service=Service(EdgeChromiumDriverManager().install()))  # Edge浏览器驱动

    # cookies: 登录网站时出现的 记录用户信息用的
    def set_cookies(self):
        """cookies: 登录网站时出现的 记录用户信息用的"""
        self.driver.get(damai_url)
        print('###请点击登录###')
        start = time.time()
        while self.driver.title.find('大麦网-全球演出赛事官方购票平台') != -1:
            if time.time() - start > LOGIN_TIMEOUT:
                print('###登录超时, 退出###')
                return False
            sleep(1)
        print('###请扫码登录###')
        start = time.time()
        while self.driver.title != '大麦网-全球演出赛事官方购票平台-100%正品、先付先抢、在线选座！':
            if time.time() - start > LOGIN_TIMEOUT:
                print('###扫码超时, 退出###')
                return False
            sleep(1)
        print('###扫码成功###')
        pickle.dump(self.driver.get_cookies(), open('cookies.pkl', 'wb'))
        print('###cookie保存成功###')
        self.driver.get(target_url)
        return True

    # 假如说我现在本地有 cookies.pkl 那么 直接获取
    def get_cookie(self):
        """假如说我现在本地有 cookies.pkl 那么 直接获取"""
        cookies = pickle.load(open('cookies.pkl', 'rb'))
        for cookie in cookies:
            cookie_dict = {
                'domain': '.damai.cn',  # 必须要有的, 否则就是假登录
                'name': cookie.get('name'),
                'value': cookie.get('value')
            }
            self.driver.add_cookie(cookie_dict)
        print('###载入cookie###')

    def login(self):
        """登录"""
        if self.login_method == 0:
            self.driver.get(login_url)
            print('###开始登录###')
        elif self.login_method == 1:
            # 创建文件夹, 文件是否存在
            if not os.path.exists('cookies.pkl'):
                self.set_cookies()  # 没有文件的情况下, 登录一下
            else:
                self.driver.get(target_url)  # 跳转到抢票页
                self.get_cookie()  # 并且登录

    def enter_concert(self):
        """打开浏览器"""
        print('###打开浏览器,进入大麦网###')
        # 调用登录
        self.login()  # 先登录再说
        self.driver.refresh()  # 刷新页面
        self.status = 2  # 登录成功标识
        print('###登录成功###')
        # 处理弹窗
        if self.isElementExist('/html/body/div[2]/div[2]/div/div/div[3]/div[2]'):
            self.driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div/div/div[3]/div[2]').click()

    # 二. 抢票并且下单
    def choose_ticket(self):
        """选票操作"""
        if self.status == 2:
            print('=' * 30)
            print('###开始进行日期及票价选择###')
            start = time.time()
            while self.driver.title.find("确认订单") == -1:
                if time.time() - start > RETRY_TIMEOUT:
                    print('###抢票超时, 退出###')
                    return
                try:
                    buybutton = self.driver.find_element(By.CLASS_NAME, 'buybtn').text
                    if buybutton == '提交缺货登记':
                        self.status = 2
                        self.driver.get(target_url)
                        sleep(REFRESH_INTERVAL)
                    elif buybutton == '立即预定':
                        self.driver.find_element(By.CLASS_NAME, 'buybtn').click()
                        self.status = 3
                    elif buybutton == '立即购买':
                        self.driver.find_element(By.CLASS_NAME, 'buybtn').click()
                        self.status = 4
                    elif buybutton == '选座购买':
                        self.driver.find_element(By.CLASS_NAME, 'buybtn').click()
                        self.status = 5
                except:
                    print('###没有跳转到订单结算界面###')
                    sleep(REFRESH_INTERVAL)
                title = self.driver.title
                if title == '选座购买':
                    self.choice_seats()
                elif title == '确认订单':
                    order_start = time.time()
                    while time.time() - order_start < 60:
                        print('正在加载.......')
                        if self.isElementExist('//*[@id="container"]/div/div[9]/button'):
                            self.check_order()
                            return
                        sleep(1)
                    print('###下单超时###')
                    return

    def choice_seats(self):
        """选择座位"""
        start = time.time()
        while self.driver.title == '选座购买':
            if time.time() - start > RETRY_TIMEOUT:
                print('###选座超时, 退出###')
                return
            # 等待座位图加载, 提示用户手动选座
            if self.isElementExist('//*[@id="app"]/div[2]/div[2]/div[1]/div[2]/img'):
                print('###请手动点击选择你想要的座位, 然后点击确定###')
                sleep(2)
                continue
            # 检查确认按钮是否出现
            if self.isElementExist('//*[@id="app"]/div[2]/div[2]/div[2]/button'):
                self.driver.find_element(By.XPATH, '//*[@id="app"]/div[2]/div[2]/div[2]/button').click()
                break
            sleep(1)

    def check_order(self):
        """下单操作"""
        if self.status in [3, 4, 5]:
            print('###开始确认订单###')
            time.sleep(1)
            try:
                # 默认选第一个购票人信息
                self.driver.find_element(By.XPATH, '//*[@id="container"]/div/div[2]/div[2]/div[1]/div/label').click()
            except Exception as e:
                print('###购票人信息选中失败, 自行查看元素位置###')
                print(e)
            # 最后一步提交订单
            time.sleep(0.5)  # 太快了不好, 影响加载 导致按钮点击无效
            self.driver.find_element(By.XPATH, '//*[@id="container"]/div/div[9]/button').click()
            time.sleep(20)

    def isElementExist(self, element):
        """判断元素是否存在"""
        flag = True
        browser = self.driver
        try:
            browser.find_element(By.XPATH, element)
            return flag
        except:
            flag = False
            return flag

    def finish(self):
        """抢票完成, 退出"""
        self.driver.quit()


if __name__ == '__main__':
    con = Concert()
    try:
        con.enter_concert()  # 打开浏览器
        con.choose_ticket()  # 选择座位
    except Exception as e:
        print(e)
        con.finish()

