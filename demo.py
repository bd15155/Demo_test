#coding=utf-8
import os,sys
from selenium import webdriver
import unittest,time
import HTMLTestRunner 
from config import setting
import xlsxwriter
import xlrd
class Fujitsu_Login(unittest.TestCase):
    def setUp(self):
        self.driver = webdriver.Firefox()
        self.driver.implicitly_wait(10)
        self.base_url = "https://mypage.fjsquare.jp/ja/fujitsu_hr/"

    #自動にスクリーンショットのフォルダー作成
    if not os.path.exists(setting.TEST_REPORT):os.makedirs(setting.TEST_REPORT + '/' + "screenshot")

    #テストケース01_01
    print('テスト開始')
    def test_fujitsu(self):
        driver = self.driver
        driver.get(self.base_url)
        driver.find_element_by_name("userid").send_keys("20491369")
        driver.find_element_by_name("password").send_keys("155155bd")
        picture_url=self.driver.get_screenshot_as_file('C:\\Users\\student1\\Desktop\\Demo_Fujitsu\\report\\screenshot\\login.png')
        driver.find_element_by_class_name("loginBtn").click()




    def tearDown(self):
        time.sleep(2)
        picture_time = time.strftime("%Y-%m-%d-%H_%M_%S", time.localtime(time.time()))
        picture_url=self.driver.get_screenshot_as_file('C:\\Users\\student1\\Desktop\\Demo_Fujitsu\\report\\screenshot\\homepage.png')
        self.driver.quit()

if __name__ == "__main__":

    suit = unittest.TestSuite()

    suit.addTest(Fujitsu_Login('test_fujitsu'))

#テスト報告生成
    filename = 'C:\\Users\\student1\\Desktop\\Demo_Fujitsu\\report\\result.html'
    fp = open(filename,'wb')

#テスト報告要素定義
    runner = HTMLTestRunner.HTMLTestRunner(stream=fp,
                                           title=u'富士通Mypageのログインテスト報告',
                                           description = u'テスト実施情報:')

#実行
runner.run(suit)
fp.close()

#スクリーンショット添付
book = xlsxwriter.Workbook('C:\\Users\\student1\\Desktop\\Demo_Fujitsu\\report\\エビデンス.xlsx')
sheet = book.add_worksheet('登録テスト_01_01')
sheet.insert_image('D4','C:\\Users\\student1\\Desktop\\Demo_Fujitsu\\report\\screenshot\\login.png')
sheet.insert_image('D54','C:\\Users\\student1\\Desktop\\Demo_Fujitsu\\report\\screenshot\\homepage.png')
book.close()
print('データキャプチャ終了')
print('EXCEL添付終了')