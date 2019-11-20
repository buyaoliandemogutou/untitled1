import unittest
import paramunittest as paramunittest
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
from common import readExcel

loginData=readExcel.readExcel().excel_data_list('login.xlsx','login')
# loginData={"username":"cgp1023"}

@paramunittest.parametrized(*loginData)
class login(unittest.TestCase):
    def setParameters(self,username):
        self.username=username

    def setUp(self):
        self.driver = webdriver.Chrome(executable_path="F:\Driver\chromedriver.exe")
        # driver.maximize_window()   窗口最大化
        self.driver.get("https://vote.na.lolesports.com/en-US#welcome")
        time.sleep(1)

    def testLogin(self):
        self.uu()

    # def login(self):
    #     #print(self.username,"开始投票")
    #     # 点击login 按钮
    #     time.sleep(2)
    #     self.driver.find_element_by_css_selector("#riotbar-account > a.riotbar-anonymous-link.riotbar-account-action").click()
    #     time.sleep(2)
    #     self.driver.find_element_by_name("username").send_keys(self.username)
    #     self.driver.find_element_by_name("password").send_keys("uu123456")
    #     btn = self.driver.find_elements_by_tag_name("button")
    #     for i in btn:
    #         if i.get_attribute("type") == "submit":
    #             i.click()
    #             time.sleep(1)
    def write_file(self,L):
        try:
            f = open("d:/input.txt", "w")
            f.write(L)
            f.write('\n')
            f.close()
        except IOError:
            print("write error;")
    def uu(self):
        # 点击login 按钮
        time.sleep(3)
        self.driver.find_element_by_css_selector("#riotbar-account > a.riotbar-anonymous-link.riotbar-account-action").click()
        time.sleep(2)
        self.driver.find_element_by_name("username").send_keys(self.username)
        self.driver.find_element_by_name("password").send_keys("uu123456")
        btn = self.driver.find_elements_by_tag_name("button")
        for i in btn:
            if i.get_attribute("type") == "submit":
                i.click()
                time.sleep(1)
        #self.login()
        time.sleep(10)
        self.driver.find_element_by_css_selector("#welcome-container > div > a").click()
        time.sleep(4)
        lpl=self.driver.find_elements_by_class_name("role")
        for i in lpl:
            if i.text=="LPL":
                i.click()
                time.sleep(2)
        time.sleep(3)
        #ming
        self.driver.find_element_by_css_selector("#page-content-container > div.content-group.vote-page.active > div.page-content.vote-region-3.active > div > div > div:nth-child(74)").click()
        #uzi
        self.driver.find_element_by_css_selector("#page-content-container > div.content-group.vote-page.active > div.page-content.vote-region-3.active > div > div > div:nth-child(75)").click()
        self.driver.find_element_by_id("complete-vote").click()
        self.driver.find_element_by_css_selector("#pgwModal > div > div > div.pm-content > div > div > a").click()
        time.sleep(2)
        print(self.username, "投票成功")
        self.write_file(self.username)


    def tearDown(self):
        self.driver.quit()
        self.driver.close()

if __name__ == "__main__":
    unittest.main()
