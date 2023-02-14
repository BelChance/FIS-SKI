from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import sys
import time
import xlwt
class Li_spider:

    def __init__(self,url,save_name):
        self.url=url
        self.save_name=save_name


    def run_spider(self):
        book = xlwt.Workbook()
        sheet = book.add_sheet('two')
        row = 0
        col = 0
        # 声明浏览器
        browser = webdriver.Chrome()
        browser.get(self.url)
        wait = WebDriverWait(browser, 3)
        try:
            #XPATH
            button = browser.find_element(By.XPATH, '//*[@id="btn_showmore"]/div/div/button')
            date = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[1]'
            place = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[2]'
            nation = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[3]/span[2]'
            disc = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[4]/div/div[1]'
            gender = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[4]/div/div[2]/div/div/div'
            rank1 = '// *[ @ id = "statistics_position"] / div[2] / div[1] / div[?] / div / div / div / div / div[1]'
            rank2 = '// *[ @ id = "statistics_position"] / div[2] / div[1] / div[?] / div / div / div / div / div[2]'
            rank3 = '// *[ @ id = "statistics_position"] / div[2] / div[1] / div[?] / div / div / div / div / div[3]'
            j = int(0)
            i = 0
            while (wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_showmore"]/div/div/button')))):
                if (i % 50 == 0):
                    try:
                        press = wait.until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="btn_showmore"]/div/div/button')))
                        time.sleep(2)
                    except NoSuchElementException:
                        print('error')
                    else:
                        press.click()
                i += 1
                number = i
                number = str(number)
                # replace
                date = date.replace("?", number)
                place = place.replace("?", number)
                nation = nation.replace("?", number)
                disc = disc.replace("?", number)
                gender = gender.replace("?", number)
                rank1 = rank1.replace("?", number)
                rank2 = rank2.replace("?", number)
                rank3 = rank3.replace("?", number)
                # ele
                ele_date = browser.find_element(By.XPATH, date)
                ele_place = browser.find_element(By.XPATH, place)
                ele_nation = browser.find_element(By.XPATH, nation)
                ele_disc = browser.find_element(By.XPATH, disc)
                ele_gender = browser.find_element(By.XPATH, gender)
                ele_rank1 = browser.find_element(By.XPATH, rank1)
                ele_rank2 = browser.find_element(By.XPATH, rank2)
                ele_rank3 = browser.find_element(By.XPATH, rank3)
                # 写入
                sheet.write(j, 0, ele_date.text)
                sheet.write(j, 1, ele_place.text)
                sheet.write(j, 2, ele_nation.text)
                sheet.write(j, 3, ele_disc.text)
                sheet.write(j, 4, ele_gender.text)
                sheet.write(j, 5, ele_rank1.text)
                sheet.write(j, 6, ele_rank2.text)
                sheet.write(j, 7, ele_rank3.text)
                j = j + 1
                # 重新赋值

                date = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[1]'
                place = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[2]'
                nation = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[3]/span[2]'
                disc = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[4]/div/div[1]'
                gender = '//*[@id="statistics_position"]/div[2]/div[1]/div[?]/div/div/a[4]/div/div[2]/div/div/div'
                rank1 = '// *[ @ id = "statistics_position"] / div[2] / div[1] / div[?] / div / div / div / div / div[1]'
                rank2 = '// *[ @ id = "statistics_position"] / div[2] / div[1] / div[?] / div / div / div / div / div[2]'
                rank3 = '// *[ @ id = "statistics_position"] / div[2] / div[1] / div[?] / div / div / div / div / div[3]'
                book.save(self.save_name)
        except NoSuchElementException:
            print('No Element' + 'data: ' + ele_date.text + "  " + date)



