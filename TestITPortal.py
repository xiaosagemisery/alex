# coding : utf-8
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import openpyxl


def get_snow_info(userid):
    driver = webdriver.Chrome()
    Task_Address = "https://xxxxxxxxxx.com/it/?id=mdtit_sc_cat_item&sys_id=51e8e6c1db022340f05f69c3ca961954"
    driver.get(Task_Address)
    driver.maximize_window()#全屏
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('//*[@id="s2id_sp_formfield_requested_for"]/a/span[2]/b'))
    driver.find_element_by_xpath('//*[@id="s2id_sp_formfield_requested_for"]/a/span[2]/b').click()
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('/html/body/div[3]/div/input'))
    driver.find_element_by_xpath('/html/body/div[3]/div/input').send_keys(userid)
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('/html/body/div[3]/ul/li/div/div[1]'))
    ullist = driver.find_elements_by_xpath('//*[@id="select2-results-2"]/li')
    print(len(ullist))
    for i in range(1,len(ullist)+1):
        useridid = driver.find_element_by_xpath('/html/body/div[3]/ul/li[' + str(i) +']/div/div[2]').text
        print(useridid)
        if useridid == userid:
            print('就是这个')
            real_user = '/html/body/div[3]/ul/li[' + str(i) +']/div'
            driver.find_element_by_xpath(real_user).click()
            break
    driver.find_element_by_xpath('//*[@id="s2id_sp_formfield_{{::field.name}}"]/a/span[2]/b').click()
    ActionChains(driver).send_keys(Keys.DOWN).perform()
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    driver.find_element_by_xpath('//*[@id="s2id_sp_formfield_who_is_the_team_owner"]/a/span[2]/b').click()
    driver.find_element_by_xpath('//*[@id="s2id_autogen9_search"]').send_keys('jinx8')
    time.sleep(10)
    ActionChains(driver).send_keys(Keys.ENTER).perform()
    driver.find_element_by_xpath('/html/body/div[1]/section/main/div/div/sp-page-row/div/div/span[2]/div/div/div[1]/div[1]/div/div[3]/form/div/sp-variable-layout/fieldset[2]/div/div/div[5]/sp-checkbox-group/fieldset/div/div[2]/div/div/div/span/span/label/input').click()
    if not driver.find_element_by_xpath('/html/body/div[1]/section/main/div/div/sp-page-row/div/div/span[2]/div/div/div[1]/div[1]/div/div[3]/form/div/sp-variable-layout/fieldset[2]/div/div/div[5]/sp-checkbox-group/fieldset/div/div[2]/div/div/div/span/span/label/input').is_selected():
        driver.find_element_by_xpath('/html/body/div[1]/section/main/div/div/sp-page-row/div/div/span[2]/div/div/div[1]/div[1]/div/div[3]/form/div/sp-variable-layout/fieldset[2]/div/div/div[5]/sp-checkbox-group/fieldset/div/div[2]/div/div/div/span/span/label/input').click()
    driver.find_element_by_xpath('//*[@id="sc_cat_item"]/div[1]/div[2]/div/div[1]/div[4]/button').click()
    WebDriverWait(driver, 160).until(lambda x: x.find_element_by_xpath('//*[@id="x0109cd6013e7d300a4f83a128144b049"]/div/div[1]/div/h1'))
    if driver.find_element_by_xpath('//*[@id="x0109cd6013e7d300a4f83a128144b049"]/div/div[1]/div/h1').text == '您的申请已经提交！':
        print('申请成功')
        time.sleep(2)
        driver.close()
    else:
        print('任务出错')



if __name__ == '__main__':
    wb = openpyxl.load_workbook('C:\win10input\EE Database for IT (version 1).xlsx')
    sheets = wb.sheetnames
    print(sheets, type(sheets))
    work_sheeet = wb.worksheets[1]
    print(work_sheeet['D'])
    ee = work_sheeet['D']
    for i in ee:
        useridname = i.value
        print(useridname)
        get_snow_info(useridname)





