from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait


def idm_config(userid, phone):
    driver = webdriver.Chrome()
    driver.get('https://identity.xxxxxx.com/sigma/app/index#/home')
    driver.maximize_window()#全屏
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('//*[@id="header-wrap"]/div[1]/ip-icon'))
    driver.switch_to.default_content()
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div/div[1]'))
    driver.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div/div[1]').click()
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[1]/div/div/md-list/div/div[2]/md-list-item[1]/div/button'))
    driver.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[1]/div/div/md-list/div/div[2]/md-list-item[1]/div/button').click()
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[2]/div/div/div/div[1]/div[1]/div[1]/div/div/div[1]/div[1]/input'))
    driver.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[2]/div/div/div/div[1]/div[1]/div[1]/div/div/div[1]/div[1]/input').send_keys(userid)
    driver.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[2]/div/div/div/div[1]/div[1]/div[1]/div/div/div[2]/button[2]').click()
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[2]/div/div/div/div[2]/div[2]/div/div[2]'))
    right_click_location = ''
    right_one = ''
    for i in range(2,8):
        try:
            right_click_location = '//*[@id="mainContent"]/div/div[1]/div/div[2]/div/div/div/div[2]/div[' + str(i) + ']/div/div[2]/div[2]/span[1]'
            driver.find_element_by_xpath(right_click_location)
        except:
            pass
        else:
            rr = driver.find_element_by_xpath(right_click_location).text
            if rr.endswith('UserID: ' + userid):
                right_one = right_click_location
            else:
                pass
    driver.find_element_by_xpath(right_one).click()
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[2]/div/div/div/div/div[1]/div/div[2]/div[2]/div[1]/div[6]/div/div/div/span[2]/input'))
    driver.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[2]/div/div/div/div/div[1]/div/div[2]/div[2]/div[1]/div[6]/div/div/div/span[2]/input').clear()
    driver.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[2]/div/div/div/div/div[1]/div/div[2]/div[2]/div[1]/div[6]/div/div/div/span[2]/input').send_keys('+' + phone)
    driver.find_element_by_xpath('//*[@id="mainContent"]/div/div[1]/div/div[2]/div/div/div/div/div[2]/div[3]/button').click()
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath('//*[@id="ng-app"]/body/div[4]/md-dialog/md-dialog-actions/button'))
    driver.find_element_by_xpath('//*[@id="ng-app"]/body/div[4]/md-dialog/md-dialog-actions/button').click()


if __name__ == '__main__':
    userid = 'xxxxxxxx'
    phone = '862199999999'
    idm_config(userid, phone)
