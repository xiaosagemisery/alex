# coding : utf-8
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

def get_snow_info(num=21):
    to_do = []
    driver = webdriver.Chrome()
    Task_Address = "https://xxxxxx.service-now.com"
    driver.get(Task_Address)
    driver.maximize_window()
    WebDriverWait(driver, 60, 0.2).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'gsft_main')))
    WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath("/html/body/table/tbody/tr[1]/td/div[1]/div/div[2]/div/div/span/div/div[4]/table/tbody/tr/td/div/table/tbody/tr[1]/td[3]/a"))
    #上方按需修改
    driver.switch_to.default_content()
    driver.find_element_by_xpath("//*[@id='gsft_nav']/div/magellan-favorites-list/ul/li[11]/div/div[1]/a/div[2]/span").click()
    driver.switch_to.frame('gsft_main')
    try:
        WebDriverWait(driver, 20).until(lambda x: x.find_element_by_xpath("/html/body/div[1]/div[1]/span/div/div[5]/table/tbody/tr/td/div/table/tbody/tr[1]/td[3]/a"))
    except:
        print('无Task')
        exit()
    else:
        pass
    for i in range(1,num):
        task_location = "/html/body/div[1]/div[1]/span/div/div[5]/table/tbody/tr/td/div/table/tbody/tr[" + str(i) + "]/td[3]/a"
        try:
            zhao_task = driver.find_element_by_xpath(task_location)
            zhao_task_text = driver.find_element_by_xpath(task_location).text
            zhao_task_link = driver.find_element_by_xpath(task_location).get_attribute('href')
        except:
            pass
        else:
            print(zhao_task_text)
            print(zhao_task_link)
            driver.get(zhao_task_link)
            request_for = driver.find_element_by_xpath("//*[@id='sc_task.request_item.u_requested_for_label']").get_attribute('value')
            print(request_for)
            driver.find_element_by_xpath("//*[@id='viewr.sc_task.request_item.u_requested_for']").click()
            WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath("//*[@id='sys_readonly.sys_user.user_name']"))
            id = driver.find_element_by_xpath("//*[@id='sys_readonly.sys_user.user_name']").get_attribute('value')
            print(id)
            site_add = driver.find_element_by_xpath('//*[@id="sys_user.location_label"]').get_attribute('value')
            site = ''
            if site_add == 'CHN-31 Shanghai Dong Yu Road':
                site = 'SH'
            elif site_add == 'CHN-11 Beijing Guanghua':
                site = 'BJ'
            elif site_add == 'CHN-51 Chengdu Tianqin':
                site = 'CD'
            elif site_add == 'CHN-31 Shanghai Tianzhou':
                site = 'SH'
            elif site_add == 'CHN-44 Guangzhou 385 Tianhe':
                site = 'GZ'
            elif site_add == 'CHN-31 Shanghai Jinhai':
                site = 'SH'
            title_detail = driver.find_element_by_xpath('//*[@id="sys_readonly.sys_user.title"]').get_attribute('value')
            print(site)
            print(title_detail)
            driver.get("https://xxxxxxx.service-now.com")
            WebDriverWait(driver, 60, 0.2).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'gsft_main')))
            WebDriverWait(driver, 60).until(lambda x: x.find_element_by_xpath("/html/body/table/tbody/tr[1]/td/div[1]/div/div[2]/div/div/span/div/div[4]/table/tbody/tr/td/div/table/tbody/tr[1]/td[3]/a"))
            driver.switch_to.default_content()
            driver.find_element_by_xpath("//*[@id='gsft_nav']/div/magellan-favorites-list/ul/li[11]/div/div[1]/a/div[2]/span").click()
            driver.switch_to.frame('gsft_main')
            try:
                WebDriverWait(driver, 30).until(lambda x: x.find_element_by_xpath("/html/body/div[1]/div[1]/span/div/div[5]/table/tbody/tr/td/div/table/tbody/tr[" + str(i+1) + "]/td[3]/a"))
            except:
                print('已无更多Task')
            else:
                pass
            to_do.append((zhao_task_text, zhao_task_link, id, request_for, site, title_detail))
    return to_do

def close_imcompleted(task_address):
    driver = webdriver.Chrome()
    driver.get(task_address)
    driver.maximize_window()  # 全屏
    WebDriverWait(driver, 60, 0.2).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'gsft_main')))
    time.sleep(2)
    Select(driver.find_element_by_xpath("//*[@id='sc_task.state']")).select_by_visible_text("Closed Complete")
    driver.find_element_by_xpath("//*[@id='sys_display.sc_task.assigned_to']").send_keys("Alex Shen")
    driver.find_element_by_xpath("//*[@id='sc_task.description']").click()
    time.sleep(5)
    # result16=WebDriverWait(driver,20,0.2).until(EC.text_to_be_present_in_element_value((By.ID,'sys_display.incident.assigned_to'), 'Alex Shen'))
    driver.find_element_by_xpath("//*[@id='sc_task.work_notes']").send_keys("User already has Extension or user doesn't need one or user's site has no IPT system")
    time.sleep(5)
    driver.find_element_by_xpath("//*[@id='sysverb_update_and_stay']").click()
    result16 = EC.text_to_be_present_in_element_value((By.XPATH, "//*[@id='sys_display.sc_task.assigned_to']"), 'Alex Shen')(driver)

    if result16:
        pass
    else:
        driver.find_element_by_xpath("//*[@id='sys_display.sc_task.assigned_to']").send_keys("Alex Shen")
        time.sleep(1)
        ActionChains(driver).send_keys(Keys.DOWN).perform()
        ActionChains(driver).send_keys(Keys.ENTER).perform()


if __name__ == '__main__':
    info = get_snow_info()
    print(len(info))
    print(info)
    print(info[0][0])



