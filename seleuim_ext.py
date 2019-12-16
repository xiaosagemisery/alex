# coding : utf-8
import time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import mysql.connector
from selenium.webdriver.support.wait import WebDriverWait
from datetime import datetime
from selenium.webdriver.common.by import By
from sendemail import office365
from idm_access import idm_config
from test4mysql import check_userext
from Testservicenow import get_snow_info
from Testservicenow import close_imcompleted

def jabber_task(user_id1, user_name1, phone_number1, site1):
    driver.find_element_by_id("设备Button").click()
    time.sleep(1)
    driver.find_element_by_xpath("//*[@id='device']/ul/li[4]/a").click()
    driver.find_element_by_xpath("//*[@id='1tbllink']").click()
    Select(driver.find_element_by_xpath("//*[@id='TKMODEL']")).select_by_visible_text("Cisco Unified Client Services Framework")
    driver.find_element_by_xpath("/html/body/table/tbody/tr/td/div/form/fieldset[3]/input").click()
    driver.find_element_by_xpath("//*[@id='NAME']").send_keys("CSF"+user_id1)
    connection_symbo2 = ' - '
    driver.find_element_by_xpath("//*[@id='DESCRIPTION']").send_keys(user_name1 + connection_symbo2 + phone_number1)
    if site1=='SH':
        Select(driver.find_element_by_xpath("//*[@id='FKDEVICEPOOL']")).select_by_visible_text("P-SHA5-Phones-NoSRST")
        Select(driver.find_element_by_xpath("//*[@id='FKCOMMONDEVICECONFIG']")).select_by_visible_text("Shanghai CDC")
        Select(driver.find_element_by_xpath("//*[@id='FKPHONETEMPLATE']")).select_by_visible_text("Standard Client Services Framework")
        Select(driver.find_element_by_xpath("//*[@id='FKCALLINGSEARCHSPACE']")).select_by_visible_text("P-SHA5-Jabber-CSS")
        Select(driver.find_element_by_xpath("//*[@id='FKLOCATION']")).select_by_visible_text("P-SHA5")
    elif site1=='CD':
        Select(driver.find_element_by_xpath("//*[@id='FKDEVICEPOOL']")).select_by_visible_text("P-CTU2-Phones-NoSRST")
        Select(driver.find_element_by_xpath("//*[@id='FKCOMMONDEVICECONFIG']")).select_by_visible_text("Chengdu CDC")
        Select(driver.find_element_by_xpath("//*[@id='FKPHONETEMPLATE']")).select_by_visible_text("Standard Client Services Framework")
        Select(driver.find_element_by_xpath("//*[@id='FKCALLINGSEARCHSPACE']")).select_by_visible_text("P-CTU2-Jabber-CSS")
        Select(driver.find_element_by_xpath("//*[@id='FKLOCATION']")).select_by_visible_text("P-CTU2")
    elif site1=='BJ':
        Select(driver.find_element_by_xpath("//*[@id='FKDEVICEPOOL']")).select_by_visible_text("P-BEI1-Phones-NoSRST")
        Select(driver.find_element_by_xpath("//*[@id='FKCOMMONDEVICECONFIG']")).select_by_visible_text("Beijing Sales CDC")
        Select(driver.find_element_by_xpath("//*[@id='FKPHONETEMPLATE']")).select_by_visible_text("Standard Client Services Framework")
        Select(driver.find_element_by_xpath("//*[@id='FKCALLINGSEARCHSPACE']")).select_by_visible_text("P-BEI1-Jabber-CSS")
        Select(driver.find_element_by_xpath("//*[@id='FKLOCATION']")).select_by_visible_text("P-BEI1")
    elif site1=='GZ':
        Select(driver.find_element_by_xpath("//*[@id='FKDEVICEPOOL']")).select_by_visible_text("P-GUA1-Phones-NoSRST")
        Select(driver.find_element_by_xpath("//*[@id='FKCOMMONDEVICECONFIG']")).select_by_visible_text("Guangzhou CDC")
        Select(driver.find_element_by_xpath("//*[@id='FKPHONETEMPLATE']")).select_by_visible_text("Standard Client Services Framework")
        Select(driver.find_element_by_xpath("//*[@id='FKCALLINGSEARCHSPACE']")).select_by_visible_text("P-GUA1-Jabber-CSS")
        Select(driver.find_element_by_xpath("//*[@id='FKLOCATION']")).select_by_visible_text("P-GUA1")
    jabber_page = driver.current_window_handle
    driver.find_element_by_xpath("//*[@id='find_fkenduser']").click()
    current_window = driver.current_window_handle #定位当前页面句柄
    all_handles = driver.window_handles   
    for handle in all_handles:
        if handle != current_window:          
            driver.switch_to.window(handle)
    Select(driver.find_element_by_xpath("//*[@id='searchField0']")).select_by_visible_text("用户 ID")
    Select(driver.find_element_by_xpath("//*[@id='searchLimit0']")).select_by_visible_text("精确等于")
    driver.find_element_by_xpath("//*[@id='searchString0']").clear()
    driver.find_element_by_xpath("//*[@id='searchString0']").send_keys(user_id1)
    driver.find_element_by_xpath("//*[@id='filterRow0']/td[7]/input").click()
    driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/table[2]/tbody/tr[2]/td[1]/input[1]").click()
    driver.find_element_by_xpath("//*[@id='7tbl']").click()
    driver.switch_to.window(jabber_page)
    Select(driver.find_element_by_xpath("/html/body/table/tbody/tr/td/div/form/table/tbody/tr/td[2]/fieldset[4]/table/tbody/tr[6]/td[2]/select")).select_by_index(2)
    if site=='SH':
        Select(driver.find_element_by_xpath("//*[@id='FKCALLINGSEARCHSPACE_REROUTE']")).select_by_visible_text("P-SHA5-Jabber-CSS")
    elif site=='CD':
        Select(driver.find_element_by_xpath("//*[@id='FKCALLINGSEARCHSPACE_REROUTE']")).select_by_visible_text("P-CTU2-Jabber-CSS")
    elif site=='BJ':
        Select(driver.find_element_by_xpath("//*[@id='FKCALLINGSEARCHSPACE_REROUTE']")).select_by_visible_text("P-BEI1-Jabber-CSS")
    elif site=='GZ':
        Select(driver.find_element_by_xpath("//*[@id='FKCALLINGSEARCHSPACE_REROUTE']")).select_by_visible_text("P-GUA1-Jabber-CSS")
    Select(driver.find_element_by_xpath("//*[@id='FKSIPPROFILE']")).select_by_visible_text("Jabber CSF SIP Profile")
    main_page1 = driver.current_window_handle
    driver.find_element_by_xpath("//*[@id='2tbllink']").click()
    try:
        ActionChains(driver).send_keys(Keys.ENTER).perform()#不指定位置直接点击回车，弹出框内容为：单击 [应用配置] 按键以使更改生效。
    except:
        driver.find_element_by_xpath("//*[@id='find_ikdevice_primaryphone']").click()
    else:
        pass
    current_window = driver.current_window_handle 
    all_handles = driver.window_handles   
    for handle in all_handles:
        if handle != current_window:          
            driver.switch_to.window(handle)
    Select(driver.find_element_by_xpath("//*[@id='searchLimit0']")).select_by_visible_text("精确等于")#添加主电话关联
    driver.find_element_by_xpath("//*[@id='searchString0']").send_keys("CSF"+user_id1)
    driver.find_element_by_xpath("//*[@id='filterRow0']/td[7]/input").click()
    driver.find_element_by_xpath("//*[@id='contentautoscroll']/form[1]/table[2]/tbody/tr[2]/td[1]/input[1]").click()
    driver.find_element_by_xpath("//*[@id='4tbl']").click()
    driver.switch_to.window(main_page1)
    try:
        driver.find_element_by_xpath("//*[@id='2tbllink']").click()
    except:
        ActionChains(driver).send_keys(Keys.ENTER).perform()
    else:
        pass
    try:
        driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/table/tbody/tr/td[1]/fieldset/table/tbody/tr[2]/td[2]/u/a").click()
    except:
        driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/table/tbody/tr/td[1]/fieldset/table/tbody/tr[2]/td[2]/u/a").click()
        driver.find_element_by_xpath("//*[@id='DNORPATTERN']").send_keys(phone_number1)
    else:
        pass
    driver.find_element_by_xpath("//*[@id='DESCRIPTION']").click()#输入目录号码后点击一处地方让界面刷新
    driver.find_element_by_xpath("//*[@id='DISPLAY']").send_keys(user_name1)
    driver.find_element_by_xpath("//*[@id='DISPLAYASCII']").click()
    driver.find_element_by_xpath("//*[@id='DISPLAYASCII']").clear()
    driver.find_element_by_xpath("//*[@id='DISPLAYASCII']").send_keys(user_name1)
    driver.find_element_by_xpath("//*[@id='E164MASK']").send_keys(phone_number1)
    former_page = driver.current_window_handle
    driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/fieldset[16]/table/tbody/tr/td[2]/input").click()
    time.sleep(2)
    current_window = driver.current_window_handle 
    all_handles = driver.window_handles  
    for handle in all_handles:
        if handle != current_window:          
            driver.switch_to.window(handle)
    Select(driver.find_element_by_xpath("//*[@id='searchField0']")).select_by_visible_text("用户 ID")
    Select(driver.find_element_by_xpath("//*[@id='searchLimit0']")).select_by_visible_text("精确等于")
    driver.find_element_by_xpath("//*[@id='searchString0']").clear()
    driver.find_element_by_xpath("//*[@id='searchString0']").send_keys(user_id1)
    driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/table[2]/tbody/tr[2]/td[1]/input[1]").click()
    driver.find_element_by_xpath("//*[@id='7tbl']").click()
    driver.switch_to.window(former_page)
    driver.find_element_by_xpath("//*[@id='1tbl']").click()
    #jabber配置文件创建完成
    driver.find_element_by_id("用户管理Button").click()
    time.sleep(1)
    driver.find_element_by_xpath("//*[@id='usermanagement']/ul/li[2]/a").click()
    Select(driver.find_element_by_name("searchField0")).select_by_visible_text("用户 ID")
    Select(driver.find_element_by_name("searchLimit0")).select_by_visible_text("精确等于")
    driver.find_element_by_xpath("//*[@id='searchString0']").clear()
    driver.find_element_by_xpath("//*[@id='searchString0']").send_keys(user_id1)
    driver.find_element_by_xpath("//*[@id='filterRow0']/td[7]/input").click()
    driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/table[2]/tbody/tr[2]/td[2]/a").click()
    driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/fieldset[5]/table/tbody/tr[1]/td[3]/input[1]").click()
    Select(driver.find_element_by_xpath("//*[@id='searchLimit0']")).select_by_visible_text("精确等于")
    driver.find_element_by_xpath("//*[@id='searchString0']").send_keys("CSF"+user_id1)
    driver.find_element_by_xpath("//*[@id='filterRow0']/td[7]/input").click()
    driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/table[2]/tbody/tr[2]/td[1]/input[1]").click()
    driver.find_element_by_xpath("//*[@id='5tbl']").click()
    driver.find_element_by_xpath("/html/body/div[2]/table[2]/tbody/tr/td[2]/div/input").click()
    primaryext = phone_number1+' '+'in'+ ' ' + 'P-REG-AllInternal'
    Select(driver.find_element_by_xpath("//*[@id='PRIMARYEXTENSION']")).select_by_visible_text(primaryext)
    driver.find_element_by_xpath("//*[@id='1tblimage']").click()
    time.sleep(3)
    print('jabber配置完成')
    #Jabber绑定完成

info = get_snow_info()
print(info)
print(info[0][0])
task_len = len(info)
for i in info:
    mydb = mysql.connector.connect(
        host="10.xx.xx.xx",
        user="root",
        passwd="xxxxxx",
        database="iptauto"
    )
    mycursor = mydb.cursor(dictionary=True)
    if i[5].startswith('Sales Rep'):
        print(i[5])
        print('用户是销售')
        close_imcompleted(i[1])
        continue
    if not i[4]:
        print('用户所在办公室没有座机')
        close_imcompleted(i[1])
        continue
    else:
        print(i[5])
        print('用户不是销售')
        Task_Address = i[1]
        site = i[4]
        user_id = i[2]
        user_name = i[3]
        sql_0 = "SELECT * FROM `main_sheet` WHERE site = '" + site + "' and userid is null order by extension limit 1"
        mycursor.execute(sql_0)
        num = mycursor.fetchall()
        print(num)
        display_name = user_name
        phone_number = num[0]['extension']
        connection_symbol = ' - '
        ext = phone_number[-5:]
        check_status = check_userext(user_id, Task_Address)
        if check_status == False:
            continue
        #检查用户是否已经有号码
        driver = webdriver.Chrome()
        driver.get("https://deskphone.xx.xxxxxx.com/ccmadmin/showHome.do")
        driver.find_element_by_xpath('//*[@id="details-button"]').click()
        driver.find_element_by_xpath('//*[@id="proceed-link"]').click()
        time.sleep(2)
        driver.maximize_window()#全屏
        driver.find_element_by_name("j_username").send_keys("sheny32")
        driver.find_element_by_name("j_password").send_keys("SHpass05")
        time.sleep(2)
        driver.find_element_by_class_name("cuesLoginButton").click()
        driver.find_element_by_id("设备Button").click()
        time.sleep(1)
        devicesetting = driver.find_element_by_xpath("//*[@id='device']/ul/li[7]/a")
        deviceprofile = driver.find_element_by_xpath("//*[@id='device']/ul/li[7]/ul/li[4]/a")
        mouse_action = ActionChains(driver)
        mouse_action.move_to_element(devicesetting).perform()#移动到设备设置
        mouse_action.move_to_element(deviceprofile).click().perform()
        driver.find_element_by_xpath("//*[@id='1tbllink']").click()
        driver.find_element_by_xpath("//*[@id='TKMODEL']").click()
        Select(driver.find_element_by_name("tkmodel")).select_by_value("404")
        driver.find_element_by_xpath("//*[@id='1tbl']").click()
        time.sleep(1)
        driver.find_element_by_xpath("//*[@id='1tbl']").click()
        driver.find_element_by_xpath("//*[@id='NAME']").send_keys(user_name)
        driver.find_element_by_xpath("//*[@id='DESCRIPTION']").send_keys(user_name + connection_symbol + phone_number)
        Select(driver.find_element_by_name("fkphonetemplate")).select_by_index(20)
        driver.find_element_by_xpath("//*[@id='2tbl']").click()
        try:
            driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/table/tbody/tr/td[1]/fieldset/table/tbody/tr[2]/td[2]/u/a").click()
        except:
            driver.get_screenshot_as_file("c:\\win10input\\"+ user_id +".png")  # 出错截屏
            print("保存配置文件时出错")
            #需要改成根据时间生成
            display_name = display_name + " " + user_id
            driver.find_element_by_xpath("//*[@id='NAME']").send_keys(" " + user_id)
            print(user_name)
            driver.find_element_by_xpath("//*[@id='2tbl']").click()
            driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/table/tbody/tr/td[1]/fieldset/table/tbody/tr[2]/td[2]/u/a").click()
        else:
            print("保存配置文件时正常")
            time.sleep(2)
        driver.find_element_by_xpath("//*[@id='DNORPATTERN']").send_keys(phone_number)
        driver.find_element_by_xpath("//*[@id='DESCRIPTION']").click()#输入目录号码后点击一处地方让界面刷新
        Select(driver.find_element_by_xpath("//*[@id='FKROUTEPARTITION']")).select_by_visible_text('P-REG-AllInternal')#无论是否已经选，都重新选择一下路由分区
        driver.find_element_by_xpath("//*[@id='DESCRIPTION']").click()#输入路由分区后点击一处地方让界面刷新
        driver.find_element_by_xpath("//*[@id='DESCRIPTION']").clear()
        driver.find_element_by_xpath("//*[@id='DESCRIPTION']").send_keys(user_name + connection_symbol + phone_number)
        driver.find_element_by_xpath("//*[@id='ALERTINGNAME']").clear()
        driver.find_element_by_xpath("//*[@id='ALERTINGNAME']").send_keys(user_name)
        driver.find_element_by_xpath("//*[@id='ALERTINGNAMEASCII']").click()
        driver.find_element_by_xpath("//*[@id='ALERTINGNAMEASCII']").clear()
        driver.find_element_by_xpath("//*[@id='ALERTINGNAMEASCII']").send_keys(user_name)
        associated_devices = driver.find_elements_by_xpath("//*[@id='SELECTEDDEVICEASSOCIATION']/option")
        for i in associated_devices:#取消所有的已关联设备
            i.click()
            driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/fieldset[2]/table/tbody/tr[9]/td/img[1]").click()
        Select(driver.find_element_by_name("fkvoicemessagingprofile")).select_by_value("45f79d69-d55e-059b-94e2-c92bf1f31a87")
        if site=='SH':
            Select(driver.find_element_by_name("fkcallingsearchspace_sharedlineappear")).select_by_visible_text("Line-Unrestricted-SHA5-CSS")
        elif site=='CD':
            Select(driver.find_element_by_name("fkcallingsearchspace_sharedlineappear")).select_by_visible_text('Line-Unrestricted-CTU2-CSS')
        elif site=='BJ':
            Select(driver.find_element_by_name("fkcallingsearchspace_sharedlineappear")).select_by_visible_text('Line-Unrestricted-BEI1-CSS')
        elif site=='GZ':
            Select(driver.find_element_by_name("fkcallingsearchspace_sharedlineappear")).select_by_visible_text('Line-Unrestricted-GUA1-CSS')
        qianzhuan_xpath = "/html/body/table/tbody/tr/td/div/form/fieldset[9]/table[1]/tbody/tr"
        time.sleep(1)
        if driver.find_element_by_xpath("//*[@id='CFBINTVOICEMAILENABLED']").is_selected():
            print("前转遇忙内线已经勾选")
        else:
            driver.find_element_by_xpath("//*[@id='CFBINTVOICEMAILENABLED']").click()
            print("前转遇忙内线本来未勾选，现在已经勾选")
        if driver.find_element_by_xpath("//*[@id='CFBVOICEMAILENABLED']").is_selected():
            print("前转遇忙外线已经勾选")
        else:
            driver.find_element_by_xpath("//*[@id='CFBVOICEMAILENABLED']").click()
            print("前转遇忙外线本来未勾选，现在已经勾选")
        if driver.find_element_by_xpath("//*[@id='CFNAINTVOICEMAILENABLED']").is_selected():
            print("前转无应答内线已经勾选")
        else:
            driver.find_element_by_xpath("//*[@id='CFNAINTVOICEMAILENABLED']").click()
            print("前转无应答内线本来未勾选，现在已经勾选")
        if driver.find_element_by_xpath("//*[@id='PFFINTVOICEMAILENABLED']").is_selected():
            print("前转无覆盖内线已经勾选")
        else:
            driver.find_element_by_xpath("//*[@id='PFFINTVOICEMAILENABLED']").click()
            print("前转无覆盖内线本来未勾选，现在已经勾选")
        if driver.find_element_by_xpath("//*[@id='PFFVOICEMAILENABLED']").is_selected():
            print("前转无覆盖外线已经勾选")
        else:
            driver.find_element_by_xpath("//*[@id='PFFVOICEMAILENABLED']").click()
            print("前转无覆盖外线本来未勾选，现在已经勾选")
        if driver.find_element_by_xpath("//*[@id='CFDFVOICEMAILENABLED']").is_selected():
            print("CTI故障时前转已经勾选")
        else:
            driver.find_element_by_xpath("//*[@id='CFDFVOICEMAILENABLED']").click()
            print("CTI故障时前转本来未勾选，现在已经勾选")
        #关于site的判断放在此处
        call_search_space=""
        if site =="SH":
            call_search_space = "P-SHA5-CFW-Local-CSS"
        elif site =="BJ":
            call_search_space = "P-BEI1-CFW-Local-CSS"
        elif site=="CD":
            call_search_space = "P-CTU2-CFW-Local-CSS"
        elif site=="GZ":
            call_search_space = "P-GUA1-CFW-Local-CSS"
        #关于site的判断放在此处
        print(call_search_space)
        Select(driver.find_element_by_name("fkcallingsearchspace_cfa")).select_by_visible_text(call_search_space)
        Select(driver.find_element_by_name("fkcallingsearchspace_scfa")).select_by_visible_text(call_search_space)
        Select(driver.find_element_by_name("fkcallingsearchspace_cfbint")).select_by_visible_text(call_search_space)
        Select(driver.find_element_by_name("fkcallingsearchspace_cfnaint")).select_by_visible_text(call_search_space)
        Select(driver.find_element_by_name("fkcallingsearchspace_pffint")).select_by_visible_text(call_search_space)
        Select(driver.find_element_by_name("fkcallingsearchspace_devicefailure")).select_by_visible_text(call_search_space)
        driver.find_element_by_xpath("//*[@id='DISPLAY']").send_keys(user_name)
        driver.find_element_by_xpath("//*[@id='DISPLAYASCII']").click()
        driver.find_element_by_xpath("//*[@id='LABEL']").send_keys(user_name + connection_symbol + ext)
        driver.find_element_by_xpath("//*[@id='1tbl']").click()
        driver.find_element_by_xpath("//*[@id='searchDiv1']/input").click()
        main_handle = driver.current_window_handle
        # print(driver.current_window_handle)
        #加一步判断呼叫搜索空间选择了没
        Select(driver.find_element_by_name("menu1")).select_by_visible_text('预订/取消预订服务')
        Select(driver.find_element_by_name("menu1")).select_by_visible_text('预订/取消预订服务')
        driver.find_element_by_xpath("//*[@id='searchDiv0']/input").click()
        current_window = driver.current_window_handle 
        all_handles = driver.window_handles 
        for handle in all_handles:
            if handle != current_window:          
                driver.switch_to.window(handle)
        Select(driver.find_element_by_name("fktelecasterservice")).select_by_visible_text('Extension Mobility')
        driver.find_element_by_xpath("//*[@id='1tbl']").click()
        driver.find_element_by_xpath("html/body/table/tbody/tr/td/div/form/table/tbody/tr/td/fieldset/input[1]").click()
        driver.find_element_by_xpath("//*[@id='2tbl']")
        driver.close()
        driver.switch_to.window(main_handle)
        driver.find_element_by_id("用户管理Button").click()
        time.sleep(1)
        driver.find_element_by_xpath("//*[@id='usermanagement']/ul/li[2]/a").click()
        Select(driver.find_element_by_name("searchField0")).select_by_visible_text("用户 ID")
        Select(driver.find_element_by_name("searchLimit0")).select_by_visible_text("精确等于")
        driver.find_element_by_xpath("//*[@id='searchString0']").send_keys(user_id)
        driver.find_element_by_xpath("//*[@id='filterRow0']/td[7]/input").click()
        time.sleep(1)
        driver.find_element_by_xpath("//*[@id='contentautoscroll']/form/table[2]/tbody/tr[2]/td[2]/a").click()
        driver.find_element_by_xpath("//*[@id='find_availableDeviceProfiles']").click()
        current_window = driver.current_window_handle 
        all_handles = driver.window_handles   
        for handle in all_handles:          
            if handle != current_window
                driver.switch_to.window(handle)
        Select(driver.find_element_by_name("searchLimit0")).select_by_visible_text("精确等于")
        driver.find_element_by_xpath("//*[@id='searchString0']").send_keys(display_name)
        driver.find_element_by_xpath("//*[@id='filterRow0']/td[7]/input").click()
        driver.find_element_by_xpath("//*[@id='contentautoscroll']/form[1]/table[2]/tbody/tr[2]/td[1]/input[1]").click()
        driver.find_element_by_xpath("//*[@id='4tbl']").click()
        driver.switch_to.window(main_handle)
        #查找完设备配置文件后点击添加所选项
        driver.find_element_by_xpath("/html/body/table/tbody/tr/td/div/form/fieldset[12]/input[1]").click()
        #*******************************************************************************
        #设置分机的步骤已经结束！
        #*******************************************************************************
    #Jabber Profile部分开始
        jabber_task(user_id, user_name, phone_number, site)
        driver.get(Task_Address)
        driver.maximize_window()#全屏
        WebDriverWait(driver, 60, 0.2).until(EC.frame_to_be_available_and_switch_to_it((By.ID, 'gsft_main')))
        Task_Time = driver.find_element_by_xpath("//*[@id='sys_readonly.sc_task.opened_at']").get_attribute('value')
        print(Task_Time)
        Task_No = driver.find_element_by_xpath("//*[@id='sys_readonly.sc_task.number']").get_attribute('value')
        print(Task_No)
        Select(driver.find_element_by_xpath("//*[@id='sc_task.state']")).select_by_visible_text("Closed Complete")
        driver.find_element_by_xpath("//*[@id='sys_display.sc_task.assigned_to']").send_keys("Alex Shen")
        time.sleep(5)
        driver.find_element_by_xpath("//*[@id='sc_task.description']").click()
        time.sleep(5)
        driver.find_element_by_xpath("//*[@id='sc_task.work_notes']").send_keys("Set " + phone_number+" and jabber")
        driver.find_element_by_xpath("//*[@id='sysverb_update_and_stay']").click()
        result16 = EC.text_to_be_present_in_element_value((By.XPATH, "//*[@id='sys_display.sc_task.assigned_to']"), 'Alex Shen')(driver)
        if result16:
            pass
        else:
            driver.find_element_by_xpath("//*[@id='sys_display.sc_task.assigned_to']").send_keys("Alex Shen")
            time.sleep(1)
            ActionChains(driver).send_keys(Keys.DOWN).perform()
            ActionChains(driver).send_keys(Keys.ENTER).perform()
        print('Task has been closed.')
        sql_1 = "UPDATE main_sheet SET userid ='" + user_id + "'," + "display_name = '" + display_name + "'" + ",task_no = " \
                + "'" + Task_No + "'" + "," + "task_date='" + Task_Time + "'" + "," + "task_add='" + Task_Address + "'" + "," + "finish_date = '" + datetime.now().strftime('%Y-%m-%d %H:%M') + "'" + " WHERE extension = '" + phone_number + "'"
        print(sql_1)
        mycursor.execute(sql_1)
        mydb.commit()
        mycursor.close()
        #————————————————————————————————————————————
        #发送邮件
        office365(user_id, user_name, phone_number, site)
        print('邮件真的发送成功了')
        #----------------------------------------------------
        #IDM修改电话号码
        idm_config(user_id, phone_number)
        #关闭浏览器
        driver.close()


