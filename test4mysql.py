import mysql.connector
from selenium.webdriver.support.wait import WebDriverWait
from datetime import datetime
from Testservicenow import close_imcompleted

def check_userext(user_id, task_address):
    mydb = mysql.connector.connect(
        host="10.xxx.xx.xxx",
        user="root",
        passwd="xxxxxx",
        database="iptauto"
    )
    mycursor = mydb.cursor(dictionary=True)
    sql_1 = "select * from main_sheet where userid = '" + user_id + "'"
    mycursor.execute(sql_1)
    myresult = mycursor.fetchall()
    if myresult:
        print('用户已经有座机号码')
        close_imcompleted(task_address )
        return False
    else:
        print('用户还没有座机号码，继续配置')
        pass
    mycursor.close()


if __name__ == '__main__':
    check_userext('xxxxxx')
