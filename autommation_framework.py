import pandas as pd
from selenium import webdriver
#import org openqa selenium support ui select
from selenium.webdriver.support.ui import Select
import time
import excel_operational
test_case_location = 'test_case/framework.xlsx'
def read_excel():
    reader = pd.read_excel(test_case_location)
    for row, column in reader.iterrows():
        sn = column["sn."]
        test_summary = column["test_summary"]
        xpath = column["xpath"]
        action = column["action"]
        value = column["value"]
        excute_flag = column["execute_flag"]
        if(excute_flag!='N'):
            action_defination(sn,test_summary,xpath,action,value)
        else:
            result = "Not tested"
            remarks = "Test was skipped due to N flag"
            print(sn,test_summary,result,remarks)
            excel_operational.write_excel(sn,test_summary,result,remarks)

def action_defination(sn,test_summary,xpath,action,value):
    try:
        if action == 'open_browser':
           driver,result,remarks = open_browser(value)
        elif action == 'open_url':
            result,remarks = open_url(value)
        elif action == 'click':
            result,remarks = click(xpath)
        elif action == 'send_value':
            result,remarks = send_value(xpath,value)
        elif action == 'wait':
            result,remarks = wait(value)
        elif action == 'select_dropdown':
            result,remarks = select_dropdown(xpath,value)
        elif action == 'verify_text':
            result,remarks = verify_text(xpath,value)
        else:
            print("Action not supported by framework")
        print(sn,test_summary,result,remarks)
        excel_operational.write_excel(sn,test_summary,result,remarks)
    except Exception as ex:
        print(ex,"Exception has occurred")
        result = 'Fail'
        remarks =ex
        print(sn, test_summary, result, remarks)
        excel_operational.write_excel(sn,test_summary,result,remarks)
def open_browser(value):
    try:
        global driver
        if value == 'Chrome':
           driver = webdriver.Chrome("E:\project1\chromedriver.exe")
           driver.maximize_window()
           result = 'PASS'
           remarks = ''
        elif value == 'firefox':
             driver = webdriver.Firefox()
             result = 'PASS'
             remarks = ''
        else:
            print(value,"Browser not supported")
            result = 'FAIL'
            remarks = 'Browser Not Supported'
        return driver,result,remarks
    except Exception as ex:
        result = "FAIL"
        remarks = ex
def wait(value):
    try:
        time.sleep(value)
        result = "PASS"
        remarks = " "
    except Exception as ex:
        result = "FAIL"
        remarks = "ex"
    return result, remarks

def open_url(value):
    try:
        driver.get(value)
        result = "PASS"
        remarks = " "
    except Exception as ex:
        result = "FAIL"
        remarks = "ex"
    return result,remarks

def click(xpath):
    try:
        driver.find_element_by_xpath(xpath).click()
        result = "PASS"
        remarks = ""
    except Exception as ex:
        result = "FAIL"
        remarks = ex
    return result, remarks
def send_value(xpath,value):
    try:
        driver.find_element_by_xpath(xpath).send_keys(value)
        result = "PASS"
        remarks = ""
    except Exception as ex:
        result = "FAIL"
        remarks = "ex"
    return result,remarks
def verify_text(xpath,value):
    output_text=driver.find_element_by_xpath(xpath).text
    try:
        assert output_text == value
    except AssertionError:
        result = "FAIL"
        remarks = "Actual value is:,"+output_text+"input value is" +value
    else:
        result = "PASS"
        remarks = ""
    return result,remarks
def select_dropdown(xpath,value):
    try:
        #driver.find_element_by_xpath(xpath).text
        #menu = Select(destination_xpath)
        #select_by_visible_text(value)
        destination_xpath = driver.find_element_by_xpath(xpath)
        menu = Select(destination_xpath)
        menu.select_by_visible_text(value)
        result = "PASS"
        remarks = ""
    except Exception as ex:
        result = "FAIL"
        remarks = "ex"
    return result, remarks
if __name__ == '__main__':
    excel_operational.remove_file()
    excel_operational.write_header()
    excel_operational.write_summary()
    read_excel()


