from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import load_workbook
from openpyxl import Workbook
from selenium.webdriver.common.by import By

wb = load_workbook(r'C:\Users\jaich\Desktop\student_data.xlsx')
ws = wb.active

options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--start-maximized')
driver = webdriver.Chrome(options=options)

driver.get("https://exam.smvec.ac.in/exam_result_ug_pg_regular_jan_2024/")

result_wb = Workbook()
result_ws = result_wb.active

flag = 0
for row in ws.iter_rows(min_row=11, values_only=True):
    if len(row) >= 2:
        regno, dob = row[:2]
        while True:
            try:
                resigter_number_field = driver.find_element(By.ID,"txtRollNo")
                dob_field = driver.find_element(By.ID,"txtDoB")
                captcha_text = driver.find_element(By.ID,"mainCaptcha")
                captcha_text_value = captcha_text.text
                captcha_field = driver.find_element(By.ID,"txtInput")

                resigter_number_field.click()
                resigter_number_field.send_keys(regno)
                dob_field.click()
                dob_field.send_keys(dob)
                captcha_field.click()
                captcha_field.send_keys(captcha_text_value)
                captcha_field.send_keys(Keys.ENTER)

                time.sleep(3)  
                
                try:
                    name_element = driver.find_element(By.XPATH,"//div[@id='printdataResults']/div[5]")
                    name = name_element.text.split(": ")[1] if ' ' in name_element.text else ''
                    sgpa_element = driver.find_element(By.XPATH,"//div[@id='printdataResults']/div[7]")
                    sgpa = sgpa_element.text.split(maxsplit=1)[1] if ' ' in sgpa_element.text else ''
                    print(name)
                    print(sgpa)
                    #Find number of subjects
                    table_element = driver.find_element(By.XPATH,"//div[@id='printdataResults']/div[6]/table/tbody")
                    tr_elements = table_element.find_elements(By.TAG_NAME, "tr")
                    tr_count = len(tr_elements)
                    row = [regno, name]
                    sub_elements_values = []
                    title_element_values = []
                    for i in range(1,tr_count+1):
                        subject_xpath_expression = "//div[@id='printdataResults']/div[6]/table/tbody/tr[{}]/td[5]".format(i)
                        title_xpath_expression = "//div[@id='printdataResults']/div[6]/table/tbody/tr[{}]/td[3]".format(i)
                        sub_i_element = driver.find_element(By.XPATH, subject_xpath_expression)
                        title_i_element = driver.find_element(By.XPATH, title_xpath_expression)
                        sub_elements_values.append(sub_i_element.text)
                        title_element_values.append(title_i_element.text)
                    # Take screenshot
                    if flag == 0:
                        width = driver.execute_script("return document.body.scrollWidth")
                        height = driver.execute_script("return document.body.scrollHeight")
                        driver.set_window_size(width, height) 
                        first_row = ["Register Number","Name"]
                        first_row.extend(title_element_values)
                        first_row.append("SGPA")
                        result_ws.append(first_row)
                        flag = 1
                    time.sleep(2)
                    print(sub_elements_values)
                    print(title_element_values)
                    row = [regno,name]
                    row.extend(sub_elements_values)
                    row.append(sgpa)
                    result_ws.append(row)
                    driver.save_screenshot(f"result/{regno}_result.png")
                    
                except Exception as e:
                    print("Error:", e)
                    continue
                break  
            except Exception as e:
                print("Error:", e)
                continue

result_wb.save("result.xlsx")

driver.quit()
