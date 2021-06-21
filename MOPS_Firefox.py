
#edition 1 (runs one time)
'''
from selenium import webdriver
import time

url ='https://mops.twse.com.tw/mops/web/stapap1_all'

driver = webdriver.Firefox(executable_path = '/usr/local/bin/geckodriver')
driver.get(url)
print('url:', url)

skind_select = driver.find_element_by_xpath('//*[@id="search"]/table/tbody/tr/td[6]/select')
skind_select.click()
skind = driver.find_element_by_xpath('//*[@id="search"]/table/tbody/tr/td[6]/select/option[16]')
skind.click()
co_id = driver.find_element_by_name('YM')
co_id.send_keys('10801')
search_co = driver.find_element_by_xpath('//*[@id="search_bar1"]/div/input')
search_co.click()
time.sleep(5)
select_co = driver.find_element_by_xpath('//*[@id="table01"]/form/table/tbody/tr[5]/td[3]/input')
select_co.click()

now_handle = driver.current_window_handle
all_handles = driver.window_handles
print('Found', len(all_handles), 'windows')
#print(now_handle)

for handle in all_handles:
    if handle != now_handle:
        driver.switch_to.window(handle)
        time.sleep(5)
        #print(driver.current_window_handle)
        value = driver.find_element_by_xpath('//*[@id="table01"]/center/table[4]/tbody/tr[7]/td[8]')
print(value.text)

print('OK!')
'''
#edition 2


from selenium import webdriver
import time
import re
import codecs
import datetime
import csv

def write_file_utf8(filename, text):
    with codecs.open(filename, 'w', encoding='utf-8') as f:
        f.write(text)
        f.close()

def log_file_utf8(filename, text):
    with codecs.open(filename, 'a', encoding='utf-8') as f:
        #append file
        f.write(text + '\n')
        f.close()
        
url ='https://mops.twse.com.tw/mops/web/stapap1_all'

driver = webdriver.Firefox(executable_path = '/usr/local/bin/geckodriver')
driver.get(url)
print('url:', url)

#calculate the amount of categories
categories = re.findall(r'\d+', '><option value="1">水泥工業</option><option value="2">食品工業</option><option value="3">塑膠工業</option><option value="4">紡織纖維</option><option value="5">電機機械</option><option value="6">電器電纜</option><option value="21">化學工業</option><option value="22">生技醫療業</option><option value="7">化學生技醫療</option><option value="8">玻璃陶瓷</option><option value="9">造紙工業</option><option value="10">鋼鐵工業</option><option value="11">橡膠工業</option><option value="12">汽車工業</option><option value="24">半導體業</option><option value="25">電腦及週邊設備業</option><option value="26">光電業</option><option value="27">通信網路業</option><option value="28">電子零組件業</option><option value="29">電子通路業</option><option value="30">資訊服務業</option><option value="31">其他電子業</option><option value="13">電子工業</option><option value="23">油電燃氣業</option><option value="14">建材營造</option><option value="15">航運業</option><option value="16">觀光事業</option><option value="17">金融保險業</option><option value="18">貿易百貨</option><option value="19">綜合企業</option><option value="20">其他</option><option value="91">')
 #print(categories)
category_counter = 0
for i in categories:
    category_counter += 1
#print(category_counter)

#write the first row
path = 'output_data.csv'
with open(path, 'w', newline='') as csvfile:
    writer = csv.writer(csvfile, delimiter = '\t')
    writer.writerow(['Datetime', 'SID', 'Date', 'id','name', 'tab_id', 'tab_name', '全體董監持股設質比例'])

irow = 0
years = ['107', '108', '109']
months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
for year in years:
    for month in months:
        date = year + month
        for i in range(category_counter):
            skind_select = driver.find_element_by_xpath('//*[@id="search"]/table/tbody/tr/td[6]/select')
            skind_select.click()
            skind_base = '//*[@id="search"]/table/tbody/tr/td[6]/select/option[%s]'%(i+2)
            skind = driver.find_element_by_xpath(skind_base)
            skind.click()
            co_id = driver.find_element_by_name('YM')
            co_id.send_keys(date)
            search_co = driver.find_element_by_xpath('//*[@id="search_bar1"]/div/input')
            search_co.click()
            time.sleep(2)
            
            co_counter = 0
            while True:
                co_name = []  
                odd = 1
                co_name_odd = driver.find_elements_by_class_name('odd')
                for i in co_name_odd:
                    co_name.insert(odd, i.text)
                    odd += 2
                even = 0
                co_name_even = driver.find_elements_by_class_name('even')
                for i in co_name_even:
                    co_name.insert(even, i.text)
                    even += 2
                for i in co_name:
                    try:
                        temp_co = i
                        co_base = '//*[@id="table01"]/form/table/tbody/tr[%s]/td[3]/input'%(co_counter+2)
                        #co_base = '//*[@id="table01"]/form/table/tbody/tr[10]/td[3]/input'
                        select_co = driver.find_element_by_xpath(co_base)
                        select_co.click()
                        co_counter += 1
                        #print(co_counter)
                        time.sleep(2)
                    
                        now_handle = driver.current_window_handle
                        all_handles = driver.window_handles
                        #print('Found', len(all_handles), 'windows')
                    
                        for handle in all_handles:
                            if handle != now_handle:
                                driver.switch_to.window(handle)
                                time.sleep(5)
                                #print(driver.current_window_handle)
                        irow += 1
                        value = driver.find_element_by_xpath('//*[@id="table01"]/center/table[4]/tbody/tr[7]/td[8]')
                        print(i, value.text)
                        ivalue = value.text
                        
                        #split id/name from original page
                        splitting_id = re.findall(r'([0-9]+)', temp_co,re.I)
                        ID = ''
                        for i in splitting_id:
                            ID = i
                        raw_name = temp_co.replace(ID, '')
                        print(ID)
                        print(raw_name)
                                               
                        #split id/name on tab 
                        compName = driver.find_element_by_xpath('//*[@id="table01"]/center/table[1]/tbody/tr/td').text
                        splitting_id = re.findall(r'([0-9]+)', compName,re.I)                        
                        tab_ID = ''
                        for i in splitting_id:
                            tab_ID = i
                        tab_name = compName.replace(tab_ID, '')
                        print(tab_ID)
                        print(tab_name)
                        
                        Stockholding_Directors_Supervisors = driver.find_element_by_xpath('//*[@id="table01"]/center/table[4]/tbody/tr[7]/td[7]').text
                        result_string = datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + '\t' + str(irow) + '\t' + str(date) + '\t' + str(ID) + '\t' + str(raw_name) + '\t' + str(tab_ID) + '\t' +  str(tab_name) + '\t' + str(Stockholding_Directors_Supervisors) + '\t' + str(ivalue)
                        print('result_string:', result_string)
                        log_file_utf8('logfile.txt', result_string)
                        path = 'output_data.csv'
                        try:
                            with open(path, 'a', newline='') as csvfile:
                                writer = csv.writer(csvfile, delimiter = '\t')
                                writer.writerow([datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), str(irow), str(date), str(ID), str(raw_name), str(tab_ID), str(tab_name), str(ivalue)])
                        except:
                           pass
                       
                        driver.close()
                        driver.switch_to.window(now_handle)
                    except:
                        break
                break
