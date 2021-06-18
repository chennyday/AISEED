from selenium import webdriver
import time
import csv


url ='https://mops.twse.com.tw/mops/web/t150sb04'

years = [107, 108, 109]

path = 'output_data.csv'
try:
    with open(path, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter = '\t')
        writer.writerow(['year', 'coID', 'co', 'date', '盈餘分配/盈虧撥補', '章程修訂', '營業報告書及財務報表'])
except:
    pass

for year in years:
    driver = webdriver.Firefox(executable_path = '/usr/local/bin/geckodriver')
    driver.get(url)
    print('url:', url)
    
    search_year = driver.find_element_by_xpath('//*[@id="year"]')
    search_year.send_keys(year)
    time.sleep(1)
    search = driver.find_element_by_xpath('/html/body/center/table/tbody/tr/td/div[4]/table/tbody/tr/td/div/table/tbody/tr/td[3]/div/div[3]/form/table/tbody/tr/td[4]/table/tbody/tr/td[2]/div/div/input')
    search.click()
    
    n = 3
    time.sleep(2)
    count = driver.find_elements_by_xpath('//*[@id="fm1"]/table/tbody/tr[1]/td/table/tbody/tr')
    print(len(count)-1)
    rowcount = (len(count))-1
    
    while n<rowcount:
        print("Executing ",n)
        try:
            coIDbase = '//*[@id="fm1"]/table/tbody/tr[1]/td/table/tbody/tr[%s]/td[1]'%(n)
            cobase = '//*[@id="fm1"]/table/tbody/tr[1]/td/table/tbody/tr[%s]/td[2]'%(n)
            datebase = '//*[@id="fm1"]/table/tbody/tr[1]/td/table/tbody/tr[%s]/td[4]'%(n)
            data1base = '//*[@id="fm1"]/table/tbody/tr[1]/td/table/tbody/tr[%s]/td[10]'%(n)
            data2base = '//*[@id="fm1"]/table/tbody/tr[1]/td/table/tbody/tr[%s]/td[12]'%(n)
            data3base = '//*[@id="fm1"]/table/tbody/tr[1]/td/table/tbody/tr[%s]/td[14]'%(n)
            
            coID = driver.find_element_by_xpath(coIDbase)
            co = driver.find_element_by_xpath(cobase)
            date = driver.find_element_by_xpath(datebase)
            data1 = driver.find_element_by_xpath(data1base)
            data2 = driver.find_element_by_xpath(data2base)
            data3 = driver.find_element_by_xpath(data3base)
            print("Completed ", n)
            n+=1
            path = 'output_data.csv'
        
            with open(path, 'a', newline='') as csvfile:
                writer = csv.writer(csvfile, delimiter = '\t')
                writer.writerow([year, coID.text, co.text, date.text, data1.text, data2.text, data3.text])
        except:
            n+=1
            pass
    driver.close()
    print("Completed year %s"%(year))
    time.sleep(1)
        
    
