from selenium import webdriver
import time
import requests
import codecs
import pdfplumber
import re
import csv
from lxml import html


url ='https://mops.twse.com.tw/mops/web/t57sb01_q5'


#idlist = ['2801']
#year = '108'
idlist = ['2801', '2809', '2812', '2816', '2820', '2823', '2832', '2834', '2836', '2838', '2845', '2849', '2850', '2851', '2852', '2855', '2867', '2880', '2881', '2882', '2883', '2884', '2885', '2886', '2887', '2888', '2889', '2890', '2891', '2892', '2897', '5876', '5880', '6005', '6024']
#years = ['106', '107', '108', '109']
years= ['108']

path = 'output_data.csv'
try:
    with open(path, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile, delimiter = '\t')
        writer.writerow(['coID', 'year', '出席董事:'])
except:
    pass

#search
for year in years:
    for coID in idlist:
        driver = webdriver.Firefox(executable_path = '/usr/local/bin/geckodriver')
        driver.get(url)
        print('url:', url)
        
        co_id = driver.find_element_by_xpath('//*[@id="co_id"]')
        co_id.send_keys(coID)
        search_year = driver.find_element_by_xpath('//*[@id="year"]')
        search_year.send_keys(year)
        
        search = driver.find_element_by_xpath('/html/body/center/table/tbody/tr/td/div[4]/table/tbody/tr/td/div/table/tbody/tr/td[3]/div/div[3]/form/table/tbody/tr/td[4]/table/tbody/tr/td[2]/div/div/input')
        search.click()
        
        now_handle = driver.current_window_handle
        all_handles = driver.window_handles
                                
                            
        for handle in all_handles:
            if handle != now_handle:
                driver.switch_to.window(handle)
                time.sleep(2)
        page = requests.get(driver.current_url)
        root = html.fromstring(page.text)
        tree = root.getroottree()
        result = root.xpath('//*[. = "股東會議事錄"]')
        for r in result:
            #print(tree.getpath(r))
            xpathtxt = tree.getpath(r)
        xpathtxt = xpathtxt[:-2]
        xpathtxt = xpathtxt.replace('table/table', 'table[2]/tbody')
        #print(xpathtxt)
        xpathtxt+='8]/a'
        #print(xpathtxt)
        target = driver.find_element_by_xpath(xpathtxt)
        target.click()
        
        driver.close()
        all_handles = driver.window_handles
        for handle in all_handles:
            if handle != now_handle:
                driver.switch_to.window(handle)
                time.sleep(2)
                
        click_pdf_link = driver.find_element_by_xpath('/html/body/center/a')
        click_pdf_link.click()
        
        #get pdf
        r = requests.get(driver.current_url)
        
        file_name_pdf = coID + '_' + year + '.pdf'
        file_name_txt = coID + '_' + year + '.txt'
        
        with open(file_name_pdf, 'wb') as f:
            f.write(r.content)
            f.close()
            
        time.sleep(2)
        driver.close()
        driver.switch_to.window(now_handle) 
              
        #pdf to txt
        pdf = pdfplumber.open(file_name_pdf)
        pages = pdf.pages
        
        all_text = ''
        for page in pdf.pages:
            text = page.extract_text()
            try:
                all_text = all_text + '\n' + text
            except:
                pass
        #print(text)
        pdf.close()
        with codecs.open(file_name_txt, 'w', encoding = 'utf-8') as f:
            f.write(all_text)
            f.close()
            
        #read txt
        with codecs.open(file_name_txt, 'r', encoding='utf-8') as f:
            content = f.read()
            f.close()
            #print(content)
        
        content_replaced = content.replace('\n', '')
        #content_replaced = content_replaced.replace(' ', '')
        #print(content_replaced)
        #re is incomplete
        match = re.search(r"出席董事：(.*?)。", content_replaced)
        try:
            print(coID, end = ' ')
            print(year)
            print(match.group(1))#0
            result = match.group(1)
        except:
            result = 'unavailable' 
            
        match = re.search(r"列席(董事)?：(.*?)。?", content_replaced)
        try:
            print(coID, end = ' ')
            print(year)
            print(match.group(1))#0
            result = match.group(1)
        except:
            result = 'unavailable'
            
       
                
        path = 'output_data.csv'
        try:
            with open(path, 'a', newline='') as csvfile:
                writer = csv.writer(csvfile, delimiter = '\t')
                writer.writerow([str(coID), str(year), result])
        except:
            pass

        driver.close()
print("Done")
