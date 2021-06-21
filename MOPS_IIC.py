import os
import codecs
import requests
import pandas as pd
from bs4 import BeautifulSoup
import time
import datetime
import random
import re
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import traceback

#aiseed_get_mops_iic.py
#python aiseed_get_mops_iic.py
#半導體業  (Subtotal: 74)
#https://mops.twse.com.tw/mops/web/t51sb01
#twse_ind_stocks = [2302, 2303, 2329, 2330, 2337, 2338, 2342, 2344, 2351, 2363, 2369, 2379, 2388, 2401, 2408, 2434, 2436, 2441, 2449, 2451, 2454, 2458, 2481, 3006, 3014, 3016, 3034, 3035, 3041, 3054, 3094, 3189, 3257, 3413, 3443, 3530, 3532, 3536, 3545, 3583, 3588, 3661, 3686, 3711, 4919, 4952, 4961, 4967, 4968, 5269, 5285, 5471, 6202, 6239, 6243, 6257, 6271, 6415, 6451, 6515, 6525, 6531, 6533, 6552, 6573, 6756, 8016, 8028, 8081, 8110, 8131, 8150, 8261, 8271]
#twse_ind_stocks = [2330, 2883]

#台灣證券交易所 - 公開資訊觀測站 - 法人說明會 (Institutional Investor Conference; IIC) 
#https://mops.twse.com.tw/mops/web/t100sb02_1
#https://mops.twse.com.tw/mops/web/ajax_t100sb02_1

headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.146 Safari/537.36'}

def log_file_utf8(filename, text):
    with codecs.open(filename, 'a', encoding='utf-8') as f:      
        f.write(text + '\n')  #append file
        f.close()

def stock_id_list(filename, text):
    with codecs.open(filename, 'a', encoding='utf-8') as f:
        for item in text:
            f.write("%s\n" % item)
        f.close()

def get_mops_iic(co_id):
    """Get MOPS IIC by Company ID.
    台灣證券交易所 - 公開資訊觀測站 - 法人說明會 (Institutional Investor Conference; IIC)
    Authors: Chenny Day and Min-Yuh Day (2021)
    Args:
        co_id: Input co_id, example: '2330'
    Returns:
        df: DataFrame output mops_iic_co_id.xlsx 
        example: 'mops_iic_2330.xlsx'
    """
    years = ['106', '107', '108', '109']
    months = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    df_i = 0
    
    for year in years:
        for month in months:
            post_data = {
                'encodeURIComponent': 1,
                'step': '1',
                'firstin': 1,
                'off': '1',
                'TYPEK': 'sii',
                'year': year,
                'month': month,
                'co_id': co_id,
            }
            #resp = requests.post('https://mops.twse.com.tw/mops/web/ajax_t100sb02_1', data = post_data, headers = headers, timeout=10)
            
            url = 'https://mops.twse.com.tw/mops/web/ajax_t100sb02_1'
            response = requests.post(url, data = post_data, headers = headers, timeout=10)
            seconds = random.uniform(1, 5)
            #print('sleep...', seconds, ' seconds', datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
            time.sleep(seconds)
            #print(response.text)
            #print('%d/%d'%(int(year), int(month)))
            print(co_id + '_' + year + '_' + month )
            try:
                html = response.text
                substring = '查無資料'   #'<font color="red">查無資料</font>' 
                search_result = re.search(substring, html)
                print('search_result', search_result)
                if not bool(search_result):
                    dfs = pd.read_html(html)
                    print('len(dfs)', len(dfs))
                    if (len(dfs)>0):
                        for df in dfs:
                            df['co_id'] = co_id
                            df['year'] = year
                            df['month'] = month
                            if (df_i == 0):
                                dfall = df
                            else:
                                dfall = dfall.append(df) 
                                #dfall = pd.concat([dfall, df], ignore_index=True)
                            print(datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
                            df_i +=1
                            print('df_i:', df_i)
                            print(df)
                            #df.to_csv('df_' + co_id + '_' + year + '_' + month + '.csv', sep = '\t')
                else:
                    print(co_id + '_' + year + '_' + month, 'Not Found')
                    dir = '/content/gdrive/My Drive'
                    dir = os.path.join(dir, 'data')
                    logfilename = os.path.join(dir, 'logfile.txt') 
                    log_file_utf8(logfilename, datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + '\tNot Found:\t' + str(co_id) + '\t' + str(year) + '\t' + str(month)) 
             
            #except:
            #    print(co_id + '_' + year + '_' + month, 'No tables found')

            except Exception as e:
                print("Error Message: " + str(e))
                print("traceback.format_exc(): ", traceback.format_exc())
                exception_count += 1
                dir = '/content/gdrive/My Drive'
                dir = os.path.join(dir, 'data')
                logfilename = os.path.join(dir, 'logfile.txt') 
                log_file_utf8(logfilename, datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + '\tError Message:\t' + str(e)) 
                log_file_utf8(logfilename, datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + '\tError co_id year month:\t' + str(co_id) + '\t' + str(year) + '\t' + str(month)) 
                pass


    print(dfall)
    
    #dir = os.getcwd()
    dir = '/content/gdrive/My Drive'
    dir = os.path.join(dir, 'data')
    output_dir = os.path.join(dir, 'aiseed', 'mops_iic')

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    output_dir_filename_csv = os.path.join(output_dir, 'mops_iic_' + co_id + '.csv')  
    output_dir_filename_xlsx = os.path.join(output_dir, 'mops_iic_' + co_id + '.xlsx')
    dfall.to_csv(output_dir_filename_csv, sep = '\t')
    dfall.to_excel(output_dir_filename_xlsx)
    return dfall

#金融保險業 (Subtotal: 35)
#半導體業  (Subtotal: 74)
#https://mops.twse.com.tw/mops/web/t51sb01
#twse_ind_stocks = [2302, 2303, 2329, 2330, 2337, 2338, 2342, 2344, 2351, 2363, 2369, 2379, 2388, 2401, 2408, 2434, 2436, 2441, 2449, 2451, 2454, 2458, 2481, 3006, 3014, 3016, 3034, 3035, 3041, 3054, 3094, 3189, 3257, 3413, 3443, 3530, 3532, 3536, 3545, 3583, 3588, 3661, 3686, 3711, 4919, 4952, 4961, 4967, 4968, 5269, 5285, 5471, 6202, 6239, 6243, 6257, 6271, 6415, 6451, 6515, 6525, 6531, 6533, 6552, 6573, 6756, 8016, 8028, 8081, 8110, 8131, 8150, 8261, 8271]
#twse_ind_stocks = [2330, 2883]
#full list:


url = 'http://isin.twse.com.tw/isin/C_public.jsp?strMode=2'
response = requests.get(url)

df = pd.read_html(response.text)[0]

df.columns = df.iloc[0]
df = df.iloc[2:949]

rawidlist = df['有價證券代號及名稱'].tolist()
twse_ind_stocks = []
for unit in rawidlist:
    for word in unit.split():
        if word.isdigit():
            twse_ind_stocks.append(word)

dir = '/content/gdrive/My Drive'
dir = os.path.join(dir, 'data')
stockidlist = os.path.join(dir, 'stockidlist.txt')
stock_id_list(stockidlist, twse_ind_stocks)


print(datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), ' Start...')
print('len(twse_ind_stocks):', len(twse_ind_stocks)) 
        
i = 0 
exception_count = 0
for stockid in twse_ind_stocks:
    try:
        co_id = str(stockid)
        df = get_mops_iic(co_id)
        if (len(df)>0):
            if (i == 0):
                df_inds = df
            else:
                df_inds = df_inds.append(df)
            i += 1
            print('df_inds i:', i)
            print(df.info())
    except Exception as e:
        print("Error Message: " + str(e))
        print("traceback.format_exc(): ", traceback.format_exc())
        exception_count += 1
        dir = '/content/gdrive/My Drive'
        dir = os.path.join(dir, 'data')
        logfilename = os.path.join(dir, 'logfile.txt')
        log_file_utf8(logfilename, datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + '\tError Message:\t' + str(e)) 
        log_file_utf8(logfilename, datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + '\tError Stockid:\t' + str(stockid)) 
        pass

print('df_inds', df_inds.info())

#dir = os.getcwd()
dir = '/content/gdrive/My Drive'
dir = os.path.join(dir, 'data')
output_dir = os.path.join(dir, 'aiseed', 'mops_iic')
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

output_dir_filename_csv = os.path.join(output_dir, 'mops_iic_' + 'C_2330_2883' + '.csv')  
output_dir_filename_xlsx = os.path.join(output_dir, 'mops_iic_' + 'C_2330_2883' + '.xlsx')
df_inds.to_csv(output_dir_filename_csv, sep = '\t')
df_inds.to_excel(output_dir_filename_xlsx)
print('exception count:', exception_count)
print(datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), ' OK!')

