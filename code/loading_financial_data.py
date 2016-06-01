"""
Created on Tue May 31 15:25:13 2016

Loading data from web
- Data from Daum finance 

@author: Minhyun Yoo
"""
from urllib.request import urlopen
from bs4 import BeautifulSoup
from datetime import datetime
from pytz import timezone
import pandas as pd

def dataframe2Excel(target, filename, sheet_str):
    # dataframe to Excel by using pandas
    writer = pd.ExcelWriter(filename, engine = 'xlsxwriter');    
    zipped = zip(target, sheet_str);    
    for data, string in zipped: 
        data.to_excel(writer, sheet_name = string);   
    writer.save();
    
def get_soup(url):
    # soup data
    html = urlopen(url);
    soup = BeautifulSoup(html, 'lxml', from_encoding='utf-8');
    
    return soup;


def value_iterator(soup, target1, target2, target3):
    dic = {};
    i = 0;
    for link in soup.find_all(target1, {target2: target3}):
        try :
            # idx name
            if (target3 == 'name'):
                val = str(link.contents[0].contents[0]);
            else:
                val = str(link.contents[0]);
        except:
            val = '';
        # float value
        if (target3 == 'idx'):
            val = float(val.replace(',', ''));            
        # save
        dic[i] = val;
        i += 1;
    return dic;


def get_wolrd_sec(wolrd_sec_url):
    soup = get_soup(wolrd_sec_url);
    
    ################### 세계증시 ######################
    ################## Quote_name ######################
    name_dic = value_iterator(soup, 'th', 'class', 'name');    
    wolrd_sec_data = pd.DataFrame({'name' : name_dic});    
    ################## Quote_values ###################### 
    val_dic = value_iterator(soup, 'td', 'class', 'idx');
    wolrd_sec_data['value'] = val_dic.values();
    ################## Quote_rate_fluc ######################
    rate_dic = value_iterator(soup, 'td', 'class', 'rate_fluc');
    wolrd_sec_data['rate_fluc'] = rate_dic.values();   
    
    return wolrd_sec_data;


def get_rate(rate_url):
    soup = get_soup(rate_url);
        
    ################### 원자재 금리 ######################
    ################## Quote_name ######################
    name_dic = value_iterator(soup, 'th', 'class', 'name');    
    rate_data = pd.DataFrame({'name' : name_dic});    
    ################## Quote_values ###################### 
    val_dic = value_iterator(soup, 'td', 'class', 'idx');
    rate_data['value'] = val_dic.values();
    ################## Quote_rate_fluc ######################
    rate_dic = value_iterator(soup, 'td', 'class', 'rate_fluc');
    rate_data['rate_fluc'] = rate_dic.values();
    
    return rate_data;
    
    
def get_fx(fx_url):
    soup = get_soup(fx_url);
    
    ################### 환율 ######################
    ################## Quote_name ######################
    name_dic = value_iterator(soup, 'td', 'class', 'name');    
    fx_data = pd.DataFrame({'name' : name_dic});    
    ################## Quote_values ###################### 
    val_dic = value_iterator(soup, 'td', 'class', 'idx');
    fx_data['value'] = val_dic.values();
    ################## Quote_rate_fluc ######################
    rate_dic = value_iterator(soup, 'td', 'class', 'rate_fluc');
    fx_data['rate_fluc'] = rate_dic.values();    
    
    return fx_data;
    

filename = 'test.xlsx';

########################### fixed urls ###########################
world_sec_url = 'http://m.finance.daum.net/m/world/index.daum?type=default';
rate_url = 'http://m.finance.daum.net/m/world/market.daum?type=cm';
fx_url = 'http://m.exchange.daum.net/mobile/exchange/exchangeMain.daum';
##################################################################

now_seoul = datetime.now(timezone('Asia/Seoul'))

# get data
rate_data = get_rate(rate_url);
fx_data = get_fx(fx_url);
wolrd_sec_data = get_wolrd_sec(world_sec_url);

dataframe2Excel([rate_data, fx_data, wolrd_sec_data], 
                filename, ['RATE','FX', 'WORLDSEC']);

print(now_seoul)