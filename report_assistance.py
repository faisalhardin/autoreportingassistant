#!/usr/bin/env python
# coding: utf-8

# In[1]:


import requests
import pandas as pd
from bs4 import BeautifulSoup
from enum import Enum, auto
import numpy as np
import urllib.request 
import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import functools as ft
import numpy as np
import os


# In[2]:


pages = {
        "merchant":"http://192.168.88.243:3000/merchant",
        "summary_report":"http://192.168.88.243:3000/summary_new",
        "channel_report":"http://192.168.88.243:3000/channel_new",
        "bill_payment":"http://192.168.88.243:3000/billpayment_new"        
        }
base_url = "http://192.168.88.243:3000/"
date_today = datetime.datetime.today().strftime('%Y%m%d')
try:
    os.remove( date_today+"/"+"report"+date_today+".txt")
except:
    pass
folder_name = "report-"+date_today
try:
    os.mkdir(folder_name)
except:
    pass


# In[3]:


def handleNonInteger(inp):
    try:
        return int(inp)
    except:
        return 0


# ## MERCHANT

# In[4]:


merchant_page = requests.get(pages["merchant"])
merchant_page


# In[5]:


soup = BeautifulSoup(merchant_page.content, 'html.parser')


# In[6]:


table = soup.find_all('div', class_="jumbotron")[0]
table_rows = table.find_all('div', class_="row")
l = []
for tr in table_rows:
    td = tr.find_all(class_='col')
    row = [tr.text for tr in td]
    
    link = tr.find('button',href=True)
    try:
        print(link['href'])
        row += [link['href'].replace(" ", "%20")]
    except:
        row += ["Link"]
    l.append(row)

df = pd.DataFrame(l[1:],columns=l[0])
df = df.rename(columns={' ': 'Status', 'CAE(%)':'CAE', 'SAR(%)':'SAR', '# of Transaction':'Num_of_trx'})
print(df)


# In[7]:


class Status(Enum):
    SKIP = "skip"
    NORMAL = "normal"
    SWITCH_CHECK = "switch check"
    SUBMIT_RC = "submit rc"
    ERR = "err"
    


# In[8]:


report = {
    Status.SWITCH_CHECK : [],
    Status.SUBMIT_RC : []
}


# In[9]:


comparable = {
    'NPG Prima' : {'CAE':85, 'SAR':90},
    'NPG Bersama' : {'CAE':85, 'SAR':90},
    'VISA Local' : {'CAE':90, 'SAR':90},
    'VISA Overseas' : {'CAE':85, 'SAR':90},
}


# In[10]:


def status(comparable, name, num_of_trx, cae, sar, basic_num=0):
#     print("here")
    print("1",type(comparable))
    print( "2", type(name))
    print("3" , type(num_of_trx))
    print("4", type(cae))
    print("5", type(sar))
    if(num_of_trx) == "":
        return ""
    elif int(num_of_trx) < basic_num:
        return Status.SKIP
    elif np.std([int(sar),comparable[name]["SAR"]], ddof=1)>5 and int(sar)<comparable[name]["SAR"]:
        return Status.SWITCH_CHECK
    elif np.std([int(cae),comparable[name]["CAE"]], ddof=1)>10 and int(cae)<comparable[name]["CAE"]:
        return Status.SUBMIT_RC
    else:
        return Status.NORMAL


# In[11]:


for i, row in df.iterrows():
    print(row["Channels"], row["Num_of_trx"], row["CAE"], row["SAR"])
    df.at[i,"Status"] = status(comparable, row["Channels"], row["Num_of_trx"], row["CAE"], row["SAR"])


# In[12]:



print(df)
for i, row in df.iterrows():
    print(row["Status"], row["Status"] == Status.SWITCH_CHECK)
    if row["Status"] == Status.SUBMIT_RC or row["Status"] == Status.SWITCH_CHECK:
        print(row, Status.SUBMIT_RC)
        report[row["Status"]].append(row["Channels"]+"\n")
        image_link = pages["merchant"]+"_detail/"+row["Link"]
        path = date_today+"-"+row['Channels'].lower().replace(" ", "-")
        print(image_link, path)
        with webdriver.Chrome('chromedriver') as driver:
            
            driver.get(image_link)
            driver.implicitly_wait(5000)
            retry = 100
            while retry > 0:
                try:
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//div[@class="container"]'))
                    )
                    time.sleep(5)
                    element.screenshot(folder_name+"/"+path+".png")
                    break
                except Exception as e:
                    retry -= 1
                    print(e)


# In[13]:


df.to_excel(folder_name+"/"+"merchant"+date_today+".xlsx")
print(report.items())
with open(folder_name+"/"+"report"+date_today+".txt","a") as file:
    file.write("A. Card Report\n")
    for key, status in report.items():
        if len(status) == 0:
            continue
        file.write(key.value+"\n")
        
        for idx, item in enumerate(status):
            file.write(str(idx+1) + ". "+ item)
            if status is Status.SUBMIT_RC:
                file.write("< insert "+path+" >")
            
        file.write("\n")


# ## Summary Report

# In[14]:


summary_table = [
    ["Mobile Banking(MBK)",95,0,0], 
    ["Personal Banking(PBK)",95,0,0], 
    ["Corporate Banking(CBS)",95,0,0], 
    ["sms banking",80,90,10], 
    ["ATM Driving", 85,90,0], 
    ["EDC",80,90,10],
    ["Phone Banking",60,90,10], 
    ["Jaringan PRIMA",85,90,0], 
    ["ATM Bersama",85,90,0], 
    ["Tiphone",85,90,0], 
    ["Nicepay",90,90,0],
    ["Finnet",90,90,0], 
    ["Euronet",90,90,0], 
    ["Maybank",90,90,10], 
    ["Visa",85,90,0], 
    ["Bimasakti",85,90,0], 
    ["Dimo",85,90,0]]

summary_dict = {}
for el in summary_table:
    summary_dict[el[0].lower()] = {
        "CAE":el[1],
        "SAR":el[2]
    }
summary_dataframe = pd.DataFrame(summary_table, columns=['Channels', 'CAE', 'SAR', 'lower_base'])
summary_dataframe["Status"] = ""
summary_dataframe["Link"] = ""


# In[15]:


summary_dataframe


# In[16]:


summary_report = requests.get(pages["summary_report"])
summary_report


# In[17]:


soup = BeautifulSoup(summary_report.content, 'html.parser')


# In[18]:


table = soup.find_all('div', class_="jumbotron")[0]
table_rows = table.find_all('div', class_="row")
l = []
for tr in table_rows:
    td = tr.find_all(class_='col')
    row = [tr.text.strip().lower() for tr in td]
    
    link = tr.find('button',href=True)
    try:
        print(link['href'])
        row += [link['href'].replace(" ", "%20")]
    except:
        row += ["Link"]
    l.append(row)
print(l[0])
df = pd.DataFrame(l[1:],columns=l[0])
df = df.rename(columns={'': 'Status', 'cae(%)':'CAE', 'sar(%)':'SAR', '# of transaction':'Num_of_trx','channels':'Channels'})
df


# In[19]:


def set_status(comparable, name, num_of_trx, cae, sar, basic_num=0):
    
    try:
        print(name, int(cae), comparable[name]["CAE"], np.std([int(cae),comparable[name]["CAE"]], ddof=1), np.std([int(cae),comparable[name]["CAE"]], ddof=1)>10, int(cae)<comparable[name]["CAE"],np.std([int(cae),comparable[name]["CAE"]], ddof=1)>10 and int(cae)<comparable[name]["CAE"] )
        if(num_of_trx) == "":
            return ""
        elif int(num_of_trx) < basic_num:
            return Status.SKIP
        elif np.std([int(sar),comparable[name]["SAR"]], ddof=1)>5 and int(sar)<comparable[name]["SAR"]:
            return Status.SWITCH_CHECK
        elif np.std([int(cae),comparable[name]["CAE"]], ddof=1)>10 and int(cae)<comparable[name]["CAE"]:
            return Status.SUBMIT_RC
        else:
            return Status.NORMAL
    except:
        return Status.ERR


# In[20]:


# for i, row in summary_dataframe.iterrows():
#     print(row["Link"])
# print(l[1:])
df


# In[21]:


summary_dataframe


# In[22]:


summary_report = {
    Status.SWITCH_CHECK : [],
    Status.SUBMIT_RC : []
}

summary_status = []
# print(summary_dataframe.columns)
for i, row in summary_dataframe.iterrows():
    table_row = df[df.Channels == row["Channels"].lower()]
    if not table_row.empty:
        xx = set_status(summary_dict, table_row["Channels"].values[0], table_row["Num_of_trx"].values[0], table_row["CAE"].values[0], table_row["SAR"].values[0], row['lower_base'])
        summary_dataframe.at[i,"Status"] = xx
        summary_dataframe.at[i,"Link"] = table_row["Link"].values[0]
        summary_dataframe.at[i,"CAE"] = table_row["CAE"].values[0]
        summary_dataframe.at[i,"SAR"] = table_row["SAR"].values[0]
        summary_dataframe.at[i,"Num_of_trx"] = table_row["Num_of_trx"].values[0]
#         print(i, df.at[i,"Link"], row["Link"])
    else:
        summary_dataframe.at[i,"Status"] = ""
        summary_dataframe.at[i,"CAE"] = 0
        summary_dataframe.at[i,"SAR"] = 0
        summary_dataframe.at[i,"Num_of_trx"] = 0
    table_row = ""


# In[23]:


cols = summary_dataframe.columns.tolist()
cols = [cols[0]] + [cols[-1]]+ cols[1:-1] 
temp = summary_dataframe[cols]
summary_dataframe = temp


# In[24]:


report_summary = {
    Status.SWITCH_CHECK : [],
    Status.SUBMIT_RC : []
}
summary_dataframe


# In[25]:


summary_dataframe.to_excel(folder_name+"/"+"summary"+date_today+".xlsx")
for i, row in summary_dataframe.iterrows():
    print(row["Status"], row["Status"] == Status.SWITCH_CHECK)
    if row["Status"] == Status.SUBMIT_RC or row["Status"] == Status.SWITCH_CHECK:
        print(row, row["Status"] )
        report_summary[row["Status"]].append(row["Channels"]+"\n")
        image_link = base_url+"summary_detail"+row["Link"]
        path = date_today+"-"+row['Channels'].lower().replace(" ", "-")
        print(image_link, path)
        with webdriver.Chrome('chromedriver') as driver:
            
            driver.get(image_link)
            driver.implicitly_wait(5000)
            retry = 100
            while retry > 0:
                try:
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//div[@class="container"]'))
                    )
                    time.sleep(5)
                    element.screenshot(folder_name+"/"+path+".png")
                    break
                except Exception as e:
                    retry -= 1
                    print(e)


# In[26]:


print(report_summary.items())
with open(folder_name+"/"+"report"+date_today+".txt","a") as file:
    file.write("B. Sumarry Report\n")
    for key, status in report_summary.items():
        if len(status) == 0:
            continue
        file.write(key.value+"\n")
        print(key.value+"\n")
        
        for idx, item in enumerate(status):
            file.write(str(idx+1) + ". "+ item)
            print(str(idx+1) + ". "+ item)
            if status is Status.SUBMIT_RC and status is Status.SWITCH_CHECK:
                file.write("< insert "+path+" >")
                print("< insert "+path+" >")
            
        file.write("\n")


# ## Channel Report

# In[27]:


channel_report = [
    ["ATM","NBALHNB",85,90,10],
    ["ATM","NWDLHNB",85,90,10],
    ["ATM","NTRHWOD",85,90,10], 
    ["ATM","NTRHWHD",85,90,10], 
    ["ATM","NBLLCHK",85,90,10], 
    ["ATM","NBLLPAY",85,90,10], 
    ["ATM","NDPSACT",85,90,10],
    ["ATM","NWDLCRL",85,90,10], 
    ["ATM","PINCHG1",85,90,10], 
    ["ATM","NBALOTR",85,90,10], 
    ["ATM","NWDLOTR",85,90,10],
    ["ATM","NTROWOD",85,90,10], 
    ["ATM","NTROWHD",85,90,10],
    ["PBK","0903A01",95,0,10],
    ["PBK","0101A01",95,0,10],
    ["PBK","0507A01",95,0,10],
    ["PBK","0501A01",95,0,10],
    ["PBK","0420A01",95,0,10],
    ["PBK","NBLLCHK",95,0,10],
    ["PBK","0520A01",95,0,10],
    ["PBK","NBLLPOS",95,0,10],
    ["MBK","0903A01",95,0,10],
    ["MBK","0101A01",95,0,10],
    ["MBK","0507A01",95,0,10],
    ["MBK","0501A01",95,0,10],
    ["MBK","0371A01",95,0,10],
    ["MBK","0372A01",95,0,10],
    ["MBK","0420A01",95,0,10],
    ["MBK","NBLLCHK",95,0,10],
    ["MBK","0520A01",95,0,10],
    ["MBK","NBLLPOS",95,0,10],
    ["CBS","1903A01",95,0,10],
    ["CBS","1507A01",95,0,10],
    ["CBS","1501A01",95,0,10],
]

channel_report_dict = {}
for el in channel_report:
    channel_report_dict[el[0].lower()+el[1].lower()] = {
        "CAE":el[2],
        "SAR":el[3]
    }
channel_report_dataframe = pd.DataFrame(channel_report, columns=['Channels', 'Transaction_Type', 'CAE', 'SAR', 'lower_base'])
channel_report_dataframe["Status"] = ""
channel_report_dataframe["Link"] = ""


# In[28]:


channel_report_dataframe


# In[29]:


scraped_channel_report = requests.get(pages["channel_report"])
soup = BeautifulSoup(scraped_channel_report.content, 'html.parser')


# In[30]:


table = soup.find_all('div', class_="jumbotron")
processed_scrapped_tables = []
for each_table in table:
    table_rows = each_table.find_all('div', class_="row")
    l = []
    for tr in table_rows:
        td = tr.find_all(class_='col')
        row = [tr.text.strip().lower() for tr in td]

        link = tr.find('button',href=True)
        try:
            print(link['href'])
            row += [link['href'].replace(" ", "%20")]
        except:
            row += ["Link"]
        l.append(row)
    print(l[0])
    df = pd.DataFrame(l[1:],columns=l[0])
    df = df.rename(columns={"":"Status",'transaction code': 'Trx_code', 'cae(%)':'CAE', 'sar(%)':'SAR', '# of transaction':'Num_of_trx', 'channel':'Channels'})
    processed_scrapped_tables.append(df)


# In[31]:


processed_scrapped_tables[1]


# In[32]:


channel_report_dataframe


# In[33]:


channel_report_dataframe[(channel_report_dataframe.Channels == "ATM") & (channel_report_dataframe.Transaction_Type == "NBALHNB")]


# In[34]:


channel_report = {
    Status.SWITCH_CHECK : [],
    Status.SUBMIT_RC : []
}

channel_status = []
Transaction_Summary = pd.DataFrame(l[1:],columns=l[0])
for i, row in channel_report_dataframe.iterrows():
    source_table = processed_scrapped_tables[1]
#     print(row["Channels"] == "ATM",row["Channels"])
    if row["Channels"] == "ATM":
        source_table = processed_scrapped_tables[0]
        table_row = source_table[source_table.Trx_code == row["Transaction_Type"].lower()]
#     print(source_table[source_table.Trx_code == row["Transaction_Type"].lower()])
    else:
        table_row = source_table[(source_table.Trx_code == row["Transaction_Type"].lower()) & (source_table.Channels == row["Channels"].lower())]
#         print(source_table.Trx_code == row["Transaction_Type"].lower(), source_table.Channel == row["Channel"].lower(), source_table.Trx_code == row["Transaction_Type"].lower() and source_table.Channels == row["Channels"].lower())
#     print(channel_report_dict, table_row["Trx_code"].values[0], table_row["Num_of_trx"].values[0], table_row["CAE"].values[0], table_row["SAR"].values[0], row['lower_base'])
#     print(table_row)
#     print(table_row["Trx_code"])
    if not table_row.empty:
        __channel = "0"
        try:
            __channel = row["Channels"].lower()
        except:
            continue
        stat  = set_status(channel_report_dict, __channel.lower()+table_row["Trx_code"].values[0], table_row["Num_of_trx"].values[0], table_row["CAE"].values[0], table_row["SAR"].values[0],10)
#         print(__channel)
        channel_report_dataframe.at[i,"Status"] = stat##
        channel_report_dataframe.at[i,"Num_of_trx"] = handleNonInteger(table_row["Num_of_trx"].values[0])
        channel_report_dataframe.at[i,"CAE"] = handleNonInteger(table_row["CAE"].values[0])
        channel_report_dataframe.at[i,"SAR"] = handleNonInteger(table_row["SAR"].values[0])
        channel_report_dataframe.at[i,"Link"] = table_row["Link"].values[0]
#         print(table_row["Num_of_trx"].values[0])
#         channel_report_dataframe.at[i,"Link"] = table_row["Link"].values[0]
    else:
        channel_report_dataframe.at[i,"Status"] = ""
        channel_report_dataframe.at[i,"Num_of_trx"] = 0
        channel_report_dataframe.at[i,"CAE"] = 0
        channel_report_dataframe.at[i,"SAR"] = 0
        channel_report_dataframe.at[i,"Link"] = ""


# In[35]:


# channel_report_dataframe["Num_of_trx"]
# processed_scrapped_tables
# print(source_table)
# processed_scrapped_tables[1]
# cols = channel_report_dataframe.columns.tolist()
# cols2 = cols[0:2] + [cols[-1]]+ cols[2:-1] 
# cols2


# In[36]:


cols = channel_report_dataframe.columns.tolist()
cols = cols[0:2] + [cols[-1]]+ cols[2:-1] 
temp = channel_report_dataframe[cols]
# channel_report_dataframe = temp
# channel_report_dataframe
# temp["Num_of_trx"] = temp["Num_of_trx"]##stype(int)
channel_report_dataframe = temp


# In[37]:


channel_report_dataframe


# In[38]:


# print(cols)
# cols = cols[:2]+[cols[-1]]+cols[2:-1]
# cols


# In[39]:


report_channel = {
    Status.SWITCH_CHECK : [],
    Status.SUBMIT_RC : []
}


# In[40]:


channel_report_dataframe.to_excel(folder_name+"/"+"channel"+date_today+".xlsx")
for i, row in channel_report_dataframe.iterrows():
    print(row["Status"], row["Status"] == Status.SWITCH_CHECK)
    if row["Status"] == Status.SUBMIT_RC or row["Status"] == Status.SWITCH_CHECK:
        print(row, row["Status"] )
        report_channel[row["Status"]].append(row["Channels"]+ "-" +row["Transaction_Type"] +"\n")
        image_link = base_url+"channel_detail"+row["Link"]
        path = date_today+"-"+row['Channels'].lower().replace(" ", "-")+"-"+row["Transaction_Type"]
        print(image_link, path)
        with webdriver.Chrome('chromedriver') as driver:
            
            driver.get(image_link)
            driver.implicitly_wait(5000)
            retry = 100
            while retry > 0:
                try:
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//div[@class="container"]'))
                    )
                    time.sleep(5)
                    element.screenshot(folder_name+"/"+path+".png")
                    break
                except Exception as e:
                    retry -= 1
                    print(e)


# In[41]:


report_channel


# In[42]:


print(report_channel.items())
with open(folder_name+"/"+"report"+date_today+".txt","a") as file:
    file.write("C.  Focus Service Report\n")
    for key, status in report_channel.items():
        if len(status) == 0:
            continue
        file.write(key.value+"\n")
        print(key.value+"\n")
        
        for idx, item in enumerate(status):
            file.write(str(idx+1) + ". "+ item)
            print(str(idx+1) + ". "+ item)
            if status is Status.SUBMIT_RC and status is Status.SWITCH_CHECK:
                file.write("< insert "+path+" >")
                print("< insert "+path+" >")
            
        file.write("\n")


# ## BILL PAYMENT METHOD

# In[43]:


bill_payment_report = [
    ["atm","telkomsel simpati prepaid",85,90,0],
    ["atm","telkomsel voucher internet",85,90,0],
    ["atm","xl prepaid",85,90,0],
    ["atm","indosat prepaid (im3)",85,90,0],
    ["atm","three (3) prepaid",85,90,0],
    ["atm","pln prabayar",85,90,0],
    ["atm","telkom group",85,90,0],
    ["atm","keb hana credit card",85,90,0],
    ["pbk","telkomsel simpati prepaid",85,0,0],
    ["pbk","telkomsel voucher internet",85,0,0],
    ["pbk","xl prepaid",85,0,0],
    ["pbk","indosat prepaid (im3)",85,0,0],
    ["pbk","three (3) prepaid",85,0,0],
    ["pbk","pln prabayar",85,0,0],
    ["pbk","keb hana credit card",85,0,0],
    ["pbk","citibank credit card",85,0,0],
    ["pbk","kereta api",85,0,0],
    ["pbk","garuda indonesia",85,0,0],
    ["pbk","pln pascabayar",85,0,0],
    ["pbk","telkom group",85,0,0],
    ["mbk","telkomsel simpati prepaid",85,0,0],
    ["mbk","telkomsel voucher internet",85,0,0],
    ["mbk","xl prepaid",85,0,0],
    ["mbk","indosat prepaid (im3)",85,0,0],
    ["mbk","three (3) prepaid",85,0,0],
    ["mbk","pln prabayar",85,0,0],
    ["mbk","keb hana credit card",85,0,0],
    ["mbk","citibank credit card",85,0,0],
    ["mbk","kereta api",85,0,0],
    ["mbk","garuda indonesia",85,0,0],
    ["mbk","pln pascabayar",85,0,0],
    ["mbk","telkom group",85,0,0],
    
]

bill_payment_dict = {}
for el in bill_payment_report:
    bill_payment_dict[el[0].lower()+el[1].lower()] = {
        "CAE":el[2],
        "SAR":el[3]
    }
bill_payment_dataframe = pd.DataFrame(bill_payment_report, columns=['Channels', 'Product', 'CAE', 'SAR', 'lower_base'])
bill_payment_dataframe["Status"] = ""
bill_payment_dataframe["Link"] = ""


# In[44]:


bill_payment_report


# In[45]:


bill_payment_report = requests.get(pages["bill_payment"])
soup = BeautifulSoup(bill_payment_report.content, 'html.parser')


# In[46]:


table = soup.find_all('div', class_="jumbotron")
processed_scrapped_tables = []
for each_table in table:
    table_rows = each_table.find_all('div', class_="row")
    l = []
    for tr in table_rows:
        td = tr.find_all(class_='col')
        row = [tr.text.strip().lower() for tr in td]

        link = tr.find('button',href=True)
        try:
            print(link['href'])
            row += [link['href'].replace(" ", "%20")]
        except:
            row += ["Link"]
        l.append(row)
    print(l[0])
    df = pd.DataFrame(l[1:],columns=l[0])
    df = df.rename(columns={"":"Status",'transaction code': 'Trx_code', 'cae(%)':'CAE', 'sar(%)':'SAR', '# of transaction':'Num_of_trx', 'channel':'Channels'})
    processed_scrapped_tables.append(df)


# In[47]:


df


# In[48]:


processed_scrapped_tables[1]


# In[49]:


bill_payment_report = {
    Status.SWITCH_CHECK : [],
    Status.SUBMIT_RC : []
}
bill_payment_status = []


# In[50]:


processed_scrapped_tables[1]


# In[51]:


source_table


# In[52]:


Transaction_Summary = pd.DataFrame(l[1:],columns=l[0])
for i, row in bill_payment_dataframe.iterrows():
    source_table = processed_scrapped_tables[1]
    if row["Channels"] == "atm":
        source_table = processed_scrapped_tables[0]
        
        table_row = source_table[(source_table["product"] == row["Product"].lower()) &
                                 (source_table.type == "payment/purchase")]
#         print(source_table.product, row["Product"].lower(), source_table.product == row["Product"].lower())
    else:
        table_row = source_table[(source_table["product"] == row["Product"].lower()) & 
                                 (source_table.Channels == row["Channels"].lower()) &
                                 (source_table.type == "payment/purchase")
                                ]
#     print(table_row)
#     print(source_table["product"])
#     print(row["Product"].lower())
#     print(source_table.type)
#     print("==============")
#     print(table_row)
    if not table_row.empty:
        __channel = "0"
        try:
            __channel = row["Channels"].lower()
        except:
            continue
        stat  = set_status(bill_payment_dict, __channel.lower()+table_row["product"].values[0], table_row["Num_of_trx"].values[0], table_row["CAE"].values[0], table_row["SAR"].values[0], 10)
        bill_payment_dataframe.at[i,"Status"] = stat##
        bill_payment_dataframe.at[i,"Num_of_trx"] = table_row["Num_of_trx"].values[0]
        bill_payment_dataframe.at[i,"CAE"] = table_row["CAE"].values[0]
        bill_payment_dataframe.at[i,"SAR"] = table_row["SAR"].values[0]
        bill_payment_dataframe.at[i,"Link"] = table_row["Link"].values[0]
    else:
        bill_payment_dataframe.at[i,"Status"] = ""
        bill_payment_dataframe.at[i,"Num_of_trx"] = 0
        bill_payment_dataframe.at[i,"CAE"] = 0
        bill_payment_dataframe.at[i,"SAR"] = 0
        bill_payment_dataframe.at[i,"Link"] = ""


# In[53]:


bill_payment_dataframe


# In[54]:


cols = bill_payment_dataframe.columns.tolist()
cols = cols[0:2] + [cols[-1]]+ cols[2:-1] 
temp = bill_payment_dataframe[cols]
# channel_report_dataframe = temp
# channel_report_dataframe
# temp["Num_of_trx"] = temp["Num_of_trx"]##stype(int)
bill_payment_dataframe = temp


# In[55]:


report_bill_payment = {
    Status.SWITCH_CHECK : [],
    Status.SUBMIT_RC : []
}


# In[56]:


bill_payment_dataframe.to_excel(folder_name+"/"+"bill_payment"+date_today+".xlsx")
for i, row in bill_payment_dataframe.iterrows():
    print(row["Status"], row["Status"] == Status.SWITCH_CHECK)
    if row["Status"] == Status.SUBMIT_RC or row["Status"] == Status.SWITCH_CHECK:
        print(row, row["Status"] )
        report_bill_payment[row["Status"]].append(row["Channels"]+ "-" +row["Product"] +"\n")
        image_link = base_url+"billpayment_detail"+row["Link"]
        path = date_today+"-"+row['Channels'].lower().replace(" ", "-")+"-"+row['Product'].lower().replace(" ", "-")
        print(image_link, path)
        with webdriver.Chrome('chromedriver') as driver:
            
            driver.get(image_link)
            driver.implicitly_wait(5000)
            retry = 100
            while retry > 0:
                try:
                    element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//div[@class="container"]'))
                    )
                    time.sleep(5)
                    element.screenshot(folder_name+"/"+path+".png")
                    break
                except Exception as e:
                    retry -= 1
                    print(e)



# In[58]:


print(report_bill_payment.items())
with open(folder_name+"/"+"report"+date_today+".txt","a") as file:
    file.write("D. Bill Payment Report\n")
    for key, status in report_bill_payment.items():
        if len(status) == 0:
            continue
        file.write(key.value+"\n")
        print(key.value+"\n")
        
        for idx, item in enumerate(status):
            file.write(str(idx+1) + ". "+ item)
            print(str(idx+1) + ". "+ item)
            if status is Status.SUBMIT_RC and status is Status.SWITCH_CHECK:
                file.write("< insert "+path+" >")
                print("< insert "+path+" >")
            
        file.write("\n")


