'''
Created on 2019/06/11
@author: yu_k
'''

import sys
import os
import time
import glob
import shutil
import datetime
import writeLog
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select

current_Dir = os.path.dirname(os.path.abspath(_file_))

# user.ini読込
strUserPath = os.path.dirname(os.path.abspath(sys.argv[0])) + "\\user.ini"
f = open(strUserPath,'r')

# ファイル終端まで全て読んだデータを返す
date = f.read()
f.close()

# 改行で区切る(改行文字そのものは戻り値のデータには含まれない)
lsUser = data.split('\n')

# .iniファイルの読み込み
for strUser in lsUser:
    if 'SsoUserId' in strUser:
        SsoUserId = strUser.replace('{}:'.format(strUser.split(':')[0]),'')
    elif 'SsoPassword' in strUser:
        SsoPassword = strUser.replace('{}:'.format(strUser.split(':')[0]),'')
    elif 'UserName' in strUser:
        UserName = strUser.replace('{}:'.format(strUser.split(':')[0]),'')
    elif 'DownloadFolder' in strUser:
        DownloadFolder = current_Dir + strUser.replace('{}:'.format(strUser.split(':')[0]),'')
    elif 'Backup_Download' in strUser:
        Backup_Download = current_Dir + strUser.replace('{}:'.format(strUser.split(':')[0]),'')
    elif 'UploadFolder' in strUser:
        UploadFolder = current_Dir + strUser.replace('{}:'.format(strUser.split(':')[0]),'')
    elif 'Backup_Upload' in strUser:
        Backup_Upload = current_Dir + strUser.replace('{}:'.format(strUser.split(':')[0]),'')
    elif 'ErrorFile_Folder' in strUser:
        ErrorFile_Folder = current_Dir + strUser.replace('{}:'.format(strUser.split(':')[0]),'')
    elif 'ChromeDriver' in strUser:
        ChromeDriver = current_Dir + strUser.replace('{}:'.format(strUser.split(':')[0]),'')
    elif 'LogFile' in strUser:
        LogFile = current_Dir + strUser.replace('{}:'.format(strUser.split(':')[0]),'')

####################################################################################

# Log記入
now = datetime.datetime.now()
LOG_File = writeLog.Log(log_File)
LOG_File.Write(now.strftime("%Y/%m/%d %H:%M:%S") + '■OPEN-UI DUKeView ダウンロード開始')

#==========================================
strUrl = 'https://www.xxxxxxxxx'
page0_User = '//*[@id="IDToken1"]'
page0_Pass = '//*[@id="IDToken2"]'
page0_Login = '//*[@id="loginFormButton"]'
page1_Logo = '//*[@id="logoIcon"]'
page1_Menu = '//*[@id="menuBtn"]'
#==========================================

# DownLoadフォルダの指定をして、Chromeを開く
chromeOptions = webdriver.ChromeOptions()
chromeOptions.add_experimental_option("prefs",{"download.default_directory": DownloadFolder})
browser = webdriver.Chrome(executable_path = ChromeDriver, chrome_options = chromeOptions)
browser.maximize_window()
browser.get(strUrl)

# page0-----------------------------------------------------------------
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.XPATH,page0_User)))
time.sleep(1)
browser.find_element_by_xpath(page0_User).send_keys(SsoUserId)
browser.find_element_by_xpath(page0_Pass).send_keys(SsoPassword)
browser.find_element_by_xpath(page0_Login).click()

# page1-----------------------------------------------------------------
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.XPATH,page1_Logo)))
time.sleep(1)

# 工程管理を選択
dotCount = len(browser.find_element_by_class_name('slick-dots').find_element_by_tag_name('li'))
strMenuNo = ''
while True:
    for i in range(dotCount-1):
        elements = browser.find_element_by_xpath(page1_Menu)
        for ele in elements:
            if ele.text == '工程管理':
                ele.click()
                time.sleep(2)
                strMenuNo = ele.get_attribute('class').split('_')[1]
                time.sleep(2)
                break
        else:
            elements2 = browser.find_element_by_tag_name('button')
            for ele2 in elements2:
                if 'Next' in ele2.text:
                    ele2.click()
                    time.sleep(2)
                    break
            continue
        break
    if strMenuNo != '':
        break

# 工事UIを選択
elements = browser.find_element_by_xpath('//*[@id="submenuId_0"]/div[1]/a'.format(strMenuNo))
for ele in elements:
    if ele.text == '工事UI':
        ele.click()
        time.sleep(2)
        break

# 各項目を選択
elements = browser.find_element_by_xpath('//*[@id="orderList"]')
select = Select(elements)
select.select_by_value('DUKeView')
time.sleep(2)

elements = browser.find_element_by_xpath('//*[@id="areaList"]')
select = Select(elements)
select.select_by_value('東京全域')
time.sleep(2)

# 田端さんアカウントであれば"コンセプトシート作成日"、それ以外のアカウントであれば"コンセプトシート作成日(田端)を選択"
elements = browser.find_element_by_xpath('//*[@id="myMesenList"]')
select = Select(elements)
ArticleId_select_Op = select.options

for alllist in ArticleId_select_Op:
    if alllist.text == 'コンセプトシート作成日(田端)':
        select.select_by_visible_text('コンセプトシート作成日(田端)')
        break
    elif alllist.text == 'コンセプトシート作成日':
        select.select_by_visible_text('コンセプトシート作成日')
        break
time.sleep(2)

# 田端さんアカウントであれば"コンセプトシート作成日:日付あり"、それ以外のアカウントであれば"コンセプトシート作成日:日付あり(田端)を選択"
elements = browser.find_element_by_xpath('//*[@id="myFilterList"]')
select = Select(elements)
ArticleId_select_Op = select.options

for alllist in ArticleId_select_Op:
    if alllist.text == 'コンセプトシート作成日:日付あり(田端)':
        select.select_by_visible_text('コンセプトシート作成日:日付あり(田端)')
        break
    elif alllist.text == 'コンセプトシート作成日:日付あり':
        select.select_by_visible_text('コンセプトシート作成日:日付あり')
        break
time.sleep(2)

# 表示ボタンを押下
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="doDisplay"]')))
browser.find_element_by_xpath('//*[@id="doDisplay"]').click()
time.sleep(1)

# ダウンロードボタン押下
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="DownLoadBtn"]')))
browser.find_element_by_xpath('//*[@id="DownLoadBtn"]').click()
time.sleep(1)

# DownLoadフォルダに格納されたことを確認し、閉じる
while True:
    lsFile = glob.glob('{}/*.*'.format(DownloadFolder))
    if len(lsFile) != 0:
        if ('.crdownload' in lsFile[0]) == False and ('.tmp' in lsFile[0]) == False:
            break
    time.sleep(1)

time.sleep(5)

# Log記入
now = datetime.datetime.now()
LOG_File = writeLog.Log(log_File)
LOG_File.Write(now.strftime("%Y/%m/%d %H:%M:%S") + '■OPEN-UI DUKeView ダウンロード完了')

browser.close()
browser.quit()

####################################################################################

# Log記入
now = datetime.datetime.now()
LOG_File = writeLog.Log(log_File)
LOG_File.Write(now.strftime("%Y/%m/%d %H:%M:%S") + '■アップロードファイル作成　開始')

# DLファイルの取得
df = (glob.glob(r"{}\*.xlsx".format(DownloadFolder)))
wb = openpyxl.load_workbook(df[0])

# DLファイルシート名の変更
ws = wb.active
ws.title = "工程管理情報"

# ULファイルとして保存
wb.save(r"{}\DUKeView_Up.xlsx".format(UploadFolder))

# ULファイルの取得
uf = (glob.glob(r"{}\*.xlsx".format(UploadFolder)))
ub = openpyxl.load_workbook(uf[0])
us = ub.active

# openpyxlで複製すると、日付が"ユーザー定義"から"シリアル値"となってしまうため、再度日付を書き込み
i = 1
for cell_obj in list(us.columns)[1]:
    if i > 1 and cell_obj.value != "#":
        cell_obj.value = cell_obj.value.strftime("%Y/%m/%d")
        i += 1
    else:
        i += 1

# ULファイルを上書き
ub.save(r"{}\DUKeView_Up.xlsx".format(UploadFolder))

# Log記入
now = datetime.datetime.now()
LOG_File = writeLog.Log(log_File)
LOG_File.Write(now.strftime("%Y/%m/%d %H:%M:%S") + '■アップロードファイル作成　完了')

####################################################################################

# DUKeViewダウンロードファイルのバックアップ
for file in glob.glob('{}/*.xlsx'.format(DownloadFolder)):
    shutil.copy(file,Backup_Download)

# ファイル名をリストで取得
files = os.listdir(DownloadFolder)
files_file = [f for f in files if os.path.isfile(os.path.join(DownloadFolder,f))]

# DUKeViewアップロードファイル名を変更
os.rename(glob.glob('{}/*.xlsx'.format(UploadFolder))[0] , UploadFolder + "\\" + files_file[0])

# DUKeViewアップロードファイルをバックアップフォルダにコピペ
for file in glob.glob('{}/*.xlsx'.format(UploadFolder)):
    shutil.copy(file ,Backup_Upload)

# DUKeViewダウンロードファイルが7個以上になったら、1番古いファイルを削除
dukefiles = glob.glob('{}/*.xlsx'.format(Backup_Download))
if len(dukefiles) > 7:
    dukefiles.sort(key=lambda x: int(os.path.getctime(x)))
    os.remove(koujifiles[0])

# Log記入
now = datetime.datetime.now()
LOG_File = writeLog.Log(log_File)
LOG_File.Write(now.strftime("%Y/%m/%d %H:%M:%S") + '■OPEN-UI DUKeView ダウンロード・アップロードファイルバックアップ　完了')

####################################################################################

# Log記入
now = datetime.datetime.now()
LOG_File = writeLog.Log(log_File)
LOG_File.Write(now.strftime("%Y/%m/%d %H:%M:%S") + '■OPEN-UI DUKeView　アップロード　開始')

#==========================================
strUrl = 'https://www.xxxxxxxxx'
page0_User = '//*[@id="IDToken1"]'
page0_Pass = '//*[@id="IDToken2"]'
page0_Login = '//*[@id="loginFormButton"]'
page1_Logo = '//*[@id="logoIcon"]'
page1_Menu = '//*[@id="menuBtn"]'
#==========================================

# DownLoadフォルダの指定をして、Chromeを開く
chromeOptions = webdriver.ChromeOptions()
chromeOptions.add_experimental_option("prefs",{"download.default_directory": ErrorFile_Folder})
browser = webdriver.Chrome(executable_path = ChromeDriver, chrome_options = chromeOptions)
browser.maximize_window()
browser.get(strUrl)


# page0-----------------------------------------------------------------
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.XPATH,page0_User)))
time.sleep(1)
browser.find_element_by_xpath(page0_User).send_keys(SsoUserId)
browser.find_element_by_xpath(page0_Pass).send_keys(SsoPassword)
browser.find_element_by_xpath(page0_Login).click()

# page1-----------------------------------------------------------------
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.XPATH,page1_Logo)))
time.sleep(1)

# 工程管理を選択
dotCount = len(browser.find_element_by_class_name('slick-dots').find_element_by_tag_name('li'))
strMenuNo = ''
while True:
    for i in range(dotCount-1):
        elements = browser.find_element_by_xpath(page1_Menu)
        for ele in elements:
            if ele.text == '工程管理':
                ele.click()
                time.sleep(2)
                strMenuNo = ele.get_attribute('class').split('_')[1]
                time.sleep(2)
                break
        else:
            elements2 = browser.find_element_by_tag_name('button')
            for ele2 in elements2:
                if 'Next' in ele2.text:
                    ele2.click()
                    time.sleep(2)
                    break
            continue
        break
    if strMenuNo != '':
        break

# 工事UIを選択
elements = browser.find_element_by_xpath('//*[@id="submenuId_0"]/div[1]/a'.format(strMenuNo))
for ele in elements:
    if ele.text == '工事UI':
        ele.click()
        time.sleep(2)
        break

# 進捗一括更新を選択
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="ProgressUpdateFileUpload"]/a/div')))
browser.find_element_by_xpath('//*[@id="ProgressUpdateFileUpload"]/a/div').click()
time.sleep(1)

# Uploadファイル選択
UploadFile = glob.glob('{}/*.*'.format(UploadFolder))
browser.find_element_by_xpath('//*[@id="file"]').send_keys(UploadFile);

# アップロードボタンを押下
WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="id_form"]/div/div/div[3]/input')))
browser.find_element_by_xpath('//*[@id="id_form"]/div/div/div[3]/input').click()
time.sleep(1)

# アップロード状態の確認
flg = 0
flg2 = 0
while flg == 0:
    time.sleep(1)

    elements = browser.find_element_by_xpath('//*[@id="fixed_table"]/table/tbody/tr['+ str(i) +']/td[2]')
    if UserName in elements.text:
        now = datetime.datetime.now()

        elements = browser.find_element_by_xpath('//*[@id="fixed_table"]/table/tbody/tr['+ str(i) +']/td[5]')
        for ele in elements:
            time.sleep(5)
            if '処理待ち' in ele.text:
                time.sleep(30)
                browser.refresh()
                time.sleep(5)
                i = 0
                break
            elif '処理中' in ele.text:
                time.sleep(30)
                browser.refresh()
                time.sleep(5)
                i = 0
                break
            elif '更新データなし' in ele.text:
                time.sleep(2)
                now = datetime.datetime.now()
                LOG_File.Write(now.strftime("%Y/%m/%d %H:%M:%S") + '■OPEN-UI DUKeView　アップロード　更新データなし終了')
                break
            elif 'エラーファイルあり' in ele.text:
                time.sleep(2)
                browser.find_element_by_xpath('//*[@id="fixed_table"]/table/tbody/tr['+ str(i) +']/td[5]/a').click()
                now = datetime.datetime.now()
                LOG_File.Write(now.strftime("%Y/%m/%d %H:%M:%S") + '■OPEN-UI DUKeView　アップロード　エラーファイルをダウンロードし終了')
                break
            elif '成功' in ele.text:
                time.sleep(2)
                now = datetime.datetime.now()
                LOG_File.Write(now.strftime("%Y/%m/%d %H:%M:%S") + '■OPEN-UI DUKeView　アップロード　更新データなし終了')
                break
    i = i + 1

time.sleep(5)
browser.close()
browser.quit()
