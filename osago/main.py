from seleniumwire import webdriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from SyncReHalka import ReHalka
import configparser
import json
import threading as th
import pandas as pd
import requests as req
import time
import math


# selenium==4.8.0
# selenium-wire==5.1.0
# telebot==0.0.5
# 



config = configparser.ConfigParser()
config.read('config.ini')

################################################################################################################################################

input_file = '123.xlsx' # Path for target input file with .xlsx
proxy_server_url = str(config["VARS"]['proxy_server_url']) # Proxy url for selenium
api_token = str(config["VARS"]['api_token']) # RehalkaApi token for capcha
proxy_username = str(config["VARS"]['proxy_username'])
proxy_password = str(config["VARS"]['proxy_password'])
proxy_ip_port = str(config["VARS"]['proxy_ip_port'])

################################################################################################################################################



df_new = pd.DataFrame(columns = ['Полис', 'Страховая компания', 'Дата оформления ОСАГО', 'Период использования ТС', 'Водители',
                               'Регион', 'КБМ', 'Стоимость ОСАГО', 'Марка, модель, категория', 'Рег. знак',
                                'VIN', 'Мощность','Цель использования', 'Страхователь', 'Собственник', 'Восстановление КБМ',
                                'Дата регистрации', 'Залогодатель', 'Залогодержатель'])
target_link = "https://kbm-osago.ru/osago/proverka-polisa-osago.html"
site_key = '2f7f78c3-233e-4567-a039-92a72d45f691'
# инициализация апи
solver = ReHalka(api_token) 
# отправка капчи на сервер
df_array = []

def get_data_by_vin(vin, browser):
    res = []
    browser.get('https://www.reestr-zalogov.ru/search/index')
    time.sleep(2)
    elem = browser.find_element(By.CSS_SELECTOR, "ul[class='nav nav-pills']")
    elem = elem.find_element(By.LINK_TEXT, "По информации о предмете залога")
    elem.click()
    time.sleep(2)
    wait = WebDriverWait(browser, 100)
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[style='margin-right: 5px;']")))
    elem = browser.find_element(By.CSS_SELECTOR, "input[style='margin-right: 5px;']")
    time.sleep(2)
    elem.click()
    elem = browser.find_element(By.ID, 'vehicleProperty.vin')
    elem.send_keys(vin)
    elem = browser.find_element(By.ID, "find-btn")
    time.sleep(2)
    elem.click()
    wait = WebDriverWait(browser, 100)
    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='search-params-tip']")))
    bs = BeautifulSoup(browser.page_source, 'html.parser')
    try:
        elem = bs.find_all("tbody", recursive = True)[0]
    except:
        return [None, None, None]
    elem = elem.find_all("td", recursive = True)
    try:
        elem1 = elem[1].text
    except:
        elem1 = None
    try:
        elem2 = elem[4].find("div", recursive = True).find("div", recursive = True).find("a", recursive = True).text
    except:
        elem2 = None
    try:
        elem3 = elem[5].find("div", recursive = True).find("div", recursive = True).find("a", recursive = True).text
    except:
        elem3 = None
    res.append(elem1)
    res.append(elem2)
    res.append(elem3)
    return res

def get_captcha_answer(solver, target_link, site_key):
    while True:
            try:
                captcha_id = solver.send_captcha(domain=target_link, site_key=site_key)
                break
            except:
                pass
    captcha_id = captcha_id.split('|')[1]
    captcha_answer = solver.get_captcha_answer(captcha_id=captcha_id)
    while 'CAPCHA_NOT_READY' in captcha_answer:
        while True:
            try:
                captcha_answer = solver.get_captcha_answer(captcha_id=captcha_id)
                break
            except:
                pass
    captcha_answer = captcha_answer.split('|')[1]
    return captcha_answer


def create_some_data(vins, j):    
    df = pd.DataFrame(columns = ['Полис', 'Страховая компания', 'Дата оформления ОСАГО', 'Период использования ТС', 'Водители',
                               'Регион', 'КБМ', 'Стоимость ОСАГО', 'Марка, модель, категория', 'Рег. знак',
                                'VIN', 'Мощность','Цель использования', 'Страхователь', 'Собственник', 'Восстановление КБМ',
                                 'Дата регистрации', 'Залогодатель', 'Залогодержатель'])
    url = 'https://kbm-osago.ru/osago/proverka-polisa-osago.html'
    url_post='https://engine.kbm-osago.ru/check_policy.php'
    headers = {
    "User-Agent": 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36',
     'X-Requested-With': 'XMLHttpRequest',
     'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 
     'origin' : 'https://kbm-osago.ru/'}
    for i in vins:
        with req.Session() as session:
            while True:
                try:
                    frame2 = {'Полис' : None, 'Страховая компания' : None, 'Дата оформления ОСАГО' : None, 'Период использования ТС' : None, 'Водители' : None,
                                            'Регион' : None, 'КБМ' : None, 'Стоимость ОСАГО' : None, 'Марка, модель, категория' : None, 'Рег. знак' : None,
                                                'VIN' : None, 'Мощность' : None,'Цель использования' : None, 'Страхователь' : None, 'Собственник' : None,
                                                'Восстановление КБМ' : None, 'Дата регистрации' : None, 'Залогодатель' : None, 'Залогодержатель' : None}
                    data = {'bsoSeries': "%D0%A5%D0%A5%D0%A5", 'vin' : None, 'datequery': None, 'isAcceptCachedData' : 1,  'g-recaptcha-response': None, 'h-captcha-response': None}
                    resp = session.get(url, headers=headers)
                    bs = BeautifulSoup(resp.text, 'html.parser')
                    date = bs.find_all("input", {"name": 'datequery' }, recursive = True)[0].get('value')
                    captcha_answer = get_captcha_answer(solver, target_link, site_key)
                    data['g-recaptcha-response'] = captcha_answer
                    data['h-captcha-response'] = captcha_answer
                    data['datequery'] = date
                    data['vin'] = str(i)
                    resp=session.post(url_post,headers=headers, data=data)
                    xz = BeautifulSoup(resp.text,'html.parser').prettify()
                    xz = json.loads(xz)
                    xz = xz['response'][0]
                    break
                except:
                    pass
            frame2['Полис'] = xz["bsoSeries"]+' '+xz["bsoNumber"]
            frame2["Страховая компания"] = xz["insurer"]
            frame2["Дата оформления ОСАГО"] = xz["bsoUpdateDate"]
            frame2["Период использования ТС"] = xz["contractStartDate"] + ' - '+ xz["contractEndDate"]
            if xz["driversCount"] == None:
                frame2["Водители"] = None
            else:
                frame2["Водители"] = str(xz["driversCount"]) + " чел."
            frame2["Регион"] = xz["policyRegion"]
            frame2["КБМ"] = xz["policyKbm"]
            frame2["Стоимость ОСАГО"] = xz["policyPrice"]+ " руб."
            frame2["Марка, модель, категория"] = xz["carMarkModel"] + f" (кат. {xz['carCategory']})"
            frame2["Рег. знак"] = xz["regNumber"]
            frame2["VIN"] = xz["vin"]
            frame2["Мощность"] = xz["power"] + " л.с."
            frame2["Цель использования"] = xz["purposeOfUse"]
            frame2["Страхователь"] = xz["insurantFio"] + f" д.р. {xz['insurantBirthdate']}"
            frame2["Собственник"] = xz["ownerFio"] + f" д.р. {xz['ownerBirthdate']}"
            frame2['Восстановление КБМ'] = f'Кажется, вы переплатили за ОСАГО {float(xz["policyPrice"]) - (float(xz["policyPrice"])*0.5)/0.75} руб. При восстановлении КБМ до 0.5, стоимость ОСАГО будет {(float(xz["policyPrice"])*0.5)/0.75} руб.'
            df.loc[len(df.index)] = frame2
    df_array[j] = df


def save_data(output_file, frame):
    df = pd.concat(frame)
    try:
        df_new = pd.read_excel(output_file)
        df = pd.concat([df_new, df])
    except Exception as e:
        print(e)
        pass
    while True:
        try:
            df.to_excel(output_file, index = False)
            break
        except:
            pass

def get_vins_n_row_count(input_file):
    vins = pd.read_excel(input_file)
    row_count = len(vins)
    vins_arr = []
    for vin in vins.values:
        vins_arr.append(list(vin)[0])
    return vins_arr, row_count

def main(input_file, output_file, id, bot):
    global df_array
    vins_temp, row_count = get_vins_n_row_count(input_file)
    row_counter = 0
    thread_number = int(config["VARS"]['thread_number'])
    if(row_count/thread_number*thread_number<1):
        thread_number = math.sqrt(row_count)
    for j in range(0, thread_number):
        if row_count%thread_number!=0 and j==thread_number-1:
            vins = vins_temp[row_counter:int(row_counter+(row_count/thread_number))+row_count%thread_number]
        else:
            vins = vins_temp[row_counter:int(row_counter+(row_count/thread_number))]
        roww_counter = 0
        thread_array = []
        df_array = []
        for i in range(0, thread_number):
            if row_count%(thread_number*thread_number)!=0 and j==thread_number-1:
                slice = vins[roww_counter:int(roww_counter+(row_count/thread_number)/thread_number)+row_count%(thread_number*thread_number)]
            else:
                slice = vins[roww_counter:int(roww_counter+(row_count/thread_number)/thread_number)]
            df_array.append(df_new)           
            thread_array.append(th.Thread(target=create_some_data, args=[slice,i]))
            roww_counter+=int((row_count/thread_number)/thread_number)
        for thread in thread_array:
            thread.start()
        for thread in thread_array:  
            while(thread.is_alive()):
                pass
        row_counter+=int(row_count/thread_number)
        save_data(output_file, df_array)
    bot.send_message(id, "Ваш файл готов!")
    bot.send_document(id, open(output_file, 'rb'))
    # get_second_data(input_file, output_file, id, bot)


def get_second_data(input_file, output_file, id, bot):
    vins, _ = get_vins_n_row_count(input_file)
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("--disable-3d-apis")
    chrome_options.add_argument('--headless=new')
    chrome_options.add_argument('ignore-certificate-errors')
    seleniumwire_options = {'proxy': {'https': f'https://{proxy_username}:{proxy_password}@{proxy_ip_port}', 'verify_ssl': False,}}
    browser = uc.Chrome(options=chrome_options, seleniumwire_options=seleniumwire_options)
    browser.fullscreen_window()
    df_counter = 1
    for vin in vins:
        new_data = get_data_by_vin(vin, browser)
        while True:
            try:
                df_new2 = pd.read_excel(output_file)
                break
            except:
                pass
        df_new2.at[df_counter,'Дата регистрации'] = new_data[0]
        df_new2.at[df_counter,'Залогодатель'] = new_data[1]
        df_new2.at[df_counter,'Залогодержатель'] = new_data[2]
        df_counter+=1
        while True:
            try:
                df_new2.to_excel(output_file, index = False)
                break
            except:
                pass
    browser.close()
    bot.send_message(id, "Вторая часть файла готова.")
    bot.send_document(id, open(output_file, 'rb'), caption="Ваш обработанный файл.")

if __name__ == '__main__':
    main(input_file, output_file="out.xlsx")



    
    


