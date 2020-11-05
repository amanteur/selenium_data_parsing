import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup as bs
import time
from openpyxl import load_workbook


def exception_search (url):
    #   exception list
    exception_list = ['12', '13', '19', '20', '21', '23', '24',
                      '27', '30', '32', '33', '39', '40', '46',
                      '49', '50', '60', '65', '67', '68', '69',
                      '70', '71', '77', '79', '80', '81', '82',
                      '84', '85']
    for i_exc in range(len(exception_list)):
        if exception_list[i_exc] in url:
            return True
        else:
            continue
    return False


def selenium_opening (base_url):
    #   defining variables #
    dict_name_url = {}




    #   selenium opening    #
    chromedriver = 'C:/chromedriver'
    options = webdriver.ChromeOptions()
    options.add_argument('headless')  # для открытия headless-браузера
    browser = webdriver.Chrome(executable_path=chromedriver, options=options)
    browser.get(base_url)

    soup_main = bs(browser.page_source, 'html.parser')
    tds = soup_main.find_all('td')

    #   looping names   #
    for td in tds:
        name = td.find('a')
        if name.string == name.empty or exception_search(name.get('href')):
            continue
        dict_name_url[name.string] = name.get('href')

    #   closing browser #
    browser.quit()
    return dict_name_url

def selenium_search(base_url,dict_name_url, period = 0, flag = False):
    #   defining variables  #
    cant_get_urls_dict = {}  # if browser cant get element 'from' this url will be added to this array
    without_info_urls_dict = {}  # URLs without info like Министерство Обороны

    names_array = np.array([])
    url_array = np.array([])
    uslugi_tel_budg_arr = np.array([])
    uslugi_tel_spec_arr = np.array([])
    uslugi_sot_budg_arr = np.array([])
    uslugi_sot_spec_arr = np.array([])
    uslugi_proc_budg_arr = np.array([])
    uslugi_proc_spec_arr = np.array([])

    # selenium opening    #
    chromedriver = 'C:/chromedriver'
    options = webdriver.ChromeOptions()
    options.add_argument('headless')  # для открытия headless-браузера
    browser = webdriver.Chrome(executable_path=chromedriver, options=options)
    browser.get(base_url)
    #   searching money #
    for key, value in dict_name_url.items():
        #   defining variables    #
        uslugi_tel_budg = 0
        uslugi_tel_spec = 0
        uslugi_sot_budg = 0
        uslugi_sot_spec = 0
        uslugi_proc_budg = 0
        uslugi_proc_spec = 0

        print(key)
        xpath = "//a[@href = '%s']" % value
        browser.find_element_by_xpath(xpath).click()    # clicking on the link name
        try:    # switching to iframe
            WebDriverWait(browser, 10).until(
                EC.frame_to_be_available_and_switch_to_it(
                    (By.XPATH, "//iframe[starts-with(@id,'fancybox-frame') and starts-with(@name,'fancybox-frame')]")
                )
            )
        except TimeoutException:
            print('oops cant switch to frame')
            time.sleep(2)
            break
        #   periods
        path_from_year = "//select[@class = 'ui-datepicker-year']/option[@value='" + str(period) + "']"
        path_to_year = "//select[@class = 'ui-datepicker-year']/option[@value='" + str(period) + "']"
        path_to_dec = "//select[@class = 'ui-datepicker-month']/option[@value='11']"
        if (flag == True):
            path_to_year = "//select[@class = 'ui-datepicker-year']/option[@value='2019']"
        try:  # clicking on the calendar
            to_button = WebDriverWait(browser, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@id='to' and @name = 'to']")
                )
            )
            to_button.click()
        except TimeoutException:
            print('oops no such element to')
            cant_get_urls_dict[key] = value
            browser.switch_to.default_content()
            browser.find_element_by_id('fancybox-close').click()
            browser.refresh()
            continue
        browser.find_element_by_xpath(path_to_year).click()
        browser.find_element_by_xpath(path_to_dec).click()
        try:
            browser.find_elements_by_xpath("//button[contains(@onclick,'hideDatepicker()')]")
            exit_button = WebDriverWait(browser, 3).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//button[contains(@onclick,'hideDatepicker()')]")
                )
            )
            exit_button.click()
        except:
            print('where is exit')
            break

        # browser.find_element_by_id('submit').click()        # ??????????
        time.sleep(1)
        try:    # clicking on the calendar
            from_button = WebDriverWait(browser, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@id='from' and @name = 'from']")
                )
            )
            from_button.click()
        except TimeoutException:
            print('oops no such element from')
            cant_get_urls_dict[key] = value
            browser.switch_to.default_content()
            browser.find_element_by_id('fancybox-close').click()
            browser.refresh()
            continue

        browser.find_element_by_xpath(path_from_year).click()
        try:
            browser.find_elements_by_xpath("//button[contains(@onclick,'hideDatepicker()')]")
            exit_button = WebDriverWait(browser, 3).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//button[contains(@onclick,'hideDatepicker()')]")
                )
            )
            exit_button.click()
        except:
            print('where is exit')
            break
        browser.find_element_by_id('submit').click()

        try:    # clicking on button in the table
            button_2212 = WebDriverWait(browser, 5).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//tr[@id='2212']")
                )
            )
            button_2212.click()
        except TimeoutException:
            print('oops no such element 2212')
            without_info_urls_dict[key] = value
            browser.switch_to.default_content()
            browser.find_element_by_id('fancybox-close').click()
            browser.refresh()
            continue

        #   parsing #
        requiredhtml = browser.page_source  # parsing html
        soup = bs(requiredhtml, 'html.parser')  # making a "soup" from this html
        trs = soup.find_all('tr', {"class": "child-2212"})
        for money in trs:                   # searching for money
            money_purpose = money.find_all('td')
            if money_purpose[0].text == ' Услуги телефонной и факсимильной связи':
                print(money_purpose[0].text)
                uslugi_tel_budg = float(money_purpose[1].text.replace(',', '.').replace(' ', ''))
                uslugi_tel_spec = float(money_purpose[2].text.replace(',', '.').replace(' ', ''))

            if money_purpose[0].text == ' Услуги сотовой связи':
                print(money_purpose[0].text)
                uslugi_sot_budg = float(money_purpose[1].text.replace(',', '.').replace(' ', ''))
                uslugi_sot_spec = float(money_purpose[2].text.replace(',', '.').replace(' ', ''))

            if money_purpose[0].text == ' Прочие услуги связи':
                print(money_purpose[0].text)
                uslugi_proc_budg = float(money_purpose[1].text.replace(',', '.').replace(' ', ''))
                uslugi_proc_spec = float(money_purpose[2].text.replace(',', '.').replace(' ', ''))


        # creating a dataframe
        names_array = np.append(names_array, key)
        url_array = np.append(url_array, "https://budget.okmot.kg" + value)
        uslugi_tel_budg_arr = np.append(uslugi_tel_budg_arr, uslugi_tel_budg)
        uslugi_tel_spec_arr = np.append(uslugi_tel_spec_arr, uslugi_tel_spec)
        uslugi_sot_budg_arr = np.append(uslugi_sot_budg_arr, uslugi_sot_budg)
        uslugi_sot_spec_arr = np.append(uslugi_sot_spec_arr, uslugi_sot_spec)
        uslugi_proc_budg_arr = np.append(uslugi_proc_budg_arr, uslugi_proc_budg)
        uslugi_proc_spec_arr = np.append(uslugi_proc_spec_arr, uslugi_proc_spec)

        #   exiting  to main menu   #
        browser.switch_to.default_content()
        browser.find_element_by_id('fancybox-close').click()
        browser.refresh()

    # #   info    #
    print('Count of not working URLs: ', len(cant_get_urls_dict))
    print('Names of not working URLs: ', cant_get_urls_dict.keys())
    print('Count of URLs without info: ', len(without_info_urls_dict))
    print('Names of URLs without info: ', without_info_urls_dict.keys())
    print('Count of names_array: ', names_array.size)
    df_main = pd.DataFrame({'Ведомство': names_array,
                            'URL': url_array,
                            'Услуги телефонной и факсимильной связи Бюджетные средства': uslugi_tel_budg_arr,
                            'Услуги телефонной и факсимильной связи Спец средства': uslugi_tel_spec_arr,
                            'Услуги сотовой связи Бюджетные средства': uslugi_sot_budg_arr,
                            'Услуги сотовой связи Спец средства': uslugi_sot_spec_arr,
                            'Прочие услуги связи Бюджетные средства': uslugi_proc_budg_arr,
                            'Прочие услуги связи Спец средства': uslugi_proc_spec_arr})

    #   closing browser #
    browser.quit()
    return df_main, cant_get_urls_dict

#   main    #
base_url = 'https://budget.okmot.kg/ru/exp_vedom/index.html?year=2018'

dict_name_url = selenium_opening(base_url)
print(len(dict_name_url))

search_result = selenium_search(base_url,dict_name_url,period=2017)
df_main = search_result[0]
dict_name_url_new = search_result[1]
print(dict_name_url_new)

# if there are troubles with opening calendar
# while len(dict_name_url_new) != 0:
#     print('in loop')
#     new_search_result = selenium_search(base_url, dict_name_url_new)
#     df_new = new_search_result[0]
#     print(df_new)
#     frames = [df_main, df_new]
#     df_main = pd.concat(frames)

#   writing main in excel file #

path = r"summary_years_1.xlsx"
book = load_workbook(path)
writer = pd.ExcelWriter(path = path, engine = 'openpyxl')
writer.book = book
df_main.to_excel(writer, sheet_name = '2019lala')

  # writing years in excel file #
# for i in range(0,5):
#     year = 2014 + i
#     print(year)
#     search_result = selenium_search(base_url, dict_name_url, period=year)
#     search_result[0].to_excel(writer, sheet_name = str(year))

writer.save()
writer.close()
print('Done!')
