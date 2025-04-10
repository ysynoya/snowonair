#获取historical
import time
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import datetime
from openpyxl import Workbook
import pandas as pd
import openpyxl
import os

from env import URL, DriverLocation, outputlocationfolder, linkslocationfolder


def get_data(driver,region):
    """
    this function get main text, score, name
    """
    print('get data...')
            # store possible open and end date
            # click historical data link
    classofopenandclosedate = 'h3 styles_h3__rJrr7'
    opendate = ''
    closedate = ''
    # averagesnowfall = ''
    # snowfalldatys = ''
    # averagebasedepth = ''
    # maxbasedepth = ''
    # biggestsnowfall = ''


    # try:
    #     datesifexist = driver.find_elements(By.CLASS_NAME,classofopenandclosedate)
    #     opendate = datesifexist[0].text
    #     closedate = datesifexist[1].text
    # except:
    #     pass
    driver.get(URL.replace('skireport','historical-snowfall'))
    print('loading page...')
    while ifPageIsFullyLoaded(driver):
        time.sleep(4)
    print('loading page...2')
    # time.sleep(2)
    classoftable = 'styles_table__Z17oT'
    # classofcolumncontent = 'h4 styles_h4__x3zzi'
    # classofhref = 'styles_titleCell__5wNFE'
    # classofsnowtype = 'styles_small__5RsX3'
    classofannualbutton = 'styles_active__oDQkX'
    classofbuttonregion = 'styles_tabs__QXi8y'
    alltables = driver.find_elements(By.CLASS_NAME,classoftable)
    FOR_LOOP_COUNT = len(alltables)
    for i in range(FOR_LOOP_COUNT):
        if(i!=0):
            button = driver.find_element(By.CLASS_NAME, classofbuttonregion).find_element(By.XPATH,'.//button[2]')
            button.click()
        rows = alltables[i].find_elements(By.TAG_NAME, 'tr')
        lst_data = []
        for row in rows[1:]:
            month = row.find_element(By.TAG_NAME,'th').text
            contents = row.find_elements(By.TAG_NAME, 'td')
            averagesnow = contents[0].text
            snowfalldays = contents[1].text
            basedepth = contents[2].text
            maxbasedepth = contents[3].text
            biggestsnowfall = contents[4].text
            lst_data.append([month,averagesnow,snowfalldays,basedepth,maxbasedepth,biggestsnowfall])
        write_to_xlsx(lst_data,i,region)
        print("id:",id,",i:",i)
    
    return opendate,closedate

def ifGDRPNotice(driver):
    # check if the domain of the url is consent.google.com
    if 'consent.google.com' in driver.current_url:
        # click on the "I agree" button
        driver.execute_script('document.getElementsByTagName("form")[0].submit()')
    return

def ifPageIsFullyLoaded(driver):
    # check if the page fully loaded including js
    return driver.execute_script('return document.readyState') != 'complete'


def find_xlsx_files(directory):
    xlsx_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".xlsx"):
                xlsx_files.append(os.path.join(root, file))
    return xlsx_files

# 示例用法


def write_to_xlsx(data,name,region):
    print('write to excel...')
    yearormonth = ''
    if name == 0:
        cols = ['Month','Average Snowfall','Snowfall Days','Average Base Depth','Max Base Depth','Biggest Snowfall']
        yearormonth = 'Month'
    else:
        cols = ['Month','Average Snowfall','Snowfall Days','Average Base Depth','Max Base Depth','Biggest Snowfall']
        yearormonth = 'Year'
    # cols = ["name", "comment", 'rating']
    today = datetime.date.today().year
    # today = datetime.datetime.now().strftime('%Y-%m-%d')
    df = pd.DataFrame(data, columns=cols)
    df.to_excel(outputlocationfolder+'\\'+region+'_'+URL.split('/')[3]+'_'+str(id)+'_'+str(today)+'_'+yearormonth+'_'+URL.split('/')[4]+'.xlsx')
    

if __name__ == "__main__":

    print('starting...')
    lst_all_areas_date = []
    # directory = "/path/to/your/folder"
    xlsx_files = find_xlsx_files(linkslocationfolder)
    for file in xlsx_files[:]:
        currentregion = file.split('_')[-1].split('.')[0]
        linksdf = pd.read_excel(file)
        links = linksdf['link'].tolist()
        id = 0
        for link in links[0:5]:
            
            if str(link).startswith('http'):
                URL = link
            
        # get browser
                options = Options()
                options.add_argument("--headless")  # show browser or not
                options.add_argument("--lang=en-US")
                options.add_experimental_option('prefs', {'intl.accept_languages': 'en,en_US'})
                options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")
                options.add_argument("--disable-blink-features=AutomationControlled")
                DriverPath = DriverLocation
                service = Service(DriverPath)
                driver = webdriver.Chrome(service=service, options=options)
                
                # driver.get(URL)
                # print('loading page...')
                # while ifPageIsFullyLoaded(driver):
                #     time.sleep(3)
                # print('loading page...2')
                # # ifGDRPNotice(driver)
                # while ifPageIsFullyLoaded(driver):
                #     time.sleep(3)
                # print('loading page...3')

                print('Getting data...')

                opendate,closedate = get_data(driver,currentregion)
                lst_all_areas_date.append([currentregion,link,opendate,closedate])
                driver.close()
            id=id+1
    # dfall = pd.DataFrame(lst_all_areas_date, columns=['region','links','opendate','closedate'])
    # dfall.to_excel('allopenandclose.xlsx')
    # write_to_xlsx(data)
    print('Done!')
