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

from env import URL, DriverLocation, outputlocation, linkslocation


def get_data(driver):
    """
    this function get main text, score, name
    """
    print('get data...')
    classoftable = 'styles_box__ukVBL'
    
    # classofcolumncontent = 'h4 styles_h4__x3zzi'
    # classofhref = 'styles_titleCell__5wNFE'
    # classofsnowtype = 'styles_small__5RsX3'
    alltables = driver.find_elements(By.CLASS_NAME,classoftable)
    FOR_LOOP_COUNT = len(alltables)
    for i in range(FOR_LOOP_COUNT):
        typename = alltables[i].find_element(By.XPATH, './/table/thead/tr/th[1]/div/div[1]/span[1]').text
        if typename == 'Closed':
            rows = alltables[i].find_elements(By.TAG_NAME, 'tr')
            lst_data = []
            for row in rows[1:]:
                contents = row.find_elements(By.TAG_NAME, 'td')
                name = contents[0].find_element(By.XPATH, './/a/span').text
                link = contents[0].find_element(By.XPATH, './/a').get_attribute('href')
                opendates = contents[1].find_element(By.XPATH, './/span').text
                timeofreport = contents[0].find_element(By.XPATH, './/a/time').text
                lst_data.append([name,link,opendates,timeofreport])
        elif typename == 'No Report Available':
            rows = alltables[i].find_elements(By.TAG_NAME, 'tr')
            lst_data = []
            for row in rows[1:]:
                contents = row.find_elements(By.TAG_NAME, 'td')
                name = contents[0].find_element(By.XPATH, './/a/span').text
                link = contents[0].find_element(By.XPATH, './/a').get_attribute('href')
                lst_data.append([name,link])
        else:
            rows = alltables[i].find_elements(By.TAG_NAME, 'tr')
            lst_data = []
            for row in rows[1:]:
                contents = row.find_elements(By.TAG_NAME, 'td')
                name = contents[0].find_element(By.XPATH, './/a/span').text
                link = contents[0].find_element(By.XPATH, './/a').get_attribute('href')
                snowfall = contents[1].find_element(By.XPATH, './/span').text
                basedepth = contents[2].find_element(By.XPATH, './/span').text
                snowtype = contents[2].find_element(By.XPATH, './/span/div').text
                opentrails = contents[3].find_element(By.XPATH, './/span').text
                openlifts = contents[4].find_element(By.XPATH, './/span').text
                timeofreport = contents[0].find_element(By.XPATH, './/a/time').text
                lst_data.append([name,link,snowfall,basedepth,snowtype,opentrails,openlifts,timeofreport])
        write_to_xlsx(lst_data,typename)
    
    return FOR_LOOP_COUNT

def ifGDRPNotice(driver):
    # check if the domain of the url is consent.google.com
    if 'consent.google.com' in driver.current_url:
        # click on the "I agree" button
        driver.execute_script('document.getElementsByTagName("form")[0].submit()')
    return

def ifPageIsFullyLoaded(driver):
    # check if the page fully loaded including js
    return driver.execute_script('return document.readyState') != 'complete'



def scrolling():
    print('scrolling...')
    from selenium.webdriver.common.action_chains import ActionChains
    from selenium.webdriver.common.keys import Keys
    import time
    from selenium.webdriver.common.by import By

    print('scrolling...')
    time.sleep(2)
    try:
        target_section = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div[2]/div/section')
    except:
        target_section = driver.find_element(By.XPATH, '/html/body/div[2]/div[6]/div[2]/div/section[1]')
        
    section_bottom = target_section.location['y'] + target_section.size['height']
    print("00_section_bottom:", section_bottom)
    count = 0
    countdifferent = 0

    # 创建 ActionChains 对象
    actions = ActionChains(driver)

    while True:
        current_scroll_y = driver.execute_script("return window.pageYOffset;")
        print("current_scroll_y:", current_scroll_y)
        # 如果页面当前滚动位置还未达到目标容器的底部，则模拟按 PageDown
        if current_scroll_y + driver.execute_script("return window.innerHeight;") < section_bottom:
            print("模拟按键 PAGE_DOWN")
            actions.send_keys(Keys.PAGE_DOWN).perform()
        else:
            # 如果滚动位置已经接近目标位置，则发送少量 ARROW_DOWN 调整
            print("模拟按键 ARROW_DOWN")
            actions.send_keys(Keys.ARROW_DOWN).perform()

        time.sleep(2)
        try:
            target_section = driver.find_element(By.XPATH, '/html/body/div[1]/div[6]/div[2]/div/section')
        except:
            target_section = driver.find_element(By.XPATH, '/html/body/div[2]/div[6]/div[2]/div/section[1]')
        section_new_bottom = target_section.location['y'] + target_section.size['height']
        print("section_new_bottom:", section_new_bottom)
        
        if section_new_bottom == section_bottom:
            count += 1
            if count > 6:
                print("已滚动到指定容器的底部，没有更多内容。")
                break
        else:
            countdifferent += 1  # 有变化时重置计数
            try:
                button_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div[5]/div[2]/div/button'))
                )
                print("按钮元素已加载:", button_element)
            except:
                pass
            count = 0
        section_bottom = section_new_bottom



def write_to_xlsx(data,name):
    print('write to excel...')
    if name == 'Closed':
        cols = ['name','link','Open Dates','time of report']
    elif name == 'No Report Available':
        cols = ['name','link']
    else:
        cols = ['name','link','Snowfall 24h','Base Depth','Snow Type','Open Trails','Open Lifts','time of report']
    # cols = ["name", "comment", 'rating']
    today = datetime.date.today()
    # today = datetime.datetime.now().strftime('%Y-%m-%d')
    df = pd.DataFrame(data, columns=cols)
    df.to_excel(outputlocation+'\\'+str(today)+'_'+name+'_'+URL.split('/')[3]+'.xlsx')
    

if __name__ == "__main__":

    print('starting...')
    linksdf = pd.read_excel(linkslocation)
    links = linksdf['Link'].tolist()
    for link in links:
        if str(link).startswith('http'):
            URL = link
        
    # get browser
            options = Options()
            # options.add_argument("--headless")  # show browser or not
            options.add_argument("--lang=en-US")
            options.add_experimental_option('prefs', {'intl.accept_languages': 'en,en_US'})
            options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")
            options.add_argument("--disable-blink-features=AutomationControlled")
            DriverPath = DriverLocation
            service = Service(DriverPath)
            driver = webdriver.Chrome(service=service, options=options)
            
            driver.get(URL)
            print('loading page...')
            while ifPageIsFullyLoaded(driver):
                time.sleep(3)
            print('loading page...2')
            ifGDRPNotice(driver)
            while ifPageIsFullyLoaded(driver):
                time.sleep(3)
            print('loading page...3')

            scrolling()
            print('Getting data...')

            data = get_data(driver)

            driver.close()

    # write_to_xlsx(data)
    print('Done!')
