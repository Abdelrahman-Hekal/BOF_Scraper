from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import undetected_chromedriver as uc
import time
import os
import re
from datetime import datetime
import pandas as pd
import warnings
import sys
import xlsxwriter
from multiprocessing import freeze_support
import calendar 
import shutil
warnings.filterwarnings('ignore')

def initialize_bot(translate):

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    #chrome_options.add_argument("--incognito")
    chrome_options.add_argument('--headless=new')
    
    # disable location prompts & disable images loading
    prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2, "profile.default_content_setting_values.notifications": 2}  
    chrome_options.page_load_strategy = 'normal'

    chrome_options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    driver.set_page_load_timeout(20)

    return driver
    
def login(driver):

    driver.get('https://sso.businessoffashion.com/login?')
    time.sleep(3)
    username = wait(driver, 4).until(EC.presence_of_element_located((By.ID, "1-email")))
    username.send_keys('brandperformancedata@gmail.com')




def scrape_articles(driver, output1, page, month, year):

    stamp = datetime.now().strftime("%d_%m_%Y")
    print('-'*75)
    print(f'Scraping The Articles Links from: {page}')
    # getting the full posts list
    links = []
    months = {month: index for index, month in enumerate(calendar.month_abbr) if month}
    full_months = {month: index for index, month in enumerate(calendar.month_name) if month}
    prev_month = month - 1
    if prev_month == 0:
        prev_month = 12

    driver.get(page)
    art_time = ''

    # handling lazy loading
    print('-'*75)
    print("Getting the previous month's articles..." )

    for _ in range(100):  
        try:
            try:
                # scrolling across the page for elements loading
                try:
                    total_height = driver.execute_script("return document.body.scrollHeight")
                    height = total_height/10
                    new_height = 0
                    for _ in range(10):
                        prev_hight = new_height
                        new_height += height             
                        driver.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                        time.sleep(0.1)
                except:
                    pass
                try:
                    try:
                        div = wait(driver, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='results-list-container']")))[-1]
                    except:
                        div = wait(driver, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class*='container']")))[-1]
                    date = wait(div, 3).until(EC.presence_of_all_elements_located((By.TAG_NAME, "time")))[-1].get_attribute('textContent').strip()
                    art_month = full_months[date.split()[1]]
                    art_year = int(date.split()[-1])
                except:
                    try:
                        date = wait(driver, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "p[class*='text-gray']")))[-1].get_attribute('textContent').strip()
                        art_month = months[date.split()[0]]
                        art_year = int(date.split()[-1])                        
                    except:    
                        date = wait(driver, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span[max-hours='120']")))[-1].get_attribute('textContent').strip()
                        art_month = months[date.split()[1]]
                        art_year = int(date.split()[-1])

                # for articles from previous year
                if art_year < year and prev_month != 12:
                    break
                # for all months except Jan
                elif art_month < prev_month and prev_month != 12 and art_year == year:
                    break
                # for Jan
                elif art_month < prev_month and prev_month == 12 and art_year < year:
                    break
            except:
                break

            # moving to the next page
            try:
                div = wait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='LoadMoreButton']")))
                button = wait(div, 3).until(EC.presence_of_element_located((By.TAG_NAME, "button")))
                driver.execute_script("arguments[0].click();", button)
                #time.sleep(3)
            except:
                try:
                    div = wait(driver, 3).until(EC.presence_of_all_elements_located((By.XPATH, "//div[@class='see-more']")))[-1]
                    button = wait(div, 3).until(EC.presence_of_element_located((By.TAG_NAME, "button")))
                    driver.execute_script("arguments[0].click();", button)
                    #time.sleep(3)
                except:
                    break

        except Exception as err:
            break

    # scraping posts urls 
    try:
        posts = wait(driver, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='list-item']")))    
    except:
        print('No posts are available')
        return

    for post in posts:
        try:
            try:
                date = wait(post, 2).until(EC.presence_of_element_located((By.TAG_NAME, "time"))).get_attribute('textContent').strip()
                art_month = full_months[date.split()[1]]
                art_year = int(date.split()[-1])
            except:
                try:
                    date = wait(post, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "p[class*='text-gray']"))).get_attribute('textContent').strip()
                    art_month = months[date.split()[0]]
                    art_year = int(date.split()[-1])                        
                except:    
                    date = wait(post, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[max-hours='120']"))).get_attribute('textContent').strip()
                    art_month = months[date.split()[1]]
                    art_year = int(date.split()[-1])

            if art_month != prev_month: continue
            if art_year < year and prev_month != 12: continue
            link = wait(post, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))[-1].get_attribute('href')
            if link not in links:
                links.append(link)
        except:
            pass

    # scraping posts
    print('-'*75)
    print('Scraping Articles...')
    print('-'*75)

    # reading previously scraped data for duplication checking
    scraped = []
    try:
        df = pd.read_excel(output1)
        scraped = df['unique_id'].values.tolist()
    except:
        pass

    n = len(links)
    data = pd.DataFrame()
    for i, link in enumerate(links):
        try:
            driver.get(link)   
        except:
            print(f'Warning: Failed to load the url: {link}')
            continue

        art_id = ''
        try:
            art_id = link.strip('/').split('/')[-1]
        except:
            pass

        if art_id in scraped: 
            print(f'Article {i+1}\{n} is already scraped, skipping.')
            continue        
        
        if art_id == '': 
            print(f'Warning: Article {i+1}\{n} has unknown ID')
            art_id = 0

        # scrolling across the page for auto translation to be applied
        try:
            total_height = driver.execute_script("return document.body.scrollHeight")
            height = total_height/30
            new_height = 0
            for _ in range(30):
                prev_hight = new_height
                new_height += height             
                driver.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                time.sleep(0.1)
        except:
            pass

        row = {}
        # article author
        en_author = ''       
        try:
            en_author = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-test*='article-byline']"))).get_attribute('textContent').replace('By', '').strip()
        except Exception as err:
            pass        
        
        # article date
        date = ''         
        try:
            date = wait(driver, 4).until(EC.presence_of_element_located((By.TAG_NAME, "time"))).get_attribute('textContent').strip().split()
        except Exception as err:
            pass
            
        # checking if the article date is correct
        try:
            art_month = int(full_months[date[1]])
            art_year = int(date[2])  
            art_day = int(date[0])       
            if art_month != prev_month: 
                print(f'skipping article with date {art_month}/{art_day}/{art_year}')
                continue
            date = f'{art_day}_{art_month}_{art_year}'
        except:
            continue    

        row['sku'] = art_id
        row['unique_id'] = art_id
        row['articleurl'] = link

        # article title
        title = ''             
        try:
            title = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1[data-test*='article-title']"))).get_attribute('textContent').strip()
        except:
            continue               
                
        row['articletitle in English'] = title            

        # article description
        des = ''             
        try:
            text = driver.page_source
            text = re.sub(r'<.*?>', '', text)
            text = text.replace("<a href=\\", '').replace("<iframe src=\\", '').replace("Subscribe to the ", '')
            elems = re.findall(r'"content":"(.*?)"', text)
            des = '\n'.join(elems).replace('"content":', '').replace('"', '').split('Learn more:')[0]
        except:
            continue               
                
        row['articledescription in English'] = des.strip('\n')
        row['articleauthor'] = en_author
        row['articledatetime'] = date            
            
        # article category
        cat = ''             
        try:
            cat = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-test='article-overline']"))).get_attribute('textContent').strip()
        except:
            pass 
            
        row['articlecategory'] = cat

        # other columns
        row['domain'] = 'BOF'
        row['hype'] = 0   

        tags = ''
        try:
            div = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-test='article-taxonomies-tags']")))
            lis = wait(div, 4).until(EC.presence_of_all_elements_located((By.TAG_NAME, "li")))
            for li in lis:
                tags += li.get_attribute('textContent').strip() + ', '
        except:
            pass

        row['articletags'] = tags.strip(', ')
        row['articleheader'] = ''

        imgs = ''
        try:
            try:
                div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-test*='headimage']")))
            except:
                div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-test*='image']")))
            elems = wait(div, 4).until(EC.presence_of_all_elements_located((By.TAG_NAME, "img")))
            for elem in elems:
                try:
                    imgs += elem.get_attribute('src') + ', '
                except:
                    pass
            imgs = imgs.strip(', ')
        except:
            pass

        row['articleimages'] = imgs
        row['articlecomment'] = ''
        row['Extraction Date'] = stamp
        # appending the output to the datafame       
        data = pd.concat([data, pd.DataFrame([row.copy()])], ignore_index=True)
        print(f'Scraping the details of article {i+1}\{n}')
           
    # output to excel
    if data.shape[0] > 0:
        data['articledatetime'] = pd.to_datetime(data['articledatetime'],  errors='coerce', format="%d_%m_%Y")
        data['articledatetime'] = data['articledatetime'].dt.date  
        data['Extraction Date'] = pd.to_datetime(data['Extraction Date'],  errors='coerce', format="%d_%m_%Y")
        data['Extraction Date'] = data['Extraction Date'].dt.date   
        df1 = pd.read_excel(output1)
        if df1.shape[0] > 0:
            df1[['articledatetime', 'Extraction Date']] = df1[['articledatetime', 'Extraction Date']].apply(pd.to_datetime,  errors='coerce', format="%Y-%m-%d")
            df1['articledatetime'] = df1['articledatetime'].dt.date 
            df1['Extraction Date'] = df1['Extraction Date'].dt.date 
        df1 = pd.concat([df1, data], ignore_index=True)
        df1 = df1.drop_duplicates()
        writer = pd.ExcelWriter(output1, date_format='d/m/yyyy')
        df1.to_excel(writer, index=False)
        writer.close()
    else:
        print('-'*75)
        print('No New Articles Found')
        
def get_inputs():
 
    print('-'*75)
    print('Processing The Settings Sheet ...')
    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '\\' in path:
        path += '\\BOF_settings.xlsx'
    else:
        path += '/BOF_settings.xlsx'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "BOF_settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        settings = {}
        urls = []
        df = pd.read_excel(path)
        cols  = df.columns
        for col in cols:
            df[col] = df[col].astype(str)

        inds = df.index
        for ind in inds:
            row = df.iloc[ind]
            link, status = '', ''
            for col in cols:
                if row[col] == 'nan': continue
                elif col == 'Category Link':
                    link = row[col]
                elif col == 'Scrape':
                    status = row[col]
                else:
                    settings[col] = row[col]

            if link != '' and status != '':
                try:
                    status = int(status)
                    urls.append((link, status))
                except:
                    urls.append((link, 0))
    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)

    return settings, urls

def initialize_output():

    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\Scraped_Data\\' + stamp
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)

    file1 = f'BOF_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
    else:
        output1 = path + "\\" + file1  

    # Create an new Excel file and add a worksheet.
    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()    

    return output1

def main():

    print('Initializing The Bot ...')
    freeze_support()
    start = time.time()
    output1 = initialize_output()
    settings, urls = get_inputs()
    month = datetime.now().month
    year = datetime.now().year
    try:
        driver = initialize_bot(False)
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()

    for url in urls:
        if url[1] == 0: continue
        link = url[0]
        try:
            scrape_articles(driver, output1, link, month, year)
        except Exception as err: 
            print(f'Warning: the below error occurred:\n {err}')
            driver.quit()
            time.sleep(2)
            driver = initialize_bot(False)

    driver.quit()
    print('-'*75)
    elapsed_time = round(((time.time() - start)/60), 4)
    input(f'Process is completed in {elapsed_time} mins, Press any key to exit.')
    sys.exit()

if __name__ == '__main__':

    main()

