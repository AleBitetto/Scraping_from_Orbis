import os
import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import time, re
from timeit import default_timer as timer
import datetime
from bs4 import BeautifulSoup
import autoit


def get_chromedriver(chromedriver_path=None, use_proxy=False, user_agent=None,
                    PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None):

    manifest_json = """
    {
        "version": "1.0.0",
        "manifest_version": 2,
        "name": "Chrome Proxy",
        "permissions": [
            "proxy",
            "tabs",
            "unlimitedStorage",
            "storage",
            "<all_urls>",
            "webRequest",
            "webRequestBlocking"
        ],
        "background": {
            "scripts": ["background.js"]
        },
        "minimum_chrome_version":"22.0.0"
    }
    """

    background_js = """
    var config = {
            mode: "fixed_servers",
            rules: {
            singleProxy: {
                scheme: "http",
                host: "%s",
                port: parseInt(%s)
            },
            bypassList: ["localhost"]
            }
        };

    chrome.proxy.settings.set({value: config, scope: "regular"}, function() {});

    function callbackFn(details) {
        return {
            authCredentials: {
                username: "%s",
                password: "%s"
            }
        };
    }

    chrome.webRequest.onAuthRequired.addListener(
                callbackFn,
                {urls: ["<all_urls>"]},
                ['blocking']
    );
    """ % (PROXY_HOST, PROXY_PORT, PROXY_USER, PROXY_PASS)

    chrome_options = webdriver.ChromeOptions()
    if use_proxy:
        pluginfile = 'proxy_auth_plugin.zip'

        with zipfile.ZipFile(pluginfile, 'w') as zp:
            zp.writestr("manifest.json", manifest_json)
            zp.writestr("background.js", background_js)
        chrome_options.add_extension(pluginfile)
    if user_agent:
        chrome_options.add_argument('--user-agent=%s' % user_agent)
    driver = webdriver.Chrome(
        executable_path=chromedriver_path,
        chrome_options=chrome_options)
    return driver



def scrape_from_Orbis(chromedriver_path=None, PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None,
                     Orbis_landing_page=None, Orbis_username=None, Orbis_pass=None, Orbis_saved_search=None,
                     Orbis_columns_set=None, maximum_chunk_size=None, time_before_closing_download=60,
                     resume_from_chunk=-1):
    
    LOG_FILE = 'Chunks_list.txt'
    SMALL_TIME_SLEEP = 0.5
    start = timer()

    # open browser and go to website
    print("Launching Chrome...", end ="")
    driver = get_chromedriver(chromedriver_path = chromedriver_path, use_proxy=True,
                              PROXY_HOST=PROXY_HOST, PROXY_PORT=PROXY_PORT, PROXY_USER=PROXY_USER, PROXY_PASS=PROXY_PASS)
    driver.implicitly_wait(30)
    driver.get(Orbis_landing_page)
    print("OK", end ="\n")

    # login
    print("Logging in...", end ="")
    driver.find_element_by_id("user").clear()
    driver.find_element_by_id("user").send_keys(Orbis_username)
    driver.find_element_by_id("pw").clear()
    driver.find_element_by_id("pw").send_keys(Orbis_pass)
    driver.find_element_by_id("user").click()
    driver.find_element_by_id("loginPage").click()
    driver.find_element_by_xpath("//div[@id='loginPage']/div[2]").click()
    driver.find_element_by_xpath("//div[@id='loginPage']/div[2]/div[3]/button").click()
    time.sleep(2)
    driver.find_element_by_xpath("//input[@value='Accept']").click()
    # check for session restart
    try:
        driver.find_element_by_xpath("//input[@value='Restart']").click()
        print("SESSION RESTART DONE", end ="\n")
    except:
        print("OK", end ="\n")

    # load search and show results
    print("Resetting search criteria...", end ="")
    driver.find_element_by_xpath("//div[@id='search-toolbar']/div/div/ul/li[2]/h4").click()
    try:   # check if "reset search" is available
        driver.find_element_by_link_text("Ricomincia").click()
        driver.find_element_by_xpath("//a[contains(text(),'Ricomincia')]").click()
        driver.find_element_by_xpath("//div[@id='search-toolbar']/div/div/ul/li[2]/h4").click()
        print("OK", end ="\n")
    except:
        print("NOT FOUND, criteria already resetted", end ="\n")
    time.sleep(2)
    print("Showing results...", end ="")
    driver.find_element_by_xpath("//div[@id='tooltabSectionload-search-section']/div/div[2]/div/div/div[2]//span[@title='"+Orbis_saved_search+"']").click()
    time.sleep(30)
    driver.find_element_by_xpath("//img[@alt='Risultati']").click()
    print("OK", end ="\n")

    # split into chunks
    total_rows = int(driver.find_element_by_id('records-count').get_attribute('data-total-records-count'))
    print("Total records found:", str(total_rows))
    rows = range(1,total_rows+1)
    chunk_index = [(rows[i:i+maximum_chunk_size][0], rows[i:i+maximum_chunk_size][-1]) for i in range(0, len(rows), maximum_chunk_size)]
    print("Splitting into", len(chunk_index), "chunks")

    # delete log file
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)
    
    # export chunks into excel files
    for chunk_i, chunk in enumerate(chunk_index):

        # save list of chunks' rows
        with open(LOG_FILE, 'a') as f:
            f.write('\nChunk ' + str(chunk_i+1) + ': ' + '-'.join([str(i) for i in chunk]))
        
        if chunk_i >= resume_from_chunk:
            print("Querying chunk", str(chunk_i+1), "/", str(len(chunk_index)), end = "\r")

            row_start = chunk[0]
            row_end = chunk[1]

            # open Excel tab
            driver.find_element_by_link_text("Excel").click()
            time.sleep(SMALL_TIME_SLEEP)
            # open Excel options
            tt = True
            while tt:     # there could be lag in opening the excel tab
                try:
                    driver.find_element_by_xpath("//div[@id='export-component-exportoptions']/div/a/img").click()
                    time.sleep(SMALL_TIME_SLEEP)
                    tt = False
                except:
                    pass        
            # uncheck "include research strategy"
            driver.find_element_by_xpath("//label[@name='component.SearchStrategy']").click()
            time.sleep(SMALL_TIME_SLEEP)
            # close Excel options
            driver.find_element_by_xpath("//div[@id='export-component-exportoptions']/div/a/img").click()
            time.sleep(SMALL_TIME_SLEEP)
            # select export list cell
            driver.find_element_by_id("component_selectedFormatId").click()
            time.sleep(SMALL_TIME_SLEEP)
            # select Orbis_columns_set
            Select(driver.find_element_by_id("component_selectedFormatId")).select_by_visible_text(Orbis_columns_set)
            time.sleep(SMALL_TIME_SLEEP)
            # select export from...to
            driver.find_element_by_id("component_RangeOptionSelectedId").click()
            time.sleep(SMALL_TIME_SLEEP)
            driver.find_element_by_id("component_RangeOptionSelectedId").click()
            time.sleep(SMALL_TIME_SLEEP)
            Select(driver.find_element_by_id("component_RangeOptionSelectedId")).select_by_visible_text(u"Un gruppo di società")
            time.sleep(SMALL_TIME_SLEEP)
            driver.find_element_by_name("component.From").click()
            time.sleep(SMALL_TIME_SLEEP)
            driver.find_element_by_name("component.From").clear()
            time.sleep(SMALL_TIME_SLEEP)
            driver.find_element_by_name("component.From").send_keys(str(row_start))
            time.sleep(SMALL_TIME_SLEEP)
            driver.find_element_by_name("component.To").click()
            time.sleep(SMALL_TIME_SLEEP)
            driver.find_element_by_name("component.To").clear()
            time.sleep(SMALL_TIME_SLEEP)
            driver.find_element_by_name("component.To").send_keys(str(row_end))
            time.sleep(SMALL_TIME_SLEEP)
            # write export file's name
            driver.find_element_by_id("component_FileName").click()
            time.sleep(SMALL_TIME_SLEEP)
            driver.find_element_by_id("component_FileName").clear()
            time.sleep(SMALL_TIME_SLEEP)
            driver.find_element_by_id("component_FileName").send_keys("Chunk_"+str(chunk_i+1).zfill(3))
            time.sleep(SMALL_TIME_SLEEP)
            # click "Export"
            driver.find_element_by_link_text("Esporta").click()
            time.sleep(time_before_closing_download)
            try:
                driver.find_element_by_xpath("//img[@alt='X']").click()
            except:
                pass
            time.sleep(10)
        
    print('\n\nTotal elapsed time:', str(datetime.timedelta(seconds=round(timer()-start))))
    print('\nBrowser can be closed. If exported files are not automatically downloaded, please check the Orbis "Export" tab in browser.')
    print('\n\nChunks list saved into "' + LOG_FILE + '"')
    
    # disconnect
    #     time.sleep(60*3)
    #     driver.find_element_by_xpath("//img[2]").click()
    #     driver.find_element_by_link_text("Disconnetti").click()
    
    

def generate_strategy(nation = "9IT", nace = "01"):
    # multiple values for "nace" must be saparated by ";"

    SAMPLE_STRATEGY = "Sample_Strategy.strategy"       # sample strategy file
    DATE_TO_REPLACE = "Tuesday, 12 April 2022 16:59:21"     # date reported in SAMPLE_STRATEGY
    NATION_TO_REPLACE = "9IT"                               # nation reported in SAMPLE_STRATEGY
    NACE_TO_REPLACE = "4941;495;52;551;553;552"             # NACE reported in SAMPLE_STRATEGY

    with open(SAMPLE_STRATEGY) as f:
        strategy_lines = f.readlines()
    
    strategy_lines = [s.replace(DATE_TO_REPLACE, datetime.datetime.now().strftime("%A, %m %B %Y %H:%M:%S")) for s in strategy_lines]
    strategy_lines = [s.replace('<Action Type="SingleAdd" ItemId="'+NATION_TO_REPLACE+'" />', '<Action Type="SingleAdd" ItemId="'+nation+'" />') for s in strategy_lines]
    strategy_lines = [s.replace('<Action Type="SingleAdd" Array="raw_semicolon_separated" ItemId="'+NACE_TO_REPLACE+'" />', '<Action Type="SingleAdd" Array="raw_semicolon_separated" ItemId="'+nace+'" />') for s in strategy_lines]

    return strategy_lines



def save_strategy(nation = "9IT", strategy_folder = "Strategy",
                  NACE_html_list = "Nace_rev2.html", output_text_separator = ";"):

    DESCRIPTION_FILE = "NACE_description.csv"
    
    # create strategy folder
    if not os.path.exists(strategy_folder):
        os.makedirs(strategy_folder)

    # read list of NACE codes and description
    HTMLFile = open(NACE_html_list, "r", encoding='utf-8')
    index = HTMLFile.read()
    soup = BeautifulSoup(index, 'html.parser')
    for script in soup(["script", "style"]):
        script.extract()  
    text = soup.get_text()
    sectors = (line.strip() for line in text.splitlines())
    sectors = [l for l in sectors if l != '']
    sectors = [l.split(" - ") for l in sectors]

    # create strategy files
    nace_list = ['"NACE"'+output_text_separator+'"Description"']
    for sec in sectors:
        nace = sec[0]
        descr = sec[1]

        # generate strategy
        strat_lines = generate_strategy(nation = nation, nace = nace)
        strat_name = 'Nation_'+nation+'_NACE_'+nace+'.strategy'
        with open(os.path.join(strategy_folder, strat_name), 'w') as f:
            f.write(''.join([s for s in strat_lines]))

        # update list of nace and description
        nace_list.append('"'+nace+'"'+output_text_separator+'"'+descr+'"')

    # save nace description
    with open(DESCRIPTION_FILE, 'w') as f:
        f.write('\n'.join(nace_list))
        
    print('- '+str(len(sectors))+' strategies generated. Saved in "'+strategy_folder+'"')
    print('- NACE sectors description saved in "'+DESCRIPTION_FILE+'"\n')
    
    
    
def upload_strategy(chromedriver_path=None, PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None,
                    Orbis_landing_page=None, Orbis_username=None, Orbis_pass=None, strategy_folder=None):
    
    # define strategy to upload
    file_list = os.listdir(strategy_folder)
    file_list = ['"' + os.path.join(os.getcwd(), strategy_folder, f) + '"' for f in file_list]
    max_string_len = max([len(x) for x in file_list]) + 1
    path_chunk_size = 259 // max_string_len     # upload window can handle max 259 chars
    upload_path = [" ".join(file_list[i:(i+path_chunk_size)]) for i in range(0, len(file_list), path_chunk_size)]
    
    # open browser and go to website
    print("Launching Chrome...", end ="")
    driver = get_chromedriver(chromedriver_path = chromedriver_path, use_proxy=True,
                              PROXY_HOST=PROXY_HOST, PROXY_PORT=PROXY_PORT, PROXY_USER=PROXY_USER, PROXY_PASS=PROXY_PASS)
    driver.implicitly_wait(30)
    driver.get(Orbis_landing_page)
    print("OK", end ="\n")

    # login
    print("Logging in...", end ="")
    driver.find_element_by_id("user").clear()
    driver.find_element_by_id("user").send_keys(Orbis_username)
    driver.find_element_by_id("pw").clear()
    driver.find_element_by_id("pw").send_keys(Orbis_pass)
    driver.find_element_by_id("user").click()
    driver.find_element_by_id("loginPage").click()
    driver.find_element_by_xpath("//div[@id='loginPage']/div[2]").click()
    driver.find_element_by_xpath("//div[@id='loginPage']/div[2]/div[3]/button").click()
    time.sleep(2)
    driver.find_element_by_xpath("//input[@value='Accept']").click()
    # check for session restart
    try:
        driver.find_element_by_xpath("//input[@value='Restart']").click()
        print("SESSION RESTART DONE", end ="\n")
    except:
        print("OK", end ="\n")

    # open search tab
    print("Opening search tab...", end ="")
    driver.find_element_by_xpath("//div[@id='search-toolbar']/div/div/ul/li[2]/h4").click()
    print("OK", end ="\n")

    # upload strategies
    for up_i, up in enumerate(upload_path):
        print("Uploading strategy chunk", str(up_i+1), "/", str(len(upload_path)), end="... ")
        
        driver.find_element_by_link_text("Importa strategia di ricerca").click()
        driver.find_element_by_id("upload-file-trigger").click()
        time.sleep(1.5)
        autoit.win_active("Open")
        time.sleep(1.5)
        autoit.control_send("Open","Edit1",up)
        autoit.control_send("Open","Edit1","{ENTER}")

        # check for existing strategy
        time.sleep(2)
        try:
            elem = driver.find_elements_by_css_selector("div.buttons.popup__buttons")
            # search for "Sovrascrivi tutto"
            stop = False
            for el in elem:
                if re.search('<a class="button overwrite-all">Sovrascrivi tutto</a>', el.get_attribute("outerHTML")):
                    driver.find_element_by_link_text("Sovrascrivi tutto").click()
                    print("MULTIPLE STRATEGIES OVERWRITTEN", end="\n")
                    stop = True
            # search for "Sovrascrivi"
            if not stop:
                driver.find_element_by_link_text("Sovrascrivi").click()
                print("SINGLE STRATEGY OVERWRITTEN", end="\n")
        except:
            print("OK", end ="\n")   
        time.sleep(1.5)           
    
    # close browser
    driver.close()

