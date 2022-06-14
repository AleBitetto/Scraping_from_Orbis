import os
import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
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
                    PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None, download_folder=None):

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
    # allow multiple download
    prefs_experim = {'profile.default_content_setting_values.automatic_downloads': 1}
    
    if use_proxy:
        pluginfile = 'proxy_auth_plugin.zip'

        with zipfile.ZipFile(pluginfile, 'w') as zp:
            zp.writestr("manifest.json", manifest_json)
            zp.writestr("background.js", background_js)
        chrome_options.add_extension(pluginfile)
    if user_agent:
        chrome_options.add_argument('--user-agent=%s' % user_agent)
    if download_folder:
        prefs_experim["download.default_directory"] = download_folder

    chrome_options.add_experimental_option("prefs", prefs_experim)
    driver = webdriver.Chrome(
        executable_path=chromedriver_path,
        chrome_options=chrome_options)
    return driver



def login_and_reset(chromedriver_path=None, use_proxy=False, user_agent=None,
                    PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None, download_folder=None,
                    Orbis_landing_page=None, Orbis_username=None, Orbis_pass=None):

    # open browser and go to website
    print("Launching Chrome...", end ="")
    driver = get_chromedriver(chromedriver_path = chromedriver_path, use_proxy=use_proxy, user_agent=user_agent,
                              PROXY_HOST=PROXY_HOST, PROXY_PORT=PROXY_PORT, PROXY_USER=PROXY_USER, PROXY_PASS=PROXY_PASS,
                              download_folder=download_folder)
    driver.implicitly_wait(5)
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
        
    return driver



def scrape_from_Orbis(chromedriver_path=None, PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None,
                     Orbis_landing_page=None, Orbis_username=None, Orbis_pass=None, Orbis_saved_search=None,
                     Orbis_columns_set=None, maximum_chunk_size=None, time_before_closing_download=60,
                     resume_from_chunk=-1, chunk_folder='', chunk_label='', loop=False, input_driver=None,
                     download_folder=None):
    
    label_sep = '-' if chunk_label != '' else ''
    LOG_FILE = os.path.join(chunk_folder, chunk_label+label_sep+'Chunks_list.txt')
    SMALL_TIME_SLEEP = 0.5
    start = timer()

    # create chunk list folder
    if not os.path.exists(chunk_folder):
        os.makedirs(chunk_folder)
        
    # create download folder
    if download_folder:
        if not os.path.exists(download_folder):
            os.makedirs(download_folder)
    
    if input_driver is None:
        
        # login and reset
        driver = login_and_reset(chromedriver_path=chromedriver_path, use_proxy=True, PROXY_HOST=PROXY_HOST,
                                 PROXY_PORT=PROXY_PORT, PROXY_USER=PROXY_USER, PROXY_PASS=PROXY_PASS,
                                 download_folder=download_folder, Orbis_landing_page=Orbis_landing_page,
                                 Orbis_username=Orbis_username, Orbis_pass=Orbis_pass)
    else:
        driver = input_driver

    # load search and show results
    print("Resetting search criteria...", end ="")
    driver.find_element_by_xpath("//div[@id='search-toolbar']/div/div/ul/li[2]/h4").click()
    try:   # check if "reset search" is available
        driver.find_element_by_link_text("Ricomincia").click()
        time.sleep(1)
        driver.find_element_by_xpath("//a[contains(text(),'Ricomincia')]").click()
        time.sleep(5)
        driver.find_element_by_xpath("//div[@id='search-toolbar']/div/div/ul/li[2]/h4").click()
        print("OK", end ="\n")
    except:
        print("NOT FOUND, criteria already resetted", end ="\n")
    time.sleep(2)
    print("Showing results...", end ="")
    driver.find_element_by_xpath("//div[@id='tooltabSectionload-search-section']/div/div/div/input[2]").click()  # search strategy in long list
    time.sleep(1)
    driver.find_element_by_xpath("//div[@id='tooltabSectionload-search-section']/div/div/div/input[2]").clear()
    driver.find_element_by_xpath("//div[@id='tooltabSectionload-search-section']/div/div/div/input[2]").send_keys(Orbis_saved_search)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//div[@id='tooltabSectionload-search-section']/div/div[2]/div/div/div[2]/ul[2]/li/div/span")))
    driver.find_element_by_xpath("//div[@id='tooltabSectionload-search-section']/div/div[2]/div/div/div[2]/ul[2]/li/div/span").click()
    time.sleep(10)
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//img[@alt='Risultati']")))
    driver.find_element_by_xpath("//img[@alt='Risultati']").click()
    print("OK", end ="\n")
    
    # split into chunks
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'records-count')))
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
        
        if chunk_i >= resume_from_chunk:
            print("Querying chunk", str(chunk_i+1), "/", str(len(chunk_index)), end = "\r")

            row_start = chunk[0]
            row_end = chunk[1]

            # open Excel tab
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.LINK_TEXT, 'Excel')))
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
            chunk_file_name = chunk_label+label_sep+"Chunk_"+str(chunk_i+1).zfill(3)
            driver.find_element_by_id("component_FileName").send_keys(chunk_file_name)
            time.sleep(SMALL_TIME_SLEEP)
            # click "Export"
            driver.find_element_by_link_text("Esporta").click()
            time.sleep(time_before_closing_download)
            try:
                driver.find_element_by_xpath("//img[@alt='X']").click()
            except:
                pass
            time.sleep(10)
        
        # save list of chunks' rows
        with open(LOG_FILE, 'a') as f:
            f.write('\nChunk ' + str(chunk_i+1) + ': ' + '-'.join([str(i) for i in chunk]))
        
    print('\n\nTotal elapsed time:', str(datetime.timedelta(seconds=round(timer()-start))))
    if not loop:
        print('\nBrowser can be closed. If exported files are not automatically downloaded, please check the Orbis "Export" tab in browser.')
        print('\n\nChunks list saved into "' + LOG_FILE + '"')
    
    if loop:
        return driver
    
    # disconnect
    #     time.sleep(60*3)
    #     driver.find_element_by_xpath("//img[2]").click()
    #     driver.find_element_by_link_text("Disconnetti").click()
    
    

def generate_strategy(nation = "9IT", nace = "01"):
    # multiple values for "nace" must be saparated by ";"

    SAMPLE_STRATEGY = "Sample_Strategy.strategy"            # sample strategy file
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
    
    # login and reset
    driver = login_and_reset(chromedriver_path=chromedriver_path, use_proxy=True, PROXY_HOST=PROXY_HOST,
                             PROXY_PORT=PROXY_PORT, PROXY_USER=PROXY_USER, PROXY_PASS=PROXY_PASS,
                             Orbis_landing_page=Orbis_landing_page, Orbis_username=Orbis_username, Orbis_pass=Orbis_pass)
    
#     # open browser and go to website
#     print("Launching Chrome...", end ="")
#     driver = get_chromedriver(chromedriver_path = chromedriver_path, use_proxy=True,
#                               PROXY_HOST=PROXY_HOST, PROXY_PORT=PROXY_PORT, PROXY_USER=PROXY_USER, PROXY_PASS=PROXY_PASS)
#     driver.implicitly_wait(5)
#     driver.get(Orbis_landing_page)
#     print("OK", end ="\n")

#     # login
#     print("Logging in...", end ="")
#     driver.find_element_by_id("user").clear()
#     driver.find_element_by_id("user").send_keys(Orbis_username)
#     driver.find_element_by_id("pw").clear()
#     driver.find_element_by_id("pw").send_keys(Orbis_pass)
#     driver.find_element_by_id("user").click()
#     driver.find_element_by_id("loginPage").click()
#     driver.find_element_by_xpath("//div[@id='loginPage']/div[2]").click()
#     driver.find_element_by_xpath("//div[@id='loginPage']/div[2]/div[3]/button").click()
#     time.sleep(2)
#     driver.find_element_by_xpath("//input[@value='Accept']").click()
#     # check for session restart
#     try:
#         driver.find_element_by_xpath("//input[@value='Restart']").click()
#         print("SESSION RESTART DONE", end ="\n")
#     except:
#         print("OK", end ="\n")

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


    
def get_number_of_firms(chromedriver_path=None, PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None,
                        Orbis_landing_page=None, Orbis_username=None, Orbis_pass=None, strategy_folder=None,
                        output_text_separator = ";"):
    
    OUT_FILE = "firms_by_NACE.csv"

    # reset firms list file
    with open(OUT_FILE, 'w') as f:
        f.write('"NACE";"Firms"')

    # get list of strategy
    strategy_list = os.listdir(strategy_folder)
    strategy_list = [s.replace(".strategy", "") for s in strategy_list]
    
    # login and reset
    driver = login_and_reset(chromedriver_path=chromedriver_path, use_proxy=True, PROXY_HOST=PROXY_HOST,
                             PROXY_PORT=PROXY_PORT, PROXY_USER=PROXY_USER, PROXY_PASS=PROXY_PASS,
                             Orbis_landing_page=Orbis_landing_page, Orbis_username=Orbis_username, Orbis_pass=Orbis_pass)
    
#     # open browser and go to website
#     print("Launching Chrome...", end ="")
#     driver = get_chromedriver(chromedriver_path = chromedriver_path, use_proxy=True,
#                               PROXY_HOST=PROXY_HOST, PROXY_PORT=PROXY_PORT, PROXY_USER=PROXY_USER, PROXY_PASS=PROXY_PASS)
#     driver.implicitly_wait(5)
#     driver.get(Orbis_landing_page)
#     print("OK", end ="\n")

#     # login
#     print("Logging in...", end ="")
#     driver.find_element_by_id("user").clear()
#     driver.find_element_by_id("user").send_keys(Orbis_username)
#     driver.find_element_by_id("pw").clear()
#     driver.find_element_by_id("pw").send_keys(Orbis_pass)
#     driver.find_element_by_id("user").click()
#     driver.find_element_by_id("loginPage").click()
#     driver.find_element_by_xpath("//div[@id='loginPage']/div[2]").click()
#     driver.find_element_by_xpath("//div[@id='loginPage']/div[2]/div[3]/button").click()
#     time.sleep(2)
#     driver.find_element_by_xpath("//input[@value='Accept']").click()
#     # check for session restart
#     try:
#         driver.find_element_by_xpath("//input[@value='Restart']").click()
#         print("SESSION RESTART DONE", end ="\n\n")
#     except:
#         print("OK", end ="\n\n")
        
    # load search and save number of firms
    tot_firms_overall = 0
    for st_i, st in enumerate(strategy_list):

        print("Searching", st, str(st_i+1), "/", str(len(strategy_list)), end="... ")

        # reset search criteria
        driver.find_element_by_xpath("//div[@id='search-toolbar']/div/div/ul/li[2]/h4").click()
        try:   # check if "reset search" is available
            driver.find_element_by_link_text("Ricomincia").click()
            driver.find_element_by_xpath("//a[contains(text(),'Ricomincia')]").click()
            time.sleep(2)
            driver.find_element_by_xpath("//div[@id='search-toolbar']/div/div/ul/li[2]/h4").click()
        except:
            pass
        time.sleep(1)

        # apply new search
        driver.find_element_by_xpath("//div[@id='tooltabSectionload-search-section']/div/div/div/input[2]").click()  # search strategy in long list
        driver.find_element_by_xpath("//div[@id='tooltabSectionload-search-section']/div/div/div/input[2]").clear()
        driver.find_element_by_xpath("//div[@id='tooltabSectionload-search-section']/div/div/div/input[2]").send_keys(st)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//div[@id='tooltabSectionload-search-section']/div/div[2]/div/div/div[2]/ul[2]/li/div/span")))
        driver.find_element_by_xpath("//div[@id='tooltabSectionload-search-section']/div/div[2]/div/div/div[2]/ul[2]/li/div/span").click()

        # get number of firms
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'records-count')))
        tot_firms = driver.find_element_by_id('records-count').get_attribute('data-total-records-count')
        tot_firms_overall += int(tot_firms)

        # append and save firms number
        with open(OUT_FILE, 'a') as f:
            f.write('\n"' + st.split('NACE_')[1] + '"' + output_text_separator + '"' + tot_firms + '"')

        print("Firms:", tot_firms)

    print('\n- Total firms:', str(tot_firms_overall))
    print('\n- Total number of firms by NACE saved in "'+OUT_FILE+'"\n')
    
    
def check_downloaded_files(chunk_folder="", download_folder="", firms_by_nace="", firms_by_nace_separator=";"):

    OUT_FILE = "firms_by_NACE_after_download.csv"

    # reset firms list file
    with open(OUT_FILE, 'w') as f:
        f.write('"NACE";"Firms"')

    # get chunks' list for each sector
    chunk_list = os.listdir(chunk_folder)

    # get expected total firms by sector
    with open(firms_by_nace) as f:
        exp_firms = f.readlines()
    exp_firms = [[s.replace('"', '').replace('\n', '') for s in x.split(firms_by_nace_separator)] for x in exp_firms[1:]]

    # loop sectors and check downloaded chunks
    missing_all = []
    for ch_i, ch in enumerate(chunk_list):

        base_chunk_name = ch.replace('s_list.txt', '_')
        sector = ch.split('-')[1].split('NACE_')[1]

        # select expected total rows
        ind = [i for i, v in enumerate(exp_firms) if v[0] == sector][0]
        exp_total_rows = int(exp_firms[ind][1])

        # read expected chunks and total rows
        with open(os.path.join(chunk_folder, ch)) as f:
            lines = f.readlines()
        total_chunk = int(lines[-1].split(':')[0].replace("Chunk ", ""))
        total_rows = int(lines[-1].split('-')[1])

        # generate expected chunks
        missing_chunks = []
        for i in range(total_chunk):
            chunk_name = base_chunk_name+str(i+1).zfill(3)+'.xlsx'
            if not os.path.exists(os.path.join(download_folder, chunk_name)):
                missing_chunks.append(chunk_name)
                missing_all.append(chunk_name)

        # create warnings
        warn_tot_rows = exp_total_rows != total_rows
        warn_missing_chunks = len(missing_chunks) > 0

        # print warnings
        if warn_tot_rows or warn_missing_chunks:
            print('\n-- Sector', sector+':')
        if warn_tot_rows:
            print('     *expected rows:', exp_total_rows, '  -> downloaded:', total_rows)
        if warn_missing_chunks:
            print('     *missing chunks:\n              ', '\n               '.join(missing_chunks))

        # append and save firms number
        with open(OUT_FILE, 'a') as f:
            f.write('\n"' + sector + '"' + firms_by_nace_separator + '"' + str(total_rows) + '"')

    print('\n\n\n- Total number of firms by NACE saved in "'+OUT_FILE+'"\n')
    
    return missing_all



def download_missing(chromedriver_path=None, PROXY_HOST=None, PROXY_PORT=None, PROXY_USER=None, PROXY_PASS=None,
                     Orbis_landing_page=None, Orbis_username=None, Orbis_pass=None, download_folder=None, missing_list=None):
    
    # login and reset
    driver = login_and_reset(chromedriver_path=chromedriver_path, use_proxy=True, PROXY_HOST=PROXY_HOST,
                             PROXY_PORT=PROXY_PORT, PROXY_USER=PROXY_USER, PROXY_PASS=PROXY_PASS,
                             download_folder=download_folder, Orbis_landing_page=Orbis_landing_page,
                             Orbis_username=Orbis_username, Orbis_pass=Orbis_pass)
    print('\n\n')
    
    # move to export tab
    driver.find_element_by_xpath("//img[@alt='Esportazioni']").click()
    time.sleep(10)
    
    # get day folder to expand
    elem = driver.find_elements_by_css_selector("div.exportsList")
    html_file = BeautifulSoup(elem[0].get_attribute("outerHTML"), 'html.parser')
    opener_subset = html_file.find_all( class_ = "opener" )

    day_list = []
    for op in opener_subset:
        try:
            day_lab = op["aria-label"]
            html_to_string = str(op)
            day_count = html_to_string.split(day_lab+' (<span class="folderExportsCount">')[1].split('</span>')[0]
            day_list.append(day_lab+' ('+day_count+')')
        except:
            pass

    # expand folder
    for dd in day_list[1:]:
        driver.find_element_by_link_text(dd).click()
        time.sleep(3)
        
    # loop missing files and download
    block_list = elem[0].get_attribute("outerHTML").split('\n\n\n')
    for ch_path in missing_list:

        ch = '-'.join(ch_path.split('-')[1:]).replace('.xlsx', '')
        print('-', ch, end=": ")
        ind = []
        for i, ss in enumerate(block_list):
            if re.search(ch, ss):
                ind.append(i)
                sub_str = ss
        if len(ind) == 1:
            # check if file is ready to download
            available = [s.split('Disponibile fino al ')[1] for s in sub_str.split('\n') if re.search('Disponibile fino al', s)][0]
            if len(available) > 0:
                # get href path for download button
                download_href = [s.split('<a class="exportDownload" href="')[1] for s in sub_str.split('\n') if re.search('<a class="exportDownload" href="', s)][0]
                download_href = download_href.split('"')[0]
                # download
                driver.find_element_by_xpath('//a[@href="'+download_href+'"]').click()
                time.sleep(5)
                print('Downloaded')
            else:
                print('Not available yet')
        elif len(ind) == 0:
            print('##### No record found #####')
        else:
            print('##### Multiple indices found #####')