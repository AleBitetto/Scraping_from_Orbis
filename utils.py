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
                      Orbis_columns_set=None, maximum_chunk_size=None):
    
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
    time.sleep(10)
    driver.find_element_by_xpath("//img[@alt='Risultati']").click()
    print("OK", end ="\n")

    # split into chunks
    total_rows = int(driver.find_element_by_id('records-count').get_attribute('data-total-records-count'))
    print("Total records found:", str(total_rows))
    rows = range(1,total_rows+1)
    chunk_index = [(rows[i:i+maximum_chunk_size][0], rows[i:i+maximum_chunk_size][-1]) for i in range(0, len(rows), maximum_chunk_size)]
    print("Splitting into", len(chunk_index), "chunks")

    # export chunks into excel files
    for chunk_i, chunk in enumerate(chunk_index):

        print("Querying chunk", str(chunk_i+1), "/", str(len(chunk_index)), end = "\r")

        row_start = chunk[0]
        row_end = chunk[1]

        driver.find_element_by_link_text("Excel").click()
        time.sleep(SMALL_TIME_SLEEP)
        driver.find_element_by_xpath("//div[@id='export-component-exportoptions']/div/a/img").click()
        time.sleep(SMALL_TIME_SLEEP)
        driver.find_element_by_xpath("//label[@name='component.SearchStrategy']").click()
        time.sleep(SMALL_TIME_SLEEP)
        driver.find_element_by_xpath("//div[@id='export-component-exportoptions']/div/a/img").click()
        time.sleep(SMALL_TIME_SLEEP)
        driver.find_element_by_id("component_selectedFormatId").click()
        time.sleep(SMALL_TIME_SLEEP)
        Select(driver.find_element_by_id("component_selectedFormatId")).select_by_visible_text(Orbis_columns_set)
        time.sleep(SMALL_TIME_SLEEP)
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
        driver.find_element_by_id("component_FileName").click()
        time.sleep(SMALL_TIME_SLEEP)
        driver.find_element_by_id("component_FileName").clear()
        time.sleep(SMALL_TIME_SLEEP)
        driver.find_element_by_id("component_FileName").send_keys("Chunk_"+str(chunk_i+1).zfill(3))
        time.sleep(SMALL_TIME_SLEEP)
        driver.find_element_by_link_text("Esporta").click()
        time.sleep(5)
        driver.find_element_by_xpath("//img[@alt='X']").click()
        time.sleep(10)
        
    print('\n\nTotal elapsed time:', str(datetime.timedelta(seconds=round(timer()-start))))
    print('\nBrowser can be closed. If exported files are not automatically downloaded, please check the Orbis "Export" tab in browser.')
    
    # disconnect
    #     time.sleep(60*3)
    #     driver.find_element_by_xpath("//img[2]").click()
    #     driver.find_element_by_link_text("Disconnetti").click()