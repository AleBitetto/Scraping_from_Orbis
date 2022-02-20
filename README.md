# Scraping_Orbis
Python notebook to scrape data from Orbis

### Installation

### Chrome Web Driver

To download data, this library relies on [`selenium`](https://selenium-python.readthedocs.io/) and [`ChromeDriver`](https://chromedriver.chromium.org/).

This requires a `chromedriver` executable which can be downloaded [here](https://chromedriver.chromium.org/downloads). Make sure that your `Chrome` version is the same as your `chromedriver` version.

`Sccraping_Orbis` assumes that the `chromedriver` executable is located at `/WebDriver` in the main folder. To supply a different path, change the variable `chromedriver_path` in the notebook.

### Credentials

Credentials must be input manually into a `config.py` file. Create a text file with the following lines:
```
Orbis_username = "username"
Orbis_pass = "password"
PROXY_USER = "username"
PROXY_PASS = "password"
```
where you need to write the Orbis username and password and Ateneo credentials (Fiscal Code and password). Then simply rename the file as `.py`

### Usage

Simply run the notebook.
Current version works with the Italian language settings and Orbis settings as described in the notebook.
