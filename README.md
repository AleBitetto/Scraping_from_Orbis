# Scraping from Orbis
Python notebook to scrape data from Orbis.

### Installation

Clone the repository with
`git clone https://github.com/AleBitetto/Scraping_from_Orbis.git`

From console navigate the cloned folder and create a new environment with:
```
conda env create -f environment.yml
conda activate scraping_Orbis
python -m ipykernel install --user --name scraping_Orbis --display-name "Python (Scraping from Orbis)"
```
This will create a `scraping_Orbis` environment and a kernel for Jupyter Notebook called `Python (Scraping from Orbis)`



### Chrome Web Driver

To download data, this notebook relies on [`selenium`](https://selenium-python.readthedocs.io/) and [`ChromeDriver`](https://chromedriver.chromium.org/).

This requires a `chromedriver` executable which can be downloaded [here](https://chromedriver.chromium.org/downloads). Make sure that your `Chrome` version is the same as your `chromedriver` version.

`Sccraping_Orbis` assumes that the `chromedriver` executable is located at `/WebDriver` in the main folder. To supply a different path, change the variable `chromedriver_path` in the notebook.

### Credentials

Credentials must be inputed manually into a `config.py` file. Create a text file with the following lines:
```
Orbis_username = "username"
Orbis_pass = "password"
PROXY_USER = "username"
PROXY_PASS = "password"
```
where you need to write the Orbis username and password and Ateneo credentials (Fiscal Code and password). Then simply rename the file as `.py`

### Usage

Simply run the [`notebook`](https://github.com/AleBitetto/Scraping_from_Orbis/blob/master/Scraping%20from%20Orbis.ipynb).

Current version works with the Italian language settings and Orbis settings as described in the notebook.
