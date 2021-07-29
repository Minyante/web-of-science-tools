# webscrapping tool for web of science

webscraper.py scrapes the web of science database for the details on the research papers related to each organization and stores it in multiple textfiles

txttocsv.py converts multiple text files in the web of science from text format (<tag> <description>) into a row-column formate and consolidates them into one large CSV file named "mergedfiles.csv"

## installation

this web scraper tool relies on selenium libraries in order to simulate chrome, selenium relies on chrome web driver which can be installed [here](https://chromedriver.chromium.org/downloads). this program was tested on ChromeDriver 92.0.4515.43 but will most likely work for other versions. your current chrome version and the downloaded chrome driver version MUST be the same.

```bash
pip install -r requirements.txt

python webscraper.py

python txttocsv.py
```

## usage

when launching the webscraper.py script, you will be prompted to enter the main directory. this directory must include the chromedriver.exe (see download above) and a file named "firmnames.xlsx" which must contain a column header "Firm Name"
  
do not interact with the automated chrome browser when running, make sure to enter credentials through the terminal 

## known issues

element not found exceptions may arise when resizing, minimizing, or switching between windows when in use 

this script may not work if the display resolution is too low as web of science renders a different webpage when the display resolution hits a certain threshold

## license
[MIT](https://choosealicense.com/licenses/mit/)
