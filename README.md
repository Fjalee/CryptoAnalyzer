# CryptoAnalyzer_Tomas
Web scraper that scrapes block explorers. This scraper can scrape:
* bscscan.com

## How to use
### Start/Exit
* Start the scraper by starting the executable file in a terminal.
* Exit the scraper by closing the terminal.

### Output
* Output folder: Crypto-Analyzer-Output.
* Output file: name example 2021_06_10_1306.csv.
    * File name is the date and time of the beginning of the scrape.
    * Recommended to open the file with software that can open csv as a "table" (eg. MS Excel).
   
* Csv file column TotalInaccurateValue: Sometimes the value on the website (eg. bscscan.com) is so big that the whole number doesn't fit in the table that is displayed on the website. If that happens the part of the number that was visible be seen will be placed in this column.
### OS
* Version for 64-bit Windows: windows_x64_executable/CryptoAnalyzer.exe
* Version for 64-bit Linux: linux_x64_executable/CryptoAnalyzer
### TIPS

Some software that open csv files may lock output's csv file from being changed by the scraper. If this occurs while scraping the scraper will continue scraping, however the new data will only be added to the file only after it is closed.

### Error handling
If something unexpected happens, the scraper will stop and produce a .log file that will log the html page that the scraper stopped on.

### App.config
In the source code file App.config some configuration may be changed by changing `value` strings:
* `<add key="OUTPUT_PATH" value="Crypto-Analyzer-Output" />` - output directory.
* `<add key="LOG_PATH" value="ErrorLog.log" />` - log directory.
* `<add key="OUTPUT_APPEND_PERIOD_IN_SECONDS" value="30" />` - How often the scraper adds data to the csv file. When the scraper is closed, the data that was scraped but not added is lost, therefore, this period shouldn't be too big.
* `<add key="SLEEP_TIME_IN_MILISECONDS" value="1000" />` - How long scraper sleeps after scraping one page. (1000 miliseconds - 1 second :) )
