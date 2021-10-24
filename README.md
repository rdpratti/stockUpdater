# Stock Spreadsheet Updater

The main purpose of this project is to automate the maintenance of a stock spreadsheet.

## Description

Given s spreadsheet used to maintain a person's stock holdings, what is the easiest method to :
1. updates the spreadsheet with up to date stock prices;
2. calculate the total value of the portfolio.

We want to develop a script that reads in a spreadsheet file (*.xlsx), with each row a stock holding.
The second column is the ticker symbol. The third column is the stock price.

The script uses the yahoo finance api to find the current stock price.
It updates column three on each stock row with that current price.

The updated spreadsheet is saved in the same directory.
The new file name has the current date to the file name.


## Getting Started

To get started, examine these resources :

1. examine the sample stock spreadsheet;
2. confirm you enivronment hhas the required dependencies;
3. execute stockUpdater.py script passing a stock spreadsheet as a parameter 

### Dependencies

See resources.txt for details
1. python 3.8
2. yfinance~=0.1.63
3. openpyxl=3.0.9

## Authors

Contributors names and contact info

Roland DePratti     roland.depratti@comcast.net

## Version History

* 0.1
    * Initial Release

## License

This project is licensed under the GNU LGPLV3 License.

## Acknowledgments

The following articles were helpful in developing this project
* [Python: How to Get Live Market Data (Less Than 0.1-Second Lag)](https://towardsdatascience.com/python-how-to-get-live-market-data-less-than-0-1-second-lag-c85ee280ed93)
* [Read and Update Excel Files](https://www.pythonpip.com/python-tutorials/how-to-read-update-excel-file-using-python/)
