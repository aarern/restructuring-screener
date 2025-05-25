# Restructuring Screener

Description:
This project pulls financial data from all of the stocks listed on NASDAQ via Python (Upload-file.xlsx), then utilizes Excel VBA and Capital IQ to analyze their bond prices and maturity dates to find discounted bonds and maturity walls, indicating if a business is in distress and if it needs corporate restructuring solutions (NASDAQ-Distressed-Watchlist.xlsm).

This script purposefully casts a wider net when finding distressed prospects because with this data, you can track if a company is moving towards bankruptcy on a quarterly basis.

If you are a student, check to see if your school pays for CapIQ and install the plugin. If not, the initial screening part will still work.

Initially, this was one large file, but it became unusable as the more there were array formulas added increased, the more CapIQ calls increased, and trying to tie VBA, CapIQ, and Python together. To fix this, I had to use fewer dynamic array formulas and instead use named ranges. I also split the initial screening call into its own Excel file that you can copy and paste over because the file would corrupt when using all three together. Also, there are two restructuring scripts because the Excel TEXTJOIN function has a character limit that is less than the combination of all NASDAQ tickers, so I had to split it into two named ranges. Also, Excel would reject another large scrape after running the first TEXTJOIN function, so I just made a second script to run the same process but with the second TEXTJOIN cell.

Update the named ranges as you rescreen for current liquidity ratios.

How to get it working:
*If the file is corrupted:
If you get a message saying "Excel cannot Open The File Because The File Format or File Extension is Not Valid," see: https://www.youtube.com/watch?v=SobYKTdwY80

If the Python script progress bar doesn't upload, then save the script and Excel file, then reopen and try again.

Pips to install: 
1. pip install pajama
2. pip install pandas openpyxl yfinance
3. pip install tqdm

Trust the macro-enabled Excel file:
Open Excel settings (Alt + F T). 
Go to the "Customize Ribbon" tab and check the "Developer" box. 
Open the "Trust Center" from the Excel settings. Select "Trust Center Settings." 
Go to the "Trusted Locations" tab and select "Add a new location." 
Browse to the location where the macro-enabled workbook will be saved. 
Enable subfolders of this location to be trusted. If using a Linux terminal, on the "Trusted Locations" page, check the box that says "Allow Trusted Locations on my network (not recommended)"

Excel Workbook Logistics:
1. Output_Bonds_1 and Output_Bonds_2 use CIQRANGE functions and if you rescreen for NASDAQ companies click the button on the "Documentation" sheet to clear the ranges under the CIQ functions, then save and reopen and CapIQ will repopulate the cells
2. The VBA script will output the unique identifier for every bond the company has issued and repeat the tickers accordingly. Once finished it is used to find the trading activity of the corporate bonds and the maturity walls of the companies. The final sorted array sorts the stocks by the % of bonds trading to zero and then the number of bonds the company has issued.t
3. For the Maturity_Walls_Test you can modify what timeframe you want all of the bonds to mature in. I chose 30 days because that seemed the most strict for finding an indisputably large maturity wall for subject companies
