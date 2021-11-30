# Refractoring  All Stocks Analysis

### Combining entire Stockmarket data over the last few years and looping all the data at one time in order to make the code more efficient to use and read.

## Steve’s Final Request for All Stocks Analysis

### Steve wanted to create a workbook to help analyze his parents choice in stock, then how it compared to other stocks.  He finally wanted to be able to compare all stocks for all years.  This required some refactoring of the previous workbook.  He wanted also wanted to include a run-time pop-up message to measure time performance on the data.

### I used the previous macro information and  changed the code after the RowCount was created.  I first created a ticker index. I then created three output arrays with tickerVolume, tickerStartingPrice, and tickerEndingPrice. These values will get the Total Daily Volume and Return on the worksheet once the codes are completed. I then used a for loop to initialize tickerVolumes to 0. Next I started a loop over all the rows within this loop, I increased the volume for the current ticker. aI then  used If/Then coding to find if  the current row is the first row within the selected tickerIndex. I Used another If/Then to find if the current row was the last row with the selected ticker. The last If/Then was used to if the next row’s ticker didn’t match, to increase the tickerIndex. Then I ended the loop. Once that Loop was completed, I made another loop  to go through the arrays (tickers) to output for Ticker, Total Daily Volume and Return using tickers, tickerVolume, tickerStartingPrice and tickerEndingPrice.

### This code allows the reader to view the year they want, quickly.  The run-time pop-up was already imbedded into the macro and the given results for both 20017 and 2018, you can see here:

<img width="1440" alt="Screen Shot 2021-11-30 at 1 02 44 PM" src="https://user-images.githubusercontent.com/93801125/144123017-ef6755f4-f2d9-4650-b0de-97d57e511ca8.png">

<img width="1440" alt="Screen Shot 2021-11-30 at 1 03 06 PM" src="https://user-images.githubusercontent.com/93801125/144123045-4d71adda-9b42-4cd4-afea-a85da159cbe0.png">

## Troubles with VBA

### I had a tough time sorting through and understanding what was being asked to input into the data.  With much research I was able to get close, but not quite where I needed to be. I did a lot of de-bugging and getting help from others to check what was wrong with the coding I did.  A good portion of it was just adding an extra letter (like ’s’ to tickers) where I shouldn’t or adding lines in the code that were unnecessary (even though the base workbook used it).  

## Results

### With the macro that was refractored, I was able to create a workbook that allows the user (Steve) to simply put in the year and the output would be viewed easily.  Losses and Gains for the stocks are color coded for quick and easy reading and understanding at a glance. The run-time pop-up allows Steve to understand how efficiently the data is being processed.  Should he want to add more years to his data, he can continue to monitor the efficiency of the macro that was created. Since the macro is connected to the buttons on
