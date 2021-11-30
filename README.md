# Refactoring  All Stocks Analysis

### Combining entire Stockmarket data over the last few years and looping all the data at one time in order to make the code more efficient to use and read.

## Steve’s Final Request for All Stocks Analysis

### Steve wanted to create a workbook to help analyze his parents choice in stock, then how it compared to other stocks.  He finally wanted to be able to compare all stocks for all years.  This required some refactoring of the previous workbook.  He wanted also wanted to include a run-time pop-up message to measure time performance on the data.

## Resluts

### I used the previous macro information and  changed the code after the RowCount was created.  I first created a ticker index. I then created three output arrays with tickerVolume, tickerStartingPrice, and tickerEndingPrice. These values will get the Total Daily Volume and Return on the worksheet once the codes are completed. I then used a for loop to initialize tickerVolumes to 0. Next I started a loop over all the rows within this loop, I increased the volume for the current ticker. aI then  used If/Then coding to find if  the current row is the first row within the selected tickerIndex. I Used another If/Then to find if the current row was the last row with the selected ticker. The last If/Then was used to if the next row’s ticker didn’t match, to increase the tickerIndex. Then I ended the loop. Once that Loop was completed, I made another loop  to go through the arrays (tickers) to output for Ticker, Total Daily Volume and Return using tickers, tickerVolume, tickerStartingPrice and tickerEnding. 

### This code allows the reader to view the year they want, quickly.  The run-time pop-up was already imbedded into the macro. You can see the run-time was better for both 2017 and 2018 here:

<img width="1440" alt="Screen Shot 2021-11-30 at 2 44 14 PM" src="https://user-images.githubusercontent.com/93801125/144132155-eb6650e5-2251-4de2-874d-4e1e7147e0ae.png">


<img width="1440" alt="Screen Shot 2021-11-30 at 1 02 44 PM" src="https://user-images.githubusercontent.com/93801125/144123017-ef6755f4-f2d9-4650-b0de-97d57e511ca8.png">


<img width="1440" alt="Screen Shot 2021-11-30 at 2 44 32 PM" src="https://user-images.githubusercontent.com/93801125/144132209-28c99188-48bf-4fa5-8397-b3ff9a5c74e7.png">


<img width="1440" alt="Screen Shot 2021-11-30 at 1 03 06 PM" src="https://user-images.githubusercontent.com/93801125/144123045-4d71adda-9b42-4cd4-afea-a85da159cbe0.png">

### As you can also see with the color coding imbedded in the code, you can easily tell which Ticker had gain and which had loss.  In bot years, "ENPH" and "RUN" both had gain, with "ENPH" doing the best between the two.  "DQ", which is where Steve's parents want to invest, Did well in 2017, but had a big loss in 2018.  Suggesting a change to either "ENPH" or "Run" would be advisable.


## Summary

###The advantages of refactoring code is better quality coding. It also a good way to maintain the code and keeping it clean and organized.  The disadvantages are that you may have to retest a lot of the functionality of the code in order to determine if the code still works, and debug any issues that come up. If the application is big, it can be harder and more time consuming to debug. When I was trying to refactor the original VBA, it helped streamline it a bit, but some of the coding I was still using from the original, had to be reformatted. Also figuring out what still applies, and what is hurting the code was a little hard to decipher.
