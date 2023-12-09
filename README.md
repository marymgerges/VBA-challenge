# VBA-challenge
12/08/2023:
I updated my Module2.vbs code to create headers headers as well as to auto fit the columns to allow a clear view of the text.

----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
For this challenge, I created a VBA script that looped through all worksheets in a workbook.
It extracted the yearly open and yearly close tickers and values by checking if the next row has the same or different tickers, and then it recorded this information for yearly change and its corresponding ticker in the appopriate cells.
Then I had it calculate the total stock volume for each ticker and place that value in the appropriate cell.
Next, I had it calculate the yearly change value. If the value was negative, I had it fill the cell with red and if it was positive, then I had it fill the cell with green.
I accounted for if the yearly open value was zero, so that I did not get an undefined result for percent change because of a zero in the denominator.
My next block of code calculated the greatest percent increase and found the corresponding ticker for it.
Similarly, I calculated the greatest percent decrease and found its corresponding ticker. Then I did the same thing for total stock volume and its ticker.
I set each of those values where they should be, and I formatted how the percent change value should be shown.
After that, I had it go to the next ticker symbol and reset the total stock volume value.
Then it would call the next iteration.
After that, I placed the values for greatest percent increase, greatest percent increase ticker, greatest percent decrease, greatest percent decrease ticker, greatest total volume, and greatest total volume ticker in the appropriate cells. 
I also formatted each of the values to the appropriate formats.
Lastly, I had it call the next sheet in the workbook so it can do the entire code again for it.
I worked with tutor Mark Fullton, who helped guide me throughout this assignment and gave me tips on certain areas I was stuck on.
