#Stock Analysis
## Overview of the Project
The purpose of this analysis was to use VBA to help Steve analyze stock for his parents, more specifically the Daqo's stock.
## Results
###Daqo's stock
In 2017 the Daqo's stock went up 199.4% which the biggest percent increase of all the stocks i ran for 2017. However, in 2018 the Daqo's stock went down 62.6% which was the biggest perecent lost of all the stock i ran for 2018. Therefore the Daqo stock seems to be very volitile, unpredictable, and a risky stock to invest in.
###Other stocks
While in 2017 the two other biggest stock percent increases were, ENPH with 129.5%, and SEDG with 184.5%. In 2018 SEDG went down 7.8% and ENPH went up 81.9%. ENPH stock seems to be the clear winner for what stock to invest in. With it going up 211.4% in two years and recording no loss. The only other stock to record no loss in both years was the RUN stock, going uo 84% in 2018 and 5.5% in 2017. That only totaling to be a 89.5% increase for two years, which is still good but 211.4% by ENPH is much better.
###How i got the Results
To get all this information i first had to download a spread sheet of stocks for the years i wanted. The spread sheet contained the stocks ticker, starting price, and ending price. I then used VBA and for loops to loop through all the stocks with the code:
For i = 0 To 11 
ticker = tickers(i)
totalVolume = 0
After i used if then statements to pull prices of tickers with the code:
If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
startingPrice = Cells(j, 6).Value
In the end i got an excel sheet that showed me the stocks percent total loss or profit. Then realized the code would be more useful if it could run for any possible year, in the future or the past. So i refactored the code to run any year possible that has data for the stocks. After refactoring the code i set up a timer to see how long the code ran for and if it was effecient, these are the times i got:![VBA_Challenge_2017](https://user-images.githubusercontent.com/94339449/148704987-978a048d-7898-4965-b52b-16668648893f.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/94339449/148704990-c115b021-afaa-484b-8c45-158468b72307.png)
As you can see the code ran at a fast and efficient paise.
##Summary
The advantages to refactoring code are:improves code, makes code easier to understand, helps find bugs, and also can make the code run faster. The disadvantage of refactoring code are:it takes time and time can equal money, could cause more bugs, refactoring can only do so much sometimes its better to just rewrite the code. In my script refactoring was the right choice because it helped my orgianize and understand my code better, made the code run faster, and the code can now run for any year that has data availible.
