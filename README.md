# **Meghan's Stock-Analysis VBA project**

## **Overview:** 
In this analysis, we have used a data set on stock price fluctuations for 2017 and 2018. Our client originally wanted the data analyzed to give him a clearer picture of the stock landscape for certain “green” stocks. After analyzing the data, we were able to pull out relevant insights for our 10 chosen stocks. This included calculating starting and ending prices and total volume of stock traded to give us a clear view of the stocks profitability for the years 2017 and 2018. After creating this sheet, we went back and refactored the code. By doing this, we are able to produce calculations in a shorter time period. 

## **Results:**
### Stock performance
Based on our data in the spreadsheet we can clearly see 2017 was a much better year for this collection of stocks over 2018. Nine out of ten produced a positive return in 2017 with 4 of this returns exceeding 100%- a very good year indeed! Compare this with 2018, and we find only 2 stocks of our original ten yielded a positive return. This massive shift should make clear that this is a volatile market for investors without many safe bets. It also leave this analyst hungry for the 2019 and 2020 data (at least) to further chart out trends here.  A key recommendation of mine would be to seek more data. It also begs further research on the differences in 2017 and 2018. Why the massive downswing in return? Did something happen in 2018 to directly affect stock performance in this sector? The answer to that question can give us more information in how we judge these stocks. 

### **Stock Recommendation**
If we are looking to pick the best stock or two in which to invest, then we can narrow it down to ENPH or RUN as our safest bets based on what we see from this data. These two are the only ones to generate a return two years running. However, even then there is much volatility. RUN has seen a tremendous gain in return going from 9.5% in 2017 to 85.2%. Will this point to continued gains in the future or is it evidence of an unstable and thereby riskier stock? ENPH’s were on yearly trend in that the percentage of return was diminished in 2018 up from a high of 136.3% in 2017 down to 97.9%. Rather than evidence of future decline, we could see this as stable returns on trend with the industry. 

### **Execution Times**
After refactoring the code, it clearly ran at greater speed proving the efficiency.  
**2017** ran in **0.5625** seconds in original script then calculated in **0.109375** seconds after refactoring.
**2018** ran in **3.210938** seconds in original script then calculated in **0.09375** seconds after refactoring.

## **Summary:** 
This challenge has been a great lesson in the realities of refactoring code and time management. On one hand, it’s clear that by going through code and reworking it for efficiency, you can produce a better-running piece of code. Putting aside extra time to refactor also allows the coder more time to understand the code and how it works. This would prove valuable for any files used often or likely to be duplicated and used for other purposes. The key disadvantage is programmer time spent. It goes back to the old adage “If it’s not broke, don’t fix it.” At times when a coder is working on a project and constrained on time, I would imagine a working code is more important than a perfectly efficient code. This comes back to the importance of using our good judgement in how we manage our projects and time. 

For me personally with the VBA script, this refactoring has given me some better understanding of how this can be used. However, the time it took me to refactor…. I’m not sure if I care that much about the speed improvement of the sheet calculations as a strong tangible benefit for the time it required of me. 
