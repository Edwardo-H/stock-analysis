# Refactoring Module 2 VBA Code and Measuring Performance
---
---
## Overview of Project
---
Steve, a recent college graduate with a finance degree, had asked for assistance with analyzing green energy stocks on behalf of his parents who are interested in investing. Steve was impressed by the original workbook that was prepared and would now like to expand his dataset to include the entire stock market.
---
### Purpose
---
Using Visual Basic for Applications (VBA) available in Excel, this project attempts to refactor the original code to once again loop through the original data and determine if the VBA script will run faster after refactoring the code. To prove the success of this endeavor, the following outputs will be provided:
- Pop-up message showing elapsed time to run analysis for 2017
- Pop-up message showing elapsed time to run analysis for 2018
---
### Results
---
Unfortunately, I was unable to complete the Module 2 Challenge. Although I felt like I understood what each line of code was attemopting to deliver, I was tripped up by Run Time Error 6: Overflow.

![RunTimeError6](https://user-images.githubusercontent.com/70344787/94387184-5c56cc80-0117-11eb-9c0a-2afdcd6dc3b2.png)

After selecting "Debug", I found the line that appeared to be causing the error.

![Debug](https://user-images.githubusercontent.com/70344787/94387262-a17afe80-0117-11eb-8c3b-527468278fe0.png)

I researched the error using Stack Overflow and other online resources but unfortunately was unable to resolve the problem. I believe it has something to do with the calculation dividing tickerStartingPrices by tickerEndingPrices and perhaps using an incorrect data type. 

![BeforeError](https://user-images.githubusercontent.com/70344787/94387514-61684b80-0118-11eb-9dd5-9ee08ff291ba.png)

The code appeared to run partially before being interrupted.



### Summary
---
