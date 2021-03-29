# Refactoring Module 2 VBA Code and Measuring Performance
---
---
## Overview of Project
---
Steve, a recent college graduate with a finance degree, had asked for assistance with analyzing green energy stocks on behalf of his parents who are interested in investing. Steve was impressed by the original workbook that was prepared for him and would now like to expand his dataset to include additional stocks.
---
### Purpose
---
Using Visual Basic for Applications (VBA) available in Excel, this project attempts to refactor the original code to once again loop through the original data and determine if the VBA script will run faster after refactoring the code. To prove the success of this endeavor, the following outputs will be provided:
- Pop-up message showing elapsed time to run analysis for 2017
- Pop-up message showing elapsed time to run analysis for 2018
---
### Results
---
https://github.com/Edwardo-H/stock-analysis/blob/master/Resources/2017%20Format.PNG

Unfortunately, I was unable to complete the Module 2 Challenge. Although I felt like I understood what each line of code was attemopting to deliver, I was tripped up by Run Time Error 6: Overflow.

![RunTimeError6](https://user-images.githubusercontent.com/70344787/94387184-5c56cc80-0117-11eb-9c0a-2afdcd6dc3b2.png)

After selecting "Debug", I found the line that appeared to be causing the error.

![Debug](https://user-images.githubusercontent.com/70344787/94387262-a17afe80-0117-11eb-8c3b-527468278fe0.png)

I researched the error using Stack Overflow and other online resources but unfortunately was unable to resolve the problem. I believe it has something to do with the calculation dividing tickerStartingPrices by tickerEndingPrices and perhaps using an incorrect data type. 

![BeforeError](https://user-images.githubusercontent.com/70344787/94387514-61684b80-0118-11eb-9dd5-9ee08ff291ba.png)

The code appeared to run partially before being interrupted.

![PartialResults](https://user-images.githubusercontent.com/70344787/94387695-eb181900-0118-11eb-81de-06b15d45e52f.png)

This image leads me to believe that the error is triggered by the code not being able to properly calculate the change in price for ticker symbol CSIQ.

### Summary
---
While I'm certain if I had more time I'd have been able to resolve the issue given the amount of resources that are available, I have to admit I'm stumped as to why my code is not working as expected. However, the module challenge still proved to be a valuable learning experience.

#### Advantages
Some advantages of refactoring code include improving the speed of the output and by streamlining the code to get rid of any unnecessary elements that may cause confusion or be misleading. Also, it could increase a programmer's productivity if they can make adjustments to a similar existing code rather than start a new one from scratch.

#### Disadvantages
On the other hand, refactoring code can be diasdavantageous if coders become to reliant on existing code thus effecting their creativity and critical thinking that might allow them to see "a new way of doing things". Also, a programmer might rely on an existing code that is ineffectice or sloppy resulting in wasting valuable time.

#### Conclusion
For this VBA script in particular, I can attest to the fact that refactoring led me to a frustrating experience. I had an understanding of the original code and it was up and running just fine. In my attempt to improve on the code, I wound creating an error that I wasn't able to debug.
