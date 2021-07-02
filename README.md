# stock-analysis
 
## Overview of Project
The purpose of this analysis project was to provide Steve with information about stocks. The worksheet is meant to inform Steve's parents about the returns and the total daily volume of the stocks in which they are considering an investment. This project has two objectives; first, to present the information about the given stocks to Steve. A second objective was to create an Excel and a VBA workbook that would enable Steve to get the information in a timely manner. Although we previously gave Steve a spreadsheet for the same purpose, retrieving data from this spreadsheet was slow and inflexible. By refactoring the code, this project aimed to decrease the time spent in gathering information. Further, the previous spreadsheet did not allow analysis of new stocks, whereas this spreadsheet allows Steve and his parents to receive the results of new stocks as well as view the data collected.
![2017 Results](https://github.com/shireenkahlon/stock-analysis/blob/main/2017%20TDV%20and%20return.png)
![2018 Results](https://github.com/shireenkahlon/stock-analysis/blob/main/2018%20TDV%20and%20Return.png)
 
 
While writing and editing the code for the refactored code, I noticed it ran a little slower than the original code. Below are screenshots of both the refactored and original code.
![2017 Original Timer](https://github.com/shireenkahlon/stock-analysis/blob/main/2017_Original_Timer.png)
![2017 Refactored Timer](https://github.com/shireenkahlon/stock-analysis/blob/main/2017_RefactoredTimer.png)
![2018 Original Timer](https://github.com/shireenkahlon/stock-analysis/blob/main/2018_Original_Timer.png)
![2018 Refactored Timer](https://github.com/shireenkahlon/stock-analysis/blob/main/2018_Refactored_Timer.png)
 
A possible reason in the difference of timing may have been how the code was written. In the refactored code, three separate loops are created to set the value of the variables, find the total daily volume, starting and ending price, and output the results. In contrast, the original code contains one nested loop to perform these same actions. Screenshots from both codes are posted below.
 
The original Code:
![Original_Code](https://github.com/shireenkahlon/stock-analysis/blob/main/Original%20code.png)
The refactored Code:
![Refactored_Code](https://github.com/shireenkahlon/stock-analysis/blob/main/Refactored%20code.png)
   
 
 
## Results
2017 and 2018 were interesting and vastly different years for the particular stocks Steve provided us with. In 2017, the stocks tended to do much better. The returns, with the exception of one, were all in the positives and some — including DQ, ENPH, FSLR, and and SEDG — had over a 100% return. In contrast, in 2018, the stocks, with the exception of one, had all negative returns. TERP had negative returns in both 2017 and 2018, whereas RUN and FSLR has positive returns in both years. Both the 2017 and 2018 stocks are attached below.
 
## Summary
 
 ### What are the advantages and disadvantages of refactoring code?
 While refactoring the code, I noticed some key general advantages and disadvantages. First, refactored code can be created so that it can be used for any data added into the worksheet - not just for the data already in hand. Second, refactoring code can save time. This second point could also be a disadvantage - refactored code could possibly take up more time in 2 ways: First, refactored code may cause VBA or any programming language to run slower. Second, refactored code could take longer to write as the original code is being analyzed while also writing the new refactored code.
 
 
 ### How do these pros and cons apply to refactoring the original VBA script?
While writing the refactored code, I ran into the pros and cons listed above - both in the process of writing the code and afterwards.
Firstly, while the original code was quicker, the refactored code would be able to help Steve and his parents if they plan to look at other stocks outside of the 12 stocks they provided. All Steve has to do is add a sheet of data into the Excel worksheet and input the year of the worksheet into the input box. For example, if he wanted to look at 2020, he could create a similar worksheet to 2017 and 2018 and input 2020 into the input box. This is a major advantage to Steve using the refactored code.
   Secondly, I found that the platform used made a difference in the speed. I tested the code on both a Mac and a Windows machine — I found that the code ran about 4 seconds faster on a Windows as compared to a Mac. While this is inconsequential in a small dataset such as this; it would make a difference in large datasets.
   There were many disadvantages as well. While having the original code as the basis, the refactored code took about ten to twelve hours to create and debug — this took a longer time of getting the data to Steve. In addition, the refactored code ended up taking longer to provide the results in the “All Stocks Analysis” worksheet which meant a longer time for Steve to wait for the results, albeit the time difference was in seconds. Refactoring code has its advantages and disadvantages; each project would have to be looked at individually when deciding whether to refactor code. This being said, if done right and planned throughly, refactoring code can possibly be a great benefit to both the data analyst and to the client.
 

