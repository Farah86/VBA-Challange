# **VBA CHALLANGE ANALYSIS**
## **1.Overview of Project:**
Steve gratuated from finance ,his parents wanted to ecourge him so they decided to be his first clients they are very passionate about green anergy such as hydro geothermal and wind they decided to invest all there money in DAQO new energy corp who make solar pans so he promised them to help do a research about all the corps who is involve in green anergy and there stocks.

## **2.Cmparing the results stocks between 2017 and 2018:**
**2a.** analysis of all stocks:
![This is an image](https://github.com/Farah86/VBA-Challange/blob/main/2018%20analysis.png)
![This is an image](https://github.com/Farah86/VBA-Challange/blob/main/2017analysis.png)
![This is an image](https://github.com/Farah86/VBA-Challange/blob/main/codes%20for%20analysis.png)
![This is an image](https://github.com/Farah86/VBA-Challange/blob/main/analysis%20loop%20through%20the%20array.png)


**2b.** refactored analysis of all stocks 
![This is an image](https://github.com/Farah86/VBA-Challange/blob/main/2017%20refactored%20analysis.png)
![This is an image](https://github.com/Farah86/VBA-Challange/blob/main/2018%20refactored%20analysis.png)
![This is an image](https://github.com/Farah86/VBA-Challange/blob/main/Refactored%20analysis.png)
![This is an image](https://github.com/Farah86/VBA-Challange/blob/main/refactored%20loop%20through%20the%20array.png)


## **3.Summary:**
**3a.Advantage and disadvantage of refactoring a code:**
Code Refactoring is an important exercise to remove code smell. It helps to find bugs, makes programs run faster, it's easier to understand the code, improves the design of software, etc. Code smell slows down the development and is prone to more defects. An adequate set of unit tests and a supportive environment should be there for code refactoring, so in our module it was to it make it easy to enhance and maintain in the future. It should not violate Open Close Principle and what i found it a little bit hard the debugging of tracking codes and finding which one is duplicated and fix it up.
    
   **3b Advantage and disadvantage of refactoring the codes on the  the execution time of both analysis:**
As I noticed refactoring the code was much easier than re writing the code and for the same resultes but for the second analysis we add an extra statment that effected our analysis results which was :
 
**If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then  
            tickerIndex = tickerIndex + 1             
                 End If**    
                 
 we got that if the tickers results that we had if it doesn't match then we increase the value by 1
 now for what i noticed
2018 resultes showed the same tocks **(ENPH =81.9%,RUN =84.0%)** the timimg was faster in the refactored analysis,and for 2017 results it showed diffrent stocks and timimg **(reafctored analysis TERP -7.2% in 0.1875 seconds)** and **(fisrt analysis ENPH 129.5% and RUN 5.5% in 1.00399 sec)** which mean refactored analysis was more detailed in showing which better between the stocks.
