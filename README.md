# stock-analysis

## Overview of Project
Steve's parents are interested in investing in green energy because they believe it is the technology of the future, so they invested all of their money into DAQO New Energy Corp without doing much research.
Steve recently earned a finance degree and promised to research the DAQO stock for his parents, but rather than investing all of the money, he wants to diversify the funds. Along with DAQO, he wants to look at a few more stocks with the help of Excel and VBA.

### Purpose
* The analysis' purpose is to assist Steve in analyzing the return on DAQO and other stocks in order to build a diversified portfolio for his parents.
* Considering researching each stock individually is time-consuming, we'll use a VBA script to automate the process and produce more reliable conclusions.


## Results
 ![VBA Stock Analysis](/Resources/VBA_Challenge_Comparative.png)

 * In **2017**, every stock returned a profit, with an average return of 67.3 percent.

      * DQ's Total Daily Volume had the _lowest_ volume of all the stocks, but the best return of _199.4 percent_.
      * Original Script took  1.417969 seconds
      
     ![VBA Timer 2017_Old](/Resources/VBA_Challenge_Timer_2017_Old.png)
      * VBA Refactored Script took  0.2304688 seconds
      
     ![VBA Timer 2017_Refactored](/Resources/VBA_Challenge_Timer_2017_Refactored.png)


 * In **2018**, Except for ENPH and RUN, all of the stocks plummeted and also had negative returns.
     - The average return in 2018 was _-8.5 percent_.
     - DQ's price has _dropped the most_.
     - Original Script took 1.433594 seconds
     
     ![VBA Timer 2018_Old](/Resources/VBA_Challenge_Timer_2018_Old.png)
      * VBA Refactored Script took  0.2148438 seconds
      
     ![VBA Timer 2018_Refactored](/Resources/VBA_Challenge_Timer_2018_Refactored.png)
      
## Summary

### Code Refactoring

Refactoring is the process of reducing redundancy within the code and optimizing its efficiency without compromising its functionality.

#### Advantages
  - Refactoring could lead to **Simplest design**
  - Refactoring could make the code **Easier to update** and **faster to execute**
  - Code that has been refactored is **cleaner** and **shorter**.

#### Disadvantages 
  - For a layman, comprehending and determining the appropriate refactoring might be **time-consuming**.
  - If all possibilities are not considered, the code may become **buggy**.
      
### Refactoring in VBA Script

#### Advantages 
  - Reduce the loop helped is processing the data faster.
  - Simplification of calculation to the number of rows for all ticker at once and less memory usage, rather than the calculation of all the rows for all the tickers.
  - The time variance in the calculation does not vary much with the performance of the PC.
#### Disadvantages 
  - It increased the complexity of the code.
  - In this Stock Analysis, the time spent creating the code took longer, but the result was 1 second reduction in time on average.
 
