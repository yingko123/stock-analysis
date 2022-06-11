# Stock Analysis Project

## Project Overview
The two goals of this project are: 
1. Analyze if there is a relationship between a stock's daily trading volume and their returns
1. Refactored code to improve performance of a subroutine 

## Results
### 1. Daily Trading Volume and Returns
After analyizing the daily trading volume and the annual returns of 2017 and 2018 for the given list of stock, I believe there is no relationship between a stock's daily trading volume and their annual returns.  Below we can see that our list of stock in general delievered higher returns during 2017 compared to 2018.  However stock performance do not seem to correlate to performance for neither years.

**2017 Stock Annual Returns vs. Total Daily Trading Volume**<br/>
<img src="Resources/2017_Vol_Rtn.PNG" width="500px">

**2018 Stock Annual Returns vs. Total Daily Trading Volume**<br/>
<img src="Resources/2018_Vol_Rtn.PNG" width="500px">

<br/>

### 2. Refactoring Subroutine
Using the indexed variable framework, I was able to breakdown a nested loop into three simple loops.  This improved the run time of the overall tasks by over 5 folds from over 1 second to roughly 200 milliseconds

<br/>

*Below is the nested loop I started with where the "j" loop was executed 12 times by the "i" loop.*

<img src="Resources/Nested.PNG" width="600px">

<br/>

*For the refactoring execercise, I broke down the nested loop by indexing my variable to elimiate the need to run through a 3000+ rows loop 12 times.*

<img src="Resources/Refactored.PNG" width="600px">

<br/>

*As a result, the run time for the 2017 and 2018 data set has improved significantly.  Run time decreased from over 1 second to roughly 200 milliseconds.*

<img src="Resources/VBA_Challenge_2017.PNG" width=300px>  <img src="Resources/VBA_Challenge_2018.PNG" width=300px>



## Summary
