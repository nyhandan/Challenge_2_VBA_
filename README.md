# Challenge 2: VBA Scripting for a Stock Analysis
VBA Challenge second attempt!
### By Dan Nyhan


## Overview of Project

### Purpose
The purpose of this analysis is to find out the returns of certain stocks over the last few years for Steve's parents. I have refactored the code to make an efficient analytical process. The goal of this analysis is to run a script in Microsoft Excel VBA that will be the most efficient in obtaining the percentage return of the stocks and the total daily volume of trading the stock. The dataset features 12 different stocks. My refactored code runs faster than the original code because it loops through all the data at once using "for" statements in VBA scripting.


### Results
#### Stock Performance
Here is the 2017 stock performance summary:

![2017 Stock Performance](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Stock_performance_2017.png)

Here is the 2018 stock performance summary:

![2018 Stock Performance](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Stock_Performance_2018.png)

Clearly, the stocks in our analysis performed much better in 2017 than 2018. It is easy to notice considering the conditional formatting denotes positive returns with a green cell color and negative returns with a red cell color. The "run" button in these images are also super convenient for Steve, because he can run the whole analysis for each year at the click of a button. 

#### Refractored Code Efficiency 
The refactored code ran much faster than the original code for both years. This is because I looped through all the data at once using extra "for" loops and "If..., Then..." statements. 

Here is how long it took the original code to run the 2017 data: 
![Original_code_time_performance 2017](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Original_Code_Runtime_2017.png)

This is how long it took the refactored code to run the 2017 data: 

![Refractored_code_time_performance 2017](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Refractored_Code_Runtime_2017.png)

Here is how long it took the original code to run the 2018 data: 
![Original_code_time_performance 2018](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Original_Code_runtime_2018.png)

This is how long it took the refactored code to run the 2018 data: 

![Refractored_code_time_performance 2018](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Refractored_Code_Runtime_2018.png)


