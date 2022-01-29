# Challenge 2: VBA Scripting for a Stock Analysis
VBA Challenge second attempt!
### By Dan Nyhan



## Purpose of the Stock Analysis
The purpose of this analysis is to find out the returns of 12 stocks in 2017 and 2018 for Steve's parents. I have refactored the code to make an efficient analytical process. The goal of this analysis is to run a script in Microsoft Excel VBA that will be the most efficient in obtaining the percentage return of the stocks and the total daily volume of trading the stock. The dataset features 12 different stocks. My refactored code runs faster than the original code because it loops through all the data at once using "for" statements in VBA scripting.


## Results 

### Stock Performance
Here is the 2017 stock performance summary:

![2017 Stock Performance](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Stock_performance_2017.png)

Here is the 2018 stock performance summary:

![2018 Stock Performance](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Stock_Performance_2018.png)

Clearly, the stocks in our analysis performed much better in 2017 than 2018. It is easy to notice, considering a conditional formatting function shows positive returns with a green cell color and negative returns with a red cell color. The "run" button in these images are also super convenient for Steve, because he can run the entire analysis for each year at the click of a button. 

### Refractored Code Efficiency 
The refactored code ran much faster than the original code for both years. This is because I looped through all the data at once using extra "for" loops and "If..., Then..." statements, as opposed to looping through one ticker at a time. This was the primary difference between the starting code and my refractored code. 

Here is how long it took the original code to run the 2017 data: 
![Original_code_time_performance 2017](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Original_Code_Runtime_2017.png)

This is how long it took the refactored code to run the 2017 data: 

![Refractored_code_time_performance 2017](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Refractored_Code_Runtime_2017.png)

Here is how long it took the original code to run the 2018 data: 
![Original_code_time_performance 2018](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Original_Code_runtime_2018.png)

This is how long it took the refactored code to run the 2018 data: 

![Refractored_code_time_performance 2018](https://github.com/nyhandan/Challenge_2_VBA_/blob/main/Challenge%202%20/Resources/Refractored_Code_Runtime_2018.png)



## Summary

### Advantages and Disadvantages to Refractoring Code
A primary advantage to refractoring code is you can greatly improve the efficiency of code that is already there, so you don't have to repeat aspects of designing the code of your analysis. As in this project, [the refactored code I used](https://courses.bootcampspot.com/courses/1018/pages/2-dot-3-2-loop-over-all-tickers?module_item_id=395590) cut the time it took to run the analysis by 1/6th! 

Another advantage to refractoring code is to add flexibility to the program in respect to the dataset. Refractoring code to accomodate for potential additions of data in the future will save a lot of headache. If you find more data for the analysis, is important to be able to jump back into a program and add new data with ease. Having to create entirely new commands or arrays would be very costly and difficult, as opposed to just adding another integer's worth of data into these commands and arrays. 

A disadvantage to refractoring code is it might make the code seem more disorganized and hard to understand, considering refractored code will add commands and other syntax to improve the code. It also may take stuff out of the original code, which could confuse people who were used to the original code. If the client for a project is inexperienced in coding, they may get easily lost with subtle differences of coding technique.

Another disadvantage to refractoring code is that it can be tricky and time-consuming to refractor code, and the original code may suffice. This would greatly vary by scenario, but refractored code may objectively be unnecessary for a project. 

### Refractoring Code Advantages and Disadvantages, in respect to this Analysis
For this project, it was much more convenient to add a conditional statement in a few places as opposed to literally starting the project from scratch. This was the  equivalent of editing a book, in which the publisher (myself) fine-tuned the shortcomings of the book (original code) instead of asking for an entirely reconstructed plotline. Refractoring code was critical to the speediness at which the analysis can be completed. Also, it would be easy to add additional stock tickers into the arrays and commands, while still maintaining the speediness.

The disadvantage to refractoring this project's code was that it took a long time, and I sacrificied more time refractoring the code than how much longer the original code takes to run. In different projects, this may not hold true, but it did for this project. Also, it is a lot harder to understand the refractored code to an inexperienced programmer, because there is a lot more going on compared to the original.
