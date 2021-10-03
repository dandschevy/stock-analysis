# Stock-Analysis Using Visual Basic for Applications

## Overview/Purpose
The subject of Module 2 was creating macros in Visual Basic for Applications (VBA).  Throughout this module, and in separate steps and macros, VBA code was created to analyze two years of stock data, create an input box for the data year, create output data based on the year chosen, format the output data (including font, cell colors and number formatting), and create a message box with the processing time.  The purpose of the Module 2 Challege, was to take those separate macros and create one efficient macro that was more concise in design and took less time to process.  This was to be accomplished using "For Loops, "Nested For Loops", "If/Then" statements, "Nested If/Then" statements and indexing.  


## Results and Summary

### Results 
A template was provided in order to serve as a roadmap for the combination of 4 separate macros created in Module 2 into one efficient macro.  The first step in the challenge assignment was to create a ticker index and set it equal to 0.  This was done in order to loop through the ticker volumes and starting/ending prices more efficiently as the data was analyzed.  The second step was to create an array for the ticker volumes, starting/ending prices.  The array allowed the creation of a simpler coding, and fewer loops, when processing the volumes and starting/ending prices.  The remaining code to create and format the output was essentially the same; however, adding it to the same macro saved time and allowed for the calculate button to be used.  

The end result was that the refactoring of the original macros saved time and used less much less code.  While approximately one second saved may not seem like a noticeable improvement, if there were millions of rows of data to analyze, the efficiencies created could result in a more material time savings.  Also, from an end-user perspective, this is much easier to use as several macros do not have to be run (fewer buttons to push) in order to complete analysis and output.  Below are the comparisons of the before and after for both 2017 and 2018 highlighting the time saved:

#### 2017 Before Refactoring
Note: Time listed for before refactoring (2017 only) does not include numerical formatting or font/cell changes.
![green_stocks_2017](https://user-images.githubusercontent.com/90434559/135773140-e8977ac8-3bff-42b3-9e5c-edad1fa8eaa4.png)

#### 2017 After Refactoring
![VBA_Challenge_2017](https://user-images.githubusercontent.com/90434559/135772996-d8a644af-fe38-4e7f-816f-564396003b45.png)

#### 2018 Before Refactoring
![green_stocks_2018](https://user-images.githubusercontent.com/90434559/135773079-1c9198fc-d151-476c-9c2d-d6164c4e870c.png)

#### 2018 After Refactoring
![VBA_Challenge_2018](https://user-images.githubusercontent.com/90434559/135773300-10b0312e-92d0-4543-bd3b-357bda46ec03.png)


### Summary
With refactoring, there are advantages and disadvantages.  In general, the advantages are a more elegant and streamlined code, less macros to run, and time savings.  The disadvantages lie in the lack of clarity or transparency of the code and issues in the ability to review the code for errors.  Additionally, another disadvantage could be the inability to revise the code easily if it has been too refactored.

For the Module 2 Challenge, the advantages of the revised code were time savings and end-user efficiency.  The disadvantage of the new code is that it would be difficult for an inexperienced coder to create it without first having created each separate macro.    

