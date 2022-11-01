# VBA of Wall Street
## Overview of Project
The goal of this project was to help our friend Steve, who just graduated with a finance degree, diversify his parents' investments. Visual Basic for Applications (VBA) was used to help automate the tasks for analyzing stock data in Excel. We initially wrote a VBA script to find the Total Daily Volume and the Yearly Return for all the stock tickers in the stock data set Steve gave us to analyze. We then refactored the code to make the VBA script run faster.

## Results
The initial VBA code looped through a manually entered array of stock tickers in the data set to calculate the total Daily Volume as well as find the starting and ending price for the first and last instance of a stock ticker respectively. This was accomplished with two nested `for` loops. The outer `for` loop iterated through each of the stock tickers that were manually entered as an array while the inner `for` loop ran through each line of the workbook and totaled the volume of stock sold each day for each iteration of the outer loop. The internal `for` loop also found the starting and ending price for each iteration of the outer loop. The last step of the outer `for` loop compiled the total volume, starting price, and ending price for each ticker to then input into a table for analysis. The initial portion of this code is below.

![Initial_VBA_Script](https://user-images.githubusercontent.com/114427019/199323088-b61da931-fdb9-49e8-ada2-376cb83c96e2.png)

When we decided to refactor the code, we added a time function to the VBA script to help us determine if the refactoring was successful in reducing the amount of time the code ran in. The initial code produced the following results:

![VBA_Challenge_2018](https://user-images.githubusercontent.com/114427019/199324387-8e29b9e0-ab6f-481d-a3ef-2c946838c120.png)

When we refactored the code, we updated the outer `for` loop to relate back to the ticker index rather than having the variable remain as "i". We also added a line at the very end of the inner `for` loop to increase the "tickerIndex" loop by one each time the iteration was complete. This meant that the code for the inner `for` loop ran much more efficiently without having the code to complete both `for` loops to analyze the next ticker's data. See the updated code below.

![Refactored_VBA_Script](https://user-images.githubusercontent.com/114427019/199326554-a697f0a4-7c9f-442f-bbe7-5fd6c1d2c216.png)

The refactoring resulted in the VBA script running faster than the initial writing of the script as pictured below.

![Refactored_VBA_Challenge_2018](https://user-images.githubusercontent.com/114427019/199326678-fb2e1072-397c-4544-a4b4-260e933064e6.png)

## Summary
### Advantages and Disadvantages of Refactoring
A major advantage of refactoring existing code is making the existing code more efficient. There are a variety of ways to make code more efficient including but not limited to taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. A big disadvantage of refactoring code is that it takes much longer than the initial writing of the code. This could happen for many reasons including the developer refactoring the code not completely understanding what the original code is doing as well as testing and re-testing the code several times to ensure the refactored code works as intended.

### Original vs Refactored VBA Script
#### Advantages
An advantage of the original VBA script was it had fewer defined variables. In the original script, the were only five defined variables whereas in the refactored script there were seven defined variables. The advantage of fewer variables is that they are easier to keep track of, and it is easier to debug the code if there are errors that relate to the definition of the variables. An obvious advantage of the refactored script is it runs in significantly less time than the original script. 

#### Disadvantages
A disadvantage of both the original and the refactored VBA scripts is that the ticker arrays were manually entered rather than self-populating based on the data set. This is means if we were to use the same VBA script on a different of stock tickers the script would be looking for the wrong tickers and therefore would not output any data analysis. In addition, the formatting section of both scripts is also not dynamic. Visualizations of our analysis is extremely helpful in understanding the data at a glance. The way the code is written in both scripts the `for` loop is not dynamic and will only format twelve rows worth of cells. This means that if we were to run the code on a data set that had more than twelve tickers, some of the ticker’s outputs would not be formatted.
