# Module 2: VBA of Wall Street Challenge

## Overview of Project
Refactor Module 2 Solution Code to run more effeciently

### Purpose
The purpose of this challenge was to update the Module 2 Solution Code using additional arrays to allow the macro to run faster.

## Results

### Stock Comparison Between 2017 and 2018
As can be seen by the screenshots of the All Stocks Charts for 2017 and 2018 below, overall, stocks in 2017 performed significantly better than those in 2018.
There were a one exceptionts:
1. Stock ticker RUN performed better in 2018 than in 2017 although it had positive returns both years.
2. Stock Ticker TERP also performed better in 2018 although it had negative returns both years.

#### All Stocks Analysis Chart (2017)                                            
![All Stocks Analysis Chart (2017)](Resources/VBA_Challenge_Chart_2017.png)      
#### All Stocks Analysis Chart (2018)
![All Stocks Analysis Chart (2018)](Resources/VBA_Challenge_Chart_2018.png)


### Time Comparisons between Original Module 2 Solution Code and the Challenge Refactored Code
Updating the code using additional arrays did decrease the amount of time required to execute the macros

Additional arrays were used to calculate the Total Volume, Starting Prices, and Ending Prices.  A tickerIndex variable was used to represent the various array indexes.  See comparison below.  Notice that `totalVolume` from the Original Module 2 Solution became `totalVolumes(tickerIndex)` with `tickerIndex` being the array index itterator.

#### Original Module 2 Solution Code Example:
```
'4.  Loop through the tickers.

        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0

'5.  Loop through rows in the data.

        Worksheets("2018").Activate
            For j = 2 To RowCount
            
'       Find the total volume for the current ticker.

                If Cells(j, 1).Value = ticker Then
                        
                        totalVolume = totalVolume + Cells(j, 8).Value
                        
                End If
                
            
'       Find the starting price for the current ticker.

                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                        
                        startingPrice = Cells(j, 6).Value
                        
                End If
```

#### Challenge Refactored Code Example:
```
'2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
            tickerVolumes(i) = 0
            
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    
    '------------------------------------Begin j loop----------------------
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                        
                        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
                        
        End If
        
```

Below are screenshots comparing macro execution times between (A) Orignial Module 2 Solution Code and (B) Challenge Refactored Code for both 2017 and 2018.

#### 2017 All Stock Analysis Macro Execution Comparison Between (A)The Original Module 2 Solution Code and (B)The Challenge Refactored Code                                            
(A)
![Original Macro Time 2017](Resources/Module_2.5.3_2017_Time_Output.png)

(B)
![Refactored Macro Time 2017](Resources/VBA_Challenge_2017.png)

#### 2018 All Stock Analysis Macro Execution Comparison Between (C)The Original Module 2 Solution Code and (D)The Challenge Refactored Code                                            
(C)

![Original Macro Time 2018](Resources/Module_2.5.3_2018_Time_Output.png)

(D)
![Refactored Macro Time 2018](Resources/VBA_Challenge_2018.png)



## Summary  
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

The written analysis contains the following structure, organization, and formatting:

There is a title, and there are multiple paragraphs (2 pt).
Each paragraph has a heading (2 pt).
There are subheadings to break up text (2 pt).
Links are working, and images are formatted and displayed where appropriate (2 pt).


## Challenges
This assignment was very challenging.
1. It was not understood what inititalizing `tickerVolumes(i)` to zero actually does as the indexes are already initialized to zero when creating the array.
   A significant amount of time was spent to gain understanding of this.
    ```
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
        For i = 0 To 11
                tickerVolumes(i) = 0
            
        Next i
     ```
2. The greatest amount of time was spent figuring out whether or not, and how, to use the tickerIndex variable to iterate array indexes:
    ### Example:
    ```
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                        
                        tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If






### Analysis of Outcomes Based on Launch Date
![Theater_Outcomes_vs_Launch_Month](Resources/Theater_Outcomes_vs_Launch_Month.png)

### Analysis of Outcomes Based on Goals
![Outcomes_vs_Goals](Resources/Outcomes_vs_Goals.png)

### Challenges and Difficulties Encountered
#### Theater_Outcomes_vs_Launch Chart
Possible Challenges Include:
* Filter
#### Outcomes_vs_Goals Chart
To avoid incorrectly typing/retyping numbers and text into formulas for the table referenced by this chart, I created a table of cells that were referenced in the formulas instead (See cells K3 through L12 for breakout of goal ranges, and cells M3 through O3 for list of outcomes below).



- What are two conclusions you can draw about the Outcomes based on Launch Date?
1. More successful campaigns were launched in May than any other month.
2. The least number of successful campaigns started in December.

- What can you conclude about the Outcomes based on Goals?
1.  A larger percentage (75.81%) of campaigns with goals below $1000 met or exceeded their goal.
2.  The lower the goal range, the more successful campaigns occur.

- What are some limitations of this dataset?
1. The dataset only includes data that Louise has gathered.  There may be additional useful data of which she is unaware.
2. The dataset only goes back to 2009
3. The data does not provide any background information about the backers which may be helpful such as address, age, gender, income brackett, etc...
4. The larger the goal range, the less data points available to make an informed decision.

- What are some other possible tables and/or graphs that we could create?
We could look at:
     * Outcomes vs. Country
     * Outcomes vs. Length of Campaign
     * Outcomes vs. Staff Pick
