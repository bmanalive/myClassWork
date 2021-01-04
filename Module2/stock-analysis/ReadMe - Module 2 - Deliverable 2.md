# Module 2: VBA of Wall Street Challenge

## Overview of Project
Refactor Module 2 Solution Code to run more effeciently

### Purpose
The purpose of this challenge was to update the Module 2 Solution Code using additional arrays to allow the macro to run faster.

## Analysis and Challenges
This assignment was very challenging.  See code below
1. It was not understood what inititalizing tickerVolumes(i) to zero actually does.  A significant amount of time was spent to gain understanding of this.
2. The greatest amount of time was spent figuring out how to use the tickerIndex variable

### Challenge Code:
Below is the Code written for the Module 2 Challenge

Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
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
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                
                        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
                        
            End If

            '3d Increase the tickerIndex.
            
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
                        
                        tickerIndex = tickerIndex + 1
                        
            End If
            
        'End If
    
    Next j
    '------------------------------------End j loop----------------------
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    'Prints the Final results
    
    For k = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
                Cells(4 + k, 1).Value = tickers(k)
                Cells(4 + k, 2).Value = tickerVolumes(k)
                Cells(4 + k, 3).Value = tickerEndingPrices(k) / tickerStartingPrices(k) - 1
        
    Next k
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

  The analysis is well described with screenshots and code (4 pt).
  
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

## Results

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
