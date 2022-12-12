# Green Stocks Analysis

## Purpose

The purpose of the Green Stocks Analysis is to utilize VBA to create a report to track the volume and return on investment (ROI) for a list of eco friendly stocks. The report is to include buttons which will make executing the macros easy for anyone to use, regardless of experience level. 

## Background

A good friend Steve recently graduated with a degree in finance,  and is ready to help his 1st clients,  his parents.  Steve's parents are interested in investing in green stocks (eco friendly),  and has asked me to assist with creating a report to track volume and return on investment for a specific list of stocks. Initially Steve's parents were only interested in 1 particular stock, DAQO New Energy Corporation, but Steve feels their portfolio should be more diversified and has asked that the report include a list of stocks that he has provided.

## Results

The original VBA code was refactored to make it more efficient by implementing a new array that stores the tickers volume, starting price and ending price. In addition an index (tickerIndex) was created to match the ticker's performance to the tickerIndex. For loops are then run which cycles through all the rows and in turn creates the output on the All Stocks Analysis tab for each stock in the list. The results of the analysis was also formatted to more easily identify stocks that are performing well versus those that are not.

#### Refactored VBA:

Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create row headers
    
    Cells(3, 1).Value = "Ticker"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    
    'Assign a value to each ticker
    
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
    'Activates worksheet based on date entered in the input box.
    
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    'Goes to the last cell in a row,  and then up to the last cell with data to return the row number.
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    'ticker index starts at 0
    
     tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    'Set initial values to 0
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
    
  
    '2b) Loop over all the rows in the spreadsheet.
    'The interator starts at row 2 and goes through the row number returned by RowCount
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'if it's the 1st row for current ticker then it sets the starting price.
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
  
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         'If it's the last row for the current ticker than it sets the ending price.
         
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
         End If
  
            '3d Increase the tickerIndex.
            'if the last row is selected,  then it adds 1 (+1) to the interation to get to the next ticker in the array.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1

        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        'Activate output worksheet- "All Stocks Analysis"
        
        Worksheets("All Stocks Analysis").Activate
        
        'Name of ticker
        
        Cells(4 + i, 1).Value = tickers(i)
        
        'Sum of the volume
        
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'Return value equals ending price divided by starting price, minus 1
        
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

        
        
    Next i
    
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

##### Results of Analysis

All Stocks (2018)		
		
![image](https://user-images.githubusercontent.com/118394620/207166037-b65eb6cf-698f-4456-8be9-70e3ef213815.png)


All Stocks (2017)		
		
![image](https://user-images.githubusercontent.com/118394620/207166117-612f8e90-eeba-49f0-ac7f-6a0ae87869e8.png)


## Summary

#### Advantages of Refactoring Code

The code is easier to read,  which in turn makes finding bugs in the code much faster.  The code is also condensed so that you are not repeating any code already written.  Refactoring also helps you to better understand the code and think about how it is working on a higher level. 

#### Disadvantages of Refactoring Code

During the refractoring process it's possible to adjust the code in such a way that the program no longer runs,  or produces an inaccurate output. 

#### Results of refactoring the Green Stocks Analysis

There was a significant decrease in the amount of time it took to run the defactored code versus the original code. 

### Original Code

![image](https://user-images.githubusercontent.com/118394620/207173164-ccba128f-a690-4f91-a914-0a371bd46748.png)


![image](https://user-images.githubusercontent.com/118394620/207172257-29b13f2b-b3e7-42b5-aefa-b607f8162e1c.png)


### Refractored Code

![image](https://user-images.githubusercontent.com/118394620/207168708-1a9f303d-fb1a-4bfd-b618-b7d0e93e71c1.png)

![image](https://user-images.githubusercontent.com/118394620/207168820-ae12e467-864f-4c0b-830c-c9ce26b3467c.png)









