# Stock-Analysis
Performing analysis on Stock data to determine the returns

## Overview of Project

### Purpose and Background

An Excel workbook was created for Steven, with a click of a button, to analyze an entire dataset. However, he wants to expand the dataset to include the entire stock market over the last few years (2017 & 2018). The original Microsft Excel VBA code works well for a dozen stocks, but may take a long time to execute for thousands of stocks. In the new Excel workbook, we refactored the code to loop through all the data in order to collect the same information. Refactoring a code can make the code more efficient, take fewer steps, use less memory, or improve logic of code to make it easier for future users to read. 

The orignal data presented include two charts (2017 & 2018) with stock information on 12 different stocks. The stock information contains: 1) Ticker (e.g. DQ), 2) Date, 3) Opening, 4) High Price, 5) Low Price, 6) Closing, 7) Adjusted Closing, and 8) Volume of the stock. The goal is to retrieve the Ticker, the Total Daily Volume, and the Return on each stock.

## Results

### Analysis of stock performance between 2017 and 2018

![VBA_Challenge_2017](https://user-images.githubusercontent.com/107021231/176988781-c9bc7b24-1749-43e5-84d1-a17101f494c5.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/107021231/176988787-f2054168-5090-4715-9c6b-dfcb0961fc4a.png)






### Analysis of original script and refactored sript 

<sub> The refactored VBA Code: </sub>

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
   For i = 0 To 11
       tickerIndex = tickers(i)
       
       
    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single, tickerEndingPrices As Single
       
       
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'If the next row’s ticker doesn’t match, increase the tickerIndex.
       Worksheets(yearValue).Activate
       tickerVolumes = 0
       
       ''2b) Loop over all the rows in the spreadsheet.
       For j = 2 To RowCount
              
           ' If the next row’s ticker doesn’t match, increase the tickerIndex.
           If Cells(j, 1).Value = tickerIndex Then
           
              '3a) Increase volume for current ticker
              tickerVolumes = tickerVolumes + Cells(j, 8).Value
        
           End If
           
           
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerStartingPrices = Cells(j, 6).Value
               
          'End If
           End If

        '3c) check if the current row is the last row with the selected ticker
        'If  Then
           If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerEndingPrices = Cells(j, 6).Value
               
          'End If
           End If
           
       Next j
       
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

           Worksheets("All Stocks Analysis").Activate
           
           Cells(4 + i, 1).Value = tickerIndex
           Cells(4 + i, 2).Value = tickerVolumes
           Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
    
            'With Range("C4:C15")
                        '.NumberFormat = "0.0%"
                        '.Value = .Value
            'End With
            

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

<sub> End Sub </sub>


  





## Summary

### 1) What are the advantages or disadvantages of refactoring code?

Based on

### 2) How do these pros and cons apply to refactoring the original VBA sript? 

Based on 
