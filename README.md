# Stock-Analysis
2.2.4
## Project Overview
This Project is to help which is the best investment for Steve's Parents. Investment they are interested in was DQ Analysis. Steve want to make sure that his parents are making the correct decision in stock market.The previous code was executed on Green stocks. For this particualr project we are refactoring the  code to analyse only 2018 and 2017 stocks. 
### Result

---
In 2018,ENPH and RUN two stocks had positive yearly Return as well as large Total Daily Volume. Both of them was outperformance than others green stocks.

---
<img width="250" alt="Screen Shot 2022-03-16 at 1 19 11 AM" src="https://user-images.githubusercontent.com/98849217/159086231-f19744a3-19e1-429a-8a11-8da2bc291bfd.png">

---

In 2017 all of stocks had positive Return except TERP (-7.2%). "DQ" made best yearly return with 199.4% but with lowest total Daily Volume (35,796,200) in 2017.

---

<img width="248" alt="Screen Shot 2022-03-16 at 1 23 04 AM" src="https://user-images.githubusercontent.com/98849217/159086167-b93d93d6-ccf9-4eb0-86a1-4756d72e4083.png"> 

---

1 Main Loop for all data and assigned tickerIndex for 12 stocks respectively.
2 Nested loop in the main loop, executed stocks original data and retrieve ticker name, startingPrices and endingPrices, and save information to each related tickerIndex.
3 nested loop in 2nd loop, in order to get volume information for each Index.
4 new loop for putting all saved output information into an analysis sheet.
Refactoring the code is very efficent in terms of time.

---
Script
'Attribute VB_Name = "Module2"
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
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                
                'If it is, then assign the current starting price to the tickerStartingPrices variable.
                
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
            
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                'If it is, then assign the current closing price to the tickerEndingPrices variable.
                
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker
        
                    tickerIndex = tickerIndex + 1
                
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
            Worksheets("All Stocks Analysis").Activate
                Cells(4 + i, 1).Value = tickers(i)
                Cells(4 + i, 2).Value = tickerVolumes(i)
                Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i) - 1)
        Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0"
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
