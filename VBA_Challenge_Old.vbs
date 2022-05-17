Sub yearValueAnalysis()
    Dim startTime As Single
    Dim endTime  As Single
yearValue = InputBox("What year would you like to run the analysis on?")
 startTime = Timer

'1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

      'Create a header row
            Cells(3, 1).Value = "Year"
            Cells(3, 2).Value = "Total Daily Volume"
            Cells(3, 3).Value = "Return"

'2) Initialize array of all tickers
    Dim tickers(11) As String
    
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
            
'3a) Initialize variables for starting price and ending price
    Dim startingPrice As Double
    Dim endingPrice As Double
    
'3b) Activate data worksheet
    Sheets(yearValue).Activate
    
'3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
            
'4) Loop through tickers
    For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0

'5) loop through rows in the data
    Sheets(yearValue).Activate
       For j = 2 To RowCount
       
                  'loop over all the rows
                  
                        If Cells(j, 1).Value = ticker Then
                
                            'increase totalVolume by the value in the current row
                            totalVolume = totalVolume + Cells(j, 8).Value
                
                        End If
                
                        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                            startingPrice = Cells(j, 6).Value
                
                        End If
                
                        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                            endingPrice = Cells(j, 6).Value
                
                        End If
    
       Next j
   'Output results
   
   Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
   Next i
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

Call formatAllStocksAnalysisTable

End Sub