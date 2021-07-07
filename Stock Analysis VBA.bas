Attribute VB_Name = "Module1"
Sub Macrocheck()

    Dim testMessage As String

    testMessage = "Hello World"
    
    MsgBox (testMessage)

End Sub

Sub Ones()

    Worksheets("All Stocks Analysis").Activate
    
Dim row As Integer
Dim col As Integer
    
    For i = 1 To 10
    
        For j = 1 To 10

        Cells(i, j).Value = i + j
        
        Next j

    Next i
    
Range("a1:j10").Clear

End Sub



Sub Checkerboard()


    'looping columns
    dataColStart = 1
    dataColEnd = 8
    For i = dataColStart To dataColEnd
    
        
        'looping rows
        dataRowStart = 1
        dataRowEnd = 8
        For j = dataRowStart To dataRowEnd
        
            
            If (i + j) Mod 2 = 0 Then
            Cells(i, j).Interior.Color = vbBlack
            
            End If
        
        Next j

    Next i


End Sub



Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQ0 (Ticker: DQ)"
    
    
    Cells(3, 1).Value = "Year"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate
    
    rowStart = 2
    RowEnd = 3013
    totalVolume = 0
    
    For i = rowStart To RowEnd
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
    
        End If
        
        'Find the starting price for the current ticker.
    
        If (Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ") Then
    
            startingprice = Cells(i, 3).Value
    
        End If
    
        'Find the ending price for the current ticker.
        If (Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ") Then
    
            endingprice = Cells(i, 6).Value
    
        End If
    
    Next i
    
    'output results
    Worksheets("DQ Analysis").Activate
        Cells(4, 1).Value = "2018"
        Cells(4, 2).Value = totalVolume
        Cells(4, 3).Value = endingprice / startingprice - 1
    
    
'Formatting Table
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Size = 13
    Range("A3:C3").Interior.ColorIndex = 37
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
End Sub


Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime As Single
    
yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer

'Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks(" + yearValue + ")"
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'Initialize an array of all tickers.
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

'Prepare for the analysis of tickers.
'Initialize variables for the starting price and ending price.

    Dim startingprice As Single
    Dim endingprice As Single
    
' Activate the data worksheet.
    Worksheets(yearValue).Activate

'Find the number of rows to loop over.

    RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
'Loop through the tickers.
    
    For i = 0 To 11
    
    ticker = tickers(i)
    totalVolume = 0

    'Loop through rows in the data.

        Worksheets(yearValue).Activate
        For j = 2 To RowCount
    
        'Find the total volume for the current ticker.
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
    
        End If
    
        'Find the starting price for the current ticker.
    
        If (Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker) Then
    
            startingprice = Cells(j, 3).Value
    
        End If
    
        'Find the ending price for the current ticker.
        If (Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker) Then
    
            endingprice = Cells(j, 6).Value
    
        End If
        
        Next j

    'Output the data for the current ticker.

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingprice / startingprice - 1
        
    Next i
    


Worksheets("All Stocks Analysis").Activate

'Formatting Table
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Size = 13
    Range("A3:C3").Interior.ColorIndex = 37
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    'we were asked to also format the price col to use a dollar sign and 2 sig dig but we don't have a price col
    
'Conditional formatting
    For i = 0 To 11
    
        If Cells(4 + i, 3) > 0 Then Cells(4 + i, 3).Interior.ColorIndex = 4
        
        If Cells(4 + i, 3) < 0 Then Cells(4 + i, 3).Interior.ColorIndex = 3
    
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    
    'another coding option would be:
    'If Cells(4, 3) > 0 Then
    'Color the cell green
    'Cells(4, 3).Interior.Color = vbGreen
    'ElseIf Cells(4, 3) < 0 Then

    'Color the cell red
    'Cells(4, 3).Interior.Color = vbRed

    'Else
    'Clear the cell color
    'Cells(4, 3).Interior.Color = xlNone

    'End If

End Sub

Sub ClearWorksheet()

    Cells.Clear
    
End Sub

Sub yearValueAnalysis()

    yearValue = InputBox("What year would you like to run the analysis on?")

End Sub



