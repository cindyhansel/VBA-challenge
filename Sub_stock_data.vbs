Sub stock_data()

'===========================================================
' Assignment 2
'
' This code analyzes the stock prices throughout a given year
' and creates a table to show the % increase and the total volume
' of each stock ticker for that year.
'
' It will then create a table that indicates the stocks that
' has the greatest % increase, the greatest % decrease, and the
' greatest total volume.
'
' Each worksheet contains data from one year of exchange. This code
' will loop through to analyze all given years.
'
'===========================================================

For Each ws In Worksheets

    '=======================================================
    ' define variables
    '=======================================================
    Dim new_ticker As Boolean
    Dim initial_open As Double
    Dim total_volume As Double
    Dim lastrow As Variant
    Dim table_index As Integer
    Dim table_lastrow As Integer
    Dim greatest_max As Double
    Dim greatest_min As Double
    Dim greatest_volume As Double
    
    '========================================================
    ' place headers and labels for tables
    '========================================================
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
        
    ' determine the last row of the data set
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    '==========================================================
    ' initialize and set variables to handle the first row of data
    ' which will be a new ticker value
    '==========================================================
    total_volume = 0
    table_index = 1
    new_ticker = True
    
    '===========================================================
    ' Begin looping through all of the rows
    '===========================================================
        For i = 2 To lastrow
            
            ' if this is a new ticker then grab the initial open value
            ' and begin a running total volume
            If new_ticker = True Then
            
                table_index = table_index + 1
                initial_open = ws.Cells(i, 3).Value
                total_volume = ws.Cells(i, 7).Value
                ws.Cells(table_index, 9).Value = ws.Cells(i, 1).Value
            
            ' if it is not a new ticker then accumulate the running total volume
            Else
            
                total_volume = total_volume + ws.Cells(i, 7).Value
                
            End If
             
            ' look ahead to see if the next row contains a new ticker
            ' if it does then record current values into the table
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                new_ticker = True
                ws.Cells(table_index, 10).Value = ws.Cells(i, 6).Value - initial_open
                ws.Cells(table_index, 11).Value = ws.Cells(table_index, 10).Value / initial_open
                ws.Cells(table_index, 12).Value = total_volume
                
            Else
                
                new_ticker = False
                
            End If
            
        Next i
        
        '============================================================
        'place the values for the last ticker in the data set
        '============================================================
        ws.Cells(table_index, 10).Value = ws.Cells(lastrow, 6).Value - initial_open
        ws.Cells(table_index, 11).Value = ws.Cells(table_index, 10).Value / initial_open
        ws.Cells(table_index, 12).Value = total_volume
    
        table_lastrow = table_index
        
        '============================================================
        ' format the appropriate columns with % and color
        '============================================================
        ws.Range("K2:K" & table_lastrow).NumberFormat = "0.00%"
        
        For i = 2 To table_lastrow
            If ws.Cells(i, 10).Value >= 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 4
            Else
                ws.Range("J" & i).Interior.ColorIndex = 3
            End If
        Next i
        
        
        '============================================================
        ' determine and assign values for the greatest table
        '============================================================
        greatest_max = ws.Cells(2, 11).Value
        greatest_min = ws.Cells(2, 11).Value
        greatest_volume = ws.Cells(2, 12).Value
        
        ' loop to find the greatest values
        For i = 3 To table_lastrow
            If ws.Cells(i, 11).Value > greatest_max Then
                greatest_max = ws.Cells(i, 11).Value
            End If
            If ws.Cells(i, 11).Value < greatest_min Then
                greatest_min = ws.Cells(i, 11).Value
            End If
            If ws.Cells(i, 12).Value > greatest_volume Then
                greatest_volume = ws.Cells(i, 12).Value
            End If
        Next i
        
        ' loop to find the corresponding rows and assign
        ' greatest % increase and corresponding ticker
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Cells(2, 17).Value = greatest_max
        For i = 2 To table_lastrow
            If ws.Cells(i, 11).Value = greatest_max Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i
            
        ' greatest % decrease and corresponding ticker
        ws.Cells(3, 17).Value = greatest_min
        For i = 2 To table_lastrow
            If ws.Cells(i, 11).Value = greatest_min Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i
                
        ' greatest total volume and corresponding ticker
        ws.Cells(4, 17).Value = greatest_volume
        For i = 2 To table_lastrow
            If ws.Cells(i, 12).Value = greatest_volume Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i
        
    '====================================================
    ' autofit new columns
    ws.Range("I1:Q1").EntireColumn.AutoFit
        
Next ws
            
End Sub


