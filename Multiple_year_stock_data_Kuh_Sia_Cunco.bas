Attribute VB_Name = "Module1"
Sub StocksPerQuarterAnalysis():

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim startRow As Long
    Dim endRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim outputRow As Long
    Dim i As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim percentChange As Double
    Dim quarterChangeRange As Range
    Dim percentChangeRange As Range
    Dim percentColumnQRange As Range
    
    ' Loop in sheets Q1, Q2, Q3, and Q4
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Create variables for greatest values
        greatestIncrease = -1
        greatestDecrease = 1
        greatestVolume = 0
        
        ' Results start in row 2 for all sheets
        outputRow = 2

        ' Create column titles in each sheet
        If ws.Cells(1, 9).Value = "" Then
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
        End If

        ' Create variables
        i = 2
        
        Do While i <= lastRow
            ' Pull the ticker symbol
            ticker = ws.Cells(i, 1).Value
            
            ' Lookup start and end of each stock value
            startRow = i
            Do While ws.Cells(i, 1).Value = ticker And i <= lastRow
                i = i + 1
            Loop
            endRow = i - 1
            
            ' Determine open price
            openPrice = ws.Cells(startRow, 3).Value
            
            ' Determine close price
            closePrice = ws.Cells(endRow, 6).Value
            
            ' Determine total volume
            totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7))) ' Column G for volume
            
            ' Put the result for ticker, quarterly change, percent change, and total volume
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (closePrice - openPrice) / openPrice
                ws.Cells(outputRow, 11).Value = percentChange
            Else
                ws.Cells(outputRow, 11).Value = 0
            End If
            ws.Cells(outputRow, 12).Value = totalVolume
            
            ' Find the greatest percent increase
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            ' Find the greatest percent decrease
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            ' Find the greatest total volume
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
            
            ' Next output is in the next row
            outputRow = outputRow + 1
            
        Loop
        
        ' Apply conditional formatting to column "Percent Change"
        Set percentChangeRange = ws.Range("K2:K" & outputRow - 1)
        With percentChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = vbGreen
        End With
        With percentChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = vbRed
        End With
        
        ' Apply conditional formatting to column "Quarterly Change"
        Set quarterChangeRange = ws.Range("J2:J" & outputRow - 1)
        With quarterChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = vbGreen
        End With
        With quarterChangeRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = vbRed
        End With
        
        ' Format "Percent Change" column as percentage with two decimal places
        percentChangeRange.NumberFormat = "0.00%"
        
        ' Put greatest values in each sheet
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume
        
        ' Make column Q as percentage but not cell Q4
        If ws.Name <> "Q4" Then
            Set percentColumnQRange = ws.Range("Q2:Q" & outputRow - 1)
            ' Do not make cell Q4 as percentage format
            If outputRow > 4 Then
                ws.Range("Q2:Q3").NumberFormat = "0.00%"
                ws.Range("Q5:Q" & outputRow - 1).NumberFormat = "0.00%"
            Else
                ws.Range("Q2:Q" & outputRow - 1).NumberFormat = "0.00%"
            End If
        End If

    Next ws

End Sub

