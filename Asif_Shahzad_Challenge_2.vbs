Attribute VB_Name = "Module1"
Sub Stock_data()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTickerSymbol As String
    Dim decreaseTickerSymbol As String
    Dim volumeTickerSymbol As String

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' ==========================
             ' Set up headers
        ' ==========================
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
     ' ===========================
         ' Initialize variables
     ' ===========================
        
        j = 2
        ticker = ws.Cells(2, 1).Value
        openPrice = ws.Cells(2, 3).Value
        totalVolume = 0
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Loop through all rows
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = ticker Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                If i = lastRow Then
                    closePrice = ws.Cells(i, 6).Value
                End If
            Else
                closePrice = ws.Cells(i - 1, 6).Value
                
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = quarterlyChange / openPrice
                Else
                    percentChange = 0
                End If
                
          ' ==============================================================
             ' Printing results, formatting and conditional Formatting
          ' ==============================================================
                
                ws.Cells(j, 9).Value = ticker
                ws.Cells(j, 10).Value = quarterlyChange
                ws.Cells(j, 11).Value = percentChange
                ws.Cells(j, 12).Value = totalVolume
                
                
                ws.Cells(j, 10).NumberFormat = "0.00"
                ws.Cells(j, 11).NumberFormat = "0.00%"
                
                If quarterlyChange > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                    ws.Cells(j, 11).Interior.ColorIndex = 4
                ElseIf quarterlyChange < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                    ws.Cells(j, 11).Interior.ColorIndex = 3
                End If
                
          ' ===============================================
             ' Check for greatest increase/decrease/volume
          ' ===============================================
                
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    increaseTickerSymbol = ticker
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decreaseTickerSymbol = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTickerSymbol = ticker
                End If
                
        ' =================================
          ' Reset variables for new ticker
        ' =================================
                
                j = j + 1
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = ws.Cells(i, 7).Value
            End If
        Next i
        
        ' ==========================================================
           ' Printing Output, Formatting cells and applying autofit
        ' ===========================================================
        ws.Cells(2, 16).Value = increaseTickerSymbol
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(3, 16).Value = decreaseTickerSymbol
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(4, 16).Value = volumeTickerSymbol
        ws.Cells(4, 17).Value = greatestVolume
        
        ' Format summary
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
        
        ' Autofit columns
        ws.Columns("I:Q").AutoFit
    Next ws
    
End Sub

