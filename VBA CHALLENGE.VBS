Sub AnalyzeStocks()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim startRow As Long
    Dim summaryRow As Long
    Dim i As Long
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        ' Clear previous results if needed (assuming the summary starts from row 2)
        ws.Range("H2:K" & ws.Cells(ws.Rows.Count, "H").End(xlUp).Row).ClearContents
        
        ' Initialize variables
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        startRow = 2
        summaryRow = 2
        volume = 0
        openPrice = ws.Cells(startRow, 3).Value
        
        ' Loop through all rows in the worksheet
        For i = 2 To lastRow
            ' Check if the ticker symbol changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value
                volume = volume + ws.Cells(i, 7).Value
                
                ' Calculate quarterly change and percent change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Output summary table using arrays
                Dim summaryData As Variant
                summaryData = Array(ticker, quarterlyChange, percentChange, volume)
                ws.Cells(summaryRow, 8).Resize(1, UBound(summaryData) + 1).Value = summaryData
                
                ' Apply conditional formatting in one go
                With ws.Cells(summaryRow, 9)
                    If quarterlyChange > 0 Then
                        .Interior.Color = RGB(0, 255, 0)
                    Else
                        .Interior.Color = RGB(255, 0, 0)
                    End If
                End With
                
                With ws.Cells(summaryRow, 10)
                    If percentChange > 0 Then
                        .Interior.Color = RGB(0, 255, 0)
                    Else
                        .Interior.Color = RGB(255, 0, 0)
                    End If
                End With
                
                ' Reset variables for the next ticker
                summaryRow = summaryRow + 1
                volume = 0
                If i <> lastRow Then
                    openPrice = ws.Cells(i + 1, 3).Value
                End If
            Else
                ' Accumulate volume
                volume = volume + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
    
    ' Calculate greatest values across all sheets
    CalculateGreatestValues
End Sub

Sub CalculateGreatestValues()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim i As Long
    
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    maxIncreaseTicker = ""
    maxDecreaseTicker = ""
    maxVolumeTicker = ""
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row
        
        ' Loop through summary data
        For i = 2 To lastRow
            If ws.Cells(i, 10).Value > maxIncrease Then
                maxIncrease = ws.Cells(i, 10).Value
                maxIncreaseTicker = ws.Cells(i, 8).Value
            End If
            
            If ws.Cells(i, 10).Value < maxDecrease Then
                maxDecrease = ws.Cells(i, 10).Value
                maxDecreaseTicker = ws.Cells(i, 8).Value
            End If
            
            If ws.Cells(i, 11).Value > maxVolume Then
                maxVolume = ws.Cells(i, 11).Value
                maxVolumeTicker = ws.Cells(i, 8).Value
            End If
        Next i
    Next ws
    
    ' Output the greatest values in Columns N, O, P
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(2, 15).Value = maxIncreaseTicker
        ws.Cells(3, 15).Value = maxDecreaseTicker
        ws.Cells(4, 15).Value = maxVolumeTicker
        
        ws.Cells(2, 16).Value = maxIncrease
        ws.Cells(3, 16).Value = maxDecrease
        ws.Cells(4, 16).Value = maxVolume
    Next ws
End Sub