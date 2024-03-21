# VBA-challenge

Hello! I'm turning in this missing homework from earlier, however my mac is very old and excel keeps crashing but I'm sure I have the code right. You can find my code in the file alphabetical_testing. Just in case it crashes when you try to open it or see it this is the code I inputed which worked when I used a friend's excel earlier on :

Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVolume = 0
        
        ' Output headers for the analysis
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Loop through each row of data
        For i = 2 To lastRow
            ' Get the ticker symbol
            ticker = ws.Cells(i, 1).Value
            
            ' Get the opening price
            openingPrice = ws.Cells(i, 3).Value
            
            ' Get the closing price
            closingPrice = ws.Cells(i, 6).Value
            
            ' Calculate the yearly change
            yearlyChange = closingPrice - openingPrice
            
            ' Calculate the percent change
            If openingPrice <> 0 Then
                percentChange = (yearlyChange / openingPrice) * 100
            Else
                percentChange = 0
            End If
            
            ' Calculate the total volume
            totalVolume = ws.Cells(i, 7).Value
            
            ' Output the information
            ws.Cells(i, 9).Value = ticker
            ws.Cells(i, 10).Value = yearlyChange
            ws.Cells(i, 11).Value = percentChange
            ws.Cells(i, 12).Value = totalVolume
            
            ' Apply conditional formatting
            If yearlyChange > 0 Then
                ws.Cells(i, 10).Interior.Color = vbGreen
            ElseIf yearlyChange < 0 Then
                ws.Cells(i, 10).Interior.Color = vbRed
            End If
            
            ' Check for maximum percent increase/decrease and total volume
            If percentChange > maxPercentIncrease Then
                maxPercentIncrease = percentChange
            ElseIf percentChange < maxPercentDecrease Then
                maxPercentDecrease = percentChange
            End If
            
            If totalVolume > maxTotalVolume Then
                maxTotalVolume = totalVolume
            End If
        Next i
        
        ' Output the stocks with greatest % increase, % decrease, and total volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("K:K")), ws.Range("K:K"), 0)).Value
        ws.Cells(3, 16).Value = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("K:K")), ws.Range("K:K"), 0)).Value
        ws.Cells(4, 16).Value = Application.WorksheetFunction.Index(ws.Range("I:I"), Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("L:L")), ws.Range("L:L"), 0)).Value
        ws.Cells(2, 17).Value = maxPercentIncrease
        ws.Cells(3, 17).Value = maxPercentDecrease
        ws.Cells(4, 17).Value = maxTotalVolume
    Next ws
End Sub
