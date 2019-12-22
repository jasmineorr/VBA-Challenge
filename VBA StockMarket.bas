Attribute VB_Name = "Module1"
Sub StockMarket()

Dim ws As Worksheet
Dim ticker As String
Dim vol As Variant
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim summaryoutput As Integer
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

On Error Resume Next

For Each ws In ThisWorkbook.Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    summaryoutput = 2

        For i = 2 To LastRow
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                vol = ws.Cells(i, 7).Value
                year_open = ws.Cells(i, 3).Value
                year_close = ws.Cells(i, 6).Value
                yearly_change = year_close - year_open
                percent_change = (year_close - year_open) / year_close
    
            ws.Cells(summaryoutput, 9).Value = ticker
            ws.Cells(summaryoutput, 10).Value = yearly_change
            ws.Cells(summaryoutput, 11).Value = percent_change
            ws.Cells(summaryoutput, 12).Value = vol
            summaryoutput = summaryoutput + 1
            vol = 0
        
        End If

    Next i
    
ws.Columns("K").NumberFormat = "0.00%"

For i = 2 To LastRow

If ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
Else
    ws.Cells(i, 10).Interior.ColorIndex = 4

End If

Next i

Next ws

End Sub
