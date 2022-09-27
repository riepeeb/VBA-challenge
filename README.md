# VBA-challenge

Sub VBA_Challenge():

    Dim ws As Worksheet
    
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    Dim ticker As String
    
    Dim volume As Double
    
    volume = 0
    
    Dim Summary_Table_Row As Integer
    
    Dim year_open As Double
    
    Dim year_close As Double
    
    Dim percent_change As Double
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    On Error Resume Next
    
    For Each ws In ThisWorkbook.Worksheets
    
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 11).Value = "Percent Change"
    
    Summary_Table_Row = 2
    
    For i = 2 To lastrow
    
    If year_open = 0 Then
    
        year_open = Cells(i, 3).Value
    End If
    
    If Cells(i - 1, 1) = Cells(i, 1) And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        year_close = ws.Cells(i, 6).Value
        yearly_change = year_close - year_open
        
    ticker = ws.Cells(i, 1).Value
    
    volume = volume + ws.Cells(i, 7).Value
    
    percent_change = (yearly_change / year_open)
    
    ws.Range("J" & Summary_Table_Row).Value = yearly_change
    
    ws.Range("I" & Summary_Table_Row).Value = ticker
    
    ws.Range("K" & Summary_Table_Row).Value = percent_change
    
    ws.Range("L" & Summary_Table_Row).Value = volume
    
    Summary_Table_Row = Summary_Table_Row + 1
    
    volume = 0
    
    Else
        volume = volume + ws.Cells(i, 7).Value
    
    End If
    
    Next i
    
    
' ----------------------------------------------- Formatting
    
    ws.Columns("K").NumberFormat = "0.00%"

    
    Next ws

End Sub
