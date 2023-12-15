VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickers()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

'declaring variables
Dim ticker_name As String
Dim ticker_volume As Double
Dim lastrow As Long
Dim percent_change As Double
Dim yearly_change As Double
Dim summary_table_row As Long
Dim year_open As Double
Dim year_close As Double
Dim max_percent As Double
Dim max_ticker As String
Dim min_percent As Double
Dim min_ticker As String
Dim max_volume_ticker As String
Dim max_volume As Double
Dim columnrange As Range


'set variables
summary_table_row = 2
ticker_volume = 0
yearly_change = 0
percent_change = 0
max_percent = 0
min_percent = 0
year_open = 0
year_close = 0
max_volume = 0
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'currently this is calculating last row and second to last row in set as yearly change, instead of first number of set
'***^solved

    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ticker_name = ws.Cells(i, 1).Value
    
        ticker_volume = ticker_volume + ws.Cells(i, 7).Value
    
        year_close = ws.Cells(i, 6).Value
    
            If year_open <> 0 Then
                yearly_change = year_close - year_open
                percent_change = (year_close / year_open) - 1
                
                    If percent_change > max_percent Then
                    max_percent = percent_change
                    max_ticker = ticker_name
                    End If
                    
                    If percent_change < min_percent Then
                    min_percent = percent_change
                    min_ticker = ticker_name
                    End If
                    
                    If ticker_volume > max_volume Then
                    max_volume = ticker_volume
                    max_volume_ticker = ticker_name
                    End If
            Else
        yearly_change = 0
        percent_change = 0
    End If
    
    ws.Range("I" & summary_table_row).Value = ticker_name
    
    ws.Range("L" & summary_table_row).Value = ticker_volume

    ws.Range("j" & summary_table_row).Value = yearly_change
    
    ws.Range("k" & summary_table_row).Value = percent_change
    
    summary_table_row = summary_table_row + 1
    
    ticker_volume = 0
    
    year_close = 0
    year_open = 0
    
    Else
    
    ticker_volume = ticker_volume + ws.Cells(i, 7).Value
    
    'store year_open only once per
    If year_open = 0 Then
        year_open = ws.Cells(i, 6).Value
    End If
    
    End If
    
Next i



ws.Range("p1").Value = "Ticker"
ws.Range("q1").Value = "Value"
ws.Range("o2").Value = "Greatest % Increase"
ws.Range("o3").Value = "Greatest % Decrease"
ws.Range("o4").Value = "Greatest Total Volume"
ws.Range("p2").Value = max_ticker
ws.Range("q2").Value = max_percent
ws.Range("p3").Value = min_ticker
ws.Range("q3").Value = min_percent
ws.Range("p4").Value = max_volume_ticker
ws.Range("q4").Value = max_volume
ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total Volume"

Set columnrange = ws.Columns("k")
columnrange.NumberFormat = "0.00%"

ws.Range("q2").NumberFormat = "0.00%"
ws.Range("q3").NumberFormat = "0.00%"
ws.Columns("j").NumberFormat = "$#,##0.00"

Next ws

End Sub

