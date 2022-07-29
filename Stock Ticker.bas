Attribute VB_Name = "Module1"
Sub Button1_Click()

Dim WS_Count As Integer
Dim ticker, symbol As String
Dim year_begin_open, year_end_close, year_change_amount As Currency
Dim year_percentage_change As Double
Dim i, x, summary_table_row As Integer
Dim volume As Double

'Set Worksheet counter equal to the number of worksheets in the workbook
WS_Count = ActiveWorkbook.Worksheets.Count

'Turn cursor to hourglass to indicate busy
Application.Cursor = xlWait

For x = 1 To WS_Count


symbol = ActiveWorkbook.Worksheets(x).Cells(2, 1).Value
year_begin_open = ActiveWorkbook.Worksheets(x).Range("C2").Value
ear_End_Close = 0
year_change_amount = 0
i = 2
summary_table_row = 2

'Create For Loop to loop through each workbook
   For i = 2 To ActiveWorkbook.Worksheets(x).Cells(Rows.Count, "A").End(xlUp).Row
   
    
    'Write out summary table column headers
    ActiveWorkbook.Worksheets(x).Cells(1, 9).Value = "Ticker"
    ActiveWorkbook.Worksheets(x).Cells(1, 10).Value = "Yearly Change"
    ActiveWorkbook.Worksheets(x).Cells(1, 11).Value = "Percent Change"
    ActiveWorkbook.Worksheets(x).Cells(1, 12).Value = "Volume"
    
    'Change color of summary table range
    'ActiveWorkbook.Worksheets(x).Range("I1:L3000").Interior.ColorIndex = 37
   
    symbol = ActiveWorkbook.Worksheets(x).Cells(i, 1).Value
    volume = volume + ActiveWorkbook.Worksheets(x).Cells(i, 7).Value
    
      'Check to see if next row is a different stock symbol.  If so, then enter If Statement and write out summary table row
      If ActiveWorkbook.Worksheets(x).Cells(i + 1, 1) <> ActiveWorkbook.Worksheets(x).Cells(i, 1) Then
      
    'Get the Year End Stock Close Price for current stock symbol
    year_end_close = ActiveWorkbook.Worksheets(x).Cells(i, 6).Value
    
    'Write out Summary Table Row
    ActiveWorkbook.Worksheets(x).Range("I" & summary_table_row).Value = symbol
    ActiveWorkbook.Worksheets(x).Range("J" & summary_table_row).Value = year_end_close - year_begin_open
    ActiveWorkbook.Worksheets(x).Range("K" & summary_table_row).Value = (year_end_close - year_begin_open) / year_begin_open
    ActiveWorkbook.Worksheets(x).Range("L" & summary_table_row).Value = volume
    
    'increase summary row counter
    summary_table_row = summary_table_row + 1
    
    'reset variables for next stock symbol
    year_end_close = 0
    year_begin_open = ActiveWorkbook.Worksheets(x).Range("C" & i + 1).Value
    volume = 0
    
    End If

    Next i
    
   Next x

    'Return cursor to normal state
    Application.Cursor = xlDefault
    MsgBox ("All Worksheets Completed.")
    End Sub






