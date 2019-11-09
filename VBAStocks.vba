Private Sub Workbook_Open()

  MsgBox "Welcome to WALL STREET TABS"

  Dim WS_Count As Integer
  Dim I As Integer

  ' Set WS_Count equal to the number of worksheets in the active
  ' workbook.
  WS_Count = ActiveWorkbook.Worksheets.Count

  ' Begin the loop.
  For I = 1 To WS_Count

      Wallstreet (ActiveWorkbook.Worksheets(I).Name)

  Next I

End Sub


Public Sub Wallstreet(stocksheet As String)

MsgBox ("Processing Sheet:  " & stocksheet)
Sheets(stocksheet).Activate

Dim ticker As String
Dim begin_year_price As Double
Dim next_row_open As Double
Dim next_row_close As Double
Dim end_year_price As Double
Dim next_ticker As String
Dim yearly_change As Double
Dim percent_change As Double
Dim stock_volume As Double
Dim total_stock_volume As Double
Dim summary_row As Long
Dim sheet_row_count As Long
Dim I As Long


Application.ActiveSheet.UsedRange
sheet_row_count = Sheets(stocksheet).UsedRange.Rows.Count


yearly_change = 0
percent_change = 0
stock_volume = 0
total_stock_volume = 0
summary_row = 2

' Initialize the Summary Column Headers
Sheets(stocksheet).Cells(1, 9).Value = "Ticker"
Sheets(stocksheet).Cells(1, 10).Value = "Yearly Change"
Sheets(stocksheet).Cells(1, 11).Value = "Percent Change"
Sheets(stocksheet).Cells(1, 12).Value = "Total Stock Volume"

'read the first row into variables before you enter the main loop
'now, I realize that I could have done this in a single loop by
' doing a + 1 on the index reference, but whatever
ticker = Sheets(stocksheet).Cells(2, 1).Value
begin_year_price = Sheets(stocksheet).Cells(2, 3).Value
next_row_close = Sheets(stocksheet).Cells(2, 6).Value
stock_volume = Sheets(stocksheet).Cells(2, 7).Value
total_stock_volume = total_stock_volume + stock_volume
' Initialize next_ticker as empty
next_ticker = ""

For I = 3 To sheet_row_count

    next_ticker = Sheets(stocksheet).Cells(I, 1).Value
    stock_volume = Sheets(stocksheet).Cells(I, 7).Value

    If next_ticker <> ticker Then

        'perform summary calcs on stock symbol
        Sheets(stocksheet).Cells(summary_row, 9).Value = ticker
        yearly_change = next_row_close - begin_year_price
        'Avoid zero divide on percentage calculation
        If begin_year_price = 0 Then
           percent_change = 0
        Else
           percent_change = (yearly_change / begin_year_price)
        End If
        Sheets(stocksheet).Cells(summary_row, 10).Value = yearly_change
        Sheets(stocksheet).Cells(summary_row, 11).Value = percent_change
        Sheets(stocksheet).Cells(summary_row, 11).NumberFormat = "0.00%"
        Sheets(stocksheet).Cells(summary_row, 12).Value = total_stock_volume
        
        'reset variables for new stock symbol
        total_stock_volume = 0
        'hit a new symbol so save off the begin year opening price
        begin_year_price = Sheets(stocksheet).Cells(I, 3).Value
        ticker = next_ticker
        summary_row = summary_row + 1

    End If
    
    total_stock_volume = (total_stock_volume + stock_volume)
    next_row_open = Sheets(stocksheet).Cells(I, 3).Value
    next_row_close = Sheets(stocksheet).Cells(I, 6).Value

Next I

' Process Last Symbol Outside of Main Loop
yearly_change = next_row_close - begin_year_price
Sheets(stocksheet).Cells(summary_row, 9).Value = ticker
Sheets(stocksheet).Cells(summary_row, 10).Value = yearly_change
Sheets(stocksheet).Cells(summary_row, 12).Value = total_stock_volume

' General Formatting

' Conditional formatting for Yearly Change Column
Dim rg As Range
Dim positive_growth As FormatCondition
Dim negative_growth As FormatCondition

Set rg = Sheets(stocksheet).Range("J2", Range("J2").End(xlDown))
rg.NumberFormat = "0.000000000"
rg.FormatConditions.Delete
Set positive_growth = rg.FormatConditions.Add(xlCellValue, xlGreaterEqual, 0)
Set negative_growth = rg.FormatConditions.Add(xlCellValue, xlLess, 0)

'define the format applied for each conditional format
With positive_growth
.Interior.Color = vbGreen
End With

With negative_growth
.Interior.Color = vbRed
End With

End Sub

