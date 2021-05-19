Sub Worksheetloop()

Dim ws As Worksheet
For Each ws In Worksheets

Dim Yearly_change As Double
Dim StockVolumn As Double
Dim Ticker_Row As Long
Dim Ticker As String
Dim PerChange As Double
Dim StkOpn As Double

Ticker_Row = 2
StockVolumn = 0
StkOpn = 2

'Pint header's for summary
ws.Range("I1").Value = "SKT"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Volume"

'loop through all tickers stock volumn
Dim LastRow As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For x = 2 To LastRow
 If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
  'Set the Ticker Name
   Ticker = ws.Cells(x, 1).Value
  'Set Yearly_change
   Yearly_change = ws.Cells(x, 6).Value - ws.Cells(StkOpn, 3)
  'Set PerChange
    If Yearly_change = 0 Or ws.Cells(StkOpn, 3).Value = 0 Then
    
    PerChange = 0
        
    Else
    
    PerChange = (Yearly_change / ws.Cells(StkOpn, 3).Value)
    
    
    End If
  'Add 1 to the Stkopen count
   StkOpn = x + 1
  'Add to the Total volume
   StockVolumn = StockVolumn + ws.Cells(x, 7).Value
   'Print the ticker name in the summary table
    ws.Range("I" & Ticker_Row).Value = Ticker
   'Print the Yearly_change in the summary table
    ws.Range("J" & Ticker_Row).Value = Yearly_change
   'Print the PerChange in the summary table
    ws.Range("K" & Ticker_Row).Value = PerChange
   'Print the ticker stock volumn in the summary table
    ws.Range("L" & Ticker_Row).Value = StockVolumn
   'Add one to the ticker table row
    Ticker_Row = Ticker_Row + 1
    'Reset the brand total
    StockVolumn = 0
   
  Else
      ' Add to the StockVolumn Total
       StockVolumn = StockVolumn + ws.Cells(x, 7).Value
       

End If

Next x

'Pint header's for bonus summary
ws.Range("O2").Value = "Greatest%Increase"
ws.Range("O3").Value = "Greatest%Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

Dim LastRow1 As Integer
LastRow1 = ws.Cells(Rows.Count, 10).End(xlUp).Row
For y = 2 To LastRow1
ws.Cells(y, 11).NumberFormat = "0.00%"

If ws.Cells(y, 10).Value >= 0 Then
ws.Cells(y, 10).Interior.ColorIndex = 4

Else
ws.Cells(y, 10).Interior.ColorIndex = 3

End If

Next y

Dim MaxP As Double, MinP As Double, MaxV As Double
Dim STKMAP As String, STKMIP As String, STKMAV As String
Dim MaxCellP As Range, MaxCellV As Range
'Set MaxP
MaxP = Application.WorksheetFunction.Max(ws.Columns("K"))
MinP = Application.WorksheetFunction.Min(ws.Columns("K"))
MaxV = Application.WorksheetFunction.Max(ws.Columns("L"))
Set MaxCellP = ws.Range("K:K").Find(MaxP, Lookat:=xlWhole)
Set MinCellP = ws.Range("K:K").Find(MinP, Lookat:=xlWhole)
Set MaxCellV = ws.Range("L:L").Find(MaxV, Lookat:=xlWhole)

ws.Range("Q2").Value = MaxP
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").Value = MinP
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = MaxV

STKMAP = Application.WorksheetFunction.Index(ws.Columns("I"), Application.WorksheetFunction.Match(ws.Range("Q2").Value, ws.Columns("K"), 0))
STKMIP = Application.WorksheetFunction.Index(ws.Columns("I"), Application.WorksheetFunction.Match(ws.Range("Q3").Value, ws.Columns("K"), 0))
STKMAV = Application.WorksheetFunction.Index(ws.Columns("I"), Application.WorksheetFunction.Match(ws.Range("Q4").Value, ws.Columns("L"), 0))
ws.Range("P2").Value = STKMAP
ws.Range("P3").Value = STKMIP
ws.Range("P4").Value = STKMAV

Next ws

MsgBox ("Completed")

End Sub

