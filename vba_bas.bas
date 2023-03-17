Attribute VB_Name = "Module1"
Sub alphabetical_testing():

'variable declaration
Dim ticker As String
Dim yearlychange As Double
Dim openyear As Double
Dim closeyear As Double
Dim percentchange As Double
Dim stockvolume As Double
Dim rowcount As Long


'variable assignment
rowcount = Cells(Rows.Count, "A").End(xlUp).Row

'ticker column i
Cells(1, 9).Value = "Ticker"

For i = 2 To rowcount
    Cells(i, 9).Value = ticker
Next i

'yearly change column j
Cells(1, 10).Value = "Yearly Change"
'year change=closeyear F - openyear C
yearlychange = Cells(i, 6).Value - Cells(i, 3).Value

'if negative=red and positive=green
'red index color=3 green index color=4
If yearlychange > 0 Then
   Interior.ColorIndex = 3
ElseIf yearlychange < 0 Then
   Interior.ColorIndex = 4
End If
   
'percent change column k
Cells(1, 11).Value = "Percent Change"


'total stock volume column l
Cells(1, 12).Value = "Total Stock Volume"



End Sub
