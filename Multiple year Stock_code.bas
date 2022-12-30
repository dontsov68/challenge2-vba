Attribute VB_Name = "Module1"
Sub Workbook()
Dim wsheet As Worksheet
Application.ScreenUpdating = False
For Each wsheet In Worksheets
wsheet.Select
Call WorkWork
Next
Application.ScreenUpdating = True
End Sub

Sub WorkWork()
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "% Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 14).Value = "Greatest % increase"
Cells(3, 14).Value = "Greatest % dicrease"
Cells(4, 14).Value = "Greatest total volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Dim n As Long
Dim i As Long
Dim k As Integer
Dim Ind As Integer
Dim Vol As Double
Dim Max As Double
Dim Min As Double
Dim Tmax As Integer
Dim Tmin As Integer
Dim Tvol As Integer
n = 2
k = 2
Max = 0
Min = 0
Do Until IsEmpty(Cells(n, 1))
Vol = Cells(n, 7).Value
Ind = 1
i = 0
Do While Cells(i + n, 1).Value = Cells(i + n + 1, 1).Value
Ind = Ind + 1
Vol = Vol + Cells(i + n + 1, 7).Value
i = i + 1
Loop
Cells(k, 9).Value = Cells(n, 1).Value
Cells(k, 10).Value = Cells(Ind + n - 1, 6).Value - Cells(n, 3).Value
If Cells(k, 10).Value > 0 Then Cells(k, 10).Interior.Color = RGB(0, 200, 0)
If Cells(k, 10).Value < 0 Then Cells(k, 10).Interior.Color = RGB(200, 0, 0)
Cells(k, 11).Value = Cells(k, 10).Value / Cells(n, 3).Value
Cells(k, 11).NumberFormat = "0.00%"
Cells(k, 12).Value = Vol
n = n + Ind
k = k + 1
Loop
p = 2
Tmax = 0
Tmin = 0
Tvol = 0
Do Until IsEmpty(Cells(p, 10))
If Cells(p, 11).Value > Max Then Max = Cells(p, 11).Value
If Cells(p, 11).Value >= Max Then Tmax = Cells(p, 11).Row
If Cells(p, 11).Value < Min Then Min = Cells(p, 11).Value
If Cells(p, 11).Value <= Min Then Tmin = Cells(p, 11).Row
If Cells(p, 12).Value > Vol Then Vol = Cells(p, 12).Value
If Cells(p, 12).Value >= Vol Then Tvol = Cells(p, 12).Row
p = p + 1
Loop
Cells(2, 16).Value = Max
Cells(3, 16).Value = Min
Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 16).NumberFormat = "0.00%"
Cells(2, 15).Value = Cells(Tmax, 9).Value
Cells(3, 15).Value = Cells(Tmin, 9).Value
Cells(4, 16).Value = Vol
Cells(4, 15).Value = Cells(Tvol, 9).Value

End Sub
