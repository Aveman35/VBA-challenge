Attribute VB_Name = "Module3"
Sub greatestvaluescol_0()

Dim ws As Worksheet

For Each ws In Worksheets

Dim gi As Double
Dim gd As Double
Dim gtv As Double

'''Next 3 rows down, formula suggested by Student Andrew Lane in Slack #02-ask-the-class
gi = Application.WorksheetFunction.Max(ws.Range("K2:K9999"))
gd = Application.WorksheetFunction.Min(ws.Range("K2:K9999"))
gtv = Application.WorksheetFunction.Max(ws.Range("L2:L9999"))

ws.Range("P" & 2).Value = gi
'''NumberFormat Formula suggested by Xpert Learning Assistant
ws.Range("P" & 2).NumberFormat = "0.00%"

Dim a As Long
 Dim endrowi As Long
  endrowi = ws.Cells(Rows.Count, "I").End(xlUp).Row
  
  For a = 2 To endrowi

If ws.Cells(a, 11) = ws.Range("P" & 2) Then

ws.Range("O" & 2) = ws.Cells(a, 9).Value

End If

Next a

ws.Range("P" & 3).Value = gd
ws.Range("P" & 3).NumberFormat = "0.00%"

 For a = 2 To endrowi

If ws.Cells(a, 11) = ws.Range("P" & 3) Then

ws.Range("O" & 3) = ws.Cells(a, 9).Value

End If

Next a

ws.Range("P" & 4).Value = gtv
ws.Range("P" & 4).NumberFormat = "0"

 For a = 2 To endrowi

If ws.Cells(a, 12) = ws.Range("P" & 4) Then

ws.Range("O" & 4) = ws.Cells(a, 9).Value

End If

Next a

Next ws

End Sub
