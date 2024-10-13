Attribute VB_Name = "Module2"
Sub quartelychangeandpercentchange()

Dim ws As Worksheet

For Each ws In Worksheets

Dim tickeri As Long
tickeri = 2

Dim QC As Double
QC = 0

Dim PC As Double
PC = 0

Dim openvalue As Double
openvalue = ws.Cells(2, 3).Value

Dim Summary_Table_RowQC As Long
  Summary_Table_RowQC = 2
  
Dim Summary_Table_RowPC As Long
  Summary_Table_RowPC = 2
  

Dim a As Long
 Dim endrowA As Long
  endrowA = ws.Cells(Rows.Count, "A").End(xlDown).Row
  For a = 2 To endrowA

 If ws.Cells(a, 1).Value <> ws.Range("I" & tickeri).Value Then
 
 QC = ws.Cells(a - 1, 6) - openvalue
 PC = (QC / openvalue)
 
  ws.Range("J" & Summary_Table_RowQC).Value = QC
  

      ' Add one to the summary table row
      Summary_Table_RowQC = Summary_Table_RowQC + 1
      
      ws.Range("K" & Summary_Table_RowPC).Value = PC
      '''NumberFormat Formula suggested by Xpert Learning Assistant
        ws.Range("K" & Summary_Table_RowPC).NumberFormat = "0.00%"
 
        Summary_Table_RowPC = Summary_Table_RowPC + 1
 
         PC = 0

 'reset openvalue to new value for next ticker
 openvalue = ws.Cells(a, 3)
 
 'move to the next ticker in the list
 tickeri = tickeri + 1

 
 
 End If

Next a

 ''''Conditional format for green and red backgrounds in PC

Dim b As Long
 Dim endrowj As Long
 
 '''Adjusted endrow formula suggested by Xpert Learning Assistant
  endrowj = ws.Cells(Rows.Count, "J").End(xlUp).Row
  
  For b = 2 To endrowj

If ws.Cells(b, 10) > 0 Then

ws.Cells(b, 10).Interior.ColorIndex = 4

    Else
    
    If ws.Cells(b, 10) < 0 Then
    ws.Cells(b, 10).Interior.ColorIndex = 3

    End If
    
    End If

Next b

Next ws


End Sub
