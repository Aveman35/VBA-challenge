Attribute VB_Name = "Module1"
Sub GetUniqueTickersandtotals()

'Found formula below on youtube via "The Excel Cave" Channel to populate unique tickers
''Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("I:I"), Unique:=True


'''''''''''''''''''''''''''''''''''''''
Dim ws As Worksheet

For Each ws In Worksheets
  
  Dim tickerresult As String

 
  Dim totalstockvolume As Double
  totalstockvolume = 0


  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

  Dim a As Long
  '''Endrow formula suggested by Xpert Learning Assistant
  Dim endrowA As Long
  endrowA = ws.Cells(Rows.Count, "A").End(xlUp).Row
  For a = 2 To endrowA


    If ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value Then

   
      tickerresult = ws.Cells(a, 1).Value

    
      totalstockvolume = totalstockvolume + ws.Cells(a, 7).Value

    
      ws.Range("I" & Summary_Table_Row).Value = tickerresult

     
      ws.Range("L" & Summary_Table_Row).Value = totalstockvolume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      
      totalstockvolume = 0

    ' If the cell immediately following a row is the same...
    Else

      ' Add
      totalstockvolume = totalstockvolume + ws.Cells(a, 7).Value

    End If

  Next a

Next ws

''''''''''''''''''''''''''''''''''''''
 

End Sub

