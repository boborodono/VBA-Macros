Attribute VB_Name = "Highlight_Alt_Rows"
Sub HighlightAlternateRows()
Dim Myrange As Range
Dim Myrow As Range
Set Myrange = Selection
For Each Myrow In Myrange.Rows
   If Myrow.Row Mod 2 = 1 Then
      Myrow.Interior.ColorIndex = 15
   End If
Next Myrow
End Sub
