Attribute VB_Name = "Close_Excel"
Sub CloseAllWorkbooks()
Dim wbs As Workbook
For Each wbs In Workbooks
wbs.Close SaveChanges:=True
Next wb
End Sub
