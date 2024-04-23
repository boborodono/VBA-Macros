Attribute VB_Name = "Backup_File"
Sub FileBackUp()
ThisWorkbook.SaveCopyAs Filename:=ThisWorkbook.Path & _
"" & Format(Date, "mm-dd-yy") & " " & _
ThisWorkbook.Name
End Sub

