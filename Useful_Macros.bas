Attribute VB_Name = "Useful_Macros"
Sub Nulls()
Attribute Nulls.VB_ProcData.VB_Invoke_Func = " \n14"

    Cells.Replace What:="[NULL]", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Cells.Replace What:="NULL", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2
End Sub

Sub Clear_Formats()
Attribute Clear_Formats.VB_ProcData.VB_Invoke_Func = " \n14"
    Cells.Select
    Selection.ClearFormats
End Sub

Sub Remove_Duplicates()
    Dim Rws As Long, Col As Long, r As Range
    Set r = Range("A1")
    Rws = Cells.Find(What:="*", after:=r, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Col = Cells.Find(What:="*", after:=r, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Application.DisplayAlerts = 0
    Application.ScreenUpdating = 0

    For x = 1 To Col
        Range(Cells(1, x), Cells(Rws, x)).RemoveDuplicates Columns:=1, Header:=xlNo
    Next x

End Sub
Sub RemoveSpaces()

    Dim myRange As Range
    Dim myCell As Range
    
    Select Case MsgBox("You Can't Undo This Action. " _
    & "Save Workbook First?", _
    vbYesNoCancel, "Alert")
    Case Is = vbYesThisWorkbook.Save
    Case Is = vbCancel
    Exit Sub
    End Select
    
    Set myRange = Selection
    For Each myCell In myRange
    If Not IsEmpty(myCell) Then
    myCell = Trim(myCell)
    End If
    Next myCell

End Sub

Sub Trim_All_Cells()

    Dim rng As Range
    Set rng = Range("A1:CA10000")
    Application.ScreenUpdating = False
    With rng
        .Value = Evaluate(Replace("If(@="""","""",Trim(@))", "@", .Address))
    End With
    Application.ScreenUpdating = True

End Sub


Sub InsertMultipleRows()

    Dim i As Integer
    Dim j As Integer

    ActiveCell.EntireRow.Select
    On Error GoTo Last
    i = InputBox("Enter number of columns to insert", "Insert Columns")
    For j = 1 To i
    Selection.Insert Shift:=xlToDown, CopyOrigin:=xlFormatFromRightorAbove
    Next j
Last:         Exit Sub
    
End Sub

Sub blankCellsWithSpace()

    Dim rng As Range
    For Each rng In ActiveSheet.UsedRange
    If rng.Value = " " Then
    rng.Style = "Note"
    End If
    Next rng

End Sub

Sub No_Zeroes()

ActiveWindow.DisplayZeros = False

End Sub

Sub FindAllBlankCells()
'PURPOSE:Add Zero to All Blank Cells within Selection

Dim BlankCells As Range

'Ensure a cell range is selected
  If TypeName(Selection) <> "Range" Then Exit Sub

'Optimize Code
  Application.ScreenUpdating = False

'Store all blank cells in variable
  On Error Resume Next
  Set BlankCells = Selection.SpecialCells(xlCellTypeBlanks)
  On Error GoTo 0

'Change the value of all blank cells
  If Not BlankCells Is Nothing Then
    
    'Display Blank Cell Count
      MsgBox "There are " & BlankCells.Count & " within cell selection."
    
    'Change All Blank Cell Values
      BlankCells.Value = "Empty"
  
  End If

End Sub

Sub Concatenate_List()

    Dim tableRange As Range
    Dim x As Integer, y As Integer
    Dim lastRow As Long

    Set tableRange = Range("A1").CurrentRegion
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row

' Copy List with Quotations and Commas
    Range("A1" & lastRow).Select
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=""'"" & TRIM(RC[-1]) & ""',"""
    Range("B1").Select
    Selection.AutoFill Destination:=Range("B1" & lastRow)
    Range("B1" & lastRow).Select
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "=TEXTJOIN("" "",TRUE,R[69]C[-2]:R[104]C[-2])"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "=LEFT(R[1]C, LEN(R[1]C)-2)"
    Range("D1").Select
    
End Sub

Sub Remove_Duplicates()
    Dim Rws As Long, Col As Long, r As Range
    Set r = Range("A1")
    Rws = Cells.Find(What:="*", after:=r, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Col = Cells.Find(What:="*", after:=r, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    Application.DisplayAlerts = 0
    Application.ScreenUpdating = 0

    For x = 1 To Col
        Range(Cells(1, x), Cells(Rws, x)).RemoveDuplicates Columns:=1, Header:=xlNo
    Next x

End Sub

Sub blankWithSpace()

    Dim rng As Range
    For Each rng In ActiveSheet.UsedRange
    If rng.Value = " " Then
    rng.Style = "Note"
    End If
    Next rng

End Sub

Sub RemoveSpaces()

    Dim myRange As Range
    Dim myCell As Range
    
    Select Case MsgBox("You Can't Undo This Action. " _
    & "Save Workbook First?", _
    vbYesNoCancel, "Alert")
    Case Is = vbYesThisWorkbook.Save
    Case Is = vbCancel
    Exit Sub
    End Select
    
    Set myRange = Selection
    For Each myCell In myRange
    If Not IsEmpty(myCell) Then
    myCell = Trim(myCell)
    End If
    Next myCell

End Sub

Sub Organize_CMS_Data_Dictionary()

Dim tableRange As Range
Dim lastRow As Long

Set tableRange = Range("A1").CurrentRegion
lastRow = ActiveSheet.Cells(Activesheets.Rows.Count, "B").End(xlUp).Row


Range ("B2:F2")



End Sub

Sub Trim_All_Cells()

Dim cell As Range
For Each cell In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants)
cell = WorksheetFunction.Trim(cell)
Next cell

End Sub

