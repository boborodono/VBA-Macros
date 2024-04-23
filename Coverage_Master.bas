Attribute VB_Name = "Coverage_Master"
Sub Coverage_Master()

    Dim tableRange As Range
    Dim lastRow As Long
    
    Set tableRange = Range("A1").CurrentRegion
    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    
'CLEAR ALL FORMATS
    Cells.Select
    Selection.ClearFormats

'ERASE "[NULL]"
    Cells.Replace What:="[NULL]", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

'FREEZE TOP TOP ROW
    Range("A2").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    
    ActiveWindow.FreezePanes = True
    
'AUTOFIT COLUMN WIDTHS
    tableRange.Select
    Selection.Columns.AutoFit
    
'FORMATTING
   
    'CENTER ALIGN
    Range("C:C,E:E,G:L"). _
        Select
    Selection.HorizontalAlignment = xlCenter

'DATES - SET SHORT DATE FORMAT
    Range("G:L").Select
    With Selection
        .NumberFormat = "yyyy/mm/dd;@"
        .ColumnWidth = 12
    End With

'$$ - RIGHT ALIGN AND ADD DECIMAL PLACES
    Range("D:D,F:F"). _
        Select
    With Selection
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0.00"
    End With

' BLACKOUT UNUSED CELLS
    Range("P1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
       .PatternTintAndShade = 0
    End With
    
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
    End With


'MAKE TEXT COLOR SOFTER
    tableRange.Select
    Selection.Font.ColorIndex = 56
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD(ROW(),2)=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
    
'FORMAT HEADER
    Range("A1:O1").Select
    With Selection
        .Style = "Heading 1"
        .Font.Size = 11
        .Font.ColorIndex = 1
        .Interior.ColorIndex = 50
    End With
    
    
'REDUCE IMAGE PROCESSING
    Application.DisplayAlerts = 0
    Application.ScreenUpdating = 0

'DON'T SHOW ZEROES
ActiveWindow.DisplayZeros = False

'SELECT A2
    Range("A2").Select


End Sub
