Attribute VB_Name = "Form_Master"
Sub Form_Master()

    Dim tableRange As Range
    Dim lastRow As Long

    lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).Row
    Set tableRange = Range("A1").CurrentRegion
    
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
    
    'CENTER ALIGN
    Range("C:D,G:I,R:R,T:T,AC:AD,AG:AG,AL:AM"). _
        Select
    Selection.HorizontalAlignment = xlCenter
    
    'RIGHT ALIGN

    
    '$$ - RIGHT ALIGN AND ADD DECIMAL PLACES
    Range("P:Q,S:S,AK:AK"). _
        Select
    With Selection
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0.00"
    End With
    
'FORMATTING
   

'SET COLUMN WIDTHS


'DATES - SET SHORT DATE FORMAT
    Range("G:H,AD:AF").Select
    With Selection
        .NumberFormat = "yyyy/mm/dd;@"
        .ColumnWidth = 13.7
        .HorizontalAlignment = xlCenter
    End With

' BLACKOUT UNUSED CELLS
    Range("AP1").Select
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
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
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
    Rows("1:1").Select
    With Selection
        .Style = "Heading 1"
        .Font.Size = 11
        .Font.ColorIndex = 1
        .Interior.ColorIndex = 46
    End With
    
'REDUCE IMAGE PROCESSING
    Application.DisplayAlerts = 0
    Application.ScreenUpdating = 0

'DON'T SHOW ZEROES
ActiveWindow.DisplayZeros = False

'SELECT A2
    Range("A2").Select


End Sub

