Attribute VB_Name = "EAS_Client_Pricing"
Sub EAS_Client_Pricing()

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
    Cells.Replace What:="NULL", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2

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
    Range("A:A,E:E,K:N,O:BC,BE:BI,BK:BV"). _
        Select
    Selection.HorizontalAlignment = xlCenter
    
    'RIGHT ALIGN
    'Range("S1:AN1,AV1:AV1,AX1:AX1").Select
    'Selection.HorizontalAlignment = xlRight
    
    '$$ - RIGHT ALIGN AND ADD DECIMAL PLACES
    'Range("N:O,R:R,T:T,V:V,X:X,Z:Z,AB:AB,AD:AD,AF:AF,AH:AH,AJ:AJ,AL:AL,AN:AQ,AX:AX"). _
    '    Select
    'With Selection
    '    .HorizontalAlignment = xlRight
    '    .NumberFormat = "#,##0.00"
    'End With
    
    
'***CALCULATIONS

'GROSS PREMIUM AMOUNT FORMULA
    Range("V2:V" & lastRow).Formula = "=SUM(RC[-2],RC[-4])"
    Range("V2").AutoFill Range("V2:V" & lastRow)
    
'PASTE VALUES
    Range("V2:V" & lastRow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("V2:V" & lastRow).Select
    Application.CutCopyMode = False
    
    
    
'AGENT COST AMOUNT FORMULA
    Range("AB2:AB" & lastRow).Formula = "=SUM(RC[-2],RC[-4],RC[-6])"
    Range("AB2").AutoFill Range("AB2:AB" & lastRow)
    
'PASTE VALUES
    Range("AB2:AB" & lastRow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AB2:AB" & lastRow).Select
    Application.CutCopyMode = False



'DEALER COST AMOUNT FORMULA
    Range("AZ2:AZ" & lastRow).Formula = "=SUM(RC[-2],RC[-4],RC[-6],RC[-8],RC[-10],RC[-12],RC[-14],RC[-16],RC[-18],RC[-20],RC[-22],RC[-24])"
    Range("AZ2").AutoFill Range("AZ2:AZ" & lastRow)
    
'PASTE VALUES
    Range("AZ2:AZ" & lastRow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AZ2:AZ" & lastRow).Select
    Application.CutCopyMode = False

'FORMATTING
   

'SET COLUMN WIDTHS
    'Range("N:O,S:AQ,AX:AX").Select
    'Selection.ColumnWidth = 13.7
    'Range("AS:AS").Select
    'Selection.ColumnWidth = 23.8
    'Range("AU:AU,AW:AW").Select
    'Selection.ColumnWidth = 13.3
    'Selection.HorizontalAlignment = xlCenter

'DATES - SET SHORT DATE FORMAT
    Range("BF:BI,BL:BS").Select
    With Selection
        .NumberFormat = "yyyy/mm/dd;@"
        .ColumnWidth = 28
        .HorizontalAlignment = xlCenter
    End With

' BLACKOUT UNUSED CELLS
    Range("CC1").Select
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
    Rows("1:1").Select
    With Selection
        .Style = "Heading 1"
        .Font.Size = 11
        .Font.ColorIndex = 46
        .Interior.ColorIndex = 9
    End With
    
'REDUCE IMAGE PROCESSING
    Application.DisplayAlerts = 0
    Application.ScreenUpdating = 0

'DON'T SHOW ZEROES
ActiveWindow.DisplayZeros = False

'SELECT A2
    Range("A2").Select


End Sub


