Attribute VB_Name = "CMS_Rates"
Sub CMS_Rates()

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
    Range("S:S,U:U,W:W,Y:Y,AA:AA,AC:AC,AE:AE,AG:AG,AI:AI,AK:AK,AM:AM,AR:AR"). _
        Select
    Selection.HorizontalAlignment = xlCenter
    
    'RIGHT ALIGN
    Range("S1:AN1,AV1:AV1,AX1:AX1").Select
    Selection.HorizontalAlignment = xlRight
    
    '$$ - RIGHT ALIGN AND ADD DECIMAL PLACES
    Range("N:O,R:R,T:T,V:V,X:X,Z:Z,AB:AB,AD:AD,AF:AF,AH:AH,AJ:AJ,AL:AL,AN:AQ,AX:AX"). _
        Select
    With Selection
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0.00"
    End With
    
    

'**CALCULATIONS
'AGENT CODES (AGENT 1)
    Range("T1:T" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("S:S").Select
    Application.CutCopyMode = False
    Range("S1:S" & lastRow).Select
    Selection.Delete Shift:=xlToLeft
    
'AGENT CODES (AGENT 2)
    Range("V1:V" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("U:U").Select
    Application.CutCopyMode = False
    Columns("U:U").Select
    Selection.Delete Shift:=xlToLeft
    
    
'AGENT CODES (AGENT 3)
    Range("X1:X" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("W:W").Select
    Application.CutCopyMode = False
    Columns("W:W").Select
    Selection.Delete Shift:=xlToLeft
    
    
'AGENT CODES (AGENT 4)
    Range("Z1:Z" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("Y:Y").Select
    Application.CutCopyMode = False
    Columns("Y:Y").Select
    Selection.Delete Shift:=xlToLeft
    
'AGENT CODES (AGENT 5)
    Range("AB1:AB" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AA:AA").Select
    Application.CutCopyMode = False
    Columns("AA:AA").Select
    Selection.Delete Shift:=xlToLeft
    
'AGENT CODES (AGENT 6)
    Range("AD1:AD" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AC:AC").Select
    Application.CutCopyMode = False
    Columns("AC:AC").Select
    Selection.Delete Shift:=xlToLeft
    
'AGENT CODES (AGENT 7)
    Range("AF1:AF" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AE:AE").Select
    Application.CutCopyMode = False
    Columns("AE:AE").Select
    Selection.Delete Shift:=xlToLeft
    
'AGENT CODES (AGENT 8)
    Range("AH1:AH" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AG:AG").Select
    Application.CutCopyMode = False
    Columns("AG:AG").Select
    Selection.Delete Shift:=xlToLeft
    
'AGENT CODES (AGENT 9)
    Range("AJ1:AJ" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AI:AI").Select
    Application.CutCopyMode = False
    Columns("AI:AI").Select
    Selection.Delete Shift:=xlToLeft
    
'AGENT CODES (AGENT 10)
    Range("AL1:AL" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AK:AK").Select
    Application.CutCopyMode = False
    Columns("AK:AK").Select
    Selection.Delete Shift:=xlToLeft
    
'AGENT CODES (AGENT 11)
    Range("AN1:AN" & lastRow).Select
    Selection.Insert Shift:=xlToRight
    Selection.SpecialCells(xlCellTypeBlanks).Select
    With Selection
        .FormulaR1C1 = "=TEXT(RC[-1],""000000"")"
    End With
    
    'PASTE VALUES
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AM:AM").Select
    Application.CutCopyMode = False
    Columns("AM:AM").Select
    Selection.Delete Shift:=xlToLeft
    
    
'ERASE "000000"
    Cells.Replace What:="000000", Replacement:="", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2


'TOTAL AMOUNT FORMULA
    Range("AO2:AO" & lastRow).Formula = "=SUM(RC[-1],RC[-3],RC[-5],RC[-7],RC[-9],RC[-11],RC[-13],RC[-15],RC[-17],RC[-19],RC[-21],RC[-23])"
    Range("AO2").AutoFill Range("AO2:AO" & lastRow)
    
'PASTE VALUES
    Range("AO2:AO" & lastRow).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("AO2:AO" & lastRow).Select
    Application.CutCopyMode = False

'FORMATTING
   

'SET COLUMN WIDTHS
    Range("N:O,S:AQ,AX:AX").Select
    Selection.ColumnWidth = 13.7
    Range("AS:AS").Select
    Selection.ColumnWidth = 23.8
    Range("AU:AU,AW:AW").Select
    Selection.ColumnWidth = 13.3
    Selection.HorizontalAlignment = xlCenter

'DATES - SET SHORT DATE FORMAT
    Range("E:H,AV:AV").Select
    With Selection
        .NumberFormat = "yyyy/mm/dd;@"
        .ColumnWidth = 17.3
        .HorizontalAlignment = xlCenter
    End With

' BLACKOUT UNUSED CELLS
    Range("AY1").Select
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
        .Font.ColorIndex = 15
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
