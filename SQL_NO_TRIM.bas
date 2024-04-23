Attribute VB_Name = "SQL_NO_TRIM"
Sub SQL_NO_TRIM()

    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    'CREATES A HELPER SHEET "DELETE" TO REMOVE DUPLICATES WHICH WILL BE REMOVED LATER
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "DELETE"
    Sheets("DELETE").Tab.ColorIndex = 30
    ActiveSheet.Range("A1").PasteSpecial xlPasteValues
    
    'REMOVES DUPLICATES
    ActiveSheet.Range("$A$1:$A$100000").RemoveDuplicates Columns:=1, Header:=xlYes
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    'CREATES SHEET TO PASTE FINAL
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "SQL NO TRIM"
    Sheets("SQL NO TRIM").Tab.ColorIndex = 49
    ActiveSheet.Range("A1").PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight
    Selection.End(xlDown).Select
    
    Selection.SpecialCells(xlCellTypeBlanks).Select
    
    ActiveWindow.DisplayGridlines = False
    
    With Selection
        .FormulaR1C1 = "=""'""&(RC[1])&""', """
    End With
    
    
    Columns("A:B").EntireColumn.AutoFit
    
    Selection.End(xlDown).Select
    
    With Selection
        .FormulaR1C1 = "=""'""&(RC[1])&""'"""
    End With
    
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "=TEXTJOIN("""",TRUE,RC[-2]:R[1000]C[-2])"
    
    Columns("A:A").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD(ROW(),2)=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    Range("B1").Select
    Columns("B:B").EntireColumn.AutoFit
    Selection.Font.Bold = True
    
    Range("B1").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    'DELETES THE FIRST SHEET CREATED CALLED "DELETE"
    Application.DisplayAlerts = False
    Sheets("DELETE").Delete
    Application.DisplayAlerts = True
    
    Range("B1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    ActiveWindow.SmallScroll ToRight:=-1
    
End Sub
