Attribute VB_Name = "table_slicers"
Sub tableSlicers()
'
' tableSlicers Macro
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$C$2:$O$72"), , xlYes).Name = _
        "Table2"
    Range("Table2[#All]").Select
    ActiveSheet.ListObjects("Table2").TableStyle = "TableStyleMedium16"
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), "Term"). _
        Slicers.Add ActiveSheet, , "Term 1", "Term", 192.75, 831.75, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), _
        "Term Miles").Slicers.Add ActiveSheet, , "Term Miles 1", "Term Miles", 230.25, _
        869.25, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), "Premium"). _
        Slicers.Add ActiveSheet, , "Premium", "Premium", 267.75, 906.75, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), _
        "Premium Tax").Slicers.Add ActiveSheet, , "Premium Tax", "Premium Tax", 305.25 _
        , 944.25, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), "CLIP Fee") _
        .Slicers.Add ActiveSheet, , "CLIP Fee", "CLIP Fee", 342.75, 981.75, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), "VAS Fee"). _
        Slicers.Add ActiveSheet, , "VAS Fee", "VAS Fee", 380.25, 1019.25, 144, 198.75
    ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.ListObjects("Table2"), "SG Admin") _
        .Slicers.Add ActiveSheet, , "SG Admin", "SG Admin", 417.75, 1056.75, 144, _
        198.75
    ActiveSheet.Shapes.Range(Array("SG Admin")).Select
End Sub


