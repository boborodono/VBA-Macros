Attribute VB_Name = "Blowup_3_Columns"
Sub Blowup_3_Columns()

    ActiveWindow.DisplayGridlines = False

    Dim c1() As Variant
    Dim c2() As Variant
    Dim c3() As Variant
    
    Dim out() As Variant
    Dim j As Long, k As Long, l As Long, n As Long

    Dim col1 As Range
    Dim col2 As Range
    Dim col3 As Range
    

    Dim out1 As Range


    Set col1 = Range("A1", Range("A1").End(xlDown))
    Set col2 = Range("B1", Range("B1").End(xlDown))
    Set col3 = Range("C1", Range("C1").End(xlDown))
   
    c1 = col1
    c2 = col2
    c3 = col3
   
    Set out1 = Range("G2", Range("K2").Offset(UBound(c1) * UBound(c2) * UBound(c3)))
    out = out1

    j = 1
    k = 1
    l = 1
    n = 1


    Do While j <= UBound(c1)
        Do While k <= UBound(c2)
            Do While l <= UBound(c3)
                out(n, 1) = c1(j, 1)
                out(n, 2) = c2(k, 1)
                out(n, 3) = c3(l, 1)
                n = n + 1
                l = l + 1
            Loop
            l = 1
            k = k + 1
        Loop
        k = 1
        j = j + 1
    Loop


    out1.Value = out
End Sub
