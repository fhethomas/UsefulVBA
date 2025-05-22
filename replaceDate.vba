Function ReplaceDates(wkSht, dateArr)
    ' dateArr in format ("2028/9","2027/8","2026/7","2025/6")
    Dim dateVar As Variant
    Dim uBInt, lBint As Integer
    Dim i As Integer
    lBint = LBound(dateArr)
    uBInt = UBound(dateArr)
    For i = lBint To uBInt
        If i > lBint Then
            wkSht.Cells.Replace What:=dateArr(i), _
                                Replacement:=dateArr(i - 1), _
                                LookAt:=xlPart, _
                                SearchOrder:=xlByRows, _
                                MatchCase:=False
        End If
    Next i
    ReplaceDates = "Function Complete"
End Function