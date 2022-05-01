Function chekku() As String
    'Excelバージョンチェック,　86_020s挿入部ここから　020u改良
    twt.Cells(1, 1).Value = "あ"  '詳細記述は初回のみ（）の方で実施
    twt.Cells(2, 1).Value = "ア"
    
    twt.Sort.SortFields.Clear
    twt.Sort.SortFields.Add Key:=Range(twt.Cells(1, 1), twt.Cells(1, 1)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With twt.Sort
        .SetRange Range(twt.Cells(1, 1), twt.Cells(2, 1))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    If twt.Cells(1, 1).Value = "ア" Then
        chekku = "新ｿｰﾄ" '(excel2019のあるver以降)"
    Else
        chekku = "旧ｿｰﾄ" '順序そのまま(excel2016以前、excel2007)"
    End If
End Function
