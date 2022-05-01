Sub saato(fbk As String, fsh As String, sas As Long, fmg1 As Long, fmr1 As Long, fmg2 As Long, fmr2 As Long, sot As Long)
    'ソート列、fmg1、fmg2　まずは昇順、１ー４列限定で
    'https://excelwork.info/excel/cellsortcollection/
    If sot = 1 Then '昇順
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Clear
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Add Key:=Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas), Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With Workbooks(fbk).Worksheets(fsh).Sort 'Sortオブジェクトに対して '並べ替える範囲を指定し↓
            .SetRange Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Worksheets(fsh).Cells(fmg2, fmr2))
            .Header = xlNo '1行目がタイトル行かどうかを指定し（規定値：xlNo）
            .MatchCase = False '大文字と小文字を区別するかどうかを指定し
            .Orientation = xlTopToBottom '並べ替えの方向(行/列)を指定し  (規定値：xlTopToBottom)
            .SortMethod = xlPinYin 'ふりがなを使うかどうかを指定し  (規定値：xlPinYin)
            .Apply '並べ替えを実行します 　省略はしない方が無難、前回のを引き継ぐらしいので
        End With
    ElseIf sot = 99 Then '降順
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Clear
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Add Key:=Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas), Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas)), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With Workbooks(fbk).Worksheets(fsh).Sort
            .SetRange Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Worksheets(fsh).Cells(fmg2, fmr2))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Else
        Call oshimai("", bfn, shn, 1, 0, "sotの引数が変です")
    End If
End Sub
