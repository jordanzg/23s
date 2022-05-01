Sub giktzg(a As Long, rog As String)     '外結事後
    Dim ii As Long, jj As Long
    '↓一ボタン毎にインクリメント（ID変わったときのみリセットされる） 。※bfn、shn側は更新されない（初回複写時の値が載ってるだけ）。
    twbsh.Cells(13, 3).Value = twbsh.Cells(13, 3).Value + 1 '7s (2,3)→(13,3)25s
    
    'ログ部ここから　30s76
    If rog <> "" Then     'if,rogがnullなら(＊だけの時）記載しない
    'MsgBox rog
        jj = 1
        Do Until Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = ""
            jj = jj + 1
            If jj = 10000 Then
                MsgBox "空白行が見つからないようです"
                Exit Sub
            End If
        Loop

        Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = 1 '項目名
        Workbooks(twn).Sheets(shog).Cells(jj, 2).Value = jj '項番

        'log
        Workbooks(twn).Sheets(shog).Cells(jj, 3).Value = "結.　外結、" & Mid(twn, 1, Len(twn) - 5) _
        & "、ｦ" & shn & "ｦ" & bfn & "、ｦ" & shn & "ｦ" & bfn & "、from" & dd1 & "to" & dd2 _
        & "、項目名、b" & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "、" & twbsh.Cells(2, 2).Value & "、" _
        & Format(Now(), "yyyymmdd_hhmmss") & "、" & bfshn.Cells(sr(8), 5).Value & "、" & Application.WorksheetFunction.Sum(bfshn.Range("A:A")) & "、" & dd2 - dd1 + 1
        '末備から、外結処理列数(min:1)、all1(行数)←(＋１では

        Workbooks(twn).Sheets(shog).Cells(jj, 4).Value = Format(Now(), "yyyymmdd")  'date
        Workbooks(twn).Sheets(shog).Cells(jj, 5).Value = Format(Now(), "yyyymmdd_hhmmss")  'timestamp
        Workbooks(twn).Sheets(shog).Cells(jj, 7).Value = bfn & "\" & shn  'to

        frmx = 9 'fromの開始列
        ii = 1
        Do Until rvsrz3(rog, ii, "ヱ", 0) = ""
            Workbooks(twn).Sheets(shog).Cells(jj, ii + frmx - 1).Value = rvsrz3(rog, ii, "ヱ", 0) 'from
            ii = ii + 1
            If ii = 200 Then
                Call oshimai("", bfn, shn, k, a, "うまくいってない2。")
            End If
        Loop
        Workbooks(twn).Sheets(shog).Cells(jj, 8).Value = ii + frmx - 2 '最右列の列に入れる値
    End If  'ログ部ここまで
    
    Application.CutCopyMode = False
    bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
    DoEvents
    Call oshimai("", bfn, shn, k, dd2, "")
End Sub
