Sub iechc(hk1 As String)  '以前はigchc
    If twbsh.Cells(3, 2).Value = "" Then  '86_017g 初回誰でも使えるように。
        twbsh.Cells(3, 2).Value = hunk2(-5, syutoku(), "1", hk1)
        If hk1 = "1" Then
            Call oshimai("", twn, "▲集計_雛形", 3, 2, "(処理終了)IDキーが作れません" & vbCrLf & "※ID:" & syutoku())
        End If
    ElseIf syutoku() = hunk2(-7, twbsh.Cells(3, 2).Value, "1", hk1) Then
        'MsgBox "正解です"
    Else  '"不正解です"
        ThisWorkbook.Activate
        Sheets("▲集計_雛形").Select
        twbsh.Cells(1, 1).Select  '緑色セル
        Call oshimai("", twn, "▲集計_雛形", 3, 2, "(処理終了)IDキー不一致" & vbCrLf & "※ID:" & syutoku())
    End If
    
    hk1 = Left(twn, 7) & "r"   'ラピド固定
    twbsh.Cells(2, 2).Value = syutoku() & "r" '新設86_016e
End Sub
