Sub hdrst(ii As Long, a As Long)  '21ｓ簡素化
    '左下ステータス表示部  100→1000へ201904
    If cnt < Int(ii / 1000) * 1000 Then  'cnt はパブリック変数
        cnt = Int(ii / 1000) * 1000
        DoEvents
        If flag = True Then Call oshimai("", bfn, shn, 1, 0, "中止しましたです") '中止ボタン処理
        Application.StatusBar = Str(cnt) & "、" & Str(a - dd1 + 1) & " / " & Str(dd2 - dd1 + 1)
    End If
End Sub
