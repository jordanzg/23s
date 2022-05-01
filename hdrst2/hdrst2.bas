Sub hdrst2(ii As Long, a As Long, ak As Long, kkk As Long, hhh As Long)
    '左下ステータス表示部
    If ak <= 0 Then ak = 100 '異常値のときは100のデフォ値(従来変わらず)で
    If cnt < Int(ii / ak) * ak Then  'cnt はパブリック変数
        If kkk <> 0 Then
            bfshn.Cells(2, 4).Value = kkk
            bfshn.Cells(3, 4).Value = hhh
        End If
        cnt = Int(ii / ak) * ak
        DoEvents
        If flag = True Then Call oshimai("", bfn, shn, 1, 0, "中止しましたよ") '中止ボタン処理
        Application.StatusBar = Str(cnt) & "、" & Str(a - dd1 + 1) & " / " & Str(dd2 - dd1 + 1)
    End If
End Sub
