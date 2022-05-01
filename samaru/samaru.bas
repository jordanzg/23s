Sub samaru(aa As Long, mr101 As String)  'サマリ値処理30ｓ59より
    'さまる列、mr(1, 1, 1)←シート名
    Dim gx As Long, gy As Long, ii As Long

    For ii = sr(0) + 4 To sr(0) + 5
        uu = 0
        gx = 0
        gy = 0
        If IsError(bfshn.Cells(ii, aa)) Then '30s66_2バグ改良
            uu = 1  'セルがエラー→値が入っている。
        ElseIf bfshn.Cells(ii, aa) <> "" Then
            uu = 1
        End If
        
        If uu = 1 Then 'ん
            '30s71_3場所若干変更、３０ｓ７４さらに変更
            If bfshn.Cells(ii, aa).Font.Color = RGB(38, 37, 38) Or bfshn.Cells(ii, aa).Font.Color = RGB(255, 0, 0) Then '黒[赤]sum
                gy = sr(8) + 5
                gx = 4
            ElseIf bfshn.Cells(ii, aa).Font.Color = RGB(1, 102, 255) Then '青数値＞０
            
                gy = sr(8) + 5 'sr(8)は対:転載列の行
                gx = 2  '4
            ElseIf bfshn.Cells(ii, aa).Font.Color = RGB(198, 90, 17) Then '茶文字列
                gy = sr(8) + 6 '17
                gx = 4 '2
            ElseIf bfshn.Cells(ii, aa).Font.Color = RGB(1, 176, 80) Then '緑nonzero
                gy = sr(8) + 6 '16
                gx = 2
            End If
        End If  'ん
             
        If gx > 0 Then  'コピペ部
             Call copipe(bfn, shn, gy, gx, gy, gx, bfn, shn, ii, aa, ii, aa, 3)  '3→FormulaR1C1(-4123向け) (脱.copy.paste)　86_020d
            
            'コピペしたセルを値に変換
            If StrConv(Left(mr101, 1), 8) <> "*" Then
                Call cpp2(bfn, shn, ii, aa, ii, aa, bfn, shn, ii, aa, 0, 0, -4163) '-4163は値をコピー　速
            Else
                '２行が＊だけの時、ここ通過してる。
                'MsgBox "ここを通るケースが未だにある。"
            End If
            bfshn.Cells(sr(0) + 3, aa).Value = Now() '活用例1タイムスタンプ入れる
        End If
    Next
End Sub
