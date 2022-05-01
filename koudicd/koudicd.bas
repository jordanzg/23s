Function koudicd(fn As String, ff As String, er0 As Currency, mr04 As String) As String
    'er0返り値は項目行の行数（返り値0も有：項無の場合）
    '①対象シートの「項目名」探し
    Dim ywe As String
    For er0 = 1 To 2000  '①
        If Workbooks(fn).Sheets(ff).Cells(er0, 1).Value = "項目名" Then
            koudicd = "項有"
            Exit For
        End If
        If Right(Workbooks(fn).Sheets(ff).Cells(er0, 1).Value, 4) = "列固有名" Then  '30s82より
            koudicd = "項固"
            Exit For
        End If
    Next
    
    If koudicd = "" Then koudicd = "項無"  '←仮
        If InStr(1, mr04, "ｦ") > 0 Then 'ｦあり[あ]
            'yweは文字
            ywe = iptfg(mr04, 3, "ｦ") '先に「ヱ」把握 iptfg・・NewVersion[い]
            If ywe <> "" And InStr(1, iptfg(mr04, 2, "ｦ"), ywe) > 0 Then 'ｦ△ヱ▲ｦヱ
                ywmoji = iptfg(iptfg(mr04, 2, "ｦ"), 1, ywe) '[う]
                yw10 = iptfg(iptfg(mr04, 2, "ｦ"), 2, ywe)  '[え]
            Else 'ヱなし
                ywmoji = iptfg(mr04, 2, "ｦ") 'yw10はnull[お]
            End If
        Else  'ｦなし
            ywmoji = mr04  'yw10はnull
        End If
        If Mid(ywmoji, 1, 1) = "ー" Then ywmoji = Mid(ywmoji, 2)
        
        '②④対象シートの「項目名」項指項有▲(ｦ△ヱ▲ｦヱ)or項零
        If ywmoji = "" Or ywmoji = "0" Then  '項零628
            If yw10 = "" Then
                er0 = 0
            ElseIf IsNumeric(yw10) Then
                er0 = Val(yw10) - 1
            Else
                Call oshimai("", bfn, shn, 1, 0, "項準bが実施できないようです。")
            End If
            koudicd = "項準b"  '後々項零に変えたい。
            'MsgBox "項準b（項零）"
        ElseIf yw10 <> "" Then   '▲(ｦ△ヱ▲ｦヱ)
            er0 = 1
            If Not IsNumeric(ywmoji) Then '場合分け622
                '従来パターン
                Do Until Workbooks(fn).Sheets(ff).Cells(er0, 1).Value = yw10
                    If er0 = 20000 Then '2000→20000
                        Call oshimai("", bfn, shn, 1, 0, "項指の項目行が見つからないようです")
                    End If
                    er0 = er0 + 1  '④
                Loop
                koudicd = "項準2"  '②　30s64　項指→項準2(旧項指)に変更
            Else        'Newパターン　項準a ここのyw10はall1の代替
                Do Until Workbooks(fn).Sheets(ff).Cells(er0, Abs(Val(ywmoji))).Value = yw10
                    If er0 = 20000 Then '2000→20000
                        Call oshimai("", bfn, shn, 1, 0, "項指の項目行が見つからないようです")
                    End If
                    er0 = er0 + 1  '④
                Loop
                koudicd = "項準a"  '②　30s64　項指→項準2に変更
            End If
        End If
        
        '③⑤対象シートの「項目名」及び項指もないが、all1が数値記載でない場合(準項目名) 終了措置628
        If koudicd = "項無" And Not IsNumeric(ywmoji) Then
            er0 = 1
            Do Until Workbooks(fn).Sheets(ff).Cells(er0, 1).Value = ywmoji
                If er0 = 20000 Then '2000→20000
                    Call oshimai("", bfn, shn, 1, 0, "項準の項目行が見つからないようです")
                End If
                er0 = er0 + 1  '⑤
            Loop
            koudicd = "項準" 'yw10は""　③
            Call oshimai("", bfn, shn, 1, 0, "項準は終了しました。項準aに移行して下さい。") '終了措置628
        End If
    If koudicd = "項無" Then er0 = 0
End Function
