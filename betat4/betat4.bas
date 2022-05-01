Sub betat4(fbk As String, fsh As String, fmg1 As Long, fmr1 As Currency, fmg2 As Long, fmr2 As Currency, tbk As String, tsh As String, tog As Long, tor As Long, er34 As String, mr_8 As String)
    'e ver
    Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).ClearContents  'betatnはcppと異なり、仕様として、クリアすることとする。
    If er34 = "pp" Then '標準型調整
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).NumberFormatLocal = "G/標準"
    ElseIf er34 = "mm" Then '文字列型調整
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).NumberFormatLocal = "@"
    ElseIf er34 = "pm" Then '通貨型調整
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).NumberFormatLocal = "#,##0;[赤]-#,##0"
    Else
        'mp
    End If
    
    If Abs(fmr1) = 0.4 Or Abs(fmr1) = 0.1 Then 'ｦｦ定文字対応、ｦーｦ行番号対応　（フィル型）
        'こちらにも(ｦｦ、0.4向け)
        If Abs(fmr1) = 0.1 Then 'ｦーｦ行番号対応　※暫定運用
            Workbooks(tbk).Sheets(tsh).Cells(tog, tor).Value = Trim$(mr_8) & Format(fmg1, "0000000") '86_019a
            If fmg2 > fmg1 Then '範囲が1行か2行しか無い場合の対処(以下同文)
                If mr_8 <> "" And mr_8 <> "000_" Then MsgBox "betat4挙動注意bb " & mr_8 '30s86_019a
                Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor).Value = Trim$(mr_8) & Format(fmg1 + 1, "0000000")
            'ここ、高速シートの2列目をコピーすることも検討し得る。
            End If
            If fmg2 > fmg1 + 1 Then  'フィルは３行以上ある場合のみ実施
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor)).AutoFill Destination:=Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor))
            End If
        Else 'ｦｦ定文字対応(通常) (0.4)
            Workbooks(tbk).Sheets(tsh).Cells(tog, tor).Value = Trim$(mr_8)
            If fmg2 > fmg1 Then
                Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor).Value = Trim$(mr_8)
            End If
            If fmg2 > fmg1 + 1 Then  'フィルは３行以上ある場合のみ実施
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor)).AutoFill Destination:=Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor))
            End If
        End If
    Else '一般対応時、含複数列ベタ型
        If er34 = "mp" Then  'セル踏襲
            Call cpp2(fbk, fsh, fmg1, Int(fmr1), fmg2, Int(fmr2), tbk, tsh, tog, tor, 0, 0, 12) '12:値と数値の書式 '遅・コピペパターン
        Else 'mm,pp,pm
            Call cpp2(fbk, fsh, fmg1, Int(fmr1), fmg2, Int(fmr2), tbk, tsh, tog, tor, 0, 0, -4163) '-4163:値
        End If
    End If
End Sub
