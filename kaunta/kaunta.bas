Function kaunta(mr() As String, qq As Long, pap5 As Long, bni As Long, er5() As Long, mr5() As String) As Long
    kaunta = 1 '減点方式
    For qap = 0 To pap5  '複数列の減点方式
        If Abs(er5(qap)) < 1 Then
            'A,０や０．４：カウント対象
        ElseIf mr5(qap) = "ヰ#N/A" Then
            If WorksheetFunction.IsNA(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) Then
                '↓エラーになるｗｈｙ？
                kaunta = 0 'カウント非対象
                Exit For '←こうしないと複数列時エラーになる。
            ElseIf Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value = "" Then
                kaunta = 0 'カウント非対象
            End If
        '以下はセルにN/Aあるとエラーになる。20200207
        ElseIf Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value = "" Then
            'B,セル空白はカウント非対象
                kaunta = 0
        '以下セルに情報あり
        ElseIf mr5(qap) = "" Then
            'C,カウント対象
        ElseIf Left(mr5(qap), 1) = "≧" Then
            'D1
            If IsDate(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) Then '日付比較新設 86_018j
                If CDate(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) < CDate(Mid(mr5(qap), 2)) Then
                    kaunta = 0 'カウント非対象
                End If
            Else '↓従来(isnumeric)
                If Val(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) < Val(Mid(mr5(qap), 2)) Then
                    kaunta = 0 'カウント非対象
                End If
            End If
        ElseIf Left(mr5(qap), 1) = "≦" Then
            'D1
            If IsDate(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) Then '日付比較新設 86_018j
                If CDate(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) > Val(Mid(mr5(qap), 2)) Then
                    kaunta = 0 'カウント非対象
                End If
            Else '↓従来(isnumeric)
                If Val(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) > Val(Mid(mr5(qap), 2)) Then
                    kaunta = 0 'カウント非対象
                End If
            End If
        ElseIf Left(mr5(qap), 1) = "ー" Or Left(mr5(qap), 1) = "ヰ" Or mr5(qap) = "n0" Then  '「ー」追加86_012y
            'E　※Dのn0→Eで処理（文字列の0にも対応可能
            '30s75 strcomp 初導入(ヰ△の△とセル側比較がうまく行かないため)
            If StrComp(Mid(mr5(qap), 2), Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value, vbBinaryCompare) = 0 Then
                kaunta = 0 'カウント非対象
            End If
        ElseIf mr5(qap) = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value Then
            'F,n0：カウント対象
        Else
            'Z,上記以外：カウント非対象　　ｦ〇ｦ□　で不一致のケースが想定される。
            kaunta = 0
        End If
    Next
End Function
