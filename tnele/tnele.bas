Function tnele(ct8 As String, ax As Long, qap0 As Long, ct3 As String, er78() As Currency, a As Long, hx As Long, bni As Long, qq As Long, fuk12 As Long, mr() As String, er() As Currency, mr8() As String) As String '転載エレメント決定　86_013j
    '転載エレメント決定　30s86_013j function化
    If fuk12 > 0 Then '30s62 ｃ：-1-2複写時
        tnele = bfshn.Cells(fuk12, ax).Value
    ElseIf fuk12 = -7 Then 'ここから以下、fuk12が0以下
        If er78(0, qap0) = 0 Then   'bx→0へ
            tnele = Trim$(CStr(Workbooks(mr(2, 0, bni)).Sheets(mr(1, 0, bni)).Cells(qq, Abs(a)).Value))
        Else
            tnele = Trim$(CStr(Workbooks(mr(2, 0, bni)).Sheets(mr(1, 0, bni)).Cells(qq, Abs(er78(0, qap0))).Value))
        End If
    '以降、bx:1
    ElseIf er78(1, qap0) = 0.4 Then 'マニュアル値(mr8)は、er8ないときだけ使われる。
        tnele = Trim$(mr8(qap0))     '30s75複数列対応化
    ElseIf Round(er(6, bni), 0) = -15 Then    '-15 ある文節エレメント抽出
        tnele = rvsrz3(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), Val(mr(2, 6, bni)), mr8(0), 0)
    ElseIf Round(er(6, bni), 0) = -14 Then '区切り数　30s86_016m
        tnele = kgcnt(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), mr8(0)) + Val(mr(2, 6, bni))
    ElseIf Round(er(6, bni), 0) = -10 Then 'naka2
        tnele = Mid(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), Val(mr(2, 6, bni)))
    ElseIf Round(er(6, bni), 0) = -9 Then 'hiduke -9
        tnele = Format(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), mr(2, 6, bni))
    ElseIf Round(er(6, bni), 0) = -8 Then 'mojihen    -13→-8
        'MsgBox "使われてますね。6行が-8"
        tnele = StrConv(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), mr(2, 6, bni))
    ElseIf Abs(er78(1, qap0)) = 0.1 Then
        tnele = Format(qq, "0000000")  '30s85_014 行番号転載対応の修正
    ElseIf er78(1, qap0) > 0.2 And mr8(qap0) <> "" And bfshn.Cells(hx, a).Value <> "" And UBound(er78(), 2) = 0 Then
        '20180813新井対応→202102複数列では実施せずの仕様に(動作おかしくなるので) 86_020j
        '何もしない（tnele=""のまま）第二因子がｖ用　 "6448"
    Else '通常時　fuk12=-8はここ通る。
        If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) < 0) Then
            If qap0 = 0 Then
                tnele = ct8  '特命条件(連載型使用時、８列加算値、86_019oこちらへ)
            Else
                tnele = ct3  '特命条件(連載型使用時、６列値（従来）)
            End If
        Else '従来の通常パターン
            tnele = Trim$(CStr(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))).Value))
        End If
    End If
End Function
