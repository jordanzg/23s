Function tszn(er9() As Currency, bni As Long, mr() As String, qq As Long, pap9 As Long, mr9() As String) As Variant  '加算(足し算)
    'ヱ対応6or8行、当列a,書き込み行hx,文節bni,P値(AM有無)、同列0多列1(使用終了)、-1-2複写,mr,er,6行目８行目ヱの数(pap9),mr9
    Dim qap(2) As Long
    'Dim nuez As Double
    qap(1) = 0
    qap(2) = pap9
    For qap(0) = qap(1) To qap(2)  '複数列取扱時、ループする。
        If Abs(er9(qap(0))) = 0.4 Or Abs(er9(qap(0))) = 0.1 Then
            nuex = Val(mr9(qap(0)))
        Else '通常時
            nuex = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er9(qap(0)))).Value 'er(11, bni)→ee→qq
        End If
        If qap(0) = 0 Then
            'ヱが０個のとき→何もしない
        Else
            nuex = nuey * nuex '複数列の２回目以降→乗算実施
        End If
        nuey = nuex
    Next
    tszn = nuex
End Function
