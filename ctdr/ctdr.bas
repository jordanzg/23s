Function ctdr(rtyu As String, tyui As String, qwer As Currency, wert As Long) As Long '　最右列を返す。'こちらは列の方
    'ブック名、シート名、er4系、当列
    '初期化 ↓対象シートの最右列　628
    ctdr = Workbooks(rtyu).Sheets(tyui).Range("A1").SpecialCells(xlLastCell).Column()
    If ctdr > 300 And Abs(qwer) < 1 Then
        Call oshimai("", bfn, shn, sr(1), wert, "対象シートが300列超え(" & ctdr & ")です")
    End If
    ctdr = ctdr + 1
    Do Until Workbooks(rtyu).Sheets(tyui).Cells(1, ctdr).EntireColumn.Hidden = False
        ctdr = ctdr + 1
    Loop  'ctrl+endの次列がhiddenだった場合の対処（85_020)
    ctdr = ctdr - 1
End Function
