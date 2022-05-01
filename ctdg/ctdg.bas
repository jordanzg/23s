Function ctdg(rtyu As String, tyui As String, qwer As Currency, wert As Long) As Long '　最終行を返す。'こちらは行の方
    'ブック名、シート名、er4系、当列
    ctdg = ctreg(rtyu, tyui)    '←こちらに凝縮、pubikouへ
    '項準b(項零)での制限事項↓
    If ctdg > 10000 And Abs(qwer) < 1 Then  '1000→10000　86_014r
        Call oshimai("", bfn, shn, sr(1), wert, "対象シートが一万行超え(" & ctdg & ")です")
    End If
End Function
