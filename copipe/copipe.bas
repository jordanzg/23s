Sub copipe(fbk As String, fsh As String, fmg1 As Long, fmr1 As Long, fmg2 As Long, fmr2 As Long, tbk As String, tsh As String, tog1 As Long, tor1 As Long, tog2 As Long, tor2 As Long, mdo As Long)
    'mdo: 1:value 、2:formula、3:formulaR1C1
    'クリップボードを使用しないコピペの実現
    If mdo = 1 Then  'Value
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)) = _
          Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Value
    ElseIf mdo = 2 Then
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)) = _
          Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Formula
    ElseIf mdo = 3 Then '-4123向け
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)).FormulaR1C1 = _
          Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).FormulaR1C1 '-4123向け
    Else
        Call oshimai("", bfn, shn, 1, 0, "まだ造成中d")
    End If
End Sub
