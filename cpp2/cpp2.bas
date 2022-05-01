Sub cpp2(fbk As String, fsh As String, fmg1 As Long, fmr1 As Long, fmg2 As Long, fmr2 As Long, tbk As String, tsh As String, tog1 As Long, tor1 As Long, tog2 As Long, tor2 As Long, mdo As Long)
        
    'mdo:12　・・・値と数値の書式　(2002以降)　　xlPasteValuesAndNumberFormats　.copy.paste
    '　　　　-4123 (-99,samaru)　数式　xlPasteFormulas　　　.copy.paste
    '　　　　-4104 (ログ部)　すべて　xlPasteAll　　.copy.paste　・・・クリップボードに情報保持される(繰返し可能)
    '　　　　-4163 (値)　　xlPasteValues　　　　　.copy.pasteしない→だから早い
    'ー4163以外は速度遅い.Copy.Paste なので
    
    If tog2 = 0 And tor2 = 0 Then '従来パターン
        'コピペルーチン　30s85_004
        If mdo = -4163 Then  '新型速度早い。
            Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1 + fmg2 - fmg1, tor1 + fmr2 - fmr1)) = _
              Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Value
    
        Else '従来型(も改良へ) コピペなので遅い(数式コピペはこれ、避けられない)。
            UserForm3.StartUpPosition = 3 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
            UserForm3.Show vbModeless
            UserForm3.Repaint
            Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Copy
            Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1).PasteSpecial Paste:=mdo
            '(参考)↑.copyのコピーメソッドは、selectチックに範囲は指定される挙動である。
            Unload UserForm3
            UserForm1.Repaint
        End If
    ElseIf fmr2 <= 0 Then  '新パターン（一行を複数行にコピペ）-99や近似高速のｦｦASC(PHONETIC())とかで使われる。
        If fmr2 = 0 Then
            If mdo = -4163 Then  '新型速度早い。
                Call oshimai("", bfn, shn, 1, 0, "まだ造成中a")
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)) = _
                Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Value
            Else '従来型(も改良へ) コピペなので遅い(数式コピペはこれ、避けられない)。
                UserForm3.StartUpPosition = 3 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
                UserForm3.Show vbModeless
                UserForm3.Repaint
                Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1)).Copy
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)).PasteSpecial Paste:=mdo
                Unload UserForm3
                UserForm1.Repaint
            End If
        ElseIf fmr2 = -1 Then
            If tog1 > tog2 Then Call oshimai("", bfn, shn, 1, 0, "tog1がtog2よりでかいa")
            UserForm3.StartUpPosition = 3 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
            UserForm3.Show vbModeless
            UserForm3.Repaint
            Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1)) = fsh
            If tog2 > tog1 Then
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1)).Copy
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor1)).PasteSpecial Paste:=mdo 'tor2は使ってない
            End If
            Unload UserForm3
            UserForm1.Repaint
        '4163分岐なし
        ElseIf fmr2 = -2 Then  'フィル(連番のみ、固定フィルはやらないに）
          '4163分岐なし
            If tog1 > tog2 Then Call oshimai("", bfn, shn, 1, 0, "tog1がtog2よりでかいb")
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1)) = fmg2
                If tog2 > tog1 Then
                    Range(Workbooks(tbk).Sheets(tsh).Cells(tog1 + 1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1 + 1, tor1)) = fmg2 + 1
                    If tog2 > tog1 + 1 Then  'フィル
                        Range(bfshn.Cells(tog1, tor1), bfshn.Cells(tog1 + 1, tor1)).AutoFill Destination:=Range(bfshn.Cells(tog1, tor1), bfshn.Cells(tog2, tor1))   'tor2は使ってない
                    End If
                End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "どうなるか未定")
            End If
        Else
        Call oshimai("", bfn, shn, 1, 0, "まだ造成中b")
    End If
    '↓86_014c追加（excel2019 対策向けtest）
    Application.CutCopyMode = False
End Sub
