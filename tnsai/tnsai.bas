Sub tnsai(ct8 As String, tst As Long, ct3 As String, er78() As Currency, a As Long, hx As Long, bni As Long, p As Long, qq As Long, fuk12 As Long, mr() As String, er() As Currency, pap7 As Long, mr8() As String)

    'ヱ対応7行8行、当列a,書込行hx(0はコメント用),文節bni,P値(AM有無)、参照行(コメント時のケース(項目行が参照行)もあり)、-1-2複写,mr,er,7行目ヱの数(pap7)
    Dim ax As Long '書込列（一方、aは当列）
    Dim tenx As String, teny As String, tenz As String
    Dim qap(3) As Long, apa7 As Long    'qap3は将来用
    Dim bx As Long, mm As Long
    'bx策定
    
    If fuk12 = -7 Then bx = 0 Else bx = 1  '-7は７行コメント用,fuk12:-1-2複写フラグのこと 通常はbx:1
    
    If er(5, bni) < 0 Or fuk12 = -7 Then apa7 = 0 Else apa7 = pap7
    '差分時あるいは7行のコメント時はpap7成分無効化(以降当ルーチンapa7使用、pap7不使用(-7以外))
    
    If fuk12 > 0 Then '-1-2複写時 fuk12 は転載元の行
        qap(1) = -1
        qap(2) = apa7
    Else
        qap(1) = 0
        If fuk12 = -7 Then  '7行のコメントの特殊時
            If pap7 <= UBound(er78(), 2) Then '普通な時
                qap(2) = 0    'pap7→0へ
            Else '普通でないとき　30s85_021新設
                'MsgBox "ここはもう来ないのでは"    '７行ヱの数＞8行ヱの数　はどこかでoshimai処理だったような。。
                Call oshimai("", bfn, shn, sr(7), a, "ここはもう来ないのでは")
                qap(2) = UBound(er78(), 2)
            End If
        ElseIf fuk12 = -8 Then
            qap(2) = 0    'UBound(er78(), 2)→0へ
        Else '通常時
            qap(2) = apa7  '（通常）UBound(er78(), 2)[８行目のヱの数]→apa7[7行目のヱの数]へ
        End If
    End If

    For qap(0) = qap(1) To qap(2)  '複数列転載時ループ(通常：7行目の複数列個分)する。単数列はループせず1回のみ。-1-2複写時、-1からループする。qap(2)までループする。
                             'qap(2)は通常は8行目ヱの数。7行コメントだけ7行ヱの数。8行ヱ数≠7行ヱ数の時意識せよ。
        '★７行ーの時はスルーへ　86_013
        If er78(0, qap(0)) = 0.1 Then
            'MsgBox "7行0.1(ー)、省略数目あり"
        ElseIf er78(0, qap(0)) = 0.4 Then
            Call oshimai("", bfn, shn, sr(7), a, "7行目0.4は今の所あり得ないかと。")
        Else
            '同列か他列か？（ax策定）転載元・転載先で使用
            If qap(0) <= apa7 Then '単数列＆複数列　apa7最小値ゼロ（：7行ヱ無し）
                'ax更新(ヱの、はみ出し分は更新しない)
                If er78(0, qap(0)) > 0 And er(5, bni) >= 0 Then '差分時でなく、7行正値の時のみ他列許容、(普通の転載)
                    ax = er78(0, qap(0))  '他列
                Else
                    '↓30s86_012w 特命条件対応版
                    If er(9, bni) > 0 And er(10, bni) > 0 And er(5, bni) >= 0 And Not (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And Int(er(7, bni)) = 0 And er(8, bni) <> 0) Then
                    If er(20, bni) = 1 Then MsgBox "yaa"
                        Call oshimai("", bfn, shn, 1, 0, "b加算ありで同列転載しようとしています。確認を。")
                    Else
                        ax = a  '差分時でも一発目だけは有効　-1の時
                    End If
                End If
            End If
    
            tenx = ""   '不要(tneleでリセットされる)だが、一応念のため
            tenx = tnele(ct8, ax, qap(0), ct3, er78(), a, hx, bni, qq, fuk12, mr(), er(), mr8())
            '↑30s86_013j function化
            
            '連載処理〔Ｂ〕・・〔Ａ〕よりこちらが先に
            If fuk12 <= 0 Then  'ｃが-1、-2以外 30s82=→<= に(コメント-7-8対応)
                If qap(0) = qap(2) Then '最終数のみ以下
                    If UBound(er78(), 2) > qap(2) Then 'はみ出ている場合のみ、以下はみ出し分連結処理〔Ｂ〕
                        For mm = qap(2) + 1 To UBound(er78(), 2)
                            If er78(0, mm) <> 0.1 Then
                                tenx = tenx & mr(2, 4, bni) & tnele(ct8, ax, mm, ct3, er78(), a, hx, bni, qq, fuk12, mr(), er(), mr8())
                            End If
                        Next
                    End If
                End If
            End If
       
            '連載処理〔Ａ〕(tenx更新)〔Ａ〕いつもの当初からの連載処理(語尾付加)
            If fuk12 <= 0 Then  'ｃが-1、-2以外
                    'ヱなし、同列、６行ゼロor-1台、８行マイナスが条件、複数列許容へ
                            '↓ノーマル連載は-2以下は許容してない。あるいは、
                    If (er(6, bni) = 0 Or Round(er(6, bni)) = -1) Or _
                        (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) < 0) Then '<0→<0.3 85s023
                        '↑特命条件連載型をorで追加
                        '↓86_018d　複数列初列で判断へ
                        If mr8(0) <> "" And (er78(bx, 0) < 0) Then  'Or er78(bx, qap(0)) = 0.1は不要だった(行番号は連載型許容してない)
                            tenx = bfshn.Cells(hx, ax).Value & tenx & mr8(0) 'あ 特命条件もこっち mr(2,8,bni)→mr8(qap(0))→mr8(0)
                        ElseIf er78(bx, 0) < -0.5 Then    '(bx, qap(0))→(bx, 0)
                            Call oshimai("", bfn, shn, sr(8), a, "もう通ることはないと思われる(※８行目は新仕様に修正して下さい)")
                            tenx = bfshn.Cells(hx, ax).Value & tenx & "、"  'い
                        Else
                            '行番号上書き型(0.1、ここでは何もしない)
                        End If
                        '↑特命条件以外での他列複数列連載型は↑この変動動作が整理されていない状況である。202005
                    End If
            End If
            
            '複数列の最終数目以外をここで書き込み、-7-8コメント時はここで実施しない(最後で)。ｰ1ｰ2複写処理はここでは行われない。
            '複数列転載行為の最終数目(単数列はその一回)は、ここでは実施せずfornextの後ろで実施(最終まとめ分のtenxがまとまってないため。)
            If tenx <> "" And hx <> 0 And fuk12 >= 0 Then  'And qap(0) < apa7 の条件撤廃　86_013q
                'tst、trt↓使えそう。-7-8コメントはここ使われていないし。
                If p = 2 And tst = 1 Then
                    With bfshn.Cells(hx, ax)  '文字列新規ちまちま(-3<0&-4<0　か連載新規),※fuk12はここ弾かれる。
                        .NumberFormatLocal = "@"
                        .Value = tenx
                    End With
                ElseIf tst = 8 And fuk12 = 0 Then  '元シートセルちまちま
                    If Abs(er78(1, qap(0))) = 0.4 Or Abs(er78(1, qap(0))) = 0.1 Then Call oshimai("", bfn, shn, sr(8), a, "この形式では踏襲型ちまちまで指定できませんa")
                    With bfshn.Cells(hx, ax)
                        .NumberFormatLocal = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap(0)))).NumberFormatLocal
                        .Value = tenx
                    End With
                ElseIf tst = 8 And er(6, bni) <> -2 And fuk12 > 0 Then   'セル複写ちまちま(fuk12)86_012k
                    With bfshn.Cells(hx, ax)
                        .NumberFormatLocal = bfshn.Cells(fuk12, ax).NumberFormatLocal
                        .Value = tenx
                    End With
                Else 'normal(感知しない、そのまま値で貼り付け、標準でもない（後で標準なり通貨なり処理）)
                    bfshn.Cells(hx, ax).Value = tenx
                End If
            End If
        End If '７行ーでない時の処理はここまで
    Next '複数列転載時、ループする。-1-2複写時、-1からループする。は、ここまで
    
    qap(0) = qap(0) - 1 '1戻す(Next後のインクリ戻す)
    'ここでのtenxは、最終まとめ分のtenxがまとまっている状態
    '単数列転載は必ずここで行われる。for側では行われない。ｰ1ｰ2複写はここで行われる。
    If hx = 0 Then MsgBox "hx=0ということがあり得るだろうか？"

    If (fuk12 = -8 Or fuk12 = -7) And tenx <> "" Then '30s79 8行コメント用
        bfshn.Cells(hx, a).ClearComments
        bfshn.Cells(hx, a).AddComment
        bfshn.Cells(hx, a).Comment.Text Text:=tenx
        bfshn.Cells(hx, a).Comment.Shape.TextFrame.AutoSize = True
    End If
End Sub
