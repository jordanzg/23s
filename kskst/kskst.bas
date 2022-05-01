Sub kskst(pap7 As Long, h As Long, er78() As Currency, er9() As Currency, mr9() As String, er3() As Currency, mr3() As String, er5() As Long, mr5() As String, c5 As Long, pap3 As Long, pap5 As Long, bni As Long, qq As Long, rrr As Long, mr() As String, er() As Currency, a As Long, cted() As Long) '高速シートは徐々にこちらへ
    Dim hirt As Variant, ii As Long, tempo As String, baba As String
    ThisWorkbook.Activate
    Sheets("高速シート_" & syutoku()).Select
    DoEvents
        
    twt.Cells.Clear
    twt.Cells.Delete Shift:=xlUp
    DoEvents
    twt.Columns("A:A").NumberFormatLocal = "@"  '一列目文字列に
    twt.Columns("G:G").NumberFormatLocal = "@"  '7列目(旧4列目)文字列に（転載前キー列用途）　30s85_027
    twt.Columns("F:F").NumberFormatLocal = "@"  '6列目も文字列に（転載前キー列用途）　30s86_014g
    twt.Columns("D:D").NumberFormatLocal = "@"  '４列目文字列に（新規分情報a※用途のみ）へ　30s85_027
'    twt.Columns("K:V").NumberFormatLocal = "@"  '86_020h
        
    rrr = qq

    If Abs(er(4, bni)) >= 1 Then  '最終行と抜け有無チェック(単列のみ)
        'ここでのhirtは対象シートの「ALL一列」の1列・・seekの最後検知用
        hirt = Range(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(1, Abs(er(4, bni))), Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(cted(0) + 1, Abs(er(4, bni)))).Value
        Do Until hirt(rrr, 1) = ""
            rrr = rrr + 1
            Call hdrst2(rrr, a, 10000, 0, 0)
        Loop
        Erase hirt
    Else
        rrr = cted(0) + 1  '項準ｂ対応　30d85_018
    End If
    DoEvents
    rrr = rrr - 1 'rrrは対象シートの最終行　qqは対象シートのベタ貼り開始行。 対象シートのデータ部がすっからかんのとき、rrr<qqとなってしまう
    cnt = 0
    
    If rrr < qq Then  '対象シートデータ部がすっからかんな時
        MsgBox "対象シートのデータ部がすっからかんです。動作は続きます"
    Else '対象シートデータ部がすっからかんでない時の処理ここから(通常はこちら)
        '2列目(行番号)の処理(フィル活用)　ベタ・ちま共用
        Call betat4(twn, "高速シート_" & syutoku(), qq, 0.1, rrr, 0.1, twn, "高速シート_" & syutoku(), qq, 2, "pp", "000_")
        
        '3列目(カウント列)に入れ込む処理◆常時入ることに
        If Abs(er(5, bni)) <> 0 Then  'カウント列の情報(複数時は一番左)→3列目へ
            '86_012r：20190515★大改造
            If pap5 = 0 And mr(2, 5, bni) = "" Then  '以前は一律この仕様↓
                Call cpp2(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er(5, bni)), rrr, Abs(er(5, bni)), twn, "高速シート_" & syutoku(), qq, 3, 0, 0, -4163) '85_027検証3
            Else '追加仕様がこちら　86_012r
                cnt = 0
                For ii = qq To rrr
                    c5 = kaunta(mr(), ii, pap5, bni, er5(), mr5())
                    If c5 = 1 Then  'カウント対象なら実施
                        If pap5 = 0 Then 'カウンタ単数列・・カウンタセルをコピペ
                            Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(ii, 3).Value = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(ii, Abs(er(5, bni)))
                        Else  'カウンタ複数列・・・1を入れる仕様
                            Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(ii, 3).Value = 1
                        End If
                    End If
                    Call hdrst2(ii, a, 1000, 0, 0)
                Next
            End If
        Else   'カウント列ゼロ時・・・高速3列に入れ込む仕様に(これまで入れて無かった)
            '1を入れ込んでる。
            Call betat4(twn, "高速シート_" & syutoku(), qq, 0.4, rrr, 0.4, twn, "高速シート_" & syutoku(), qq, 3, "pp", "1")
        End If
        
        '一列目作成(pap3単数時・複数時場合分け)
        If pap3 = 0 Then '単数列時（ pap3=0 ）：ベタ貼り仕様

            'こっちは高速シート演算使ってない。 hiro→hirt
            hirt = Range(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er3(0))), Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(rrr + 1, Abs(er3(0))))
            For ii = 1 To 1 + rrr - qq
                If hirt(ii, 1) = "" Then
                    hirt(ii, 1) = "ー(情報空白行)ー" '013tこちらへ
                Else
                    hirt(ii, 1) = hirt(ii, 1) & ""      '数値→文字列とさせる技　86_012
                End If
            Next
            Range(Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(qq, 1), Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(rrr, 1)).Value = hirt
            '↑hirtの一番下(ダミー行)は無視されるだけ
            Erase hirt
            Application.Cursor = xlWait
            Application.Cursor = xlDefault
        Else  '複数列時（ pap3>0 ） babaの暫定変数は後々恒久措置検討(既存使い回せないか？)。
                'ｍｍ：文字列型、pm：通貨型、
            If er3(0) <> 0.1 Then '（６→11列目に転載）第1列でのー(：行番号転載)は実施しない。
                Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er3(0)), rrr, Abs(er3(0)), twn, "高速シート_" & syutoku(), qq, 11, "mm", mr3(0))
            End If
            
            For ii = 1 To pap3 '（７→１２列目以降に転載）第2列からはこちら　行番号転載もあればする。
                If er3(ii) = 0.1 And mr3(ii) <> "" Then MsgBox "betat4挙動注意b" '30s86_019m
                Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, er3(ii), rrr, er3(ii), twn, "高速シート_" & syutoku(), qq, 11 + ii, "mm", mr3(ii))
            Next

            tempo = "R[0]C[1]"
            baba = tempo
            For ii = 1 To pap3 '数式の素生成
                tempo = "&" & """" & mr(2, 4, bni) & """" & "&" & "R[0]C[" & LTrim(Str(ii + 1)) & "]"
                baba = baba & tempo
            Next
                
            Application.Calculation = xlCalculationAutomatic   '数式計算方法自動に(次行演算の為)
            twt.Cells(qq, 10).FormulaR1C1 = "=" & baba  '10列目1行目に数式の素を注入

            Call copipe(twn, "高速シート_" & syutoku(), qq, 10, qq, 10, twn, "高速シート_" & syutoku(), qq + 1, 10, rrr, 10, 3) '3→FormulaR1C1 '86_019s
        
            '↓ようやく10列目を１列目にコピー（数式→値化）
            Call betat4(twn, "高速シート_" & syutoku(), qq, 10, rrr, 10, twn, "高速シート_" & syutoku(), qq, 1, "mm", "")
            
            twt.Columns("J:K").ClearContents  '10列11列クリア(12列以降は特にクリアはしてない)
        End If  '一列目作成ここまで
                   
 '複文節工事予定ここまで　86_018m
        
        '1列目→11列目(カタカナ、半角化)
        
        twt.Columns("K:K").NumberFormatLocal = "G/標準"
        If hrkt = 16 Then '30s86_020s
            baba = "ASC(PHONETIC(R[0]C[-10]))"
        Else
            baba = "ASC(R[0]C[-10])"
        End If
        
        twt.Cells(qq, 11).FormulaR1C1 = "=" & baba  '10列目1行目に数式の素を注入
        Call copipe(twn, "高速シート_" & syutoku(), qq, 11, qq, 11, twn, "高速シート_" & syutoku(), qq + 1, 11, rrr, 11, 3) '3→FormulaR1C1
        twt.Columns("J:J").NumberFormatLocal = "@"
            
        '↓11列目を１0列目にコピー（数式→値化）
        Call betat4(twn, "高速シート_" & syutoku(), qq, 11, rrr, 11, twn, "高速シート_" & syutoku(), qq, 10, "mm", "")
        twt.Columns("K:K").ClearContents  '11列クリア
        
        'ヴ→ｳﾞ(・・ASC(PHONETIC で処理されないので。strconv24では、ヴ→ｳﾞされるので、ここまでやって初めてASC(PHONETICとstrconv24が等価になる。
        Range(Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(qq, 10), Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(rrr, 10)).Replace What:="ヴ", _
            Replacement:="ｳﾞ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False

        Application.Calculation = xlCalculationManual  '再計算再び手動に
        
        '特命条件↓　7列(旧4列)ベタ・事前ソート入れる。
        If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0) Then
            '8行目(転載列・最初数列)を7列目にベタっと転載
            Call cpp2(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er(8, bni)), rrr, Abs(er(8, bni)), twn, "高速シート_" & syutoku(), qq, 7, 0, 0, -4163)
            '8行目(転載列・最後数列)を6列目にベタっと転載
            Call cpp2(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er78(1, UBound(er78(), 2))), rrr, Abs(er78(1, UBound(er78(), 2))), twn, "高速シート_" & syutoku(), qq, 6, 0, 0, -4163)
            
            '加算列の転載(9列)　86_013d↓
            If er9(0) <> 0.1 Then
                Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er9(0)), rrr, Abs(er9(0)), twn, "高速シート_" & syutoku(), qq, 9, "pp", mr9(0))
            Else '６行ーなら、こっち(30s86_019o　運用終了)
                Call oshimai("", bfn, shn, sr(6), a, "６行ーの運用は終了しました。")
                'Call betat4(twn, "高速シート_" & syutoku(), qq, 0.4, rrr, 0.4, twn, "高速シート_" & syutoku(), qq, 9, "pp", "1")
                'MsgBox "6行ーです。"
            End If
            Call saato(twn, "高速シート_" & syutoku(), 3, qq, 1, rrr, 10, 99) '3列目降ソート　99は降順の意 1列～10範囲で3列目を、
            Call saato(twn, "高速シート_" & syutoku(), 7, qq, 1, rrr, 10, 1) '7列目(旧4列目)昇ソート
            Call saato(twn, "高速シート_" & syutoku(), 10, qq, 1, rrr, 10, 1) '昇ソート 1→10列目(Excel2019対策・平・片同一視されない)
            
            ii = rrr
            cnt = 0
            
            '3列目消込プログラム　　　'rrr→rrr+1(空白コピペ用) 86_013v
            hirt = Range(Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(qq, 1), Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(rrr + 1, 10)).Value

            Do Until ii = qq  '  30s86_012s　qq行(データ開始行)は以下の操作やらない。qq+1行までが対象
                '1列目は常に情報がある。
                If hirt(-qq + 1 + ii, 3) <> "" Then '3列目情報ありなら、以下 hirt へ
                    If hirt(-qq + 1 + ii, 7) = "" Then '3列情報有,7列空なら、3列空白(①)、8列何もしない
                        hirt(-qq + 1 + ii, 3) = hirt(rrr - qq + 2, 3) '←セル空白化
                    Else  '3列＆4列共に情報あり　8、9
                        If hirt(-qq + 1 + ii, 8) = "" And hirt(-qq + 1 + ii, 9) <> "" Then hirt(-qq + 1 + ii, 8) = hirt(-qq + 1 + ii, 9)
                            
                            '↓処理行と上の行の(1→10列＆7列)が一致ならば以下（肝）7
                        If hirt(-qq + 1 + ii, 10) = hirt(-qq + 1 + ii - 1, 10) And hirt(-qq + 1 + ii, 7) = hirt(-qq + 1 + ii - 1, 7) Then  ',1)→,10)
                            hirt(-qq + 1 + ii, 3) = hirt(rrr - qq + 2, 3) '←空白化
                            hirt(-qq + 1 + ii, 7) = hirt(rrr - qq + 2, 3) '←空白化
                            hirt(-qq + 1 + ii, 6) = hirt(rrr - qq + 2, 3) '←空白化
                            
                            If hirt(-qq + 1 + ii, 8) <> "" Or hirt(-qq + 1 + ii - 1, 9) <> "" Then '86_014g if条件追加
                                hirt(-qq + 1 + ii - 1, 8) = hirt(-qq + 1 + ii, 8) + hirt(-qq + 1 + ii - 1, 9) '8列上行＝8列当行+9列上行
                                hirt(-qq + 1 + ii, 8) = hirt(rrr - qq + 2, 3) '処理後、8列当行は空白化
                            End If
                        End If
                    End If
                ElseIf hirt(-qq + 1 + ii, 7) <> "" Then '3列情報ナシで7列ありならこちら
                    hirt(-qq + 1 + ii, 7) = hirt(rrr - qq + 2, 3) '7列空白化
                    hirt(-qq + 1 + ii, 6) = hirt(rrr - qq + 2, 3) '6列空白化
                End If
                ii = ii - 1
                Call hdrst2(rrr - ii, a, 10000, 0, 0)
            Loop  '3列目消込プログラムここまで
            '※ここでのiiはqq(開始行)
            
            Range(Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(qq, 1), Workbooks(twn).Sheets("高速シート_" & syutoku()).Cells(rrr + 1, 10)).Value = hirt
            Erase hirt
            
            Application.Cursor = xlWait
            
            '↓開始行特別処置A　8列先頭行　'↓９列nullで８列ゼロ入り込む阻止用
            If twt.Cells(qq, 3).Value <> "" And twt.Cells(qq, 7).Value <> "" And twt.Cells(qq, 8).Value = "" And twt.Cells(qq, 9).Value <> "" Then
                twt.Cells(qq, 8).Value = twt.Cells(qq, 9).Value  'qq=iiですね
            End If
            
            '↓開始行特別処置B　6列、7列先頭行　（無くてもバグらないかもしれないが）
            If twt.Cells(qq, 3).Value = "" And twt.Cells(qq, 7).Value <> "" Then
                twt.Cells(qq, 7).Value = twt.Cells(rrr + 10, 3).Value  '←空白化
                twt.Cells(qq, 6).Value = twt.Cells(rrr + 10, 3).Value  '←空白化
            End If
            
            '↓８列→3列目にコピー（数式→値化）パターンA～Cどの場合も途中段階として
            Call betat4(twn, "高速シート_" & syutoku(), qq, 8, rrr, 8, twn, "高速シート_" & syutoku(), qq, 3, "pp", "")  '8(旧5)
            Application.Cursor = xlDefault
        End If '特命条件処理（４列ベタ）ここまで
        
        cnt = 0
        rrr = rrr + 1  'rrrは最終行の次行(all1的には空白になった行)

        'ロック識別子挿入
        If h >= k Then
            twt.Cells(rrr, 1).Value = bfshn.Cells(h, Abs(er(2, bni)))  'mghz(strconv24済)の最下行
            twt.Cells(rrr, 10).Value = StrConv(bfshn.Cells(h, Abs(er(2, bni))), 8 + hrkt) 'mghzの最下行　1→10列目(Excel2019対策
            twt.Cells(rrr, 2).Value = "000_0000000"
            rrr = rrr + 1
        End If

        cnt = 0
        rrr = rrr - 1
        'この時点のrrrは高速シートのデータ終了行(含ロック因子)、qqは相変わらずデータ開始行(対象シート及び高速シート)
        
        '元来の高速シート昇順降順はこっち
        Call saato(twn, "高速シート_" & syutoku(), 3, qq, 1, rrr, 10, 99)  '3列(ｶｳﾝﾄ列)降順
            '↑これがないと、特命条件の連載が数の多い順にならない。特命条件のルーチンい移してもよいが。　20210309記
        Call saato(twn, "高速シート_" & syutoku(), 10, qq, 1, rrr, 10, 1)   '昇ソート 1→10列目(Excel2019対策・平・片同一視されない)
    End If '対象シートデータ部がすっからかんでない時の処理ここまで

    Workbooks(bfn).Activate  '86_017h　上からこちらへ（正常性確認中）
    Sheets(shn).Select
    DoEvents

End Sub
