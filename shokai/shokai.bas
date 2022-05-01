Sub 初回のみ()
    Dim hk1 As String, mghx As Long, xlvrn As String
    
    kyosydou  '共通初動
 
    xlvrn = chekku 'Excelバージョンチェック
    twt.Cells(3, 1).Value = "ｿｰﾄ後"
    twt.Cells(1, 2).Value = "あ"
    twt.Cells(2, 2).Value = "ア"
    twt.Cells(3, 2).Value = "ｿｰﾄ前"
    twt.Cells(1, 3).Value = "ア"
    twt.Cells(2, 3).Value = "あ"
    twt.Cells(3, 3).Value = "新仕様(片仮名が上へ)"
    twt.Cells(1, 4).Value = "あ"
    twt.Cells(2, 4).Value = "ア"
    twt.Cells(3, 4).Value = "旧仕様(片・平同一視)"
    twt.Cells(1, 6).Value = "↓textjoin結果"  '30s86_021e
    twt.Cells(2, 6).Value = "ｦｦTEXTJOIN(""、"",TRUE,A2,B2)"
    twt.Cells(3, 6).Value = "↑「あ、ア」なら正常"
    twt.Cells(4, 6).Value = "　「#NAME?」ならこのOfficeではサポートされていない(Excel再起動でサポートされることもある)"
    'ｦｦ→=
    Range(twt.Cells(2, 6), twt.Cells(2, 6)).Replace What:="ｦｦ", _
        Replacement:="=", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    If hrkt = 16 Then MsgBox "hrkt=16(phonetic使用[従来型])です。キーに「ヶ」があると重くなり、誤作動の可能性あり。2013でも2019でも"
    If xlvrn = "旧ｿｰﾄ" And hrkt = 0 Then MsgBox "旧ｿｰﾄ(2013)でhrkt=0(phonetic不使用)です。キーに同じひらがなカタカナある時(「あ」「ア」など)誤作動の危険あり注意。"
    
    If Not (shn = "▲集計_雛形" And bfn = twn) Then
        bfshn.Cells(sr(0) + 1, 3).Value = ""    '13→sr(0)+1
        bfshn.Cells(sr(0) + 2, 3).Value = ""     '14→sr(0)+2
    End If

    Call iechc(hk1)

    mghx = Application.Match("。", Range(twbsh.Cells(1, 1), twbsh.Cells(1, 5000)), 0) 'mghxはマクロファイルの「。」の列,mghzは当シートの、
    
    jj = 6 'j→jj 30s82
    ii = 19967 'i→ii 30s82
    Do Until jj >= mghz
        If bfshn.Cells(2, jj).Value <> "" Then
            If Not IsError(Application.Match(bfshn.Cells(2, jj).Value, Range(bfshn.Cells(2, jj + 1), bfshn.Cells(2, mghz)), 0)) Then
                Range(bfshn.Cells(2, Application.Match(bfshn.Cells(2, jj).Value, Range(bfshn.Cells(2, jj + 1), bfshn.Cells(2, mghz)), 0) + jj), bfshn.Cells(2, Application.Match(bfshn.Cells(2, jj).Value, Range(bfshn.Cells(2, jj + 1), bfshn.Cells(2, mghz)), 0) + jj)).Select
                Call oshimai("", bfn, shn, 2, Int(jj), "文字重複セルあり（" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & "セル）")
            End If
        End If
        If bfshn.Cells(2, jj).Value <> "" Then '17s
            If bfshn.Cells(3, jj).Value > ii Then ii = bfshn.Cells(3, jj).Value
        End If
        jj = jj + 1
        Application.StatusBar = "重複確認、" & Str(jj) & " / " & Str(mghz) '9s
    Loop

    '再計算を一旦自動に
    Application.Calculation = xlCalculationAutomatic
    Application.ExtendList = False 'データ範囲拡張:オフ（：右隣セルが勝手に書式変わられるのを阻止）"
    
    'オートフィルタが設定されているかどうか判断＆解除
    If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    Application.EnableAutoComplete = False  'オートコンプリート
       
    ThisWorkbook.Activate
    
    jj = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = ""
        jj = jj + 1
        If jj = 10000 Then
            MsgBox "空白行が見つからないようです"
            Exit Sub
        End If
    Loop
    
    Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = 1 '項目名
    Workbooks(twn).Sheets(shog).Cells(jj, 2).Value = jj  '項番
    
    Workbooks(twn).Sheets(shog).Cells(jj, 3).Value = _
    "初.　初回、" & _
    twn _
    & "、ｦ" & shn & "ｦ" & bfn _
    & "、ｦ▲集計_雛形" & "ｦ" & twn _
    & "、from" & dd1 & "to" & dd2 _
    & "、項目名、b" & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "、" & twbsh.Cells(2, 2).Value & "、" _
    & Format(Now(), "yyyymmdd_hhmmss") & "、" & bfshn.Cells(sr(8), 5).Value & "、0、0"
    Workbooks(twn).Sheets(shog).Cells(jj, 4).Value = Format(Now(), "yyyymmdd") 'date
    Workbooks(twn).Sheets(shog).Cells(jj, 5).Value = Format(Now(), "yyyymmdd_hhmmss") 'timestamp
    Workbooks(twn).Sheets(shog).Cells(jj, 7).Value = bfn & "\" & shn 'to
    Workbooks(twn).Sheets(shog).Cells(jj, 8).Value = 9  '最右列(複写は固定値)
    Workbooks(twn).Sheets(shog).Cells(jj, 9).Value = twn & "\▲集計_雛形"   'from
    'ログ部ここまで
    
    Call cpp2(twn, "▲集計_雛形", 14, 1, 18, 5, bfn, shn, sr(0) - 1 + 5 - 2, 1, 0, 0, -4104) 'サマリ関数周辺丸ごとコピペなので４１０４
    Call cpp2(twn, "▲集計_雛形", 15, mghx, 18, mghx + 2, bfn, shn, sr(0) - 1 + 5 - 1, mghz, 0, 0, -4104) '同、サマリ関数周辺丸ごとコピペ(mghz側)
    Call cpp2(twn, "▲集計_雛形", 11, 5, 15, 6, bfn, shn, sr(0) - 1, 5, 0, 0, -4122)  '同、サマリ関数周辺丸ごとコピペ(mghz側)
    
    Workbooks(bfn).Activate '↑の.copy後、これをここに入れると、セルが複数個所選択されている妙な映りは解消されるっぽい
    Sheets(shn).Select
    
    Range(bfshn.Cells(2, mghz - 2), bfshn.Cells(21, mghz)).Borders.Color = RGB(191, 191, 191)  '←　=-4210753　 30s86_017p　'30s86_012i
    
    '再計算を手動に
    Application.Calculation = xlCalculationManual

    With Application.AutoCorrect      'オートコレクトさせない　３０ｓ５２
        .TwoInitialCapitals = False
        .CorrectSentenceCap = False
        .CapitalizeNamesOfDays = False
        .CorrectCapsLock = False
        .ReplaceText = False
        .DisplayAutoCorrectOptions = True
    End With

    bfshn.Cells(sr(0) - 1, 5).Select   '緑色セル
    Range(bfshn.Cells(gg1, dd1), bfshn.Cells(gg2, dd2)).Select '選択範囲は戻す bfshn被せた
    Call oshimai(syutoku(), bfn, shn, 1, 0, "初回処理完了")
End Sub
