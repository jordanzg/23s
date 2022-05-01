Sub 外部結合()   '一番上↑にパブリック変数あり。見逃し注意
    Dim f As String, xsheet As Worksheet, xbook As Workbook, wsfag As Boolean
    Dim bun As Long, bni As Long
    Dim mr() As String, er() As Currency
    Dim nkg As Long, kahi As Long, cted(1) As Long, rrr As Long 'ppp→rrr 86_014r
    Dim pap() As Long   'pap配列変数化(86_019d)
    Dim er2() As Currency, er3() As Currency, er5() As Long, er78() As Currency, er9() As Currency, er34 As String
    Dim mr2() As String, mr3() As String, mr5() As String, mr8() As String, mr9() As String
    Dim a As Long, pkt As Long, nn As Long, n1 As Long, qq As Long, pqp As Long
    Dim am1 As String, am2 As String, h As Long, m As Long, k0 As Long, h0 As Long, n2 As Long, kg1 As String
    Dim qap As Long, ii As Long, jj As Long, trt As Long, tst As Long, dif As Long, axa As Long
    Dim saemp '←今も型が設定されていない
    Dim hirt As Variant, hiru As Variant, tameshi As Range, ct8 As String, tempo As String, baba As String
    Dim kasan As Variant, c5 As Long, c99 As String, c98 As String, ct3 As String, hk1 As String, rog As String
    Dim zzz As String, zyyz() As String, xxxx() As String, zxxz() As String, nifuku As Long
  
    '再計算を一旦自動に
    Application.Calculation = xlCalculationAutomatic
    'おまじない(「コードの実行が中断されました」対処)
    Application.EnableCancelKey = xlDisabled
    UserForm1.StartUpPosition = 1 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
    UserForm1.Show vbModeless
    UserForm1.Repaint

    kyosydou  '共通の初動
 
    bfshn.Cells(sr(3), 5).Value = chekku     'Excelソート仕様表記　30s86_020t　Excelバージョンチェック
    bfshn.Cells(sr(5), 5).Value = "hrkt_" & hrkt     'Excelソート仕様表記　30s86_020t　Excelバージョンチェック
    
    If bfshn.Cells(sr(3), 5).Value = "旧ｿｰﾄ" And hrkt = 0 Then MsgBox "旧ｿｰﾄ(2013)でhrkt=0(phonetic不使用)です。キーに同じひらがなカタカナある時(「あ」「ア」など)誤作動の危険あり注意。"
    
    twt.Cells.Clear

    Workbooks(bfn).Activate
    bfshn.Select
        
    If dd1 = 0 Then Call oshimai("", bfn, shn, 1, 0, "dd1がゼロです")

    'その日の初回チェック
    ii = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(ii, 1).Value = ""
        ii = ii + 1
        If ii = 50000 Then
            MsgBox "空白行が見つからないようです"
            Exit Sub
        End If
    Loop
    
    If Workbooks(twn).Sheets(shog).Cells(ii - 1, 4).Value <> Val(Format(Now(), "yyyymmdd")) Then
        Call oshimai("", bfn, shn, 1, 0, "その日の初回は、最初に[FIRST]ボタンを押して下さい。")
    End If
    
    If bfshn.Cells(sr(0) - 1, 5) = "" Then
        Call oshimai("", bfn, shn, sr(0) - 1, 5, "★集計名を入力して下さい。")
    End If

    Call iechc(hk1) '旧igchc(hk1)
    hk1 = ""
    flag = False

    ii = 1
    '当シート側のall1探し
    Do Until bfshn.Cells(ii, 1).Value = "all1"
        If IsNumeric(bfshn.Cells(ii, 1).Value) And bfshn.Cells(ii, 1).Value <> "" Then '17s
            Call oshimai("", bfn, shn, ii, 1, "一列目は数値を入れないで下さい")
        End If
        ii = ii + 1
        If ii = 100 Then
            Call oshimai("", bfn, shn, 1, 0, "当シート「all1」が見つからないようです")
        End If
    Loop

    k = ii + 1     'k確定'kはデータ開始行(サンプル行ではなくなった。)
    bfshn.Cells(1, 2).Value = k     'データ開始行

    'この時点でのiiは、当シートの「all1」記載行
    Do Until bfshn.Cells(ii, 1).Value = ""
        ii = ii + 1
    Loop
    'ここでのiiは当シートall1列の空白になった行、データ無しの場合はデータ開始行
    
    'オートフィルタが設定されてれば、解除
    If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    nn = 0 'セル空白チェックフラグ
    If dd1 <= 5 Then Call oshimai("", bfn, shn, 1, 0, "タテヨコ終了、6列目以降が対象です")    'よ
    If dd2 >= mghz Then Call oshimai("", bfn, shn, 1, 0, "枠外が選択されてます")    '86_108i

    
For a = dd1 To dd2 '選択範囲列分の繰り返し　ら
    
    Application.Goto bfshn.Cells(k, a - 3), True
    bfshn.Cells(k, a).Select
    
    kg1 = "" 'リセット　kg2はmr(2,3,bni)のredimでリセットされる。
    bni = 1  'リセット
    bun = 1  'リセット
    nkg = 0 'リセット
    nifuku = 0
    
    '複文節用区切り文字の確定　fnywti→rvsrz3を流用
    kg1 = Mid(rvsrz3(bfshn.Cells(sr(7), a).Value, 3, "ｦ", 0), 1, 1)  'kg2定義は先 30s73:nkg2→0

    If bfshn.Cells(sr(6), a).Value <= -90 Then
        bun = 1 '-90台は強制1
    ElseIf kg1 = "" Then
        nkg = 1 'ゐ区切り無指定(kg1="")は文節ゼロ（区切りしない）
    Else '複文節数(bun)確定のためのルーチン(kg1<>"")
        For ii = 1 To 8
            Do Until rvsrz3(bfshn.Cells(sr(ii), a).Value, bni, kg1, 0) = ""
                bni = bni + 1
                If bni > 120 Then Call oshimai("", bfn, shn, 4, 2, "bni120超え")
            Loop
            bni = bni - 1
            If ii = 2 And bni > 1 Then
                nifuku = 1
                MsgBox "2行目での複文節あり。注意を。"  '解禁へ　86_016v
            End If
            If bun < bni Then bun = bni '（この時点でbun確定）
            bni = 1
        Next
    End If

'◆Ａ◆単文節
    bni = 1 'リセット
    ReDim pap(9, bun) 'As long  30s86_019c
    ReDim er(11, bun) 'As Currency 10→11 30s83
    ReDim mr(4, 11, bun) 'As String 30s81第３因子化 10→11 30s83 ,mr(3→mr(4：ｦ6ｦ→ｦFｦ 用
    
    '再計算を手動に
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    For ii = 1 To 7 'mr(0,記述　　ii=2からでも大丈夫と思われる。
        mr(0, ii, 1) = rvsrz3(bfshn.Cells(sr(ii), a).Value, 1, kg1, nkg)  'n行目の全因子　bni:1
    Next
    
    '1文節目特注
    mr(1, 1, 1) = rvsrz3(mr(0, 1, 1), 2, "ｦ", 2)  'シート名
    mr(2, 1, 1) = rvsrz3(mr(0, 1, 1), 3, "ｦ", 0)  'ファイル名
    
    mr(2, 7, 1) = kg1 'ゐ　７行目第二因子
    If StrConv(Left(mr(1, 1, 1), 1), 8) = "*" Then 'り　30s57左一文字が＊だけの時（＊～の時は～がmr(2,0,1)に入る）
        bfshn.Cells(sr(0), a).Value = bfshn.Cells(sr(0) + 3, a).Value
        bfshn.Cells(sr(0) + 1, a).Value = bfshn.Cells(sr(0) + 4, a).Value
        bfshn.Cells(sr(0) + 2, a).Value = bfshn.Cells(sr(0) + 5, a).Value
    Else  'り(*だけでない通常時、外結最後部まで続く)
    
    For ii = 1 To 7  '空白確認（bni=1）
        If bfshn.Cells(sr(ii), a).Value = "" Then nn = sr(ii)
    Next
    
    If bfshn.Cells(sr(6), a).Value > -90 Then     'ここでのiiは8 -99は8行目確認しない
        If bfshn.Cells(sr(8), a).Value = "" Then nn = sr(8) 'ii→8　（同値）
    End If
    If nn > 0 Then Call oshimai("", bfn, shn, nn, a, "外結設定情報が空欄の所があります")
'◇Ａ◇単文節ここまで↑

'◆Ｂ◆複文節for（：準備編）ふあ↓　文節毎のforがここから始まる
    For bni = 1 To bun
        nn = 0 'セル空白チェックフラグ→？
        mr(2, 0, bni) = bfn '30s82
        mr(1, 0, bni) = shn  '30s82
        For ii = 1 To 7
            mr(0, ii, bni) = rvsrz3(bfshn.Cells(sr(ii), a).Value, bni, kg1, nkg)  'n行目の全因子
            If mr(0, ii, bni) = "" Then mr(0, ii, bni) = mr(0, ii, bni - 1)  '2文節目以降、空欄なら前節コピペ
        Next
        
        If bfshn.Cells(sr(6), a).Value <= -90 Then
            mr(1, 1, bni) = shn  'シート名　'-90台はセル見ず、bfshn強制　セルは日付書式など自由に書けられる。
            mr(2, 1, bni) = bfn  'ファイル名
        Else
            mr(1, 1, bni) = rvsrz3(mr(0, 1, bni), 2, "ｦ", 2)  'シート名 ""にはならない
            mr(2, 1, bni) = rvsrz3(mr(0, 1, bni), 3, "ｦ", 0)  'ファイル名30s73:nkg2→0
        End If
        If Left(StrConv(mr(1, 1, bni), 8), 1) = "\" Then mr(1, 1, bni) = shn '￥→シート名変換
        If mr(1, 1, bni) = "" Then Call oshimai("", bfn, shn, sr(1), a, "対象シート名(" & bni & "文節目)が空欄です")
        If mr(2, 1, bni) = "" Then mr(2, 1, bni) = bfn
        mr(4, 1, bni) = mr(1, 1, bni)    '86_014s これ、抜けてた。
        wsfag = False   'ファイル・シート有無chk
        Do Until wsfag = True
            For Each xbook In Workbooks
                If xbook.Name = mr(2, 1, bni) Then wsfag = True
            Next xbook
            If wsfag = False Then
'                If MsgBox("ファイルが確認できませんが、" & vbCrLf & "続行しますか？", 289, "ファイル不明") = vbCancel Then 'キャンセル時
                    '↑この手法（Yes,No選択方式）はここでは意味なかったので、このやり方は中止。↓こちらへ
                    Call oshimai("", bfn, shn, sr(1), a, "実施中止しました。" & vbCrLf & "ファイルが確認できません")
'                end if
            End If
        Loop
    
        wsfag = False
        For Each xsheet In Workbooks(mr(2, 1, bni)).Sheets
            If xsheet.Name = mr(1, 1, bni) Then wsfag = True
        Next xsheet
        If wsfag = False Then Call oshimai("", bfn, shn, sr(1), a, mr(1, 1, bni) & " のシートが不明です(" & bni & "文節目)")
    
        nn = 0 '一旦リセット
            
            For ii = 2 To 6  '(ヱ)はここで定まる。
                mr(2, ii, bni) = rvsrz3(mr(0, ii, bni), 3, "ｦ", 0)
            Next
            mr(3, 4, bni) = rvsrz3(mr(0, 4, bni), 4, "ｦ", 0) '第3因子（工事中）
            
            If bni >= 2 Then mr(2, 7, bni) = mr(2, 7, bni - 1) '戻った(ゐ)強制前節コピー

            'mr(2, 11, bni)・・・項準とかが入る。↓er(11,bni)はリファラ変数
            mr(2, 11, bni) = koudicd(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), mr(0, 4, bni))
            
            mr(1, 2, bni) = yhwat1(bfn, shn, 2, sr(2), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 2, bni))
            mr(4, 2, bni) = zhwat1(bfn, shn, 2, sr(2), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 2, bni))

            For ii = 3 To 6
                mr(1, ii, bni) = yhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(ii), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, ii, bni))
                mr(4, ii, bni) = zhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(ii), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, ii, bni))
            Next
            
            mr(1, 7, bni) = yhwat1(bfn, shn, 2, sr(7), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 7, bni))
            mr(4, 7, bni) = zhwat1(bfn, shn, 2, sr(7), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 7, bni))
       
            For ii = 2 To 7
                er(ii, bni) = Val(mr(1, ii, bni)) 'ヱ対応時：val関数は数字と認識出来る所までを数値変換する。要は第一列目を。
            Next
    
        If nn > 0 Then Call oshimai("", bfn, shn, nn, a, "（処理中止）" & vbCrLf & "数値以外の情報があります。")
        If er(5, 1) >= 0 And er(7, 1) < 0 And er(6, 1) > -90 Then Call oshimai("", bfn, shn, 1, 0, "（処理中止：文法エラー）" & vbCrLf & "er(5,0)>=0　で　er(7,0)<0　です。")  'bni→1
    
        If er(6, bni) > -90 Then '8番目（転載列）の処理
            c98 = bfshn.Cells(sr(8), a).Value  '26ｓ
            mr(0, 8, bni) = rvsrz3(bfshn.Cells(sr(8), a).Value, bni, kg1, nkg)
            mr(2, 8, bni) = rvsrz3(mr(0, 8, bni), 3, "ｦ", 0) ' '30s73:nkg2→0
            mr(1, 8, bni) = yhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2有_30s56
            mr(4, 8, bni) = zhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2有_30s56
        
            If bni >= 2 Then '2文節目以降からというのがミソ、1文節目は下記
                If mr(0, 8, bni) = "" Then mr(0, 8, bni) = mr(0, 8, bni - 1)
                If mr(2, 8, bni) = "" Then mr(2, 8, bni) = mr(2, 8, bni - 1) 'mr(2,8,0)誤動作阻止 2文節目以降、空欄なら前節コピペ
                mr(1, 8, bni) = yhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2有_30s56
                mr(4, 8, bni) = zhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2有_30s56
            End If
            er(8, bni) = Val(mr(1, 8, bni))
            pap(8, 0) = kgcnt(mr(1, 8, bni), mr(2, 4, bni)) '30s86_019  papx→pap(x,0) 以下同じ ,0)は既存改築分(いじらない,papzouseiしない)
            
'            pap(8, bni) = kgcnt(mr(1, 8, bni), mr(2, 4, bni)) '86_019e  bni:1～　,bni)は新規造成分
            '↓等価
            Call papzousei(pap(), mr(), 8, bni) '86_019e  bni:1～　,bni)はpapzouseiは新規造成分が対象(既存改築分はやらない)
            
            If Round(er(6, bni)) <> -2 And er(8, bni) < 0 Then   '複数列解禁86_013f
                If pap(8, 0) = 0 Then
                    If mr(2, 8, bni) = "" Then mr(2, 8, bni) = "、" '区切り文字デフォは「、」（-15でも適用）
                Else
                    If rvsrz3(mr(2, 8, bni), 0 + 1, mr(2, 4, bni), 0) = "" Then mr(2, 8, bni) = "、" & mr(2, 8, bni) '30s86_018c
                End If
            End If
        
            If er(6, bni) <= -3 And er(8, bni) < 0 Then Call oshimai("", bfn, shn, 1, 0, "「c=-3以下は重複連なり型(e<0)は実行できないです。")
            
            '30s86_012s追加↓
            If er(5, bni) < 0 And er(7, bni) = 0 Then Call oshimai("", bfn, shn, sr(7), a, "差分時7行0は実施されなくなりました。")
            
            '30s82d追加↓                      er(5,1)→er(5,bni) 86_010
            If (er(6, 1) = -1 Or er(6, 1) = -2) And er(5, bni) >= 0 And kg1 <> "" And rvsrz3(bfshn.Cells(sr(7), a).Value, 2, kg1, 0) <> "" Then Call oshimai("", bfn, shn, nn, a, "6行-1-2の時でnot差分時(通常時)は7行複文節不可です。")
 
            '↓順変更　30s86_012w
            If Round(er(6, bni)) = -2 Then
                mr(0, 9, bni) = mr(0, 8, bni)
                mr(2, 9, bni) = mr(2, 8, bni)
                mr(1, 9, bni) = mr(1, 8, bni)
                er(9, bni) = Val(mr(1, 9, bni))
            ElseIf er(6, bni) > 0 Then
                mr(0, 9, bni) = mr(0, 6, bni)
                mr(2, 9, bni) = mr(2, 6, bni)
                mr(1, 9, bni) = mr(1, 6, bni)
                er(9, bni) = Val(mr(1, 9, bni))
            End If

            If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0) Then
                'MsgBox "社数抽出該当" '←特命条件追加30s86_012w
                bfshn.Cells(sr(6), 5).Value = "特命" '←20191118　こちらへ　86_015c
            Else
                bfshn.Cells(sr(6), 5).Value = ""
            End If
          
            If Not (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) > 0) Then '特命条件変更30s86_012w、追加30s86_012s
            '   6行ー２でない、　　　　　　かつ　８行０でない、　　　かつ　（５行＋ あるいは　６行０以下）　ならば
                If (Not Round(er(6, bni)) = -2) And (Not er(8, bni) = 0) And (er(5, bni) >= 0 Or er(6, bni) <= 0) Then '新仕様(30s83)
                '10行目処理(er(10,0)は実質の当シート転載列) (特命条件の連載型は実施する30s86_012w)
                    mr(0, 10, bni) = mr(0, 8, bni)
                    mr(2, 10, bni) = mr(2, 8, bni)
                    mr(1, 10, bni) = mr(1, 8, bni)
                    er(10, bni) = Val(mr(1, 10, bni))
                    If pap(8, 0) = 0 Then
                        If mr(2, 8, bni) <> "" And er(10, bni) = 0.1 Then
                            er(10, bni) = -0.1
                            er(8, bni) = -0.1
                        End If
                    Else
                        If rvsrz3(mr(2, 8, bni), 0 + 1, mr(2, 4, bni), 0) <> "" And er(10, bni) = 0.1 Then
                            er(10, bni) = -0.1
                            er(8, bni) = -0.1
                        End If
                    End If
                End If
            End If '特命条件追加30s86_012s
        End If

        nn = 0 '一旦リセット

        'log処理部
        ii = 1
        If StrConv(Left(mr(0, 1, bni), 1), 8) <> "*" And bfn <> twn Then '*接頭辞はログ生成対象外
            Do Until ii = 0
                If rvsrz3(rog, ii, "ヱ", 0) = "" Then
                    rog = mr(2, 1, bni) & "\" & mr(1, 1, bni) & "ヱ" & rog  'rogはその外結実施範囲での対象シートファイルの集合体(ヱで連結)
                    ii = 0
                ElseIf rvsrz3(rog, ii, "ヱ", 0) = mr(2, 1, bni) & "\" & mr(1, 1, bni) Then
                    ii = 0
                Else
                    ii = ii + 1
                End If
                If ii = 200 Then Call oshimai("", bfn, shn, k, a, "うまくいってない。")
            Loop
        End If

        If er(2, bni) > 0 And er(9, bni) > 0 And er(10, bni) <> 0 And er(7, bni) = 0 And er(5, bni) >= 0 Then '↓特命条件考慮
            Call oshimai("", bfn, shn, 1, 0, "処理中止。c加算ありで同列転載しようとしています。確認を。")
        End If
        Call papzousei(pap(), mr(), 2, bni) '86_019e　papzouseiは、新規区画分(bni:1～)を造成
    Next
'◇Ｂ◇複文節ふあ（準備編）ここまで↑

'◆Ｃ◆単文節（本番プレ）ここから↓
    bni = 1 '1文節目で判断、実施部分'うあ

    '前回データコピー
    If StrConv(Left(bfshn.Cells(sr(1), a).Value, 2), 8) <> "**" Then '**はやらないへ024_検証2
        bfshn.Cells(sr(0), a).Value = bfshn.Cells(sr(0) + 3, a).Value
        bfshn.Cells(sr(0) + 1, a).Value = bfshn.Cells(sr(0) + 4, a).Value
        bfshn.Cells(sr(0) + 2, a).Value = bfshn.Cells(sr(0) + 5, a).Value
    End If
    h = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1 + k - 2 '現状の最下行(以下伸びていく)※データ無しの時は項目行となり、h=k-1となるので注意
    
    '体裁(teisai)ここから
    trt = 0 'リセット trt・・・タイプの定義　trt:セルのタイプ
    'trt：0初期ﾘｾｯﾄ,-9・-99型,-2当列加算型,-1当列転載型,1強制文字列型
    'trt：-9・-99型 ,-1当列転載型、-2当列加算型(含特命加算型)、1強制文字列型(含特命連載型)
    'trt定義部
    If er(6, bni) <= -90 Then '-99はこちら
        trt = -9
        '↓特命条件（加算or連載型）
    ElseIf er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0 Then
        If er(8, bni) > 0 Then
            trt = -2 '特命加算型
        Else
            trt = 1 '特命連載型
        End If
    ElseIf er(6, bni) > 0 Or Round(er(6, bni)) = -2 Then
        trt = -2  '当列加算型 差分時も
    ElseIf er(5, bni) >= 0 Then
        If er(10, bni) < 0 Then
            trt = 1   '強制文字列型
            'MsgBox "強制文字列です。"
        Else
            trt = -1    '当列転載型
            'MsgBox "連載じゃないです。"
        End If
    ElseIf er(5, bni) < 0 Then  '差分文字比較時
        If mr(2, 6, bni) = "1" Then
            Call oshimai("", bfn, shn, sr(6), a, "差分時の6行op1の運用は終わっている")
        ElseIf mr(2, 6, bni) = "-1" Then
            trt = -1 '差分を文字(ﾊﾞﾘｭｰ)で表現 →転載型
        Else
            trt = -1 '差分を数(0,1)で表現、-2→-1へ(最初文字、あとで数値のため)
        End If
    End If
    If trt = 0 Then Call oshimai("", bfn, shn, 1, 0, " trt = 0 ？")  'trtゼロ状態で無いことの確認念のため
        
    'tst定義部ここから
    tst = -1 'リセット　tst:セルの型
    'tst : -1初期ﾘｾｯﾄ , ｰ2通貨,   0標準,   1文字列,　 7一括踏襲型 , 8ちまちま
        
    If trt = -9 Then
        If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then
            tst = 1     '文字列型
        ElseIf (er(4, bni) < 0 Or er(4, bni) = 0.1) Then
            tst = -2     '通貨型
        ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then '対象シート名を模倣　a11_1列→対象シートに変更
            tst = 7 '一括踏襲型(模倣型)
        Else
            tst = 0 '標準型
        End If
    ElseIf trt = -2 Then
        If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  'ーー
            tst = 0     '標準型
        ElseIf (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  '＋ー
            tst = 7     '一括型
        ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then  'ー＋
                        '何もしない(tst = -1)
                        '↓こちらへ転向 30s86_014f
            tst = 8     '一括型
            MsgBox "ちまちま(tst = 8)"
        Else                                              '＋＋
            tst = -2    '通貨型
        End If
    ElseIf trt = 1 Then '強制文字列
        tst = 1 '文字列型
    ElseIf trt = -1 Then
        If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  'ーー
            tst = 1     '文字列型
        ElseIf (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  '＋ー
            tst = 7     '一括型
        ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then  'ー＋
            tst = 8     'ちまちまセル踏襲型
            MsgBox "ちまちま(tst = 8)"
        Else                                              '＋＋
            tst = 0    '標準型
        End If
    End If
    
    If tst = -1 Then Call oshimai("", bfn, shn, 1, 0, " tst = -1 ？")
    
    If h >= k And StrConv(Left(bfshn.Cells(sr(1), a).Value, 2), 8) <> "**" Then '既存行無い時はこちら実行しないへ,「＊＊あ」も実行しない
        '選択列の既存データクリア(事前既存行)
        With Range(bfshn.Cells(k, a), bfshn.Cells(h, a))
            .ClearContents
            .ClearComments '30s79コメントもクリア
            .Interior.Pattern = xlNone '30s79
            .NumberFormatLocal = "G/標準"  '30s83デフォルトでまず
        End With


        '当列事前の体裁一律化 86_013d trt,tst化
        If trt = -9 Then '-99はこちら   86_020r記→　◇E◇で同じこと実施するようになったので、こちらは不要かも
            If tst = 1 Then '←当列転載文字型
                Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "@"
            ElseIf tst = -2 Then  '通貨型
                Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[赤]-#,##0"
            ElseIf tst = 7 Then '対象シート名を模倣　a11_1列→対象シートに変更
                Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = bfshn.Cells(sr(1), a).NumberFormatLocal 'sr(4)→sr(1)へ30s68
            End If
        '新仕様-99以外 文字列指名or連載　er(7, bni) > 0は-1の複数列対応
        ElseIf tst = 1 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "@"
        ElseIf tst = -2 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        ElseIf tst = 0 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "G/標準"
        End If
        '体裁(teisai)ここまで 86_012g
        
        If er(2, 1) < 0 Or kgcnt(mr(1, 2, 1), mr(2, 4, bni)) > 0 Then  '30s45右端クリアへ
            With Range(bfshn.Cells(k, mghz), bfshn.Cells(h, mghz + 2)) '3目列目まで削除へ30s80
                .ClearContents
                .NumberFormatLocal = "@" '文字列
            End With
            With Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(h, mghz + 1))  '2列目は数値用
                .ClearContents
                .NumberFormatLocal = "G/標準"
            End With
        End If
    End If
    
    '4列目に要素転記(-90台も実施、８行目別途) ※文字列変換済み
    For ii = 1 To 7
        hk1 = "" '第3因子対応
        If StrConv(Left(mr(0, ii, bni), 1), 8) = "*" Then  '30s75（*有無も追加）
            bfshn.Cells(sr(ii), 4).Value = "*ｦ" & mr(4, ii, bni) & "ｦ" & mr(2, ii, bni) & hk1
        Else
            bfshn.Cells(sr(ii), 4).Value = "ｦ" & mr(4, ii, bni) & "ｦ" & mr(2, ii, bni) & hk1
        End If
    Next

    '8行内容を4列目に要素転記　※文字列変換済み
    If mr(2, 8, bni) = Chr(Val("&H" & "0A")) Then '行の膨らみ阻止16s
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then
            bfshn.Cells(sr(8), 4).Value = "*ｦ" & mr(4, 8, bni) & "ｦ(LF)"
        Else
            bfshn.Cells(sr(8), 4).Value = "ｦ" & mr(4, 8, bni) & "ｦ(LF)"
        End If
    Else
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then
            bfshn.Cells(sr(8), 4).Value = "*ｦ" & mr(4, 8, bni) & "ｦ" & mr(2, 8, bni)
        Else
            bfshn.Cells(sr(8), 4).Value = "ｦ" & mr(4, 8, bni) & "ｦ" & mr(2, 8, bni)
        End If
    End If
    
    If er(6, bni) < -90 Then
        '85_024検証6　-90台で５行目ヱ対応用
        pap(5, 0) = kgcnt(mr(1, 5, bni), mr(2, 4, bni)) '5行目ヱの数 30s48
        ReDim er5(pap(5, 0))
        ReDim mr5(pap(5, 0))
        er5(0) = Val(mr(1, 5, bni)) 'ヱでない時向け　erx()は通貨型なのでvalを被せざるを得ない。
        mr5(0) = mr(2, 5, bni)
        If pap(5, 0) > 0 Then 'mr(2, 4, bni) <> ""
            For ii = 0 To pap(5, 0)
                er5(ii) = Val(rvsrz3(mr(1, 5, bni), ii + 1, mr(2, 4, bni), 0))
                mr5(ii) = rvsrz3(mr(2, 5, bni), ii + 1, mr(2, 4, bni), 0)
            Next
        End If
    End If
'～～～～～
   
  bfshn.Cells(sr(0), 4).Value = bni & "/" & bun & "文節"   '30s86_012新設
  bfshn.Cells(sr(0) + 1, 4).Value = rvsrz3(bfshn.Cells(1, a).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0) & "列"
   
'◇Ｃ◇単文節（本番プレ）ここまで↑
 
 'るA 以降-99or*,**はここ通過
  If er(6, bni) > -90 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then

'◆Ｄ◆複文節for（：本番実行編）ふい↓　文節毎のforがここから始まる ループ部
    For bni = 1 To bun    '上に移行へ　86_016w
    
    '5列目より左転載防止 86_012_m
    If er(7, bni) >= 1 And er(7, bni) <= 5 And er(5, bni) >= 0 Then
        Call oshimai("", bfn, shn, 1, 0, "5列目より左に転載しないで下さい。")
    End If
    
    'ループ前の初期値
    n1 = k  'n1：前回突合処理対象行_当シートでの、kはデータ開始行（固定）m2→n1
    pap(2, 0) = kgcnt(mr(1, 2, bni), mr(2, 4, bni)) '86_016w mr(1, 2, 1)→mr(1, 2, bni)
    ReDim er2(pap(2, 0))
    ReDim mr2(pap(2, 0))
    
    er2(0) = Val(mr(1, 2, bni)) 'mr2活性化(30s86_017a)　mr(1, 2, 1)→mr(1, 2, bni)
    mr2(0) = mr(2, 2, bni)
    If pap(2, 0) > 0 Then
        For ii = 0 To pap(2, 0)
            er2(ii) = Val(rvsrz3(mr(1, 2, bni), ii + 1, mr(2, 4, bni), 0))
            mr2(ii) = rvsrz3(mr(2, 2, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
        
    '下側からこちらへ　86_020c
    pap(3, 0) = kgcnt(mr(1, 3, bni), mr(2, 4, bni)) '3行目ヱの数
    ReDim er3(pap(3, 0))
    ReDim mr3(pap(3, 0))
    er3(0) = Val(mr(1, 3, bni)) 'ヱでない時向け　erx()は通貨型なのでvalを被せざるを得ない(以下同)。
    mr3(0) = mr(2, 3, bni)
    If pap(3, 0) > 0 Then
        For ii = 0 To pap(3, 0)
            er3(ii) = Val(rvsrz3(mr(1, 3, bni), ii + 1, mr(2, 4, bni), 0))
            mr3(ii) = rvsrz3(mr(2, 3, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
        
    '近似高速種別確認　619 当面単文節で(∵近似可否er2(pap(2,0))のため。)
    mr(1, 11, bni) = "" '一旦リセット 純高速/近似高速/ノーマル　が入る。

    If er(2, 1) < 0 Then '　bni→1に書換え(ここの条件分岐全体を)　86_016w
        If er(6, 1) < 0 Then mr(1, 11, 1) = "近似高速" Else mr(1, 11, 1) = "純高速"
    Else
        mr(1, 11, 1) = "ノーマル"
    End If 'bni:2以降はmr(1,11,bni)nullなので注意
    
    bfshn.Cells(sr(2), 5).Value = mr(1, 11, 1)     'ノーマル・純・近似表記５列　bni→1
    'mghz　４行目に区切り文字を入れる mghzの要素生成で使用
    With bfshn.Cells(sr(4), mghz)
        .NumberFormatLocal = "@"  '文字列として入れる
        .Value = mr(2, 4, bni)
    End With
    
    If (er(2, 1) < 0 Or pap(2, 0) > 0) Then   '2行目にヱがある時か高速時実施(1文節目で判断)、1文節目ヱある時→全文節ヱ適用される。
        If k <= h Then '当シート既存情報あり
            If StrConv(bfshn.Cells(sr(1), a), 8) = "\" And er(6, 1) >= 0 And mr(1, 2, bni) <> mr(1, 3, bni) And pap(2, 0) > 0 Then
                    'Call oshimai("", bfn, shn, sr(2), a, "\の時で仮想キー使用＆6行>0のときは2行3行一致が必要です。") '無限ループ防止s
                    'MsgBox "2行3行不一致(無限ループの可能性有)" '86_016r　86_021b解除、無限ループするなら戻す
            End If
                
            'mghz2列の情報埋め込み（値と数値の書式で貼付)
            If er(2, 1) < 0 Then '高速（単数列も複数列も）
                    
                bfshn.Cells(k, mghz).Select '動作可視化

                DoEvents    '86_019r
                Application.Calculation = xlCalculationAutomatic    '数式計算方法自動に　'新形式　85_007
                    
                tempo = "R[0]C[" & LTrim(Str(Abs(er2(0)) - mghz)) & "]"
                baba = tempo
                    
                For ii = 1 To pap(2, 0) '数式の素生成
                    tempo = "&" & "R" & LTrim(Str(Abs(sr(4)))) & "C[0]" & "&" & "R[0]C[" & LTrim(Str(Abs(er2(ii)) - mghz)) & "]"
                    baba = baba & tempo
                Next

                Range(bfshn.Cells(k, mghz), bfshn.Cells(h, mghz)).NumberFormatLocal = "G/標準"    'mghz列一旦標準へ（数式入れるため）
                    
                bfshn.Cells(sr(8), mghz).FormulaR1C1 = "=" & baba & "&" & """" & """"     '←数式右端に「&""」を付加(空白セルが「0」になる対策)　30s86_020s
                    
                '↓関数貼り付け化　86_020g　8行目→当シートに
                Call copipe(bfn, shn, sr(8), mghz, sr(8), mghz, bfn, shn, k, mghz, h, mghz, 3)  '3→FormulaR1C1
                    
                '値に変換
                Call copipe(bfn, shn, k, mghz, h, mghz, bfn, shn, k, mghz, h, mghz, 1)  '1→Valueコピペ　旧cpp2、-4163
                '30s86_020s
                If hrkt = 16 Then baba = "ASC(PHONETIC(R[0]C[-1]))" Else baba = "ASC(R[0]C[-1])"
                    
                bfshn.Cells(k, mghz + 1).FormulaR1C1 = "=" & baba
                '↓関数貼り付け化　86_020g
                Call copipe(bfn, shn, k, mghz + 1, k, mghz + 1, bfn, shn, k + 1, mghz + 1, h, mghz + 1, 3) '3→FormulaR1C1
                'mghz+1→mghz、値化
                Call copipe(bfn, shn, k, mghz + 1, h, mghz + 1, bfn, shn, k, mghz, h, mghz, 1)    '1→Value
                    
                'mghz+1列(行番号)の処理(フィル活用)　※ここにあるのはmghz+1の数式をさっさと消すため
                bfshn.Cells(k, mghz + 1).Value = k
                If h > k Then
                    Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(k, mghz + 1)).AutoFill Destination:=Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(h, mghz + 1)), Type:=xlFillSeries '←連続する数値
                End If
                    
                'ヴ→ｳﾞ(・・ASC で処理されないので。strconv24では、ヴ→ｳﾞされるので、ここまでやって初めてASCとstrconv8が等価になる。
                Range(Workbooks(bfn).Sheets(shn).Cells(k, mghz), Workbooks(bfn).Sheets(shn).Cells(h, mghz)).Replace What:="ヴ", _
                    Replacement:="ｳﾞ", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False
                    
                Range(bfshn.Cells(k, mghz), bfshn.Cells(h, mghz)).NumberFormatLocal = "@"  'mghz列一旦文字列へ戻す（数式入れ、値化したので）
                
                DoEvents
                Application.Calculation = xlCalculationManual  '再計算再び手動に（重くなるため）30s66
                     
                If pap(2, 0) > 0 Then '30s86_020c　020q高速専用として復帰　２行複数列で下側がall空白chk 30s86_020a
                    ii = h
                    Do Until bfshn.Cells(ii, mghz) <> StrConv(wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3), 8 + hrkt) '←使ってるのmr(2,4,bni)だけ
                        ii = ii - 1
                    Loop
                    If ii < h Then
                        MsgBox "mghz下側に「" & StrConv(wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3), 8 + hrkt) & "」発生したので、"
                        MsgBox "mghzクリア入ります挙動注意_(2行複数列)" & ii + 1 & "～" & h
                        Range(bfshn.Cells(ii + 1, mghz), bfshn.Cells(h, mghz)).ClearContents
                    Else
                        'MsgBox "mghzクリアはありませんでした(2行複数列)"
                    End If
                End If
            Else '低速（単数列も複数列も）
                bfshn.Cells(sr(8), mghz).Value = "(不使用(ちまちまpap(2,0)≠0)"
                For ii = k To h
                    bfshn.Cells(ii, mghz).Value = wetaiou(bfn, shn, ii, er2(), mr(2, 4, bni), mr(1, 11, 1), mr2(), 2) 'mghz
                    bfshn.Cells(ii, mghz + 1).Value = ii 'mghz+1処理
                    Call hdrst(ii, a)         '左下ステータス表示部
                Next
                
                '30s86_020m 低速にも適用させるため、こちらにも　２行複数列で下側がall空白chk 30s86_020a
                If pap(2, 0) > 0 Then '30s86_020c　　020q低速専用として開設
                    ii = h
                    Do Until bfshn.Cells(ii, mghz) <> wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) '←使ってるのmr(2,4,bni)だけ
                        ii = ii - 1
                    Loop
                    If ii < h Then
                        MsgBox "mghz下側に「" & wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) & "」発生したので、"
                        MsgBox "mghzクリア入ります挙動注意_(2行複数列)" & ii + 1 & "～" & h
                        Range(bfshn.Cells(ii + 1, mghz), bfshn.Cells(h, mghz)).ClearContents
                    Else
                        'MsgBox "mghzクリアはありませんでした(2行複数列)"
                    End If
                End If
            End If
            cnt = 0
        End If '当シート既存情報あり
        
        If er(2, 1) > 0 Then er(2, bni) = mghz Else er(2, bni) = -mghz '←86_016w
    End If
    
    '(ｈ変更)非独立集計　（「ど」対応あべこべに） '＊ど」にも対応 2行複数列にも対応となっている。ゐ複文節にも対応
    If Not (InStr(1, mr(0, 2, bni), "ｦ") > 0 And InStr(1, rvsrz3(mr(0, 2, bni), 1, "ｦ", 0), "ど") > 0) Then  'h変える
        ii = h
        Do Until ii = k - 1
            If bfshn.Cells(ii, Abs(er(2, 1))) <> "" Then Exit Do
            ii = ii - 1
        Loop
        h = ii
    Else  'h変えない
        MsgBox "ど対応"
    End If
    
    DoEvents
    
    ct3 = ""
    ct8 = "" '86_019o
    With Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)) 'フィルタ解除する（コピペでの、はしょられ防止）
        If .FilterMode Then .ShowAllData
    End With

    cted(0) = ctdg(mr(2, 1, bni), mr(1, 1, bni), er(4, bni), a)
   ' If cted(0) > 500000 Then Call oshimai("", bfn, shn, 1, 0, "対象シート行数50万行超え")  '30s86_021e
    
    cted(1) = ctdr(mr(2, 1, bni), mr(1, 1, bni), er(4, bni), a)

    If h >= k And (mr(1, 11, 1) = "純高速" Or mr(1, 11, 1) = "近似高速") Then   'mghzソート
        '昇順キー(mghz列)を昇順に=xlAscending 昇順処理は近似高速(一文節のみ)・純高速(全文節)共に実施へ
        Call saato(bfn, shn, mghz, k, mghz, h, mghz + 1, 1)
    End If
    
    k0 = k  '複文節一旦リセット仕様へ
    h0 = h '複文節一旦リセット仕様へ
   
    If h >= k And (mr(1, 11, 1) = "近似高速") Then
        am1 = "" '近似時は初回値強制nullへ。(k行採用できないため)
        n1 = k
    ElseIf er(2, 1) < 0 Then  '高速時はこちら
        am1 = bfshn.Cells(k, Abs(er(2, bni))).Value  '30s83こちらで復活
        n1 = bfshn.Cells(k, Abs(er(2, bni)) + 1).Value
    Else  '通常
        am1 = bfshn.Cells(k, Abs(er(2, bni))).Value  '30s83こちらで復活
        n1 = k
    End If
    'ｓか」

    pap(5, 0) = kgcnt(mr(1, 5, bni), mr(2, 4, bni)) '5行目ヱの数
    pap(9, 0) = kgcnt(mr(1, 9, bni), mr(2, 4, bni)) 'pap6から変更
    
    ReDim er5(pap(5, 0))
    ReDim mr5(pap(5, 0))
    ReDim er9(pap(9, 0))
    ReDim mr9(pap(9, 0))
    
    er9(0) = Val(mr(1, 9, bni))
    mr9(0) = mr(2, 9, bni)

    If pap(9, 0) > 0 Then
        For ii = 0 To pap(9, 0)
            er9(ii) = Val(rvsrz3(mr(1, 9, bni), ii + 1, mr(2, 4, bni), 0))
            mr9(ii) = rvsrz3(mr(2, 9, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
    
    er5(0) = Val(mr(1, 5, bni))
    mr5(0) = mr(2, 5, bni)
    If pap(5, 0) > 0 Then
        For ii = 0 To pap(5, 0)
            er5(ii) = Val(rvsrz3(mr(1, 5, bni), ii + 1, mr(2, 4, bni), 0))
            mr5(ii) = rvsrz3(mr(2, 5, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
    
    pap(8, 0) = kgcnt(mr(1, 8, bni), mr(2, 4, bni)) '8行目ヱの数  ＊の時も仮で
   
    'pap(8,0)再定義（＊のとき）30ｓ70
    If StrConv(Left(rvsrz3(mr(0, 8, bni), 1, "ｦ", 0), 1), 8) = "*" Then '30s74改良
        paq8 = pap(8, 0) / 2 'paq8は半数(：＊のグルーピング数　.5もあり得る)
        qap = 0
        For ii = 0 To Int(paq8)
            If ii = Int(paq8) And paq8 - Int(paq8) = 0 Then '最終周かつpap(8,0)偶数(孤立)
                qap = qap + 1
            Else
                fma = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 1, mr(2, 4, bni), 0)) 'from列
                If fma = 0.1 Then fma = cted(1)  '85_020
                tob = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 2, mr(2, 4, bni), 0)) 'to列
                If tob = 0.1 Then tob = cted(1)  '85_020
                qap = qap + Abs(fma - tob) + 1
            End If
        Next
        pap(8, 0) = qap - 1 'pap(8,0)再定義完了
    End If
            '↓-1:複数列比較で使用30s79、0:7側、1：8側→終了
    ReDim er78(-1 To 1, -1 To pap(8, 0)) '←pap(7,0)<=pap(8,0)　という前提,-1はc=-1-2の時だけ使用(tensai側で)
    ReDim mr8(-2 To pap(8, 0)) '30s75
    
    er78(1, 0) = Val(mr(1, 8, bni))
    
    If er(10, bni) = -0.1 Then er78(1, 0) = er(10, bni) '30s86_018d 追加（行番号連載補正対応）
    
    If pap(8, 0) > 0 Then
        mr8(0) = rvsrz3(mr(2, 8, bni), 0 + 1, mr(2, 4, bni), 0)

        If StrConv(Left(bfshn.Cells(sr(8), a).Value, 1), 8) = "*" Then
            qaap = 0
            For ii = 0 To Int(paq8) '周
                fma = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 1, mr(2, 4, bni), 0)) 'from列
                If fma = 0.1 Then fma = cted(1)
                er78(1, qaap) = fma
                If ii < Int(paq8) Or paq8 - Int(paq8) = 0.5 Then   '最終周でない、あるいは最終周でtoあり
                    tob = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 2, mr(2, 4, bni), 0)) 'to列
                    If tob = 0.1 Then tob = cted(1)  '85_020
                    If fma > tob Then sa8 = -1 Else sa8 = 1
                    For qap = qaap + 1 To qaap + 1 + Abs(fma - tob) - 1 'from＝toはfor特性上、実施されない
                        er78(1, qap) = er78(1, qaap) + (sa8) * (qap - qaap)
                    Next
                End If
                qaap = qap
            '＊の時は、mr8(qap)は当面不使用とする　'30s75
            Next
        Else
            For ii = 1 To pap(8, 0) 'これまで通り
                er78(1, ii) = Val(rvsrz3(mr(1, 8, bni), ii + 1, mr(2, 4, bni), 0))
                mr8(ii) = rvsrz3(mr(2, 8, bni), ii + 1, mr(2, 4, bni), 0)
            Next
        End If
           
        '30s86_017v 特例措置（これ連載でないときも実施されているようである。問題のときはプログラム改善）86_020j
        If mr(2, 8, bni) = mr(2, 4, bni) Then
            MsgBox "連載時の特例措置mr8(0)→" & mr(2, 8, bni) & "へ"
            mr8(0) = mr(2, 8, bni)
        End If
    Else
        mr8(0) = mr(2, 8, bni)
    End If
    
    'pap(8,0)ここまで。ここからpap(7,0)
    'fma,tobはpap(7,0)として再リセットされて使用される(pap(8,0)のが踏襲されて使用されない)
    
    soroeru = 0
    er78(0, -1) = 0 '30s62 -1はc=-1-2の時だけ使用(tensai側で)
    er78(0, 0) = Val(mr(1, 7, bni)) '(1,0)→(0,0)修正30s61_4
    
    pap(7, 0) = kgcnt(mr(1, 7, bni), mr(2, 4, bni)) '7行目ヱの数  ＊の時も仮で
    If pap(8, 0) > 0 Then 'er78を定める。
        '７行目
        If StrConv(Left(bfshn.Cells(sr(7), a).Value, 1), 8) = "*" Then  '＊の周回ルーチン
            If Val(rvsrz3(mr(1, 7, bni), pap(7, 0) + 1, mr(2, 4, bni), 0)) = 0.1 Then '7行右が「－」
                soroeru = 1
                paq7 = (pap(7, 0) - 1) / 2
                '一番右に「－」アリ→soroeruビットを立てて、ーは無い前提(：pap(7,0)-1)でpaq7を設定
            Else
                paq7 = pap(7, 0) / 2 'paq7は半数(：＊のグルーピング数　.5もあり得る)
            End If
            
            qaap = 0
            For ii = 0 To Int(paq7) '＊ペアグループ毎で周回
                fma = Val(rvsrz3(mr(1, 7, bni), ii * 2 + 1, mr(2, 4, bni), 0)) 'from列
                er78(0, qaap) = fma  'ペアグループ毎の一個目
                qap = qaap  'ここでのqapは現fmaの配列位置(0,2,,,)
                pap(7, 0) = qaap
                If ii < Int(paq7) Or paq7 - Int(paq7) = 0.5 Then
                '(右のーは無い仮定での)最終周でない、あるいは最終周でtoあり
                '※右がーの処理は下のsoreoeru=1　の所で実施される。
                    tob = Val(rvsrz3(mr(1, 7, bni), ii * 2 + 2, mr(2, 4, bni), 0)) 'to列
                    
                    If qaap + 1 + Abs(fma - tob) - 1 > pap(8, 0) Then
                        Call oshimai("", bfn, shn, 1, 0, "転載列数：7行目>8行目です。確認を。")
                    End If
                    
                    If fma > tob Then sa7 = -1 Else sa7 = 1
                    For qap = qaap + 1 To qaap + 1 + Abs(fma - tob) - 1 '30ｓ70fromto導入
                        er78(0, qap) = er78(0, qaap) + (sa7) * (qap - qaap)
                    Next
                    pap(7, 0) = qap - 1
                    qaap = qap 'ここでのqap,qaapは、ある＊グループ精算後のnextのfma入れ込む位置
                End If
            Next
            
            If pap(7, 0) > pap(8, 0) Then Call oshimai("", bfn, shn, 1, 0, "pap(7,0)>pap(8,0)です。確認を。")
 
            If soroeru = 1 Then '7行右が「－」・・pap(8,0)と揃える
                For qaap = pap(7, 0) + 1 To pap(8, 0)
                    er78(0, qaap) = er78(0, pap(7, 0)) + (qaap - pap(7, 0))
                Next
                If er78(0, qaap - 1) >= mghz Then
                    Call oshimai("", bfn, shn, 1, Int(er78(0, qaap - 1)), "mghzはみ出てます。" & vbCrLf _
                    & rvsrz3(bfshn.Cells(1, Int(er78(0, qaap - 1))).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0) & "列までデータあり")
                End If
                pap(7, 0) = pap(8, 0)
            End If
        Else '普通のルーチン
            If pap(7, 0) > pap(8, 0) Then Call oshimai("", bfn, shn, 1, 0, "pap(7,0)>pap(8,0)です。確認を。")
            For ii = 1 To pap(8, 0)
                er78(0, ii) = Val(rvsrz3(mr(1, 7, bni), ii + 1, mr(2, 4, bni), 0))
            Next
        End If
    End If
'pap系の値策定はここまで。以降でpapの値はないと思われる。

    '↓mr8(pap(7,0))→mr8(0) 86_017v
    If Len(mr8(0)) > 1 And (er(8, bni) < 0 Or Round(er(6, bni), 0) = -15 Or Round(er(6, bni), 0) = -14) Then   '-14追加202005
        If Val("&H" & mr8(0)) <> 0 Then
            'MsgBox "16進変換アリ " & mr8(0) & "→" & Chr(Val("&H" & mr8(0)))
            mr8(0) = Chr(Val("&H" & mr8(0)))
        Else
            MsgBox "16進変換サレズ：" & mr8(0)
        End If
        '16進変換対象は、マニュアル連載のmr8だけである。
    End If

    If (er(5, bni) < 0 And pap(7, 0) <> pap(8, 0)) Then Call oshimai("", bfn, shn, sr(7), a, "列数が一致しません（差分の複数列比較）")
    
    '4列目に要素転記(８行目別途) ※２文節以降も実施,4列目反映へ
    For ii = 1 To 7
        hk1 = "" '第3因子対応
        If mr(3, ii, bni) <> "" Then
        'MsgBox "第3因子:" & mr(3, ii, bni)
        hk1 = "ｦ" & mr(3, ii, bni)
        End If
        If StrConv(Left(mr(0, ii, bni), 1), 8) = "*" Then  '30s75（*有無も追加）　'mr(1,→mr(4 へ
            bfshn.Cells(sr(ii), 4).Value = "*ｦ" & mr(4, ii, bni) & "ｦ" & mr(2, ii, bni) & hk1
        Else
            bfshn.Cells(sr(ii), 4).Value = "ｦ" & mr(4, ii, bni) & "ｦ" & mr(2, ii, bni) & hk1
        End If
    Next

    '8行内容を4列目に要素転記　※文字列変換済み
    If mr(2, 8, bni) = Chr(Val("&H" & "0A")) Then '行の膨らみ阻止16s
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then  'mr(1,→mr(4 へ
            bfshn.Cells(sr(8), 4).Value = "*ｦ" & mr(4, 8, bni) & "ｦ(LF)"
        Else
            bfshn.Cells(sr(8), 4).Value = "ｦ" & mr(4, 8, bni) & "ｦ(LF)"
        End If
    Else
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then
            bfshn.Cells(sr(8), 4).Value = "*ｦ" & mr(4, 8, bni) & "ｦ" & mr(2, 8, bni)
        Else
            bfshn.Cells(sr(8), 4).Value = "ｦ" & mr(4, 8, bni) & "ｦ" & mr(2, 8, bni)
        End If
    End If
    
    bfshn.Cells(sr(0), 4).Value = bni & "/" & bun & "文節"
    bfshn.Cells(sr(4), 5).Value = mr(2, 11, bni)    'koum（項準とか）
    cnt = 0  '左下カウンタリセット
    qq = er(11, bni)

    '対象シートの「1」探し 30s81 上側から引っ越し
    If Left(mr(2, 11, bni), 2) = "項無" And qq = 0 Then qq = 1 '項無の場合（⑧）
    If Left(mr(2, 11, bni), 2) <> "項準" Then '⑨項目行ある場合はその次行、ない場合は１行目から探す
        Do Until Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(4, bni))).Value = "all1" Or Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(4, bni))).Value = 1
            qq = qq + 1
            If qq = 2000 Then Call oshimai("", bfn, shn, k, a, "対象シートのall1列の「1」が見つからないようです。")
        Loop
        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(4, bni))).Value = "all1" Then qq = qq + 1
    Else '項準：次行から⑦　項順2、項順aも
        qq = qq + 1
    End If
    
    If pap(8, 0) > 0 Then
        '7行分 一旦仮完成、保留　30s79 複数列コメント挿入
        If pap(7, 0) > 0 Then '←参照シートに項目行が無い場合を除く
            Call tnsai(ct8, tst, ct3, er78(), a, sr(7), bni, 1, k - 1, -7, mr(), er(), pap(7, 0), mr8())
        End If
        '8行完成、一旦保留
        If er(11, bni) > 0 And pap(8, 0) > 0 Then '←参照シートに項目行が無い場合を除く
            If mr(2, 11, bni) = "項固" Then
                Call tnsai(ct8, tst, ct3, er78(), a, sr(8), bni, 1, qq - 1, -8, mr(), er(), 0, mr8())
            Else
                Call tnsai(ct8, tst, ct3, er78(), a, sr(8), bni, 1, Int(er(11, bni)), -8, mr(), er(), 0, mr8())
            End If
        End If
    End If

    '◆高速シート作成30s82f
        If mr(1, 11, 1) = "純高速" Then
            Call kskst(pap(7, 0), h, er78(), er9(), mr9(), er3, mr3(), er5(), mr5(), c5, pap(3, 0), pap(5, 0), bni, qq, rrr, mr(), er(), a, cted())
            hirt = Range(twt.Cells(1, 1), twt.Cells(rrr + 1, 9)).Value '※rrr→rrr+1(項無或いはデータすっからかん対策)　86_018i
        End If '高速シート作成ここまで
    
    ii = qq 'rrrは高速シートのデータ終了行(含ロック因子)を温存へ、iiが増えていく
    cnt = 0 'カウンタリセット
    pqp = 0 'ロックオン可否リセット

    If mr(1, 11, 1) = "純高速" Or mr(1, 11, 1) = "近似高速" Then
        hiru = Range(bfshn.Cells(1, mghz), bfshn.Cells(h, mghz + 1)).Value
    End If

    Application.Goto bfshn.Cells(k, a - 3), True  '86_020n
    bfshn.Cells(k, a).Select
    
    '◆ここから行毎
    Do While ii <= cted(0) '628より
        ct3 = ""
        ct8 = "" '86_019o
        If Abs(er(4, bni)) >= 1 Then
            If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(ii, Abs(er(4, bni))).Value = "" Then
                Exit Do  '対象シートのall_1列
            End If
        End If

        If mr(1, 11, 1) = "純高速" Then  '高速ロックオン判定
            If hirt(ii, 2) = "000_0000000" Then  '純高速時は高速シート参照へ　30s86_012r   0→0000000  86_019c
                UserForm4.StartUpPosition = 2 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
                UserForm4.Show vbModeless
                UserForm4.Repaint
                bfshn.Cells(sr(2), 5).Value = "純高ﾛｯｸ" '20191118
                k0 = h0  'これがポイント
                pqp = 1
                Unload UserForm4
                UserForm1.Repaint
            End If
            
            ct3 = hirt(ii + pqp, 6)  'ここでct3注入(特命条件で使用、転載エレメント)
            ct8 = hirt(ii + pqp, 8)  '   30s86_019o(特命条件で使用、転載エレメント(値の方))
            
            '↓これするために、５行０でも、高速シート３列目に１を埋め込む。へ。
            If hirt(ii + pqp, 3) = "" Then
                c5 = 0
            Else  '↓c5最終選考(純高速)追加(86_020f)
                                           '↓使用されてるのmr(2,4,bni)だけ
                If pap(2, 0) > 0 And wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) = hirt(ii + pqp, 1) Then '､､､､だけ→不採用
                    c5 = 0
                    MsgBox "2行目複数列の環境で対象列がall空白、出ました、無視されます(純高速)。避けたければ2行目単数列で。"
                Else '採用
                    c5 = 1
                    qq = Val(rvsrz3(hirt(ii + pqp, 2) & "", 2, "_", 0)) '高速シート→対象シートの行に換算,配列変数導入、c5=1のみに適用へ  86_019c
                End If
            End If
        
        Else  '純高速以外　ct3は必ず""
            c5 = kaunta(mr(), ii, pap(5, 0), bni, er5(), mr5()) 'ここでのiiは高速シートの0カウント行
             '（ここでのc5 = 1は、まだ仮採用状態）
            If c5 = 1 Then  'で、←c5最終選考(not純高速)　'↓使用されてるのmr(2,4,bni)だけ
                If pap(2, 0) > 0 And wetaiou("", "", 0, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) = wetaiou(mr(2, 1, bni), mr(1, 1, bni), ii, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) Then
                    c5 = 0 '採用予定→不採用へ　86_020f追加
                    MsgBox "2行目複数列の環境で対象列がall空白、出ました、無視されます(not純高速)。避けたければ2行目単数列で。"
                Else '採用予定→本採用へ
                    qq = ii 'c5=1確定時のみに適用へ(一応バグ)　86_015b　従来はこちらでけであった
                End If
            End If
        End If
        
        'カウント対象であれば実施（で無ければ飛ばす）
        If c5 = 1 Then
            'ヱ複文節の返り値を返す(○ヱ▽ヱ◇形式)。複文節で無くても常時使われる(am2)
            If mr(1, 11, 1) = "純高速" Then  '高速ロックオン判定
                am2 = hirt(ii + pqp, 1)   '30s85_027検証6
            Else
                am2 = wetaiou(mr(2, 1, bni), mr(1, 1, bni), qq, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) '30s51
            End If
            
            If am2 = "" Then am2 = "ー(情報空白行)ー"
            hk1 = ""

            '突合キーunicode→?になる対策開始　86_016q
            If InStr(am2, "?") = 0 And InStr(StrConv(am2, 8 + hrkt), "?") > 0 Then MsgBox "uniあり注意（" & am2  '86_020s　24→8 + hrkt

            'p判定--■--■--■--　p→pkt　86_014n
            If mr(1, 11, 1) = "純高速" Or mr(1, 11, 1) = "近似高速" Then  '高速ロックオン判定
               pkt = kskup(am1, am2, n1, n2, h, er(2, bni), er(6, bni), k0, h0, pap(2, 0), er2(), mr(1, 11, 1), pqp, er5(0), er3(), hiru)
            Else '低速
               pkt = tskup(am1, am2, n1, n2, h, er(2, bni), er(6, bni), k0, h0, pap(2, 0), er2(), mr(1, 11, 1), pqp, er5(0), er3())
            End If

            kahi = 0 '(加算可否判別フラグ) zyou→kahiへ　86_014j
            kasan = 0 '16s
            
            If pkt = -1 Then Exit Do
            If pkt = -2 Then '★629(85_001) pap(2,0)はゼロである。　高速ベタ
                MsgBox "高速ベタ"
                '特命条件での高速ベタはどうなるのでしょうか？↓今は強制終了
                If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0) Then Call oshimai("", bfn, shn, 1, 0, "特命条件、大丈夫？")
                
                If rrr - qq <= 1 Then '対象シートデータ2行以下なら高速ベタでなく通常方法に戻す(p=2)
                     pkt = 2
                Else
                    MsgBox "表空白ベタ貼り転載開始します。" & vbCrLf & _
                    "行数：(予定)" & rrr - qq + 1 & vbCrLf & _
                    "列数：(予定)" & 1 _
                    , 64, "ベタ貼り「転載」開始"
                    
                    '一列二列、５列のオートフィル跡地
                    Call betat4(bfn, shn, k, 0.4, k + rrr - qq, 0.4, bfn, shn, k, 1, "pp", "1")
                    
                    If Not StrConv(Left(mr(0, 2, bni), 1), 8) = "*" Then 'ー２にも2行＊概念反映へ
                        Call betat4(bfn, shn, k, 0.4, k + rrr - qq, 0.4, bfn, shn, k, 2, "pp", "c")
                    End If
                    
                    Call cpp2("", Now(), 0, 0, 0, -1, bfn, shn, k, 5, k + rrr - qq, 5, -4163) '5列目（now）これに凝縮 速
                    
                    '突合列　'↓文字列での値化（速達のため）
                    Call betat4(twn, "高速シート_" & syutoku(), qq, 1, rrr, 1, bfn, shn, k, Abs(er2(0)), "mm", "")
                    
                    '当列
                    If mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then 'ｦｦ定数加算対応
                        Call betat4(bfn, shn, k, 0.4, k + rrr - qq, 0.4, bfn, shn, k, a, "pm", mr(2, 9, bni))
                    ElseIf (er(8, bni) = 0 Or er(7, bni) <> 0) And er(6, bni) > 0.5 Then  '加算処理部(同列転載なら実行しない)
                        Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Int(er(6, bni)), rrr, Int(er(6, bni)), bfn, shn, k, a, "pm", "")
                    End If
                   
                    If pap(7, 0) <> pap(8, 0) Then MsgBox "p=-2でpap(7,0) <> pap(8,0)です.留意を(処理は続行)"
                        'MsgBox "横方向ベタ貼りです"
                    If er(8, bni) <> 0 Then
                        For jj = 0 To pap(7, 0) 'pap(7,0),pap(8,0)＝0の時は一回ポッキリ実施
                            qap = jj
                            If er78(1, jj) <> 0 Then
                                If er78(0, jj) = 0 Then
                                    axa = a
                                    If er(6, bni) > 0 And er(5, bni) >= 0 Then
                                        Call oshimai("", bfn, shn, 1, 0, "処理中止。a加算ありで同列転載しようとしています。確認を。")
                                    End If
                                Else
                                    axa = er78(0, jj)
                                End If
                            End If
                    
                            If pap(7, 0) > 0 And jj <> pap(7, 0) Then
                                Do While er78(0, jj + 1) - er78(0, jj) = 1 And er78(1, jj + 1) - er78(1, jj) = 1
                                    jj = jj + 1   ' for内のjjを増やす。
                                    If jj = pap(7, 0) Then Exit Do
                                Loop
                                'If qap <> jj Then MsgBox "横方向ベタ貼りです"
                            End If
                            DoEvents
                            Application.StatusBar = "横ベタ中、" & Str(jj) & " / " & Str(pap(7, 0)) & " 、 " & Str(Abs(qap - jj) + 1) & "列"
                
                            If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then '文字列時
                                er34 = "mm" '文字列指定(ーー)
                            ElseIf er(3, bni) > 0.2 And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then '通貨時
                                er34 = "pm"  '通貨指定（＋ー）
                            ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then
                                er34 = "mp"  'セル踏襲 "ap"→"mp（ー＋）"　遅・コピペパターン
                            Else
                                er34 = "pp"  '（＋＋）
                            End If
                            If Abs(er78(1, qap)) = 0.1 Then MsgBox "betat4挙動注意a" '86_019a
                            Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, er78(1, qap), rrr, er78(1, jj), bfn, shn, k, axa, er34, mr8(jj))
                        Next
                    End If
                    Exit Do
                End If
                Application.StatusBar = False
            End If  'p=-2ここまで

            '突合先にAMある時(p=1)の処理　n2：当シート処理行　qq：対象シート取込行
            If pkt = 1 Then
                If Round(er(6, bni)) = -2 Then
                    'c=-2の加算処理(ね、ひ)
                    'If er(7, bni) <> 0 And er(5, bni) >= 0 Then
                    '↓30s86_020n bni→1
                    If er(7, bni) <> 0 And er(5, 1) >= 0 Then
                        Call oshimai("", bfn, shn, 1, 0, "「-2」で他列操作をしようとしています。修正を。")
                    End If
                    '項準適用廃止30s64
                    If mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then
                        kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                        kahi = 1
                    ElseIf Round(er(9, bni)) <> 0 Then
                        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(9, bni))).Value <> "" Then
                            kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                            kahi = 1
                        End If
                    End If
                    If kahi = 1 Then bfshn.Cells(n2, a).Value = bfshn.Cells(n2, a).Value + kasan
                ElseIf er(8, bni) <> 0 And Not (er(5, bni) < 0 And er(6, bni) > 0) Then
                    '転載処理(連載型) 強制同列・(上書型) え、う　他列(上書型) さ、せ、ち、と (強制同列)
                    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) > 0) Then  '30s86_012w '特命加算型の方
                    '特命条件　特命時(加算型)はここはスルー
                    'MsgBox "iei"
                    Else  'これまではこちら↓　特命時(連載型)もこちら
                        If er(6, bni) = -5 Or er(6, bni) = -7 Then
                            bfshn.Cells(n2, a).Value = hunpan(mr(2, 1, bni), mr(1, 1, bni), qq, er(3, bni), er(4, bni), er(5, bni), er(6, bni), er(8, bni), mr(2, 6, bni), hk1)
                            If hk1 <> "" And er(7, bni) <> 0 Then bfshn.Cells(n2, a).Value = hk1
                        Else '特命条件(連載型)では↓こちらを通る。
                            Call tnsai(ct8, tst, ct3, er78(), a, n2, bni, pkt, qq, 0, mr(), er(), pap(7, 0), mr8()) '30s62一元化
                        End If
                    End If
                End If
                
                '加算処理(同列転載なら実行しない) か、え、き、あ
                If (er(10, bni) = 0 Or Abs(er(7, bni)) > 0.2) And (er(6, bni) > 0 Or Round(er(6, bni)) = -1) Then
                    If (er(5, bni) >= 0 And er(6, bni) < 0 And er(7, bni) <> 0) Or _
                    (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And Int(er(7, bni)) = 0 And er(8, bni) <> 0) Then
                        kasan = 1  '6行-1で他列転載時 or 特命条件時86_14s
                        kahi = 1
                    ElseIf mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then 'ｦｦ定数加算対応
                        '特命状況の加算型はここを通るケースある（既存p=1なら）→通らないへ86_14s
                        kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                        kahi = 1
                    ElseIf Round(er(9, bni)) > 0 Then
                        '特命状況の加算型はここを通るケースある（既存p=1なら）少→通らないへ86_14s
                        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = "" _
                            Or (mr(2, 9, bni) = "n0" And Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = 0) Then
                            '加算対象セルが空白or0なら加算処理実施しない
                        Else
                            '特命状況の加算型はここを通るケースある（既存p=1なら）多→通らないへ86_14s
                            kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                            kahi = 1
                        End If
                    End If
                    
                    If kahi = 1 Then
                        If er(2, 1) < 0 And n2 > h0 Then  '高速かつ新規の既存
                            hirt(n2 - h0 + 1, 5) = hirt(n2 - h0 + 1, 5) + kasan    '5列目処理
                            '(高速シート５列目使用、高速かつ新規の既存、（特命条件もこのケースあり）
                        Else
                            bfshn.Cells(n2, a).Value = bfshn.Cells(n2, a).Value + kasan
                            '特命条件でこのケースもある　最初から既存の時
                        End If
                    End If
                End If
                n1 = n2
            End If '(p=1ここまで)
            
            '突合先にAM無い時（追加）(p=2)　n2：当シート処理行　qq：対象シート取込行
            If pkt = 2 And er(6, bni) >= 0 Then
                '低速新規はちまちま1,2,5列追加へ（「ど」集計時の対応仕様）最初に　86_011q
                If mr(1, 11, 1) = "ノーマル" And bfshn.Cells(n2, 1).Value <> 1 Then
                    bfshn.Cells(n2, 1).Value = 1
                    If Not StrConv(Left(mr(0, 2, bni), 1), 8) = "*" Then '30s73
                        bfshn.Cells(n2, 2).Value = "c"
                    End If
                    bfshn.Cells(n2, 5).Value = Now() 'タイムスタンプ追記
                End If
            
                If bfshn.Cells(n2, Abs(er(2, bni))).Value <> "" Then
                    Call oshimai("", bfn, shn, 1, 0, "新規予定行に既に情報があります。確認を。")
                End If
                    
                If bfshn.Cells(n2 + 1, a).Value <> "" Then Call oshimai("", bfn, shn, 1, 0, "新規次行に既に情報があります。確認を。")

                If er(2, 1) < 0 Then  '85_007 高速時場合分け
                        hirt(n2 - h0 + 1, 4) = am2      '4列目処理
                Else  'ノーマル（複数列・単列）
                    If pap(2, 0) > 0 Then
                        '86_012j追加バグ対応
                        With bfshn.Cells(n2, Abs(er(2, bni)))  'キー追加（mghz)ノーマル複数列時
                            .NumberFormatLocal = "@"
                            .Value = am2
                        End With
                    Else  '
                        With bfshn.Cells(n2, Abs(er2(0)))  'キー追加（mghzのケースは発生しない。er(2,x)ではなくer(x)なので
                            .NumberFormatLocal = "@"
                            .Value = am2
                        End With
                    End If
                End If
                
                If pap(2, 0) > 0 Then   '２行目ヱ対応(複数列のとき) ※全部の複数列が対象　mghzではない
                    For jj = 0 To pap(2, 0) 'qap→jj
                        If Round(Abs(er2(jj))) > 0 Then  '30s70 0.4(null,ー)は無視とする(エラー防止のため)
                            With bfshn.Cells(n2, Abs(er2(jj)))
                                .NumberFormatLocal = "@"
                                .Value = rvsrz3(am2, jj + 1, mr(2, 4, bni), 0)
                            End With
                        End If
                    Next
                End If
             
                '転載処理部 連結型 あ、い、し、そ、つ、な(読点あり)強制同列・上書型 え、う （他列転載）・上書型 さ、せ 強制同列
                If er(8, bni) <> 0 And Not (er(5, bni) < 0 And (Round(er(6, bni)) = -2 Or er(6, bni) > 0)) Then
                    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) > 0) Then  '特命加算型の方
                        '特命条件追加30s86_012s　'特命時転載はしてはいけませんです。　'特命条件　特命時(加算型)はここはスルー
                    Else  'これまではこちら↓　特命時(連載型)はこちら　特命条件(連載型)では↓こちらを通る。
                        Call tnsai(ct8, tst, ct3, er78(), a, n2, bni, pkt, qq, 0, mr(), er(), pap(7, 0), mr8()) '30s62一元化
                    End If
                End If

                '加算処理部(同列転載なら実行しない) か、え、き,あ
                If (er(10, bni) = 0 Or er(7, bni) <> 0) And er(6, bni) > 0 Then  '特命条件対応版へ
                    '特命条件加算型はここでも処理される。　'項準適用廃止30s64
                    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And Int(er(7, bni)) = 0 And er(8, bni) <> 0) Then
                        '特命条件一律これで　30s86s
                        kahi = 1
                        kasan = 1
                    ElseIf mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then
                        kahi = 1
                        kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                    ElseIf Round(er(9, bni)) > 0 Then
                        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = "" Or _
                                (mr(2, 9, bni) = "n0" And _
                            Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = 0) Then
                        Else
                            kasan = tszn(er9(), bni, mr(), qq, pap(9, 0), mr9())
                            kahi = 1
                        End If
                    End If
                    If kahi = 1 Then
                        If er(2, 1) < 0 Then
                            hirt(n2 - h0 + 1, 5) = kasan
                            '特命条件こっちっぽい　高速なので
                        Else
                            bfshn.Cells(n2, a).Value = kasan
                        End If
                    End If
                End If
                    
                If er(2, 1) > 0 Then h0 = n2 'h0も更新(高速時以外) ※一文節目で判断30s22
                '１列２列５列　ちまちま→文節最後にまとめてで
                h = h + 1
            
            End If '(p=2ここまで)
            
            If (pkt = 2 Or pkt = 1) Then  '-1-2新規以外の時 新設
                am1 = am2
                n1 = n2
            End If
        End If 'カウント対象時行うここまで

        ii = ii + 1
        If er(2, 1) < 0 Then
            Call hdrst2(ii, a, 10000, k0, h0)
        Else
            Call hdrst2(ii, a, 1000, k0, h0) '201904　100→1000
        End If
    Loop '◇ここから行毎ここまで
  
    'ここでのqq・・対象シートで最終でカウント対象とした行
    bfshn.Cells(2, 4).Value = k0 'Loop終了後行う30s86_002より
    bfshn.Cells(3, 4).Value = h0 'Loop終了後行う30s86_002より
   
    If mr(1, 11, 1) = "純高速" Or mr(1, 11, 1) = "近似高速" Then Erase hiru
    
    If mr(1, 11, 1) = "純高速" Then '高速シートAーE列入れ込み戻し、当列、突合列へベタ
        Range(twt.Cells(1, 1), twt.Cells(rrr + 1, 5)).Value = hirt  '※項無対策(rrr→rrr+1)
        Erase hirt

        'ノーマル同様、先に1,2,5列処理へ
        bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
    
        '一列２列５列新規増分分のまとめて埋め込み　高速時のみに(低速は↑で実施済み)
        If bfshn.Cells(1, 4).Value - 1 + k - 1 < h Then  '「ど」考慮型
            '２列目（ｃ）こちらが先
            If Not StrConv(Left(mr(0, 2, bni), 1), 8) = "*" Then 'ｐ＝ー２にも2行＊概念反映へ
                Call betat4(bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 0.4, h, 0.4, bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 2, "pp", "c")
            End If
            '5列目（now）
            Call cpp2("", Now(), 0, 0, 0, -1, bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 5, h, 5, -4163) '5列目（now）これに凝縮　-1→-2
            '一列目（１）は最後で
            Call betat4(bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 0.4, h, 0.4, bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 1, "pp", "1")
            DoEvents
        End If
        
        If h0 < h Then  '新規あるときのみ ←h1→h0へ。準高速onlyなので　And pap(2,0)撤廃
            '突合列新規ベタ(単数列時のみ)
            If pap(2, 0) = 0 Then
                Call betat4(twn, "高速シート_" & syutoku(), 2, 4, h - h0 + 1, 4, bfn, shn, h0 + 1, Abs(er2(0)), "mm", "")
            End If
            
            '当列新規ベタ（加算時のみ）※同列転載なら実行しない)
            '↓特命対応
            If (er(10, bni) = 0 Or er(7, bni) <> 0) And er(6, bni) > 0 Then
                '当列控え
                Call betat4(twn, "高速シート_" & syutoku(), 2, 5, h - h0 + 1, 5, bfn, shn, h0 + 1, a, "pm", "")
                '※この後、高速シート５列目は、↓のstrconv24用途で使用
            End If
            
            'mghz列新規ベタ 30s86_011
            Application.Calculation = xlCalculationAutomatic    '数式計算方法自動に　'新形式　85_007
            DoEvents '86_019r

            Application.Calculation = xlCalculationManual  '再計算再び手動に（重くなるため）30s66
            
        End If
    End If
    
    cnt = 0
    Call hdrst(ii, a)   'exitdoを考慮し、ここにも
    bfshn.Cells(1, a).Value = cted(0)
    '↓不要かもだが、念のため
    bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1

    'サマリ値処理（転載分）p=-2は、転載も実施
    If er(5, bni) >= 0 And er(6, bni) >= 0 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then
        '再計算を一旦自動に戻す
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        If pap(7, 0) > pap(8, 0) Then ii = pap(8, 0) Else ii = pap(7, 0)
        'もともとp=-2の転載処理
        jj = ii  'jjは最終節
        
        bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
        
        For ii = 0 To jj 'pap(7,0),pap(8,0)＝0の時は一回ポッキリ実施
            Application.StatusBar = "サマル中、" & Str(ii) & " / " & Str(jj)
            If er78(0, ii) > 0.2 And er78(0, ii) <> a Then  '86_013f 7行0.4対処→0.1対処
                Call samaru(Int(er78(0, ii)), mr(1, 1, 1)) '複数列さまる
            End If
        Next
        '再計算を再び手動に
        Application.Calculation = xlCalculationManual
        Application.DisplayStatusBar = True
    End If
    
    Next
'◇Ｄ◇複文節ふい（実行編）ここまで↑

'★★＃papchk予定地＃
Call papchk(pap(), 0, bun)

'◆Ｅ◆単文節ふい（事後編）ここから↓
    bni = 1 '1文節目の値で判断、実行 86_020p　bun→1戻す。一括踏襲のところは、bni→bunへ置換
    mg2 = 0 'mg2の変数はとりあえずpublicで入れている。整理後ほど

    'ここに -1-2 の空白埋め込みルーチン　複数列対応30s61_4
    If er(6, bni) = -1 Or er(6, bni) = -2 Or er(6, bni) <= -3 Then
        If nifuku = 1 Then
            'MsgBox "2文節目での複文節あり。注意を。"  '解禁へ　86_016v
            Call oshimai("", bfn, shn, 1, 0, "2文節目複文節での指定は許容されていません")
        End If
      
      'c<=-1から変更(C<=-3 の空白埋め込みは近似高速時のみやるに復活へ　85_014
        DoEvents
        cnt = 0
        k0 = k  'k0一旦リセット(高速時、またここから増えていく)
        
        If mr(1, 11, bni) = "近似高速" Then  '近似型
            If er(6, bni) = Round(er(6, bni), 0) Then '-15.1→複写対応しない
                mg2 = 1
                For ii = h To k Step -1  'i→ii
                    Do Until bfshn.Cells(ii, mghz).Value <> ""  '使われていない？
                        ii = ii - 1
                    Loop
                    '次行(前行)一致＆aに情報あり　'↓こちらへ(strconv「2」被せる)
                    If StrConv(bfshn.Cells(ii, mghz).Value, 2) = StrConv(bfshn.Cells(ii + 1, mghz).Value, 2) Then
                        If bfshn.Cells(bfshn.Cells(ii + 1, mghz + 1).Value, a).Value <> "" Then
                            Call tnsai(ct8, tst, ct3, er78(), a, bfshn.Cells(ii, mghz + 1).Value, bni, 1, 0, bfshn.Cells(ii + 1, mghz + 1).Value, mr(), er(), pap(7, 0), mr8())
                        End If
                    End If
                    Call hdrst(h - ii, a)  '左下ステータス表示部
                Next
                cnt = 0
            End If
        ElseIf er(2, bni) < 0 Then
            Call oshimai("", bfn, shn, sr(2), a, "旧高速使用終了です。")
        Else 'ノーマル
            If er(6, bni) = -1 Or er(6, bni) = -2 Then
                For ii = k + 1 To h  'i→ii
                    Do Until bfshn.Cells(ii, Abs(a)).Value = NullString
                        ii = ii + 1
                    Loop
                    If bfshn.Cells(ii, Abs(er(2, bni))).Value = bfshn.Cells(ii - 1, Abs(er(2, bni))).Value Then '前行と同じ場合
                        If bfshn.Cells(ii - 1, a).Value <> "" Then '-1他列転載時、これまでのようにmatchてっぺんに必ず情報があるとは限らなくなったので、その対処）
                            Call tnsai(ct8, tst, ct3, er78(), a, ii, bni, 1, 0, ii - 1, mr(), er(), pap(7, 0), mr8()) '30s81qq無効化
                        End If
                    Else
                        '高速時は↓以下通らない(キーが昇順の前提であるため。必ずエラーになる)。
                        If Not IsError(Application.Match(bfshn.Cells(ii, Abs(er(2, bni))).Value, Range(bfshn.Cells(k0, Abs(er(2, bni))), bfshn.Cells(ii - 1, Abs(er(2, bni)))), 0)) Then  'match使用
                            m = Application.WorksheetFunction.Match(bfshn.Cells(ii, Abs(er(2, bni))).Value, Range(bfshn.Cells(k0, Abs(er(2, bni))), bfshn.Cells(ii - 1, Abs(er(2, bni)))), 0) 'h→ii-1に修正30s24
                            If bfshn.Cells(k0 + m - 1, a).Value <> "" Then  '-1他列転載時、これまでのようにmatchてっぺんに必ず情報があるとは限らなくなったので、その対処
                                Call tnsai(ct8, tst, ct3, er78(), a, ii, bni, 1, 0, k0 + m - 1, mr(), er(), pap(7, 0), mr8()) '30s63集約化
                            End If
                        End If
                    End If
                    Call hdrst(ii, a)    '左下ステータス表示部
                    bfshn.Cells(2, 4).Value = k0 'Cells(13, 3)→Cells(2, 4)
                Next
                cnt = 0
            End If
        End If
    End If
    cnt = 0

    '面取り
    '↓連載型特命条件も追加へ　86_013d
    If ((er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) < 0) Or _
        (er(9, bni) = 0 And er(10, bni) < 0)) And _
        Not ((er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1)) Then '(当列)最後の「、」を取る。
        
        Application.Cursor = xlWait '85_026
      
        If h >= k Then 'データ無の時は実施しないへ　86_018f
            hirt = Range(bfshn.Cells(k, a), bfshn.Cells(h + 1, a)).Value 'h→h+1 variant で配列、ただしセルが1x1だけの時、配列として認識してくれない対策　30s86_017v
            For ii = k To h
                'hirtでは、ｋ→１、ｈ→ｈ－ｋ＋１　、forのiiに、-k+1をかぶせる　ii→ii-k+1
                '↓86_017m(連載複数文字対応) ↓86_017v(修正・区切り文字は複数列一番右→一番左に仕様変更へ（この方が自然）
                If Right(hirt(ii - k + 1, 1), Len(mr8(0))) = mr8(0) Then hirt(ii - k + 1, 1) = Left(hirt(ii - k + 1, 1), Len(hirt(ii - k + 1, 1)) - Len(mr8(0)))
                If hirt(ii - k + 1, 1) = "" Then hirt(ii - k + 1, 1) = NullString '　""を空白に（カウントさせないため）←ここはケガの功名にならなかった。
                Call hdrst(ii, a)   '左下ステータス表示部
            Next
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).Value = hirt '新バ
            Erase hirt '新バ
        End If
        Application.Cursor = xlDefault
    End If
    cnt = 0
    
    '結合
    If (er(5, bni) < 0) And er(6, bni) <= 0 And Round(er(6, bni)) <> -2 And er(7, bni) < 0 And Not IsNumeric(bfshn.Cells(sr(7), a).Value) Then  '2s
        For ii = k To h  'i→ii
            bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value & bfshn.Cells(ii, Abs(er(7, bni))).Value '文字結合
        Next
        Call hdrst(ii, a)    '左下ステータス表示部
    End If
        
    '差分処理(er(5,0)<0)
    If (er(5, bni) < 0) Then  '2s
        cnt = 0
        axa = 0
        dif = 1
        If IsNumeric(bfshn.Cells(sr(7), a).Value) Then
            axa = a + er(7, bni) '数値：位置が相対値
        Else
            axa = Abs(er(7, bni)) '文字：位置が絶対値
            If er(7, bni) < 0 Then dif = -1 '和・結合実施フラグ
        End If
        bfshn.Cells(sr(0) + 3, 3).Value = ""

        Application.Cursor = xlWait
        For ii = k To h
            saemp = 0
            If pap(7, 0) <> 0 Or (bfshn.Cells(ii, a).Value <> "" Or bfshn.Cells(ii, axa).Value <> "") Then '両方空欄なら実施しない(：単列なな以外実施)
                If IsNumeric(bfshn.Cells(ii, axa).Value) Or IsDate(bfshn.Cells(ii, axa).Value) Then
                    saemp = bfshn.Cells(ii, axa).Value
                End If
                '文字比較(新対応)※従来型は下部elseifへ
                If er(6, bni) <= 0 And Round(er(6, bni)) <> -2 And dif = 1 Then
                    zzz = bfshn.Cells(ii, a).Value '既存のデータ箱
                    ReDim zxxz(pap(7, 0)) 'pap(7,0)=8の前提（複数差分比較時）
                    ReDim xxxx(pap(7, 0)) 'pap(7,0)=8の前提（複数差分比較時）
                    ReDim zyyz(pap(7, 0)) 'pap(7,0)=8の前提（複数差分比較時）20190614、013q
                    
                    Range(Cells(ii, a), Cells(ii, a)).ClearContents  '一旦空白に　86_020bからこの記述
                        
                    zyyz() = Split(zzz, mr(2, 4, bni))  'Split関数で対象シート側個々要素を一気に入れる
                    
                    For jj = 0 To pap(7, 0) '比較列を個々見ていく ii→jjへ
                        xxxx(jj) = bfshn.Cells(ii, er78(0, jj)).Value '当列側比較対象列
                        'kahi(各々のセル比較結果の策定)
                        If zzz = "" Then
                            kahi = -1
                        '↓ここでエラー
                        ElseIf zyyz(jj) = "" And xxxx(jj) = "" Then 'zyz→zyyz(jj)
                            kahi = 2  'なしなし
                        ElseIf zyyz(jj) = xxxx(jj) Then
                            kahi = 0
                        Else
                            kahi = 1
                        End If
                            
                        If mr(2, 6, bni) = "-1" Then 'op:1　差分は具体的差分表記を羅列パターン )1→-1に変更
                            If kahi <> -1 Then  '-1は羅列すらしない
                                If kahi = 1 Then '差分情報付記
                                    zxxz(jj) = zyyz(jj)  'セル転記は最後↓、join関数にて
                                Else  '情報掲載せず（kahi=0,2)
                                End If
                            End If
                        Else 'op:2、無、c、cc（差分は数値フラグ表記パターン）
                            If mr(2, 6, bni) = "-2" Then  '数値フラグを羅列で表記　※セル転記は最後↓、join関数にて　2→-2に変更
                                '数値フラグ設定
                                If kahi <> 2 Then
                                    zxxz(jj) = CStr(kahi) '=LTrim(Str(kahi)) -1,0,1が刻まれる
                                    '挿入箇所(ヒヅケセルあたりのコメント）
                                End If
                            Else 'op:無、c、cc　 '数値フラグを合算で表記
                                If kahi <> 2 Then  '比較結果２→合算操作自体を行わない
                                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value + kahi  'セルに結果記載（合計値）
                                        
                                    'differentセルの色付け
                                    If kahi = 1 And Left(mr(2, 6, bni), 1) = "c" Then '色付け(c,cc)
                                        bfshn.Cells(ii, a).Interior.Color = RGB(254, 254, 238) '← = 15662846　→　30s86_017p
                                        bfshn.Cells(ii, er78(0, jj)).Interior.Color = RGB(254, 254, 238) '← = 15662846　30s86_017p
                                        bfshn.Cells(ii, er78(0, jj)).ClearComments
                                        If mr(2, 6, bni) = "cc" Then 'さらにコメント付加（比較元側）(cc)処理重い
                                            bfshn.Cells(ii, er78(0, jj)).AddComment
                                            bfshn.Cells(ii, er78(0, jj)).Comment.Text Text:=zyyz(jj) 'zyz→zyyz(jj)
                                            bfshn.Cells(ii, er78(0, jj)).Comment.Shape.TextFrame.AutoSize = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    '羅列型(１か２)はここで記載↓join関数使用
                    If ((mr(2, 6, bni) = "-1" And zzz <> "") Or mr(2, 6, bni) = "-2") Then
                        bfshn.Cells(ii, a).Value = Join(zxxz, mr(2, 4, bni))
                    End If
                    '文字比較(新対応)201708追加ここまで
                ElseIf er(6, bni) <= 0 And Round(er(6, bni)) <> -2 And dif = -1 Then
                    '文字結合→移行によりここでは何もしないことに
                ElseIf dif = -1 And er(8, bni) >= 0 Then  '1m =0→>=0 30s69バク修正
                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value + saemp '和
                ElseIf dif = -1 And er(8, bni) < 0 Then  '<-0.5→<0　30s69バク修正
                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value * saemp '積
                ElseIf er(8, bni) >= 0 Then 'e=0→e>=0修正16ｓ
                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value - saemp '差
                ElseIf er(8, bni) < 0 Then  '1m　<-0.5→<0　30s69バク修正
                    bfshn.Cells(ii, a).Value = saemp - bfshn.Cells(ii, a).Value '差(反転)
                End If
            End If
            '比較の移行先はここ←？
            Call hdrst(ii, a)              '左下ステータス表示部
        Next  '差分のある一行処理ルーチンはここまで
        
        Application.Cursor = xlDefault
        cnt = 0
    End If
  End If 'るA　-99or*,**は通過、ここまで
  
' 86_020p　戻す。一括踏襲のところは、bni→bunへ置換
  '終了時の文字調整(強制orちまちま→不要、通常の加算や転載or突合元情報の最終行の書式に合わせる時のみ要、)
  If er(6, bun) >= 0 Or (er(6, bun) < 0 And tst = 7 And trt <> -9) Then 'c<0のvlookup型も実施しないに(tst7,-99以外)　86_013d
        If tst = -2 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        ElseIf tst = 0 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "G/標準"
        ElseIf tst = 7 Then
            If trt = -2 Then  '加算型の一括踏襲
                If Abs(er(9, bun)) = 0.4 Or Abs(er(9, bun)) = 0.1 Then
                    MsgBox "加算固定値(0.4、0.1ー)なので一括踏襲は行われません。"
                Else
                    If qq > 0 Then
                    Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = Workbooks(mr(2, 1, bun)).Sheets(mr(1, 1, bun)).Cells(qq, Abs(er(9, bun))).NumberFormatLocal
                    Else
                    MsgBox "一括踏襲はスルーします(情報無いので)"  '86_016z
                    End If
                End If
            ElseIf trt = -1 Then   '転載型の一括踏襲
                If er(10, bun) < 0.5 Then
                    MsgBox "連載、ｦｦ○(：転載情報書式が存在しない) は対象外"
                Else
                    If qq > 0 Then
                    Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = Workbooks(mr(2, 1, bun)).Sheets(mr(1, 1, bun)).Cells(qq, Abs(er(10, bun))).NumberFormatLocal
                    Else
                    MsgBox "一括踏襲はスルーとなります(情報無いので)""  '86_016z"
                    End If
                End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "こういうケースがあるのだろうか？")
            End If
        End If

        '折り返して全体を表示しない (LF対処)16s　一律適用へ18s
        If trt <> -2 Then Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).WrapText = False
  End If
'◇Ｅ◇ここまで(not98)
        
'～～-99､-98用ここから～～
  If er(6, bni) <= -90 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then 'るC　20190214　*ありは実施しないに
    
        '独立集計(ｈ変更)
        'こちらの「ど」の対応はどうするか要検討
        If InStr(1, mr(0, 2, bni), "ｦ") > 0 And InStr(1, rvsrz3(mr(0, 2, bni), 1, "ｦ", 0), "ど") > 0 Then
            ii = h
            Do Until ii = k - 1  'ケツから上方向へサーチ
                If bfshn.Cells(ii, Abs(er(2, 1))) <> "" Then Exit Do  'bfshn.を被せた
                ii = ii - 1
            Loop
            h = ii
        End If
                    
        DoEvents
        Application.Calculation = xlCalculationAutomatic    '数式計算方法自動に　'新形式　85_007

        Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "G/標準"    'mghz列一旦標準へ（数式入れるため）
    

    
    
        '数式でコピペ
        If mr(2, 6, bni) = "c" Then '85_024_まだ使用されていない。
            MsgBox "使われている？"
            bfshn.Cells(sr(8), a).Copy  'ほぼ避けられない
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).PasteSpecial Paste:=xlPasteAll 'すべて -4104
        Else '従来型 -4123→xlPasteFormulas→.copy.paste
            Call copipe(bfn, shn, sr(8), a, sr(8), a, bfn, shn, k, a, h, a, 3)  '3→FormulaR1C1(脱.copy.paste)　86_020d
        End If
    
        '30s86_021cこちらへ
        If tst = 1 Then '←当列転載文字型
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "@"
        ElseIf tst = -2 Then  '通貨型
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        ElseIf tst = 7 Then '対象シート名を模倣　a11_1列→対象シートに変更
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = bfshn.Cells(sr(1), a).NumberFormatLocal 'sr(4)→sr(1)へ30s68
        End If
    
        If er(6, bni) = -99 Then 'セルを値に変換
            Call cpp2(bfn, shn, k, a, h, a, bfn, shn, k, a, 0, 0, -4163) '何と""→Nullstringされる(ケガの功名、浄化作用)
        End If
        '30s86_021c上へ
'        If tst = 1 Then '←当列転載文字型
'            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "@"
'        ElseIf tst = -2 Then  '通貨型
'            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[赤]-#,##0"
'        ElseIf tst = 7 Then '対象シート名を模倣　a11_1列→対象シートに変更
'            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = bfshn.Cells(sr(1), a).NumberFormatLocal 'sr(4)→sr(1)へ30s68
'        End If

        bfshn.Cells(sr(8), 4).Value = "(-99系)"  '86_019q
        bfshn.Cells(sr(8), a).ClearComments   '86_019q　既存コメントは消していく仕様へ
        bfshn.Cells(1, a).Value = h  'h0→h 30s64
    
        DoEvents
        Application.Calculation = xlCalculationManual  '再計算再び手動に（重くなるため）30s66
    
        cnt = 0
        If er(5, bni) <> 1 Then  'カウント基準列が空白行はこちらも空白に 30s68
            Application.Cursor = xlWait
            For ii = k To h
                If kaunta(mr(), ii, pap(5, 0), bni, er5(), mr5()) = 0 Then '30s85_024検証6
                    'bfshn.Cells(ii, a).Value = bfshn.Cells(sr(0) + 2, 1).Value
                    '↓30s86_020b
                    Range(Cells(ii, a), Cells(ii, a)).ClearContents
                End If
                Call hdrst(ii, a)            '左下ステータス表示部
            Next
            Application.Cursor = xlDefault
        End If
    
        cnt = 0
        If er(7, bni) = 0 Then '30s84_617
            Application.Cursor = xlWait
            yayuyo = 0
            For ii = k To h  'h
                If IsError(bfshn.Cells(ii, a)) Then '30s68_2バグ改良
                    yayuyo = yayuyo + 1 '85_005
                    If yayuyo = 10 Then Call oshimai("", bfn, shn, k, a, "やゆよ10回")
                ElseIf bfshn.Cells(ii, a) = "" Then
                    Range(Cells(ii, a), Cells(ii, a)).ClearContents    '""→空白　30s86_020b　NullStringからこちらへ
                End If
                Call hdrst(ii, a)            '左下ステータス表示部
            Next
        End If
        Application.Cursor = xlDefault
  '～～-99用ここまで～～
  End If  'るC
    Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).WrapText = False  '86_016t（折り返ししないAPIのXML膨張防止）こっちに
    
  '左下ステータス表示部　ここだけ特殊バージョン
  DoEvents
  If flag = True Then Call oshimai("", bfn, shn, k, a, "中止しました")    '中止ボタン処理
  Application.StatusBar = Str(cnt) & "、" & Str(a - dd1 + 1) & " / " & Str(dd2 - dd1 + 1)
  cnt = 0
 
End If 'り(not"*")
     
    '再計算を自動に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
     
    'サマリ値処理（転載分）ここでも(-1,-2のみ)
    If er(5, bni) >= 0 And er(6, bni) > -90 And er(6, bni) < 0 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then
        If pap(7, 0) > pap(8, 0) Then jj = pap(8, 0) Else jj = pap(7, 0) 'ii→jj
        
        bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
        
        For ii = jj To 0 Step -1
            If er78(0, ii) > 0.2 And er78(0, ii) <> a Then Call samaru(Int(er78(0, ii)), mr(1, 1, 1))  '複数列さまる
        Next
    End If
    bfshn.Cells(sr(0) + 4, a).Select
    
    Call samaru(a, mr(1, 1, 1))  '当列さまる
    bfshn.Cells(sr(0) + 3, a).Value = Now() '活用例1タイムスタンプ入れる
    bfshn.Cells(sr(0) + 2, 4).Value = Now() '活用例1タイムスタンプ入れる sr(0) + 3,→sr(0) + 2　３０ｓ７４
    '↓一行毎にインクリメント（ID変わったときのみリセットされる） 。※bfn、shn側は更新されない（初回複写時の値が載ってるだけ）。
    twbsh.Cells(14, 3).Value = twbsh.Cells(14, 3).Value + 1 '7s (3,3)→(14,3)25s
Next '選択範囲列分の繰り返し　ら
'ここでの「a」は、d+1である。
Call giktzg(a, rog)  'ログ部別プロシージャー化_201905
   
End Sub '外部結合ここまで
