Sub 複写()  ' 選択範囲を別シートにコピー
    Dim a, j, x As Integer   'g1→gg1,g→gg2(ローカル撤廃,グローバル化)　30s74
    Dim i As Long, jj(2) As Long, kk
    Dim shemei, wd, dg1, dg2 As String, hk1 As String, pasu As String
    Dim se_name As String, fimei As String, c99 As String
    
    'おまじない(「コードの実行が中断されました」対処)
    Application.EnableCancelKey = xlDisabled
    pasu = ActiveWorkbook.Path
    
    kyosydou
    If dd1 = 0 Then Call oshimai("", bfn, shn, 1, 0, "dd1がゼロです")
    
    'その日の初回チェック
    j = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(j, 1).Value = ""
        j = j + 1
        If j = 50000 Then
            MsgBox "空白行が見つからないようです"
            Exit Sub
        End If
    Loop
    If Workbooks(twn).Sheets(shog).Cells(j - 1, 4).Value <> Val(Format(Now(), "yyyymmdd")) Then
        Call oshimai("", bfn, shn, 1, 0, "その日の初回は、最初に[初回]ボタンを押して下さい。")
    End If
    
    If bfshn.Cells(sr(0) - 1, 5) = "" Then
        Call oshimai("", bfn, shn, sr(0) - 1, 5, "集計名を入力して下さい(緑色セル)")
    End If

    Call iechc(hk1)  '旧igchc(hk1)
    
    a = Len(bfn)
    zikan = Format(Now(), "yymmdd_hhmmss")
    shemei = bfshn.Cells(1, 3).Value & "_" & zikan
    se_name = bfshn.Cells(1, 3).Value
    If bfshn.Cells(1, 5).Value = "←" Then '30s81
        fimei = shemei
    ElseIf bfshn.Cells(1, 5).Value <> "" Then
        fimei = bfshn.Cells(1, 5).Value & "_" & zikan
    End If
    
    '再計算を自動に
    Application.Calculation = xlCalculationAutomatic
    j = 6
    i = 19967
    Do Until j >= mghz
        If bfshn.Cells(2, j).Value <> "" Then
            If Not IsError(Application.Match(bfshn.Cells(2, j).Value, Range(bfshn.Cells(2, j + 1), bfshn.Cells(2, mghz)), 0)) Then
                Range(bfshn.Cells(2, Application.Match(bfshn.Cells(2, j).Value, Range(bfshn.Cells(2, j + 1), bfshn.Cells(2, mghz)), 0) + j), bfshn.Cells(2, Application.Match(bfshn.Cells(2, j).Value, Range(bfshn.Cells(2, j + 1), bfshn.Cells(2, mghz)), 0) + j)).Select
                Call oshimai("", bfn, shn, 2, Int(j), "文字重複セルあり（" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & "セル）")
            End If
        End If
        If bfshn.Cells(2, j).Value <> "" And IsNumeric(bfshn.Cells(3, j).Value) Then '30s76 若干改良(３行目文字の時は無視)
            If bfshn.Cells(3, j).Value > i Then i = bfshn.Cells(3, j).Value
        End If
        j = j + 1
        Application.StatusBar = "重複確認、" & Str(j) & " / " & Str(mghz) '9s
    Loop
    i = i + 1
    Application.StatusBar = False
    If gg2 <= 3 Then
      '---Unicode部ここから---   '再計算を手動に
        Application.Calculation = xlCalculationManual
        For j = dd1 To dd2
            Application.StatusBar = "埋め込み中、" & Str(j - dd1 + 1) & " / " & Str(dd2 - dd1 + 1) '9s
            If gg2 = 3 And bfshn.Cells(2, j).Value = "" And bfshn.Cells(3, j).Value = "" Then '30s76
                bfshn.Cells(3, j).Value = i
                i = i + 1
            End If
            If gg1 <= 2 And bfshn.Cells(2, j).Value = "" And bfshn.Cells(3, j).Value <> "" Then
                If IsError(Application.Match(ChrW(bfshn.Cells(3, j).Value), Range(bfshn.Cells(2, 6), bfshn.Cells(2, mghz)), 0)) Then
                    bfshn.Cells(2, j).Value = ChrW(bfshn.Cells(3, j).Value)
                Else
                    Range(bfshn.Cells(2, Application.Match(ChrW(bfshn.Cells(3, j).Value), Range(bfshn.Cells(2, 6), bfshn.Cells(2, mghz)), 0) + 5), bfshn.Cells(2, Application.Match(ChrW(bfshn.Cells(3, j).Value), Range(bfshn.Cells(2, 6), bfshn.Cells(2, mghz)), 0) + 5)).Select
                    Call oshimai("", bfn, shn, 2, Int(j), "挿入予定文字「" & ChrW(bfshn.Cells(3, j).Value) & "」：" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & "セルと重複")
                End If
            ElseIf gg1 = 3 And bfshn.Cells(2, j).Value <> "" And bfshn.Cells(3, j).Value = "" Then
                bfshn.Cells(3, j).Value = AscW(bfshn.Cells(2, j).Value)
                If bfshn.Cells(3, j).Value < 0 Then bfshn.Cells(3, j).Value = bfshn.Cells(3, j).Value + 65536
                bfshn.Cells(3, j).Value = "手" & bfshn.Cells(3, j).Value
            End If
        Next
        '再計算を自動に
        Application.Calculation = xlCalculationAutomatic
        bfshn.Cells(3, 3).Value = i
        Application.StatusBar = False
      '---Unicode部ここまで---
    ElseIf gg2 = sr(8) And gg1 = sr(8) Then
       '8行コメントコピー 30s86_012_a
        If dd1 = dd2 And bfshn.Cells(sr(6), dd1).Value = -99 Then
            If TypeName(ActiveCell.Comment) = "Comment" Then 'コメント有りの場合
                c99 = ActiveCell.Comment.Text  'コメント内容
                If Left(c99, 1) = "=" Then
                    If MsgBox("セルをコメントに掲載の式:" & vbCrLf & c99 & vbCrLf & "に、置き換えていいですか？", 289, "事前かくにん") = vbOK Then 'ok時
                        c99 = Replace(c99, "ｦｦ", "=")
                        ActiveCell.Value = c99
                        ActiveCell.Replace What:="ｦｦ", Replacement:="=", LookAt:=xlPart, _
                            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                            ReplaceFormat:=False
                    End If
                Else
                    MsgBox "実施対象外"
                End If
            End If
        End If
    Else  '本来
        '上側からこちらに引っ越し　86_014e
        If bfshn.Cells(1, 3) = "" Then
            Call oshimai("", bfn, shn, 1, 3, "複写名を入力して下さい(紫色セル)")
        ElseIf Len(bfshn.Cells(1, 3)) > 14 Then
            Call oshimai("", bfn, shn, 1, 3, "複写名は14文字以内に収めて下さい(" & Len(bfshn.Cells(1, 3)) & ")")
        End If
                
        'Aa：当シート、ここから（メイン部）
        Selection.Copy
        'Aa：データ開始行採取（後のフィルタリング、ウィンドウの固定のため）
        j = 1
        Do Until bfshn.Cells(j, 1).Value = 1 Or bfshn.Cells(j, 1).Value = "all1"
            j = j + 1
            If j = 200 Then
                MsgBox "「1」が見つからないようです"
                Exit Sub
            End If
        Loop
        If bfshn.Cells(j, 1).Value = 1 Then k = j Else k = j + 1 'kはデータ開始行(サンプル行ではなくなった)
        'Aa：当シート、ここまで
        'Ab：新シート、ここから
        Worksheets.Add  '新シート生成
        ActiveSheet.Name = shemei
    
        Range(Cells(2, 3), Cells(2, 3)).Select
        'Ab：貼り付け（値と数値の書式で貼付）
        Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats '85_24検証3から12→数式と数値の書式で貼付 へ
        
        'Ab：貼り付け（コメントを貼付）30s79
        Selection.PasteSpecial Paste:=xlPasteComments  '-4144
        'Ab：書式を貼り付け
        Selection.PasteSpecial Paste:=xlPasteFormats '-4122
        
        'Ab：罫線編集
        Selection.Borders.Color = RGB(191, 191, 191) '←= -4210753 　　30s86_017p
            
        '85_024検４当シート
        If gg1 < k Then
            If gg1 > sr(8) Then x = gg1 Else x = sr(8) + 1 'まずここでx使用(値のコピー開始行)
            bfshn.Select
            Range(Cells(x, dd1), Cells(k - 1, dd2)).Copy
            
            'Bbop：新シート
            Workbooks(bfn).Sheets(shemei).Select
            Range(Cells(x - gg1 + 2, 3), Cells(x - gg1 + 2, 3)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats  '12 'Bb1op：貼り付け（値と数値の書式で貼付）
        End If
            
        Range(Cells(1, 1), Cells(1, dd2 - dd1 + 3)).Select  '一行目セルうす緑に
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(179, 255, 231) 'うす緑　元15204275　30s86_017p
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
        Range(Cells(1, 2), Cells(1, dd2 - dd1 + 3)).Select  '一行目unicode文字色ほぼ透明に
        With Selection.Font
            .Color = RGB(226, 239, 218)  '元-2428958 30s86_017p
            .TintAndShade = 0
        End With
    
        Cells.FormatConditions.Delete      '条件付き書式解除
    
        x = k - gg1 + 1 'Ab：x：転載シートのフィルタリング基準行

        If gg2 < k - 1 Then x = gg2 - gg1 + 1 + 1 '特例パターン
        
        If x > 0 Then  'x>0でないパターンはないかと。ｘは２以上
            Range(Cells(2, 1), Cells(x, 2)).Select  'Ab：1列2列タテ上部をうす緑に
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(179, 255, 231) 'うす緑　'元15204275　 30s86_017p
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        
        Selection.ColumnWidth = 2  'Ab：1列2列幅調整
        Range(Cells(1, 2), Cells(1, 2)).Select
        Cells(1, 2).Value = "."   'if文撤廃、30s74
    
        x = k - gg1 + 1 'Ab：x：転載シートのフィルタリング基準行(再)
    
        'Ab：新シート、ここまで
        If bfshn.Cells(1, 5).Value <> "" Then '不要かと625
            'dw = ""  '625以降新バージョン
            '↓30s86_019j、30s86_019m
            If bfshn.Cells(2, 3).Value <> Left(twn, 7) And bfshn.Cells(2, 3).Value <> syutoku() & "r" And bfshn.Cells(2, 3).Value <> "" Then
                dw = bfshn.Cells(2, 3).Value  '625以降新バージョン
            Else 'これまでのパターン
                dw = passwordGet(10)
            End If
            
            UserForm2.Show vbModal 'PWありボタン押した場合、dwにPWが格納される。
            wd = dw
            If Application.Version < 16 And fmt = "csv" Then
                Call oshimai("", bfn, shn, 1, 0, "excel2013以前はcsv(utf-8)で保存はできません)")
            End If
        End If
        'Ba：当シート、左端タテ系部
        bfshn.Select
    
        If x > 1 And k - gg2 < 1 Then  'k - gg2 < 2→1 (項目行下端は対象外に)
    
            Range(Cells(gg1, 1), Cells(gg2, 2)).Copy     '1列2列タテ大雑把コピー(all1とカウンタcの所)
    
            'Bbop：新シート
            Workbooks(bfn).Sheets(shemei).Select '１列２列大雑把貼り付け(装飾系あとで)
            Range(Cells(2, 1), Cells(2, 1)).Select
            'Bb1op：貼り付け（値と数値の書式で貼付）
            Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats  '12
            With Selection.Font
                .Color = RGB(128, 128, 128) '30s86_017p ねずみ色に(=8421504)
                .TintAndShade = 0
            End With
    
            'Bb2op：オートフィルタ設定
            Rows(x).AutoFilter
            
            'Bb2op：ウィンドウ枠の固定
            Range(Cells(x + 1, 3), Cells(x + 1, 3)).Select
            ActiveWindow.FreezePanes = True
        
            Range(Cells(2, 1), Cells(x, 2)).ClearContents   '30s74　Bb2op：1列2列タテ上部をうす緑に
            
            Cells(x, 1).Value = "項目"
            Cells(x, 2).Value = "ｃ：集計対象　　"
        
            Range(Cells(x, 1), Cells(x, 2)).Select
            With Selection 'Bb2op：1列2列の項目行は上寄せ
                .VerticalAlignment = xlTop
                .Orientation = -90
            End With
            Selection.Font.Size = 9 'Bb2op：1列2列の項目行フォント調整
       
            If x > 3 Then
                Range(Cells(2, 1), Cells(x - 1, 1)).Select 'Bb2op：1列2列の上部フォントほぼ透明化
                With Selection.Font
                    .Color = RGB(179, 255, 231) '←　= -1572941　30s86_017p
                    .TintAndShade = 0
                End With
            End If
    
            If gg1 < sr(0) + 6 Then  'subtotalコピペ　N/A対応もコピペへ
                
                'subtotal事前
                'i(mghzコピペ開始用)定義(gg1：選択範囲開始行に左右される)
                
                If gg1 = sr(0) + 5 Then 'Bb2op 30s77改良
                    i = sr(0) + 4 '3行特殊
                    jj(1) = 1
                ElseIf gg1 < sr(0) + 5 Then
                    i = sr(0) + 3  '４行標準
                    jj(1) = i - gg1 + 2 + 1
                End If
        
                'subtotal本番、Bb2op_opC
                
                For j = mghz + 1 To mghz Step -1  '85_024検証9
                    'Bb2op_opCa：ここから、当シート(subtotalコピペ)
                    bfshn.Select     'B：右端subtotalコピー
                    Range(bfshn.Cells(i, j), bfshn.Cells(sr(0) + 6, j)).Copy
                
                    'Bb2op_opCb：ここから、新シート（subtotalコピペ）
                    Workbooks(bfn).Sheets(shemei).Select

                    If gg1 > sr(0) + 3 Then
                        Range(Cells(1, 2), Cells(1, 2)).Select 'はみ出しペースト
                    Else
                        Range(Cells(i - gg1 + 2, 2), Cells(i - gg1 + 2, 2)).Select '標準ペースト
                    End If
                    ActiveSheet.Paste
                
                    Range(Cells(jj(1), 2), Cells(jj(1) + 1, 2)).Columns.AutoFit
                
                    For jj(0) = jj(1) To jj(1) + 1
                        jj(2) = Range(Cells(jj(0), 2), Cells(jj(0), 2)).Font.Color
                        
                        Range(Cells(jj(0), 2), Cells(jj(0), 2)).Copy
                                                                
                        For ii = jj(1) To jj(1) + 1
                            For kk = 3 To dd2 - dd1 + 3
                                If WorksheetFunction.IsErr(Cells(ii, kk).Value) Then 'N/A条件追加,DIV/0条件追加（IsNA→IsErr）30s86_020x
                                    If Cells(ii, kk).Font.Color = jj(2) Then
                                        Range(Cells(ii, kk), Cells(ii, kk)).Select
                                        ActiveSheet.Paste
                                    End If
                                ElseIf Cells(ii, kk).Font.Color = jj(2) And Cells(ii, kk) <> "" Then '30s77null条件追加(従来)
                                        Range(Cells(ii, kk), Cells(ii, kk)).Select
                                        ActiveSheet.Paste
                                End If
                            Next
                        Next
                    Next
                Next
            'Bb2op_opCここまで
            End If
        'Bb2opここまで
        End If
        
        Range(Cells(2, 1), Cells(2, 1)).Select
        'Bb：新シート、ここまで
    
        'D：ここから、当シート(体裁調整)
        'Da
        bfshn.Select
    
        Range(Cells(1, 1), Cells(sr(0) + 5, mghz - 1)).Select  'Da1：1-18行(標準)を調整
        With Selection.Font  'Da：フォント調整
            .Name = "ＭＳ Ｐゴシック"  'Da：←これでWin10影響(メイリオなど)影響受けなくなる。
            .Size = 9
        End With
    
        Range(Cells(sr(0) + 3, 5), Cells(sr(0) + 3, mghz - 1)).Select 'Da1：標15行調整
        With Selection.Font  'Da：フォント、色調整
            .Name = "Haettenschweiler"
            .Size = 10
        End With
        Selection.NumberFormatLocal = "m/d h:mm"
    
        Range(Cells(sr(0), 5), Cells(sr(0), mghz - 1)).Select  'Da1：標12行調整
        With Selection.Font  'フォント、色調整
            .Name = "Haettenschweiler"
            .Size = 10
        End With
        Selection.NumberFormatLocal = "m/d h:mm"
        
        'Da1：前歴集計値調                                              '赤が出る事あり注意↓
        Range(Cells(sr(0) + 1, 5), Cells(sr(0) + 2, mghz - 1)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        'Da1：前歴集計値調                                              '赤が出る事あり注意↓
        Range(Cells(sr(0) + 4, 5), Cells(sr(0) + 5, mghz - 1)).NumberFormatLocal = "#,##0;[赤]-#,##0"
    
        '↓86_014m
        Range(Cells(sr(0), 3), Cells(sr(0) + 1, 5)).Select 'Da1：赤セル部フォント再調整1＆2
        With Selection.Font
            .Name = "ＭＳ Ｐゴシック"
            .Size = 9
        End With
        Selection.NumberFormatLocal = "G/標準" '←withに入れるとエラーになる。
    
        'Da2：当シート列幅をコピー
        Range(Cells(2, dd1), Cells(2, dd2)).Copy '列幅コピる（2行目にて）
        
        Range(Cells(gg1, dd1), Cells(gg1, dd1)).Select  '開始セル（：左上）
        dg1 = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) '開始セル転記
        Range(Cells(gg2, dd2), Cells(gg2, dd2)).Select  '終了セル（：右下）
        dg2 = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) '終了セル転記
        Range(Cells(gg1, dd1), Cells(gg2, dd2)).Select '選択範囲は戻す
        'Da：当シート、ここまで
        
        'Db：ここから、新シート（列幅のみコピペ）
        Workbooks(bfn).Sheets(shemei).Select
        
        Range(Cells(1, 3), Cells(1, 3)).PasteSpecial Paste:=xlPasteColumnWidths '8'列幅を貼り付け
    
        If x > 1 And k - gg2 < 2 Then  'Dbop_Unicodeの貼り付け(項目行(k-1)が含まれていることが貼り付けの条件)
        'Dbop：x：フィルタリング基準行
            Range(Cells(1, 3), Cells(1, 3)).PasteSpecial Paste:=xlPasteValues '-4163
        End If
        
        'Db
        Cells(1, 1).Value = "複.　複写、" & twn & "、" & shemei & "、ｦ" & shn _
        & "ｦ" & bfn & "、" & dg1 & dg2 & "、列固有名"
        Range(Cells(2, 2), Cells(2, 2)).Select
    
        'ログ部ここから→kyosydouへ移設(Activate処理撤廃)
        j = 1
        Do Until Workbooks(twn).Sheets(shog).Cells(j, 1).Value = ""
            j = j + 1
            If j = 10000 Then
                MsgBox "空白行が見つからないようです"
                Exit Sub
            End If
        Loop

        Workbooks(twn).Sheets(shog).Cells(j, 1).Value = 1 '項目名
        Workbooks(twn).Sheets(shog).Cells(j, 2).Value = j  '項番
        
        Workbooks(twn).Sheets(shog).Cells(j, 3).Value = Workbooks(bfn).Sheets(shemei).Cells(1, 1).Value & "、z" & wd & "z" & "b" _
        & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "、" & twbsh.Cells(2, 2).Value & "、" & _
        Format(Now(), "yyyymmdd_hhmmss") & "、" & bfshn.Cells(sr(8), 5).Value & "、" & gg2 - gg1 + 1 & "、" & dd2 - dd1 + 1
        
        Workbooks(twn).Sheets(shog).Cells(j, 4).Value = Format(Now(), "yyyymmdd") 'date
        Workbooks(twn).Sheets(shog).Cells(j, 5).Value = Format(Now(), "yyyymmdd_hhmmss") 'timestamp
        If fimei = "" Then
            Workbooks(twn).Sheets(shog).Cells(j, 7).Value = shemei 'to
        Else
            Workbooks(twn).Sheets(shog).Cells(j, 7).Value = fimei & "." & fmt & "\" & shemei 'to
        End If
        Workbooks(twn).Sheets(shog).Cells(j, 8).Value = 9  '最右列(複写は固定値)
        Workbooks(twn).Sheets(shog).Cells(j, 9).Value = bfn & "\" & shn  'from
        
        'ログ部ここまで
 
        'シートを別ファイルとしても複製(E1セル情報有の時)30s81
        If bfshn.Cells(1, 5).Value <> "" Then  '抜本改良　30s86_017x~z
            '627メモ帳のPW→Excel新シートへ
            If Left(fmt, 3) <> "csv" Then 'csv系はPWシート作らない xlsxPWなしも一旦PWシートは作る(office2007にファイルのテーマ変えるため)
                
                ThisWorkbook.Activate
                Sheets("高速シート_" & syutoku()).Select
                Sheets("高速シート_" & syutoku()).Copy

'                Sheets("高速シート_" & syutoku()).Name = shemei & "のPW"
                Sheets("高速シート_" & syutoku()).Name = "仮ですaaa"    '30s86_021b
                Worksheets.Add
                ActiveSheet.Name = shemei & "のPW"
                
                Application.DisplayAlerts = False
                Worksheets("仮ですaaa").Delete
                Application.DisplayAlerts = True

                '86_020v　改良
                Cells(2, 1).Value = "★可能ならURLは、日付指定されたものに書き換える。"
                Cells(3, 1).Value = "　その方が親切"
                Cells(4, 1).Value = "☆そもそもマイドライブでない時は、下記文言類ごっそり消す。"
                Cells(5, 1).Value = "-------------------------------------------------------------------------------"
                Cells(6, 1).Value = "フォルダの場所〔Google・マイドライブ〕"
                Cells(7, 1).Value = "※ 本メール宛先の方だけに開示してます。"
                Cells(8, 1).Value = "https://drive.google.com/drive/folders/1Cz_vb9zUBAs-PurxPLIubX-cFu_heUDS"
                
                Cells(10, 1).Value = "フォルダ名：" & Format(Now(), "yyyymmdd")
                Cells(11, 1).Value = "ファイル名： " & fimei & "." & fmt
                Cells(12, 1).Value = "ファイルのＰＷ：" & wd
                
                Cells(14, 1).Value = "※フォルダは一時的な共有置き場です。一定期間経過後､適宜削除します。"
                Cells(15, 1).Value = "　ＤＬした上で使用して下さい。"
                Cells(16, 1).Value = "-------------------------------------------------------------------------------"
                
                Cells(18, 1).Value = "シート名：" & shemei
                Cells(19, 1).Value = "↑★必要なら、含める。"
                Range(Cells(2, 1), Cells(19, 1)).Select
'                If wd <> "" Then Shell "c:\windows\system32\notepad.exe", vbNormalFocus 'PW用メモ帳立ち上げ 86_020z　廃止へ

                '30s79追加、新ファイルのフォントを游ゴシックではなく、MSPゴシック仕様に
                If Application.Version < 16 Then
                    'MsgBox "excel2013以前です"
                Else                  'MsgBox "excel2016以降です"
                    If IsNumeric(syutoku()) Then
                        '↓レノボのPC
                        ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                        "C:\Program Files (x86)\Microsoft Office\Root\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml")
                    Else
                        '↓fmvのPC
                        'ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                            "C:\Program Files\WindowsApps\Microsoft.Office.Desktop_16051.12228.20364.0_x86__ああああ\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml" _
                            )
                        '20200115にRev上がったかと。まだこの頃は、store版
                        'ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                            "C:\Program Files\WindowsApps\Microsoft.Office.Desktop_16051.12325.20288.0_x86__ああああ\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml" _
                            )
                        '202004のRev　(store版からDL版へ)
                        ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                            "C:\Program Files (x86)\Microsoft Office\Root\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml" _
                            )
                    End If
                End If
            Else 'csv系
                Workbooks(bfn).Activate
                Sheets(shemei).Select
                Sheets(shemei).Copy
            End If
            
            'ファイルを所定の様式・所定のファイル名で保存
            If fmt = "xlsx" Then
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlOpenXMLWorkbook, Password:=wd, CreateBackup:=False    'xlsxフォーマット
            ElseIf fmt = "csv" Then
                Rows("1:1").Select
                Selection.Delete Shift:=xlUp
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlCSVUTF8, CreateBackup:=False                          'csv(utf-8)フォーマット
            ElseIf fmt = "csvsjis" Then
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlCSV, CreateBackup:=False                              'csvフォーマット
            Else
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlExcel12, Password:=wd, CreateBackup:=False            'xlsbフォーマット
            End If
            
            If Left(fmt, 3) <> "csv" Then
            
                'Eb　（従来：対象のシートを新規ファイルとしてコピーしていた）
                Workbooks(bfn).Activate
                Sheets(shemei).Select
                Sheets(shemei).Copy Before:=Workbooks(fimei & "." & fmt).Sheets(shemei & "のPW")
            
                Workbooks(fimei & "." & fmt).Activate
                Sheets(shemei & "のPW").Select
                Sheets(shemei & "のPW").Move 'PWシートは新ファイルとして移す
            
                If wd = "" Then
                    Sheets(shemei & "のPW").Name = "不要"  '※PWなしxlsx→捨てファイル
                Else '30s86_020a PWシートをUTF-8のテキストとして保存を追加(メモ帳立ち上げは廃止)
                    ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei & "のPW.txt", _
                    FileFormat:=xlUnicodeText, CreateBackup:=False
                End If
            
                Workbooks(fimei & "." & fmt).Activate
                ActiveWorkbook.Save      '上書保存
            
            End If
        End If
    End If
    Application.CutCopyMode = False
End Sub  '複写ここまで
