Sub kyosydou()
    '共通（外結、初回、複写）共通の初動をまとめる。
    kinkyu

    Dim ii As Long, nm As Variant
    Dim xsheet As Worksheet 'xWsheet→xsheet

    Application.CutCopyMode = False
    DoEvents
    gg1 = Selection.Row    '選択開始行
    gg2 = Selection.Rows(Selection.Rows.Count).Row   '選択終了行
    dd1 = Selection.Column    '選択開始列
    dd2 = Selection.Columns(Selection.Columns.Count).Column    '選択終了列
    
    If dd1 = 0 Then MsgBox "dd1がゼロですね"
    bfn = ActiveWorkbook.Name 'bfn,shnはパブリック
    shn = ActiveSheet.Name
    Set bfshn = Workbooks(bfn).Worksheets(shn) '30ｓ63より。当シート・当ファイル
    
    twn = ThisWorkbook.Name 'マクロファイル名そのもの(～.xlsm)
    Set twbsh = Workbooks(twn).Worksheets("▲集計_雛形") '30ｓ74より
    
    nm = Array("", "対象ｼｰﾄ名", "当：突合列", "対：突合列", "対：ｵｰﾙ1列", "対：ｶｳﾝﾄ列", "対：加算列･他", "当：転載列", "対：転載列", "対：実質加算列", "当：実質転載列") '30s63,array化
    If IsError(Application.Match(nm(1), Range(bfshn.Cells(1, 2), bfshn.Cells(200, 2)), 0)) Then Call oshimai("", bfn, shn, 4, 2, "（処理中止）「" & nm(1) & "」が見つかりません")
    For ii = 1 To 8
        sr(ii) = WorksheetFunction.Match(nm(ii), Range(bfshn.Cells(1, 2), bfshn.Cells(200, 2)), 0)
        If sr(0) < sr(ii) Then sr(0) = sr(ii)
    Next
    sr(0) = sr(8) + 1  '下側から移設
    
    If IsError(Application.Match("。", Range(bfshn.Cells(1, 1), bfshn.Cells(1, 5000)), 0)) Then
        Call oshimai("", bfn, shn, 1, 0, "「" & shn & "」シート右上に「。」がありません。入れて下さい。")
    Else
        mghz = Application.Match("。", Range(bfshn.Cells(1, 1), bfshn.Cells(1, 5000)), 0)
    End If
    shog = "log_" & syutoku() & "_" & Format(Date, "yyyymm")
    
    'ログシート有無chk、30s82、初回のみ→ここに移設
    For Each xsheet In ThisWorkbook.Sheets
        If xsheet.Name = shog Then xFlag = True 'boolean型の初期値はfalse
    Next xsheet
    If xFlag = True Then ' 該当のシートがある場合の処理
        '（何もしない）
    Else ' 該当のシートがない場合の処理 '
        Workbooks(twn).Activate '30s83
        Worksheets.Add
        ActiveSheet.Name = shog
        nm = Array("", "項目名", "項番", "log", "date", "timestamp", "メモ", "to", "最右列", "from9") '30s83,array化
        For ii = 1 To 9
            Workbooks(twn).Sheets(shog).Cells(1, ii).Value = nm(ii)
        Next
        Workbooks(bfn).Activate          'こちらへ（いかが？）
        bfshn.Select
    End If
    
    xFlag = False
    For Each xsheet In ThisWorkbook.Sheets     '転記有無chk、30s82e
        If xsheet.Name = "高速シート_" & syutoku() Then xFlag = True
    Next xsheet

    If xFlag = True Then ' 該当のシートがある場合の処理
        Set twt = Workbooks(twn).Worksheets("高速シート_" & syutoku()) '30s82f
        twt.Cells.Clear
    
        ThisWorkbook.Activate
        twbsh.Select '30s86_017q　変なエラー解消用(フォントをテーマの色にすると再起動する件)
                                  '↑100%解消している訳ではない
        Sheets("高速シート_" & syutoku()).Activate
        Rows("6:6").Select
        Selection.Delete Shift:=xlUp
        DoEvents
    
        Range(Cells(2, 3), Cells(2, 3)).Select
        Workbooks(bfn).Activate       '86_017h
        bfshn.Select
        DoEvents
    Else ' 該当のシートがない場合の処理 '
        Worksheets.Add
        ActiveSheet.Name = "高速シート_" & syutoku()
        Set twt = Workbooks(twn).Worksheets("高速シート_" & syutoku()) '30s82f
        Workbooks(bfn).Activate
        bfshn.Select
    End If
    
    bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1 '30s83 =sum(a:a)+1からこちらへ
    Exit Sub
End Sub
