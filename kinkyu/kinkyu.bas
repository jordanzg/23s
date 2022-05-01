Sub kinkyu()  'これがある事で強制終了、データ消滅を阻止できているので、これは消さない。
    On Error GoTo myError
    bfn = ActiveWorkbook.Name 'bfn,shnはパブリック
    shn = ActiveSheet.Name
    '↓この2行でエラーが起こる
    ThisWorkbook.Activate
    Sheets("▲集計_雛形").Activate
    DoEvents
    Workbooks(bfn).Activate          'こちらへ（いかが？）
    Sheets(shn).Activate
    DoEvents
    Exit Sub
myError:
    MsgBox "エラーです。終わります。"
    End
End Sub
