Sub oshimai(msx As String, fn As String, ff As String, ii As Long, aa As Long, msbv As String)
         '↑msxは使われてないですね→再使用へ　201912
    Unload UserForm1
    Unload UserForm3
    Workbooks(fn).Activate '30s76追加（外結ログ対策）
    DoEvents
    Worksheets(ff).Select  '30s76追加（外結ログ対策）
    If aa > 0 And ii > 0 Then
        Workbooks(fn).Sheets(ff).Cells(ii, aa).Select
        If msx = "" Then
            If msbv <> "" Then MsgBox msbv & vbCrLf & "(選択セル)"
        Else
            If msbv <> "" Then MsgBox msbv & vbCrLf & "(選択セル)", 289, msx
        End If
    Else
        If msx = "" Then
            If msbv <> "" Then MsgBox msbv
        Else
            If msbv <> "" Then MsgBox msbv, 289, msx
        End If
    End If
    '名前の定義の削除
    Dim nmn As Name
    For Each nmn In ActiveWorkbook.Names
        On Error Resume Next  ' エラーを無視。
        nmn.Delete
    Next
    'フィルタ再
    If k > 1 Then bfshn.Rows(k - 1).AutoFilter         '一度つけて、
    
    Application.Calculation = xlCalculationAutomatic  '再計算自動に戻す
    Application.StatusBar = False
    If aa > 0 And ii > 0 Then Workbooks(fn).Sheets(ff).Cells(ii, aa).Select
    Application.Cursor = xlDefault
    End
End Sub
