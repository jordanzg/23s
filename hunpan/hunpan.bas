Function hunpan(fn As String, f As String, ii As Long, g As Currency, e7 As Currency, e5 As Currency, c As Currency, e As Currency, mc As String, hk1 As String) As String
    Dim et5 As String, et As String
    '暗号鍵の文字列
    If mc = "1" Then 'mc・・６行op、鍵の種類
        et5 = "1"
    ElseIf mc = "e5" Then  '使われている　5行目カウント基準列にある文字列が鍵(a234567　とか)
        et5 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(e5))))
    ElseIf mc = "e" Then  '8行目転載列にある文字列が鍵(a234567　とか)　使われていない
        et5 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(e))))
    Else     '3行目転載列にある文字列が鍵(a234567　とか)　使われていない
        et5 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(g))))  '
    End If
    et = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(e))))  '暗号鍵確定
    hunpan = hunk2(c, et, et5, hk1)    '暗号鍵
End Function
