Function yhwat1(fn As String, f As String, ee As Currency, yhs As Long, a As Long, bni As Long, kg1 As String, nkg As Long, kg2 As String, bugyo As String) As String
    '対_ファイル名、対_シート名、対_項目行、当_行番号、当_列番号、文節数、区切り文字(ゐ)、無し区切りビット、区切り文字(ヱ)、その文節の原本情報（＊も含む）
    '仮想キー　◇ヱ◇ヱヱ◇対応30s43(doloop→fornextに）
    Dim ymj1 As String, ymj2 As String, mh As Long, yap As Long '←ヱの個数
    Dim pm As Long, ii As Long, wenk As Long  'ー無有　、ヱの順番目(kgwe→ii)　、ヱの区切り可否
    Dim chwat1 As String, chwat2 As String '返り値の区切り毎生成物　、返り値の素（蓄積型）
    
    If kg2 = "" Then wenk = 1 'wenk=1：ヱ区切り無し、wenk=0 ：ヱ区切り有り
    pm = 1
    If Right(fn, 4) = ".xls" Then mh = 256 Else mh = 2000 '85_009　.xlsにも対応
    ymj1 = rvsrz3(bugyo, 2, "ｦ", 2) 'nkg：２、色んな行で＊あヱい　形式（＊ありヲ不使用）を許容（30s73記）
    If bni >= 2 And bugyo = "" Then  '30s75 ymj1→bugyo に(第二文節以降がｦｦの時、前節情報踏襲になるバグ対処)
        yhwat1 = ""   '←yhwat1が文字関数にならざるを得ない根源
    Else 'bni=1 のとき、若しくは、ymj1情報ありのとき
        yap = kgcnt(ymj1, kg2)  '←ヱの個数
        For ii = 1 To yap + 1
            pm = 1  '30s81_7追加（バグ、pmはfor毎にリセットしなければならない）
            ymj2 = rvsrz3(ymj1, ii, kg2, wenk)
            If Mid(ymj2, 1, 1) = "ー" Then
                pm = -1
                If ymj2 = "ー" Then ymj2 = "" Else ymj2 = Mid(ymj2, 2)
            End If
            chwat1 = ""
            If IsNumeric(ymj2) Or IsDate(ymj2) Then
                chwat1 = ymj2
            ElseIf ymj2 = "" Then
                chwat1 = pm * 0.4  'ｦｦ→戻り値0.4、ｦーｦ→戻り値-0.4が入る
                If chwat1 = -0.4 Then chwat1 = 0.1 '85_024 -0.4→0.1
            ElseIf ee > 0 Then
                If IsError(Application.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)) Then
                    Call oshimai("", bfn, shn, yhs, a, "（処理中止）" & vbCrLf & "シート名：" & f & " 上の" & vbCrLf & "項目名「" & ymj2 & "」が見つかりません")
                Else
                    chwat1 = pm * Application.WorksheetFunction.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)
                End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "「項目名」が「なし」で文字検索しています")
            End If
            If ii = 1 Then chwat2 = chwat1 Else chwat2 = chwat2 & kg2 & chwat1
        Next
        yhwat1 = chwat2
    End If
End Function
