Function wetaiou(fn As String, f As String, ii As Long, erx() As Currency, kg2 As String, spd As String, mrx() As String, gyu As Long) As String
    'ファイル名、シート名、行数、列数(複数)、区切り文字、高速可否、列第二因子(複数)、２行目か３行目か
    Dim am2 As String, am3 As String, qap As Long, jj As Long
    If fn = "" Then '30s86_020a
        If UBound(erx()) > 0 Then 'ヱがあるときのみ（二発目以降）。単文節は通らない。
            jj = UBound(erx())   '2.11(従来通り）　３もある？
            For qap = 1 To jj
                am2 = am2 & kg2 & am3
            Next
        End If
    Else '既存ver
        'ヱ複数列の返り値を返す(○ヱ▽ヱ◇形式)。複数列で無くても常時使われる(am2)
        If Abs(erx(0)) < 1 Then
            If erx(0) > 0.3 Then
                If mrx(0) = "" Then '第二因子あればそちらの記載優先で　85_027検証8
                    am2 = "" '単数列はここはあり得ない。複文節でかつ列指定無い時(0.4)が該当。
                Else
                    am2 = mrx(0)  '85_027検証8　３行目第二因子
                End If
            End If
        ElseIf spd = "純高速" Or spd = "近似高速" Then  '単文節時or複数列一発目(0番目)高速
            am2 = CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(0))).Value)  '高速時はtrim実施しないで揃えるへ　85＿027検証6
        Else '単文節時or複数列一発目(0番目)低速 ※返り値null(セルが空白返り値)の場合もある。
            am2 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(0))).Value))
        End If
    
        If UBound(erx()) > 0 Then 'ヱがあるときのみ（二発目以降）。単文節は通らない。
            jj = UBound(erx())   '2.11(従来通り）　３もある？
            For qap = 1 To jj
                If Abs(erx(qap)) < 1 Then
                    If erx(qap) < 0.3 Then
                        am3 = Format(ii, "0000000") '「ー」の時(0.1)、行番号をキーとする。
                    ElseIf mrx(qap) = "" Then '(0.4)　New投入　85_027検証8
                        am3 = "" '従来(0.4)※３行目第二因子なし　※従来型
                    Else
                        am3 = mrx(qap)  '※３行目第二因子あり　※New　85_027検証8
                    End If
                ElseIf spd = "純高速" Or spd = "近似高速" Then    '高速時はtrim実施しないで揃えるへ　85＿027検証6
                    am3 = CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(qap))).Value)
                Else  '低速は従来通り
                    am3 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(qap))).Value))
                End If
                am2 = am2 & kg2 & am3
            Next
        End If
    End If
    wetaiou = am2
End Function
