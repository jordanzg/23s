Function rvsrz3(cef As String, bni As Long, kgr As String, nkg As Long) As String '30s03
    '対象文字列(A:*ｦあｦい)、文節数、区切り文字、区切りしないフラグ 変数二文字化　86_014k
    'bni=0 (文節ゼロ)→そのままフルで返す
    Dim ipc As Long, bb As Long, cc As Long, prd As String
    If kgr = "" Then
        'Call oshimai("", bfn, shn, 1, 0, "このケース、あるのかな？")　'→ありました
        prd = "、"
    Else
        prd = kgr
    End If
    If nkg = 1 Then
        rvsrz3 = cef  '区切りなしフラグ有効→そのまま返す
    'nkg=2→ｦなし時の(bni)2文節(ｦ◇ｦ)対処ルーチン。※ｦの解析の時(sr)しか使われない。
    ElseIf bni = 2 And nkg = 2 And StrConv(Left(cef, 1), 8) <> "*" And InStr(1, cef, "ｦ") = 0 Then 'nkg:2→ｦ対応　あ（ヲ無し）　bni2限定に,色んな行で使用 30s81接頭ど対応
        rvsrz3 = cef 'bni=1→2　30ｓ61
    ElseIf bni = 2 And nkg = 2 And StrConv(Left(cef, 1), 8) = "*" And InStr(1, cef, "ｦ") = 0 Then  '＊あ（ヲ無し）、＊　bni2限定に,色んな行で使用 30s81 接頭ど対応
        If Len(cef) > 1 Then
            If StrConv(Mid(cef, 2, 1), 8) = "*" Then '**
                If Len(cef) = 2 Then
                    Call oshimai("", bfn, shn, 1, 0, "「**」は使用されないです。" & vbCrLf & "「**あ」形式でよろしく")
                Else 'New 建設中　＊＊あ
                    rvsrz3 = Mid(cef, 3) '＊＊あ　→あ」を返す
                End If
            Else '従来型
                rvsrz3 = Mid(cef, 2) '＊あ　→あ」を返す
            End If
        Else
            rvsrz3 = cef '＊
        End If
    Else 'nkg=0, nkg=2：ヲ有り（先頭＊含む）
        If bni > 0 Then
            Do
                ipc = ipc + 1
                bb = cc
                cc = InStr(bb + 1, cef, prd)
            Loop Until cc = 0 Or ipc = bni 'それ以降該当なしor規定文節到達で抜ける。
        End If
        If cc = 0 And ipc = bni Then  'ジャスト規定文節で区切無しに　０文節もこちらに入る。
            rvsrz3 = Mid(cef, bb + 1)
        ElseIf cc = 0 And ipc < bni Then  '規定文節未到達(よって該当文節は"")
            rvsrz3 = ""
        Else '規定文節で区切り文字もまだある。
            rvsrz3 = Mid(cef, bb + 1, cc - bb - 1)
        End If
    End If
End Function
