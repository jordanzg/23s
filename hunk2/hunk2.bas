Function hunk2(cc As Currency, et As String, et5 As String, hk1 As String) As String
    '暗号復号、対象文字、鍵、エラービット 変数二文字化　86_014k
    Dim mjs As Long '転載文字の文字数
    Dim gjs As Long '突合文字の文字数
    Dim pp As Long, vv As Long
    Dim hh As String '変換用一文字
    Dim uu As String '変換後文字
    Dim tt As Long '変換unicodeシフト定数
    Dim jj As Long '範囲チェック
    Dim qq As Long '余り
    Dim mb As Long 'まぶし係数
    Dim ww As Long, tn As Long, ss As Long

    uu = ""
    tn = 0
    gjs = Len(et5) '暗号鍵の文字数
    mjs = Len(et) '暗号対象文字の文字数
    vv = 0
    For pp = 1 To gjs
        ss = AscW(Mid(et5, pp, 1))
        vv = vv + ss
    Next
    ww = vv Mod 10  '暗号鍵のまぶしバイアス値
    vv = 0
    
    For pp = 1 To mjs 'ppは転載文字の、ある文字目
        qq = (pp - tn) Mod gjs + 1  '暗号鍵の文字数循環
        ss = AscW(Mid(et5, qq, 1))
        mb = (ss + pp - tn + ww) Mod 50 'まぶし係数(０～＋－４９)
        tt = 25000 - 30 + 99 * mb '仮のtt
        If cc = -5 Then
            jj = 0
            tt = tt
        ElseIf cc = -7 Then
            jj = tt
            tt = (-1) * tt
        End If
        If Mid(et, pp, 1) = "、" Then
            hh = "、"
            tn = pp
        ElseIf (AscW(Mid(et, pp, 1)) >= jj + 32 And AscW(Mid(et, pp, 1)) <= jj + 126) Then
            hh = ChrW(AscW(Mid(et, pp, 1)) + (tt))
        Else
            vv = 1
        End If
        If vv = 0 Then uu = uu & hh
    Next

    If vv = 0 Then
        hunk2 = uu
    Else
        hunk2 = "規定外文字あり：" & et
        hk1 = "1"
    End If
End Function
