Function iptfg(jmj As String, czc As Long, ww As String) As String
    Do Until jj >= czc '←>を入れてるのは無限ループ防止
        fzf = tzt
        tzt = InStr(fzf + 1, jmj, ww)
        jj = jj + 1
    Loop
    If tzt = 0 Then tzt = Len(jmj) + 1 '4オク対応
    iptfg = Mid(jmj, fzf + 1, tzt - fzf - 1)
End Function
