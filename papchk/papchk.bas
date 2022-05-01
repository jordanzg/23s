Sub papchk(pap() As Long, nmb As Long, bun As Long)
    Dim ii As Long
    For ii = 1 To bun
    '    If pap(nmb, 0) <> pap(nmb, ii) Then MsgBox "pap(nmb," & ii & ")不一致"  '当面、nmbは不活性で
        If pap(2, 0) <> pap(2, ii) Then MsgBox "pap(2," & ii & ")不一致"
    Next
End Sub
