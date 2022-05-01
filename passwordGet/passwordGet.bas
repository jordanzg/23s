Function passwordGet(knt As Integer) As String
    Dim ii As Integer, aa As Integer, bb As String
    'PW文字数は２文字以上ないとエラーになる。 変数二文字化　86_014k
    For ii = 4 To knt
        aa = 0
        Do Until aa = 1 'Randomize
            bb = rndchr("nsuu")
            If InStr(1, passwordGet, bb) = 0 Then aa = 1
        Loop
        passwordGet = passwordGet & bb
    Next ii
End Function
