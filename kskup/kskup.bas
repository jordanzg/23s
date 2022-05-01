Function kskup(am1 As String, am2 As String, n1 As Long, n2 As Long, h As Long, b As Currency, c As Currency, k0 As Long, h0 As Long, pap2 As Long, er2() As Currency, spd As String, pqp As Long, e5 As Long, er3() As Currency, hiru As Variant) As Long  'じあ
    'p(kskup)判定新仕様、n3は近似用　,(引数)m2→n1へ リファラ：n3、n2,h0,k0, pqp追加624,e5&er3()追加629
    Dim m As Long, n3 As Long
    Dim tskosk As Long 'strconv用(定速：2(従来通り)、高速：26(New85_006))
    Dim n3a As Long
    '◆高速側(純・近似)専用。低速側はtskupへ
    tskosk = 10 + hrkt '30s86_020s
    kskup = 0 'リセット(不要だが)
    n3 = 0 'ゼロスタート
    krpm2 = 0 '高速ロック＆p=-2判定フラグ

    If h < k Then '当シートに何も無い場合(初回のみ通過ゾーン) kあ
        If c < 0 Then
            kskup = -1    '-1-2この時点でexitdo(p:-1とする。)
            MsgBox "表空白の-1-2です(p=-1,exitdo、このまま終了されます)。"
        Else  '※c>=0が前提となる。
            kskup = 2
            n2 = h + 1
            
            If er2(0) < 0 Then  '高速ロックオン(表空白)です。"
                UserForm4.StartUpPosition = 2 '1　エクセルの中央、　2　画面の中央、　3　画面の左上
                UserForm4.Show vbModeless
                UserForm4.Repaint
                bfshn.Cells(sr(2), 5).Value = "純高ﾛｯｸ(表空白)" 'そのうち
                If pap2 = 0 And Abs(e5) < 1 And UBound(er3()) > 0 Then 'p=-2の判定　６２９
                    For jj = 1 To UBound(er3())
                        If er3(jj) = 0.1 Then krpm2 = 1  'er3(jj)<0→=0.1 へ
                    Next
                    If krpm2 = 1 Then kskup = -2 'p=-2の確定　６２９
                End If
                Unload UserForm4
                UserForm1.Repaint
            End If
        End If
    ElseIf StrConv(am2, tskosk) = StrConv(am1, tskosk) Then '(アドバラピド)前回一致
        kskup = 1
        n2 = n1
    ElseIf spd = "純高速" And (k0 > h0 Or pqp = 1) Then  '(アドバラピド)高速ロックオン状態　pqp追加624
        If c < 0 Then
            kskup = -1    '-1-2この時点でexitdo(p:-1とする。)
            MsgBox "ここはもう通らないはず(-1-2)以後当シート側情報ないです(p=-1)。"
        Else
            kskup = 2
            n2 = h + 1
        End If
    '(アドバラピド)次行一致 c = Round(c, 0)追加85＿026（：-15.１実施しない、-15実施する）
    ElseIf Round(c) <> -1 And Round(c) <> -2 And c = Round(c, 0) And (spd = "純高速" Or spd = "ノーマル") And n1 < h0 And _
            StrConv(am2, tskosk) = StrConv(bfshn.Cells(n1 + 1, Abs(b)).Value, tskosk) Then
        kskup = 1
        If b < 0 Then '純高速時
            n3 = n1 + 1 'n2は今回突合処理対象行(仮)
            n2 = hiru(n3, 2) '近似の読み替え ８５＿027検証ｃ
            k0 = n3
        Else 'ノーマル
            n2 = n1 + 1
        End If
    Else
        'p=0
    End If
    
    If kskup = 0 Then 'まだ決まらず(p=0)→マッチング実施
        If k0 > h0 Then Call oshimai("", bfn, shn, 1, 0, "k0>h0でmatchにいくようなことはあってはならない。")
        If IsError(Application.Match(StrConv(am2, tskosk), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 1)) Then '近似でもエラーは発生する。
            m = 0
        Else
            m = Application.WorksheetFunction.Match(StrConv(am2, tskosk), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 1) '仮策定
            If StrConv(am2, tskosk) <> StrConv(bfshn.Cells(k0 + m - 1, Abs(b)), tskosk) Then '一致してなければ二段階へ
                m = 0
            Else 'ok近似実施
                kskup = 1
                n3 = k0 + m - 1 'n2は今回突合処理対象行(仮)
                n2 = hiru(n3, 2) '近似の読み替え ８５＿027検証ｃ
                If spd = "純高速" Then k0 = n3
            End If
        End If
        If spd = "純高速" And m = 0 Then '新規で純高速は当処理行う
            kskup = 2
            n2 = h + 1
        End If
    End If
End Function
