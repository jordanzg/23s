Function tskup(am1 As String, am2 As String, n1 As Long, n2 As Long, h As Long, b As Currency, c As Currency, k0 As Long, h0 As Long, pap2 As Long, er2() As Currency, spd As String, pqp As Long, e5 As Long, er3() As Currency) As Long  'じあ
    'p(tskup)判定新仕様、n3は近似用　,(引数)m2→n1へ リファラ：n3、n2,h0,k0, pqp追加624,e5&er3()追加629
    Dim m As Long, n3 As Long
    Dim tskosk As Long 'strconv用(定速：2(従来通り)、高速：26(New85_006))
    '◇低速側専用
    tskosk = 2  '大文字→小文字化
    tskup = 0 'リセット(不要だが)
    n3 = 0 'ゼロスタート
    krpm2 = 0 '高速ロック＆p=-2判定フラグ
    If h < k Then  '当シートに何も無い場合(初回のみ通過ゾーン) kあ
        If c < 0 Then
            tskup = -1    '-1-2この時点でexitdo(p:-1とする。)
            MsgBox "表空白の-1-2です(p=-1,exitdo、このまま終了されます)。"
        Else  '※c>=0が前提となる。
            tskup = 2
            n2 = h + 1
            h0 = n2 '(=k0,高速時ロックオン) h0処理は基本tskupより後、ここだけ特例(hk逆転イレギュラー解消措置
            If er2(0) < 0 Then
                MsgBox "こちらは低速専用になりました。高速でこちら通るのはおかしい。"
            End If
        End If
    ElseIf LCase(am2) = LCase(am1) Then '(アドバラピド)前回一致　'←86_016q(uni対策)　StrConv(am2→LCase(am2)
        tskup = 1
        n2 = n1
    ElseIf spd = "純高速" And k0 = h0 And b < 0 And pqp = 1 Then '(アドバラピド)高速ロックオン状態　pqp追加624
        MsgBox "こちらは低速専用になりました。純高速でこちら通るのはおかしい。"
    
    '(アドバラピド)次行一致 c = Round(c, 0)追加85＿026（：-15.１実施しない、-15実施する）
    ElseIf Round(c) <> -1 And Round(c) <> -2 And c = Round(c, 0) And (spd = "純高速" Or spd = "ノーマル") And n1 < h0 And _
            LCase(am2) = LCase(bfshn.Cells(n1 + 1, Abs(b)).Value) Then    '←86_016q(uni対策)　StrConv(am2→LCase(am2)
        tskup = 1
        If b < 0 Then '純高速時
            MsgBox "こちらは低速専用になりました。高速でこちら通るのはおかしい。"
        Else 'ノーマル
            n2 = n1 + 1
        End If
    Else
        'p=0
    End If
    If tskup = 0 Then 'まだ決まらず(p=0)→マッチング実施
        If c < 0 And (spd = "近似高速") Then  'spd = "旧近似高速" Or spd = "旧近似ノーマル" Or　は除外 85_006
            MsgBox "こちらは低速専用になりました。近似高速でこちら通るのはおかしい。"
        Else 'ノーマルor純　 am2→strcnv化　85_008
            If IsError(Application.Match(LCase(am2), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 0)) Then  '新規2
                If c >= 0 Then '-1-2はやらないに（p=0のまま終了へ　'←86_016q(uni対策)　StrConv(am2→LCase(am2)
                    tskup = 2
                    n2 = h + 1
                End If
            Else
                tskup = 1
                m = Application.WorksheetFunction.Match(LCase(am2), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 0)  '完全一致
                If spd = "純高速" Then '高速時
                    MsgBox "こちらは低速専用になりました。純高速でこちら通るのはおかしい。"
                Else 'not高速時
                    n2 = k0 + m - 1
                End If
            End If
        End If
    End If
End Function
