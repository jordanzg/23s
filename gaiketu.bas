Attribute VB_Name = "gaiketu"
Public flag As Boolean 'UserForm1�A�g�̂��ߗvpublic
Public dw As String, fmt As String 'UserForm2�A�g�̂��ߗvpublic
Dim kurai As Currency, cnt As Long, k As Long, xFlag As Boolean, shog As String
Dim bfn As String, shn As String, bfshn As Worksheet, twbsh As Worksheet, twt As Worksheet
Dim dd1 As Long, dd2 As Long, gg2 As Long, gg1 As Long, mghz As Long, mg2 As Long
Dim twn As String, sr(8) As Long  'xWsheet As Worksheet,
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function kgcnt(cef As String, kgr As String) As Long
    '������Ɋ܂܂���؂蕶���̐���Ԃ�
    'cef:������Akgr:��؂蕶��
    Dim cunt As Long
    cunt = 0
    If kgr <> "" Then
        Do Until InStr(bb + 1, cef, kgr) = 0
            cc = InStr(bb + 1, cef, kgr)
            cunt = cunt + 1
            bb = cc
        Loop
    End If
    kgcnt = cunt
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function ctreg(rtyu As String, tyui As String) As Long
    '�ŏI�s��Ԃ�(ctrl+end���s��hidden�̏ꍇ�̑Ώ���)
    ctreg = Workbooks(rtyu).Worksheets(tyui).Range("A1").SpecialCells(xlLastCell).Row()
    ctreg = ctreg + 1
    Do Until Workbooks(rtyu).Sheets(tyui).Cells(ctreg, 1).EntireRow.Hidden = False
        ctreg = ctreg + 1
    Loop  'ctrl+end�̎��s��hidden��������Ahidden���ꂽ�ŏI�s��Ԃ��B
    ctreg = ctreg - 1
End Function
'�[���[���[�ȏ�A��pubikoued
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function syutoku() As String 'publicfunction��function��
    syutoku = Environ("Username")
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub kyosydou()
    '���ʁi�O���A����A���ʁj���ʂ̏������܂Ƃ߂�B

    Dim ii As Long, nm As Variant
    Dim xsheet As Worksheet 'xWsheet��xsheet

    Application.CutCopyMode = False

    DoEvents
    
    gg1 = Selection.Row    '�I���J�n�s
    gg2 = Selection.Rows(Selection.Rows.Count).Row   '�I���I���s
    dd1 = Selection.Column    '�I���J�n��
    dd2 = Selection.Columns(Selection.Columns.Count).Column    '�I���I����
    
    If dd1 = 0 Then MsgBox "dd1���[���ł���"
    
    bfn = ActiveWorkbook.Name 'bfn,shn�̓p�u���b�N
    shn = ActiveSheet.Name
    Set bfshn = Workbooks(bfn).Worksheets(shn) '30��63���B���V�[�g�E���t�@�C��
    
    twn = ThisWorkbook.Name '�}�N���t�@�C�������̂���(�`.xlsm)
    Set twbsh = Workbooks(twn).Worksheets("���W�v_���`") '30��74���
    
    nm = Array("", "�Ώۼ�Ė�", "���F�ˍ���", "�΁F�ˍ���", "�΁F���1��", "�΁F���ė�", "�΁F���Z�񥑼", "���F�]�ڗ�", "�΁F�]�ڗ�", "�΁F�������Z��", "���F�����]�ڗ�") '30s63,array��
    If IsError(Application.Match(nm(1), Range(bfshn.Cells(1, 2), bfshn.Cells(200, 2)), 0)) Then Call oshimai("", bfn, shn, 4, 2, "�i�������~�j�u" & nm(1) & "�v��������܂���")
    For ii = 1 To 8
        sr(ii) = WorksheetFunction.Match(nm(ii), Range(bfshn.Cells(1, 2), bfshn.Cells(200, 2)), 0)
        If sr(0) < sr(ii) Then sr(0) = sr(ii)
    Next
    sr(0) = sr(8) + 1  '��������ڐ�
    
    If IsError(Application.Match("�B", Range(bfshn.Cells(1, 1), bfshn.Cells(1, 5000)), 0)) Then
        Call oshimai("", bfn, shn, 1, 0, "�u" & shn & "�v�V�[�g�E��Ɂu�B�v������܂���B����ĉ������B")
    Else
        mghz = Application.Match("�B", Range(bfshn.Cells(1, 1), bfshn.Cells(1, 5000)), 0)
    End If

    shog = "log_" & syutoku() & "_" & Format(Date, "yyyymm")
    '���O�V�[�g�L��chk�A30s82�A����̂݁������Ɉڐ�
    For Each xsheet In ThisWorkbook.Sheets
        If xsheet.Name = shog Then xFlag = True 'boolean�^�̏����l��false
    Next xsheet

    If xFlag = True Then ' �Y���̃V�[�g������ꍇ�̏���
        '�i�������Ȃ��j
    Else ' �Y���̃V�[�g���Ȃ��ꍇ�̏��� '
        Workbooks(twn).Activate '30s83
        Worksheets.Add
        ActiveSheet.Name = shog
        nm = Array("", "���ږ�", "����", "log", "date", "timestamp", "����", "to", "�ŉE��", "from9") '30s83,array��
        For ii = 1 To 9
            Workbooks(twn).Sheets(shog).Cells(1, ii).Value = nm(ii)
        Next
    
        Workbooks(bfn).Activate          '������ցi�������H�j
        bfshn.Select
    End If
    
    xFlag = False
    For Each xsheet In ThisWorkbook.Sheets     '�]�L�L��chk�A30s82e
        If xsheet.Name = "�����V�[�g_" & syutoku() Then xFlag = True
    Next xsheet

    If xFlag = True Then ' �Y���̃V�[�g������ꍇ�̏���
        Set twt = Workbooks(twn).Worksheets("�����V�[�g_" & syutoku()) '30s82f
        twt.Cells.Clear
    
    '�G���[�Ή����؁@86_017f
    DoEvents
    
    ThisWorkbook.Activate
    Sheets("�����V�[�g_" & syutoku()).Select
    Rows("6:6").Select
    Selection.Delete Shift:=xlUp
    DoEvents
    
    Range(Cells(2, 3), Cells(2, 3)).Select
    Workbooks(bfn).Activate       '86_017h
    bfshn.Select
    DoEvents
    
    Else ' �Y���̃V�[�g���Ȃ��ꍇ�̏��� '
        Worksheets.Add
        ActiveSheet.Name = "�����V�[�g_" & syutoku()
        Set twt = Workbooks(twn).Worksheets("�����V�[�g_" & syutoku()) '30s82f
        Workbooks(bfn).Activate
        bfshn.Select
    End If
    bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1 '30s83 =sum(a:a)+1���炱�����
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub �O������()   '��ԏな�Ƀp�u���b�N�ϐ�����B����������
    Dim f As String, xsheet As Worksheet, xbook As Workbook, wsfag As Boolean
    Dim bun As Long, bni As Long
    Dim mr() As String, er() As Currency
    Dim nkg As Long, kahi As Long, cted(1) As Long, rrr As Long 'ppp��rrr 86_014r
    Dim pap2 As Long, pap3 As Long, pap5 As Long, pap7 As Long, pap8 As Long, pap9 As Long
    Dim er2() As Currency, er3() As Currency, er5() As Long, er78() As Currency, er9() As Currency, er34 As String
    Dim mr2() As String, mr3() As String, mr5() As String, mr8() As String, mr9() As String
    Dim a As Long, pkt As Long, n As Long, n1 As Long, qq As Long, pqp As Long
    Dim am1 As String, am2 As String, h As Long, m As Long, k0 As Long, h0 As Long, n2 As Long, kg1 As String
    Dim qap As Long, ii As Long, jj As Long, trt As Long, tst As Long, dif As Long, axa As Long
    Dim saemp ', sou '�������^���ݒ肳��Ă��Ȃ�
                       '���^�ݒ��
    Dim hirt As Variant, hiru As Variant, tameshi As Range
    Dim kasan As Variant, c5 As Long, c99 As String, c98 As String, ct3 As String, hk1 As String, rog As String
    Dim zzz As String, zyyz() As String, xxxx() As String, zxxz() As String, nifuku As Long
  
    '�Čv�Z����U������
    Application.Calculation = xlCalculationAutomatic
    '���܂��Ȃ�(�u�R�[�h�̎��s�����f����܂����v�Ώ�)
    Application.EnableCancelKey = xlDisabled
    UserForm1.StartUpPosition = 1 '1�@�G�N�Z���̒����A�@2�@��ʂ̒����A�@3�@��ʂ̍���
    UserForm1.Show vbModeless
    UserForm1.Repaint

    kyosydou  '���ʂ̏���

    If dd1 = 0 Then Call oshimai("", bfn, shn, 1, 0, "dd1���[���ł�")

    '���̓��̏���`�F�b�N
    ii = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(ii, 1).Value = ""
        ii = ii + 1
        If ii = 50000 Then
            MsgBox "�󔒍s��������Ȃ��悤�ł�"
            Exit Sub
        End If
    Loop
    
    If Workbooks(twn).Sheets(shog).Cells(ii - 1, 4).Value <> Val(Format(Now(), "yyyymmdd")) Then
        Call oshimai("", bfn, shn, 1, 0, "���̓��̏���́A�ŏ���[FIRST]�{�^���������ĉ������B")
    End If
    
    If bfshn.Cells(sr(0) - 1, 5) = "" Then
        Call oshimai("", bfn, shn, sr(0) - 1, 5, "���W�v������͂��ĉ������B")
    End If

    Call iechc(hk1) '��igchc(hk1)
    
    hk1 = ""
    flag = False

    ii = 1
    '���V�[�g���̂P�ł͂Ȃ�all1�T��
    Do Until bfshn.Cells(ii, 1).Value = "all1"
        If IsNumeric(bfshn.Cells(ii, 1).Value) And bfshn.Cells(ii, 1).Value <> "" Then '17s
            Call oshimai("", bfn, shn, ii, 1, "���ڂ͐��l�����Ȃ��ŉ�����")
        End If
        ii = ii + 1
        If ii = 100 Then
            Call oshimai("", bfn, shn, 1, 0, "���V�[�g�uall1�v��������Ȃ��悤�ł�")
        End If
    Loop

    k = ii + 1     'k�m��'k�̓f�[�^�J�n�s(�T���v���s�ł͂Ȃ��Ȃ����B)
    bfshn.Cells(1, 2).Value = k '�f�[�^�J�n�s

    '���̎��_�ł�ii�́A���V�[�g�́uall1�v�L�ڍs
    Do Until bfshn.Cells(ii, 1).Value = ""
        ii = ii + 1
    Loop
    '�����ł�ii�͓��V�[�gall1��̋󔒂ɂȂ����s�A�f�[�^�����̏ꍇ�̓f�[�^�J�n�s
    
    '�I�[�g�t�B���^���ݒ肳��Ă�΁A����
    If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    n = 0 '�Z���󔒃`�F�b�N�t���O
    
    If dd1 <= 5 Then Call oshimai("", bfn, shn, 1, 0, "�^�e���R�I���A6��ڈȍ~���Ώۂł�")    '��
    'q = 0   (�����p����Ă��Ȃ����ۂ��̂Ŕp�~��)

For a = dd1 To dd2 '�I��͈͗񕪂̌J��Ԃ��@��
    kg1 = "" '���Z�b�g�@kg2��mr(2,3,bni)��redim�Ń��Z�b�g�����B
    bni = 1  '���Z�b�g
    bun = 1  '���Z�b�g
    nkg = 0 '���Z�b�g
 nifuku = 0
    
    '�����ߗp��؂蕶���̊m��@fnywti��rvsrz3�𗬗p
    kg1 = Mid(rvsrz3(bfshn.Cells(sr(7), a).Value, 3, "�", 0), 1, 1)  'kg2��`�͐� 30s73:nkg2��0

    If bfshn.Cells(sr(6), a).Value <= -90 Then
        bun = 1 '-90��͋���1
    ElseIf kg1 = "" Then
        nkg = 1 '���؂薳�w��(kg1="")�͕��߃[���i��؂肵�Ȃ��j
    Else '�����ߐ�(bun)�m��̂��߂̃��[�`��(kg1<>"")
        For ii = 1 To 8
            Do Until rvsrz3(bfshn.Cells(sr(ii), a).Value, bni, kg1, 0) = ""
                bni = bni + 1
                If bni > 120 Then Call oshimai("", bfn, shn, 4, 2, "bni120����")
            Loop
            bni = bni - 1
            If ii = 2 And bni > 1 Then
                nifuku = 1
                MsgBox "2�s�ڂł̕����߂���B���ӂ��B"  '���ւց@86_016v
            '    Call oshimai("", bfn, shn, 1, 0, "2���ߖڂł̕����ߎw��͋��e����Ă��܂���")
            End If
            If bun < bni Then bun = bni '�i���̎��_��bun�m��j
            bni = 1
        Next
    End If

'���`���P����
    bni = 1 '���Z�b�g
    ReDim er(11, bun) 'As Currency 10��11 30s83
    ReDim mr(4, 11, bun) 'As String 30s81��R���q�� 10��11 30s83 ,mr(3��mr(4�F�6����F� �p
    
    '�Čv�Z���蓮��
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = True
    
    For ii = 1 To 7 'mr(0,�L�q�@�@ii=2����ł����v�Ǝv����B
        mr(0, ii, 1) = rvsrz3(bfshn.Cells(sr(ii), a).Value, 1, kg1, nkg)  'n�s�ڂ̑S���q�@bni:1
    Next
    
    '1���ߖړ���
    mr(1, 1, 1) = rvsrz3(mr(0, 1, 1), 2, "�", 2)  '�V�[�g��
    mr(2, 1, 1) = rvsrz3(mr(0, 1, 1), 3, "�", 0)  '�t�@�C����
    
    mr(2, 7, 1) = kg1 '��@�V�s�ڑ����q
    If StrConv(Left(mr(1, 1, 1), 1), 8) = "*" Then '��@30s57���ꕶ�����������̎��i���`�̎��́`��mr(2,0,1)�ɓ���j
        bfshn.Cells(sr(0), a).Value = bfshn.Cells(sr(0) + 3, a).Value
        bfshn.Cells(sr(0) + 1, a).Value = bfshn.Cells(sr(0) + 4, a).Value
        bfshn.Cells(sr(0) + 2, a).Value = bfshn.Cells(sr(0) + 5, a).Value
    Else  '��(*�����łȂ��ʏ펞�A�O���Ō㕔�܂ő���)
    
    For ii = 1 To 7  '�󔒊m�F�ibni=1�j
        If bfshn.Cells(sr(ii), a).Value = "" Then n = sr(ii)
    Next
    
    If bfshn.Cells(sr(6), a).Value > -90 Then     '�����ł�ii��8 -99��8�s�ڊm�F���Ȃ�
        If bfshn.Cells(sr(8), a).Value = "" Then n = sr(8) 'ii��8�@�i���l�j
    End If
    If n > 0 Then Call oshimai("", bfn, shn, n, a, "�O���ݒ��񂪋󗓂̏�������܂�")
'���`���P���߂����܂Ł�

'���a��������for�i�F�����ҁj�ӂ����@���ߖ���for����������n�܂�
    For bni = 1 To bun
        n = 0 '�Z���󔒃`�F�b�N�t���O���H
        mr(2, 0, bni) = bfn '30s82
        mr(1, 0, bni) = shn  '30s82
        For ii = 1 To 7
            mr(0, ii, bni) = rvsrz3(bfshn.Cells(sr(ii), a).Value, bni, kg1, nkg)  'n�s�ڂ̑S���q
            If mr(0, ii, bni) = "" Then mr(0, ii, bni) = mr(0, ii, bni - 1)  '2���ߖڈȍ~�A�󗓂Ȃ�O�߃R�s�y
        Next
        
        If bfshn.Cells(sr(6), a).Value <= -90 Then
            mr(1, 1, bni) = shn  '�V�[�g���@'-90��̓Z�������Abfshn�����@�Z���͓��t�����Ȃǎ��R�ɏ�������B
            mr(2, 1, bni) = bfn  '�t�@�C����
        Else
            mr(1, 1, bni) = rvsrz3(mr(0, 1, bni), 2, "�", 2)  '�V�[�g�� ""�ɂ͂Ȃ�Ȃ�
            mr(2, 1, bni) = rvsrz3(mr(0, 1, bni), 3, "�", 0)  '�t�@�C����30s73:nkg2��0
        End If
        If Left(StrConv(mr(1, 1, bni), 8), 1) = "\" Then mr(1, 1, bni) = shn '�����V�[�g���ϊ�
        If mr(1, 1, bni) = "" Then Call oshimai("", bfn, shn, sr(1), a, "�ΏۃV�[�g��(" & bni & "���ߖ�)���󗓂ł�")
        If mr(2, 1, bni) = "" Then mr(2, 1, bni) = bfn
        mr(4, 1, bni) = mr(1, 1, bni)    '86_014s ����A�����Ă��B
        wsfag = False   '�t�@�C���E�V�[�g�L��chk
        Do Until wsfag = True
            For Each xbook In Workbooks
                If xbook.Name = mr(2, 1, bni) Then wsfag = True
            Next xbook
            If wsfag = False Then
'                If MsgBox("�t�@�C�����m�F�ł��܂��񂪁A" & vbCrLf & "���s���܂����H", 289, "�t�@�C���s��") = vbCancel Then '�L�����Z����
                    '�����̎�@�iYes,No�I������j�͂����ł͈Ӗ��Ȃ������̂ŁA���̂����͒��~�B
                    Call oshimai("", bfn, shn, sr(1), a, "���{���~���܂����B" & vbCrLf & "�t�@�C�����m�F�ł��܂���")
'                end if
            End If
        Loop
    
        wsfag = False
        For Each xsheet In Workbooks(mr(2, 1, bni)).Sheets
            If xsheet.Name = mr(1, 1, bni) Then wsfag = True
        Next xsheet
        If wsfag = False Then Call oshimai("", bfn, shn, sr(1), a, mr(1, 1, bni) & " �̃V�[�g���s���ł�(" & bni & "���ߖ�)")
    
        n = 0 '��U���Z�b�g
   
        If kurai = 2.1 Then  '�G���g���[�d�l er����
            For ii = 2 To 7
                If IsNumeric(bfshn.Cells(sr(ii), a).Value) Then
                    er(ii, bni) = bfshn.Cells(sr(ii), a).Value
                    mr(1, ii, bni) = er(ii, bni) '30s84_5�ǉ�
                Else
                    n = sr(ii)
                End If
            Next
            mr(2, 11, bni) = "����"  'koumdicd�̓G���g���[�ł͎��{���Ȃ��ɁB30s84_4
            er(11, bni) = 0
        Else  '�G���g���[�ȊO
            For ii = 2 To 6  '(��)�͂����Œ�܂�B
                mr(2, ii, bni) = rvsrz3(mr(0, ii, bni), 3, "�", 0)
            Next
            mr(3, 4, bni) = rvsrz3(mr(0, 4, bni), 4, "�", 0) '��3���q�i�H�����j
            
            If bni >= 2 Then mr(2, 7, bni) = mr(2, 7, bni - 1) '�߂���(��)�����O�߃R�s�[

            'mr(2, 11, bni)�E�E�E�����Ƃ�������B��er(11,bni)�̓��t�@���ϐ�
            mr(2, 11, bni) = koudicd(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), mr(0, 4, bni))
            
            mr(1, 2, bni) = yhwat1(bfn, shn, 2, sr(2), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 2, bni))
            mr(4, 2, bni) = zhwat1(bfn, shn, 2, sr(2), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 2, bni))

            For ii = 3 To 6
                mr(1, ii, bni) = yhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(ii), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, ii, bni))
                mr(4, ii, bni) = zhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(ii), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, ii, bni))
            Next
            
            mr(1, 7, bni) = yhwat1(bfn, shn, 2, sr(7), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 7, bni))
            mr(4, 7, bni) = zhwat1(bfn, shn, 2, sr(7), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 7, bni))
       
            For ii = 2 To 7
                er(ii, bni) = Val(mr(1, ii, bni)) '���Ή����Fval�֐��͐����ƔF���o���鏊�܂ł𐔒l�ϊ�����B�v�͑���ڂ��B
            Next
        End If
    
        If n > 0 Then Call oshimai("", bfn, shn, n, a, "�i�������~�j" & vbCrLf & "���l�ȊO�̏�񂪂���܂��B")
        If er(5, bni) >= 0 And er(7, bni) < 0 And er(6, bni) > -90 Then Call oshimai("", bfn, shn, 1, 0, "�i�������~�F���@�G���[�j" & vbCrLf & "er(5,0)>=0�@�Ł@er(7,0)<0�@�ł��B")
        If er(5, bni) < 0 And Round(kurai) = 2 Then Call oshimai("", bfn, shn, 1, 0, "�i�������~�F�x�[�V�b�N�j" & vbCrLf & "�J�E���g��񂪃}�C�i�X�ł��B")
    
        If er(6, bni) > -90 Then '8�Ԗځi�]�ڗ�j�̏���
            c98 = bfshn.Cells(sr(8), a).Value  '26��
            mr(0, 8, bni) = rvsrz3(bfshn.Cells(sr(8), a).Value, bni, kg1, nkg)
            mr(2, 8, bni) = rvsrz3(mr(0, 8, bni), 3, "�", 0) ' '30s73:nkg2��0
            mr(1, 8, bni) = yhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2�L_30s56
            mr(4, 8, bni) = zhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2�L_30s56
        
            If bni >= 2 Then '2���ߖڈȍ~����Ƃ����̂��~�\�A1���ߖڂ͉��L
                If mr(0, 8, bni) = "" Then mr(0, 8, bni) = mr(0, 8, bni - 1)
                If mr(2, 8, bni) = "" Then mr(2, 8, bni) = mr(2, 8, bni - 1) 'mr(2,8,0)�듮��j�~ 2���ߖڈȍ~�A�󗓂Ȃ�O�߃R�s�y
                mr(1, 8, bni) = yhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2�L_30s56
                mr(4, 8, bni) = zhwat1(mr(2, 1, bni), mr(1, 1, bni), er(11, bni), sr(8), a, bni, kg1, nkg, mr(2, 4, bni), mr(0, 8, bni)) 'kg2�L_30s56
            End If
            er(8, bni) = Val(mr(1, 8, bni))
            pap8 = kgcnt(mr(1, 8, bni), mr(2, 4, bni)) '30s75 pap8������ɂ�
            
            If Len(mr(2, 8, bni)) > 1 And (er(8, bni) < 0 Or Round(er(6, bni), 0) = -15) Then  '-15�K�p�����ǉ�
                'MsgBox "8�s�����q�F(&H)" & mr(2, 8, bni) & "��" & Chr(Val("&H" & mr(2, 8, bni))) & "��"
                'Call oshimai("", bfn, shn, 1, 0, "�����͋x�~���Ă݂܂�")
                mr(2, 8, bni) = Chr(Val("&H" & mr(2, 8, bni)))
            End If
            
            If Round(er(6, bni)) <> -2 And er(8, bni) < 0 Then   '���������86_013f
                If mr(2, 8, bni) = "" Then mr(2, 8, bni) = "�A" '��؂蕶���f�t�H�́u�A�v�i-15�ł��K�p�j
            End If
        
            If er(6, bni) <= -3 And er(8, bni) < 0 Then Call oshimai("", bfn, shn, 1, 0, "�uc=-3�ȉ��͏d���A�Ȃ�^(e<0)�͎��s�ł��Ȃ��ł��B")
            
            '30s86_012s�ǉ���
            If er(5, bni) < 0 And er(7, bni) = 0 Then Call oshimai("", bfn, shn, sr(7), a, "������7�s0�͎��{����Ȃ��Ȃ�܂����B")
            
            '30s82d�ǉ���                      er(5,1)��er(5,bni) 86_010
            If (er(6, 1) = -1 Or er(6, 1) = -2) And er(5, bni) >= 0 And kg1 <> "" And rvsrz3(bfshn.Cells(sr(7), a).Value, 2, kg1, 0) <> "" Then Call oshimai("", bfn, shn, n, a, "6�s-1-2�̎���not������(�ʏ펞)��7�s�����ߕs�ł��B")
 
            '�����ύX�@30s86_012w
            If Round(er(6, bni)) = -2 Then
                mr(0, 9, bni) = mr(0, 8, bni)
                mr(2, 9, bni) = mr(2, 8, bni)
                mr(1, 9, bni) = mr(1, 8, bni)
                er(9, bni) = Val(mr(1, 9, bni))
            ElseIf er(6, bni) > 0 Then
                mr(0, 9, bni) = mr(0, 6, bni)
                mr(2, 9, bni) = mr(2, 6, bni)
                mr(1, 9, bni) = mr(1, 6, bni)
                er(9, bni) = Val(mr(1, 9, bni))
            End If

            If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0) Then
                'MsgBox "�А����o�Y��" '�����������ǉ�30s86_012w
                bfshn.Cells(sr(6), 5).Value = "����" '��20191118�@������ց@86_015c
            Else
                bfshn.Cells(sr(6), 5).Value = ""
            End If
          
            If Not (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) > 0) Then '���������ύX30s86_012w�A�ǉ�30s86_012s
            '   6�s�[�Q�łȂ��A�@�@�@�@�@�@���@�W�s�O�łȂ��A�@�@�@���@�i�T�s�{ ���邢�́@�U�s�O�ȉ��j�@�Ȃ��
                If (Not Round(er(6, bni)) = -2) And (Not er(8, bni) = 0) And (er(5, bni) >= 0 Or er(6, bni) <= 0) Then '�V�d�l(30s83)
                '10�s�ڏ���(er(10,0)�͎����̓��V�[�g�]�ڗ�) (���������̘A�ڌ^�͎��{����30s86_012w)
                    mr(0, 10, bni) = mr(0, 8, bni)
                    mr(2, 10, bni) = mr(2, 8, bni)
                    mr(1, 10, bni) = mr(1, 8, bni)
                    er(10, bni) = Val(mr(1, 10, bni))
                    If mr(2, 8, bni) <> "" And er(10, bni) = 0.1 Then er(10, bni) = -0.1
                End If
            End If '���������ǉ�30s86_012s
        End If

        n = 0 '��U���Z�b�g

        'log������
        ii = 1
        If StrConv(Left(mr(0, 1, bni), 1), 8) <> "*" And bfn <> twn Then '*�ړ����̓��O�����ΏۊO
            Do Until ii = 0
                If rvsrz3(rog, ii, "��", 0) = "" Then
                    rog = mr(2, 1, bni) & "\" & mr(1, 1, bni) & "��" & rog  'rog�͂��̊O�����{�͈͂ł̑ΏۃV�[�g�t�@�C���̏W����(���ŘA��)
                    ii = 0
                ElseIf rvsrz3(rog, ii, "��", 0) = mr(2, 1, bni) & "\" & mr(1, 1, bni) Then
                    ii = 0
                Else
                    ii = ii + 1
                End If
                If ii = 200 Then Call oshimai("", bfn, shn, k, a, "���܂������ĂȂ��B")
            Loop
        End If

        If er(2, bni) > 0 And er(9, bni) > 0 And er(10, bni) <> 0 And er(7, bni) = 0 And er(5, bni) >= 0 Then '�����������l��
            Call oshimai("", bfn, shn, 1, 0, "�������~�Bc���Z����œ���]�ڂ��悤�Ƃ��Ă��܂��B�m�F���B")
        End If
    Next
'���a�������߂ӂ��i�����ҁj�����܂Ł�

'���b���P���߁i�{�ԃv���j�������火
    bni = 1 '1���ߖڂŔ��f�A���{����'����

    '���������m�F�@���߂P�Ŕ��f�ibni=1�j
    If er(2, bni) < 0 Then  'And q = 0 Then
        If Round(kurai) = 2 Then MsgBox ("�x�[�V�b�N�F�ᑬ�ƂȂ�܂��B")
        'q = 1 �iq�͎g���Ă��Ȃ����ۂ��̂Łj
    End If

    '�x�[�V�b�N���̍������ᑬ�]��
    If Round(kurai) = 2 Then er(2, bni) = Abs(er(2, bni))

    '�O��f�[�^�R�s�[
    If StrConv(Left(bfshn.Cells(sr(1), a).Value, 2), 8) <> "**" Then '**�͂��Ȃ���024_����2
        bfshn.Cells(sr(0), a).Value = bfshn.Cells(sr(0) + 3, a).Value
        bfshn.Cells(sr(0) + 1, a).Value = bfshn.Cells(sr(0) + 4, a).Value
        bfshn.Cells(sr(0) + 2, a).Value = bfshn.Cells(sr(0) + 5, a).Value
    End If
    h = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1 + k - 2 '����̍ŉ��s(�ȉ��L�тĂ���)���f�[�^�����̎��͍��ڍs�ƂȂ�Ah=k-1�ƂȂ�̂Œ���
    
'�����ɓ���B30s86_006
'    If h < ctdg(bfn, shn, 1, a) Then
'        'MsgBox ctdg(bfn, shn, 1, a)
''������ii�g�pok�m�F�ς�
'        ii = ctdg(bfn, shn, 1, a)
'        Do Until ii = h  '�P�c���������փT�[�`
'            If Cells(ii, Abs(er(2, 1))) <> "" Then Exit Do
'            ii = ii - 1
'        Loop
'        If ii > h Then
'            MsgBox "���₳�Ȃ���"
'            h = ii
'        End If
'    End If
'�����܂łɓ����

    '�̍�(teisai)��������
    trt = 0 '���Z�b�g trt�E�E�E�^�C�v�̒�`�@trt:�Z���̃^�C�v
    'trt�F0����ؾ��,-9�E-99�^,-2������Z�^,-1����]�ڌ^,1����������^
    'trt�F-9�E-99�^ ,-1����]�ڌ^�A-2������Z�^(�ܓ������Z�^)�A1����������^(�ܓ����A�ڌ^)
    'trt��`��
    If er(6, bni) <= -90 Then '-99�͂�����
        trt = -9
        '�����������i���Zor�A�ڌ^�j
    ElseIf er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0 Then
        If er(8, bni) > 0 Then
            trt = -2 '�������Z�^
        Else
            trt = 1 '�����A�ڌ^
        End If
    ElseIf er(6, bni) > 0 Or Round(er(6, bni)) = -2 Then
        trt = -2  '������Z�^ ��������
    ElseIf er(5, bni) >= 0 Then
        If er(10, bni) < 0 Then
            trt = 1   '����������^
            'MsgBox "����������ł��B"
        Else
            trt = -1    '����]�ڌ^
            'MsgBox "�A�ڂ���Ȃ��ł��B"
        End If
    ElseIf er(5, bni) < 0 Then  '����������r��
        If mr(2, 6, bni) = "1" Then
            Call oshimai("", bfn, shn, sr(6), a, "��������6�sop1�̉^�p�͏I����Ă���")
        ElseIf mr(2, 6, bni) = "-1" Then
            trt = -1 '�����𕶎�(��ح�)�ŕ\�� ���]�ڌ^
        Else
            trt = -1 '������(0,1)�ŕ\���A-2��-1��(�ŏ������A���ƂŐ��l�̂���)
        End If
    End If
    If trt = 0 Then Call oshimai("", bfn, shn, 1, 0, " trt = 0 �H")  'trt�[����ԂŖ������Ƃ̊m�F�O�̂���
        
    'tst��`����������
    tst = -1 '���Z�b�g�@tst:�Z���̌^
    'tst : -1����ؾ�� , �2�ʉ�,   0�W��,   1������,�@ 7�ꊇ���P�^ , 8���܂���
        
    If trt = -9 Then
        If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then
            tst = 1     '������^
        ElseIf (er(4, bni) < 0 Or er(4, bni) = 0.1) Then
            tst = -2     '�ʉ݌^
        ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then '�ΏۃV�[�g����͕�@a11_1�񁨑ΏۃV�[�g�ɕύX
            tst = 7 '�ꊇ���P�^(�͕�^)
        Else
            tst = 0 '�W���^
        End If
    ElseIf trt = -2 Then
        If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  '�[�[
            tst = 0     '�W���^
        ElseIf (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  '�{�[
            tst = 7     '�ꊇ�^
        ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then  '�[�{
                        '�������Ȃ�(tst = -1)
                        '��������֓]�� 30s86_014f
            tst = 8     '�ꊇ�^
            MsgBox "���܂���(tst = 8)"
        Else                                              '�{�{
            tst = -2    '�ʉ݌^
        End If
    ElseIf trt = 1 Then '����������
        tst = 1 '������^
    ElseIf trt = -1 Then
        If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  '�[�[
            tst = 1     '������^
        ElseIf (er(4, bni) < 0 Or er(4, bni) = 0.1) Then  '�{�[
            tst = 7     '�ꊇ�^
        ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then  '�[�{
            tst = 8     '���܂��܃Z�����P�^
            MsgBox "���܂���(tst = 8)"
        Else                                              '�{�{
            tst = 0    '�W���^
        End If
    End If
    
    If tst = -1 Then Call oshimai("", bfn, shn, 1, 0, " tst = -1 �H")
    
    If h >= k And StrConv(Left(bfshn.Cells(sr(1), a).Value, 2), 8) <> "**" Then '�����s�������͂�������s���Ȃ���,�u�������v�����s���Ȃ�
        '�I���̊����f�[�^�N���A(���O�����s)
        With Range(bfshn.Cells(k, a), bfshn.Cells(h, a))
            .ClearContents
            .ClearComments '30s79�R�����g���N���A
            .Interior.Pattern = xlNone '30s79
            .NumberFormatLocal = "G/�W��"  '30s83�f�t�H���g�ł܂�
        End With

        '���񎖑O�̑̍وꗥ�� 86_013d trt,tst��
        If trt = -9 Then '-99�͂�����
            If tst = 1 Then '������]�ڕ����^
                Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "@"
            ElseIf tst = -2 Then  '�ʉ݌^
                Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[��]-#,##0"
            ElseIf tst = 7 Then '�ΏۃV�[�g����͕�@a11_1�񁨑ΏۃV�[�g�ɕύX
                Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = bfshn.Cells(sr(1), a).NumberFormatLocal 'sr(4)��sr(1)��30s68
            End If
        '�V�d�l-99�ȊO ������w��or�A�ځ@er(7, bni) > 0��-1�̕�����Ή�
        ElseIf tst = 1 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "@"
        ElseIf tst = -2 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[��]-#,##0"
        ElseIf tst = 0 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "G/�W��"
        End If
        '�̍�(teisai)�����܂� 86_012g
        
        If er(2, 1) < 0 Or kgcnt(mr(1, 2, 1), mr(2, 4, bni)) > 0 Then  '30s45�E�[�N���A��
            With Range(bfshn.Cells(k, mghz), bfshn.Cells(h, mghz + 2)) '3�ڗ�ڂ܂ō폜��30s80
                .ClearContents
                .NumberFormatLocal = "@"
            End With
            With Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(h, mghz + 1))  '2��ڂ͐��l�p
                .ClearContents
                .NumberFormatLocal = "G/�W��"
            End With
        End If
    End If
    
    '4��ڂɗv�f�]�L(-90������{�A�W�s�ڕʓr) ��������ϊ��ς�
    For ii = 1 To 7
        hk1 = "" '��3���q�Ή�
        If StrConv(Left(mr(0, ii, bni), 1), 8) = "*" Then  '30s75�i*�L�����ǉ��j
            bfshn.Cells(sr(ii), 4).Value = "*�" & mr(4, ii, bni) & "�" & mr(2, ii, bni) & hk1
        Else
            bfshn.Cells(sr(ii), 4).Value = "�" & mr(4, ii, bni) & "�" & mr(2, ii, bni) & hk1
        End If
    Next

    '8�s���e��4��ڂɗv�f�]�L�@��������ϊ��ς�
    If mr(2, 8, bni) = Chr(Val("&H" & "0A")) Then '�s�̖c��ݑj�~16s
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then
            bfshn.Cells(sr(8), 4).Value = "*�" & mr(4, 8, bni) & "�(LF)"
        Else
            bfshn.Cells(sr(8), 4).Value = "�" & mr(4, 8, bni) & "�(LF)"
        End If
    Else
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then
            bfshn.Cells(sr(8), 4).Value = "*�" & mr(4, 8, bni) & "�" & mr(2, 8, bni)
        Else
            bfshn.Cells(sr(8), 4).Value = "�" & mr(4, 8, bni) & "�" & mr(2, 8, bni)
        End If
    End If
    
    If er(6, bni) < -90 Then
        '85_024����6�@-90��łT�s�ڃ��Ή��p
        pap5 = kgcnt(mr(1, 5, bni), mr(2, 4, bni))  '5�s�ڃ��̐� 30s48
        ReDim er5(pap5)
        ReDim mr5(pap5)
        er5(0) = Val(mr(1, 5, bni)) '���łȂ��������@erx()�͒ʉ݌^�Ȃ̂�val��킹����𓾂Ȃ��B
        mr5(0) = mr(2, 5, bni)
        If pap5 > 0 Then 'mr(2, 4, bni) <> ""
            For ii = 0 To pap5
                er5(ii) = Val(rvsrz3(mr(1, 5, bni), ii + 1, mr(2, 4, bni), 0))
                mr5(ii) = rvsrz3(mr(2, 5, bni), ii + 1, mr(2, 4, bni), 0)
            Next
        End If
    End If
    
'�`�`�`�`�`
   
  bfshn.Cells(sr(0), 4).Value = bni & "/" & bun & "����"   '30s86_012�V��
  bfshn.Cells(sr(0) + 1, 4).Value = rvsrz3(bfshn.Cells(1, a).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0) & "��"
   
'���b���P���߁i�{�ԃv���j�����܂Ł��@�@������Ɉ����z��
 
 '��A �ȍ~-99or*,**�͂����ʉ�
  If er(6, bni) > -90 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then

'���c��������for�i�F�{�Ԏ��s�ҁj�ӂ����@���ߖ���for����������n�܂� ���[�v���@������Ɉ����z��
    
    For bni = 1 To bun    '��Ɉڍs�ց@86_016w
    
    '5��ڂ�荶�]�ږh�~ 86_012_m
    If er(7, bni) >= 1 And er(7, bni) <= 5 And er(5, bni) >= 0 Then
        Call oshimai("", bfn, shn, 1, 0, "5��ڂ�荶�ɓ]�ڂ��Ȃ��ŉ������B")
    End If
    
    '���[�v�O�̏����l
    
    n1 = k  'n1�F�O��ˍ������Ώۍs_���V�[�g�ł́Ak�̓f�[�^�J�n�s�i�Œ�jm2��n1
   
    pap2 = kgcnt(mr(1, 2, bni), mr(2, 4, bni)) '86_016w mr(1, 2, 1)��mr(1, 2, bni)
    ReDim er2(pap2)
    ReDim mr2(pap2)
    
    er2(0) = Val(mr(1, 2, bni)) 'mr2������(30s86_017a)�@mr(1, 2, 1)��mr(1, 2, bni)
    mr2(0) = mr(2, 2, bni)
    If pap2 > 0 Then
        For ii = 0 To pap2
            er2(ii) = Val(rvsrz3(mr(1, 2, bni), ii + 1, mr(2, 4, bni), 0))
            mr2(ii) = rvsrz3(mr(2, 2, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
        
    '�ߎ�������ʊm�F�@619 ���ʒP���߂�(��ߎ���er2(pap2)�̂��߁B)
    mr(1, 11, bni) = "" '��U���Z�b�g ������/�ߎ�����/�m�[�}���@������B

    If er(2, 1) < 0 Then '�@bni��1�ɏ�����(�����̏�������S�̂�)�@86_016w
        If er(6, 1) < 0 Then mr(1, 11, 1) = "�ߎ�����" Else mr(1, 11, 1) = "������"
    Else
        mr(1, 11, 1) = "�m�[�}��"
    End If 'bni:2�ȍ~��mr(1,11,bni)null�Ȃ̂Œ���
    
    bfshn.Cells(sr(2), 5).Value = mr(1, 11, 1)     '�m�[�}���E���E�ߎ��\�L�T��@bni��1
        
    If (er(2, 1) < 0 Or pap2 > 0) Then    '2�s�ڂɃ������鎞�����������{(1���ߖڂŔ��f)�A1���ߖڃ����鎞���S���߃��K�p�����B
        If k <= h Then '���V�[�g������񂠂�
            
            If er(6, 1) < 0 And pap2 = 1 And er2(pap2) = 0.1 Then  'A�ߎ��P���m��A�ׂ��������(���xup�ړI) And k <=h �O��(���)
                Call oshimai("", bfn, shn, sr(2), a, "A���ߎ��g�p�I��")
            ElseIf er(6, 1) < 0 And pap2 > 1 And er2(pap2) = 0.1 Then   'B:�ߎ������m��
                Call oshimai("", bfn, shn, sr(2), a, "B���ߎ��g�p�I��")
            Else  'C�ʏ�^(not�ߎ�)�E��2���񖄂ߍ��݁B�������ߎ�������������(������񂪂���Ƃ�)
                If StrConv(bfshn.Cells(sr(1), a), 8) = "\" And er(6, 1) >= 0 And mr(1, 2, bni) <> mr(1, 3, bni) And pap2 > 0 Then
                    'Call oshimai("", bfn, shn, sr(2), a, "\�̎��ŉ��z�L�[�g�p��6�s>0�̂Ƃ���2�s3�s��v���K�v�ł��B") '�������[�v�h�~s
                    MsgBox "2�s3�s�s��v(�������[�v�̉\���L)" '86_016r
                End If
                'mghz2��̏�񖄂ߍ��݁i�l�Ɛ��l�̏����œ\�t)
                If pap2 = 0 Then  '�ׂ��������
                    Application.Calculation = xlCalculationAutomatic    '�����v�Z���@�����Ɂ@'�V�`���@85_007
                    bfshn.Cells(sr(8), mghz).Value = "��ASC(PHONETIC(" & bfshn.Cells(sr(8), Abs(er(2, bni))).Address(RowAbsolute:=False, ColumnAbsolute:=False) & "))"
                    bfshn.Cells(sr(8), mghz).Replace What:="��", Replacement:="=", LookAt:=xlPart, _
                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                        ReplaceFormat:=False 'strconv24�Ɠ���
                    
                    '�����ŃR�s�y �V�p�^�[��
                    Call cpp2(bfn, shn, sr(8), mghz, 0, 0, bfn, shn, k, mghz, h, mghz, -4123)  'xlPasteFormulas�x
                    
                    '�l�ɕϊ�
                    Call cpp2(bfn, shn, k, mghz, h, mghz, bfn, shn, k, mghz, 0, 0, -4163) 'xlPasteValues��
                    
                    '������
                    Range(Workbooks(bfn).Sheets(shn).Cells(k, mghz), Workbooks(bfn).Sheets(shn).Cells(h, mghz)).Replace What:="��", _
                        Replacement:="��", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                        ReplaceFormat:=False
                    
                    '�� �� ��
                    bfshn.Cells(sr(8), mghz).Replace What:="=", Replacement:="��", LookAt:=xlPart, _
                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                        ReplaceFormat:=False
                    DoEvents
                    Application.Calculation = xlCalculationManual  '�Čv�Z�Ăю蓮�Ɂi�d���Ȃ邽�߁j30s66

                    'mghz+1��(�s�ԍ�)�̏���(�t�B�����p)
                    bfshn.Cells(k, mghz + 1).Value = k

                    If h > k Then
                        bfshn.Cells(k + 1, mghz + 1).Value = k + 1
                        If h > k + 1 Then
                            Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(k + 1, mghz + 1)).AutoFill Destination:=Range(bfshn.Cells(k, mghz + 1), bfshn.Cells(h, mghz + 1))
                        End If
                    End If
                
                Else '���܂���(pap2��0)
                    bfshn.Cells(sr(8), mghz).Value = "(�s�g�p(���܂���pap2��0)"
                    For ii = k To h
                        'mghz��̃R�s�y�s�P�ʂ�
                        If er(2, 1) < 0 Then   '�V�݁@85_006, ������ strconv24�킹��p�^�[����
                            bfshn.Cells(ii, mghz).Value = StrConv(wetaiou(bfn, shn, ii, er2(), mr(2, 4, bni), mr(1, 11, 1), mr2(), 2), 24)
                        Else '���]���p�^�[��(strconv26�Ȃ�)�@�ᑬ���͏]���̂܂܂�
                            
                            bfshn.Cells(ii, mghz).Value = wetaiou(bfn, shn, ii, er2(), mr(2, 4, bni), mr(1, 11, 1), mr2(), 2) 'mghz
                        
                        End If
                        bfshn.Cells(ii, mghz + 1).Value = ii 'mghz+1����
                        Call hdrst(ii, a)         '�����X�e�[�^�X�\����
                    Next
                End If
                cnt = 0
            End If
        End If '���V�[�g������񂠂�
      
        If er(2, 1) > 0 Then er(2, bni) = mghz Else er(2, bni) = -mghz '��86_016w
    
    End If
    
    '(���ύX)��Ɨ��W�v�@�i�u�ǁv�Ή����ׂ��ׂɁj '���ǁv�ɂ��Ή�
    If Not (InStr(1, mr(0, 2, bni), "�") > 0 And InStr(1, rvsrz3(mr(0, 2, bni), 1, "�", 0), "��") > 0) Then
        ii = h
        Do Until ii = k - 1
            If bfshn.Cells(ii, Abs(er(2, 1))) <> "" Then Exit Do
            ii = ii - 1
        Loop
        h = ii
    Else
        MsgBox "�ǑΉ�"
    End If
    
    DoEvents
    
    ct3 = ""
    With Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)) '�t�B���^��������i�R�s�y�ł́A�͂�����h�~�j
        If .FilterMode Then .ShowAllData
    End With

    cted(0) = ctdg(mr(2, 1, bni), mr(1, 1, bni), er(4, bni), a)
    cted(1) = ctdr(mr(2, 1, bni), mr(1, 1, bni), er(4, bni), a)

    If h >= k And (mr(1, 11, 1) = "������" Or mr(1, 11, 1) = "�ߎ�����") Then   'mghz�\�[�g
        '�����L�[(mghz��)��������=xlAscending ���������͋ߎ�����(�ꕶ�߂̂�)�E������(�S����)���Ɏ��{��
        Call saato(bfn, shn, mghz, k, mghz, h, mghz + 1, 1)
    End If
    
    k0 = k  '�����߈�U���Z�b�g�d�l��
    h0 = h '�����߈�U���Z�b�g�d�l��
   
    If h >= k And (mr(1, 11, 1) = "�ߎ�����") Then
        am1 = "" '�ߎ����͏���l����null�ցB(k�s�̗p�ł��Ȃ�����)
        n1 = k
    ElseIf er(2, 1) < 0 Then  '�������͂�����
        am1 = bfshn.Cells(k, Abs(er(2, bni))).Value  '30s83������ŕ���
        n1 = bfshn.Cells(k, Abs(er(2, bni)) + 1).Value
    Else  '�ʏ�
        am1 = bfshn.Cells(k, Abs(er(2, bni))).Value  '30s83������ŕ���
        n1 = k
    End If
    '�����v
    pap3 = kgcnt(mr(1, 3, bni), mr(2, 4, bni))  '3�s�ڃ��̐��@�@�ȉ���
    pap5 = kgcnt(mr(1, 5, bni), mr(2, 4, bni))  '5�s�ڃ��̐� 30s48
    pap9 = kgcnt(mr(1, 9, bni), mr(2, 4, bni))  'pap6����ύX
    
    ReDim er3(pap3)
    ReDim mr3(pap3)
    ReDim er5(pap5)
    ReDim mr5(pap5)
    ReDim er9(pap9)
    ReDim mr9(pap9)
    
    er3(0) = Val(mr(1, 3, bni)) '���łȂ��������@erx()�͒ʉ݌^�Ȃ̂�val��킹����𓾂Ȃ��B
    mr3(0) = mr(2, 3, bni)
    If pap3 > 0 Then 'mr(2, 4, bni) <> "" 's55�����߂�
        For ii = 0 To pap3
            er3(ii) = Val(rvsrz3(mr(1, 3, bni), ii + 1, mr(2, 4, bni), 0))
            mr3(ii) = rvsrz3(mr(2, 3, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
    
    er9(0) = Val(mr(1, 9, bni))
    mr9(0) = mr(2, 9, bni)

    If pap9 > 0 Then 'mr(2, 4, bni) <> ""
        For ii = 0 To pap9
            er9(ii) = Val(rvsrz3(mr(1, 9, bni), ii + 1, mr(2, 4, bni), 0))
            mr9(ii) = rvsrz3(mr(2, 9, bni), ii + 1, mr(2, 4, bni), 0) '30s78�ǉ�
        Next
    End If
    
    er5(0) = Val(mr(1, 5, bni)) '���łȂ��������@erx()�͒ʉ݌^�Ȃ̂�val��킹����𓾂Ȃ��B
    mr5(0) = mr(2, 5, bni)
    If pap5 > 0 Then 'mr(2, 4, bni) <> ""
        For ii = 0 To pap5
            er5(ii) = Val(rvsrz3(mr(1, 5, bni), ii + 1, mr(2, 4, bni), 0))
            mr5(ii) = rvsrz3(mr(2, 5, bni), ii + 1, mr(2, 4, bni), 0)
        Next
    End If
    
    pap8 = kgcnt(mr(1, 8, bni), mr(2, 4, bni))  '8�s�ڃ��̐�  ���̎�������
   
    'pap8�Ē�`�i���̂Ƃ��j30��70
    If StrConv(Left(rvsrz3(mr(0, 8, bni), 1, "�", 0), 1), 8) = "*" Then '30s74����
        paq8 = pap8 / 2  'paq8�͔���(�F���̃O���[�s���O���@.5�����蓾��)
        qap = 0
        For ii = 0 To Int(paq8)
            If ii = Int(paq8) And paq8 - Int(paq8) = 0 Then '�ŏI������pap8����(�Ǘ�)
                qap = qap + 1
            Else
                fma = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 1, mr(2, 4, bni), 0)) 'from��
                If fma = 0.1 Then fma = cted(1)  '85_020
                tob = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 2, mr(2, 4, bni), 0)) 'to��
                If tob = 0.1 Then tob = cted(1)  '85_020
                qap = qap + Abs(fma - tob) + 1
            End If
        Next
        pap8 = qap - 1 'pap8�Ē�`����
    End If
            '��-1:�������r�Ŏg�p30s79�A0:7���A1�F8�����I��
    ReDim er78(-1 To 1, -1 To pap8) '��pap7<=pap8�@�Ƃ����O��,-1��c=-1-2�̎������g�p(tensai����)
    ReDim mr8(-2 To pap8) '30s75
    
    er78(1, 0) = Val(mr(1, 8, bni))
    
    If pap8 > 0 Then
        mr8(0) = rvsrz3(mr(2, 8, bni), 0 + 1, mr(2, 4, bni), 0)   '30s75
        If StrConv(Left(bfshn.Cells(sr(8), a).Value, 1), 8) = "*" Then
            qaap = 0
            For ii = 0 To Int(paq8) '��
                fma = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 1, mr(2, 4, bni), 0)) 'from��
                If fma = 0.1 Then fma = cted(1)  '85_020
                er78(1, qaap) = fma
                If ii < Int(paq8) Or paq8 - Int(paq8) = 0.5 Then   '�ŏI���łȂ��A���邢�͍ŏI����to����
                    tob = Val(rvsrz3(mr(1, 8, bni), ii * 2 + 2, mr(2, 4, bni), 0)) 'to��
                    If tob = 0.1 Then tob = cted(1)  '85_020
                    If fma > tob Then sa8 = -1 Else sa8 = 1
                    For qap = qaap + 1 To qaap + 1 + Abs(fma - tob) - 1 'from��to��for������A���{����Ȃ�
                        er78(1, qap) = er78(1, qaap) + (sa8) * (qap - qaap)
                    Next
                End If
                qaap = qap
            '���̎��́Amr8(qap)�͓��ʕs�g�p�Ƃ���@'30s75
            Next
        Else
            For ii = 1 To pap8 '����܂Œʂ�
                er78(1, ii) = Val(rvsrz3(mr(1, 8, bni), ii + 1, mr(2, 4, bni), 0))
                mr8(ii) = rvsrz3(mr(2, 8, bni), ii + 1, mr(2, 4, bni), 0) '30s75
            Next
        End If
    Else
        mr8(0) = mr(2, 8, bni) '30s75
    End If

    'pap8�����܂ŁB��������pap7
    'fma,tob��pap7�Ƃ��čă��Z�b�g����Ďg�p�����(pap8�̂����P����Ďg�p����Ȃ�)
    
    soroeru = 0
    er78(0, -1) = 0 '30s62 -1��c=-1-2�̎������g�p(tensai����)
    er78(0, 0) = Val(mr(1, 7, bni)) '(1,0)��(0,0)�C��30s61_4
    
    pap7 = kgcnt(mr(1, 7, bni), mr(2, 4, bni))  '7�s�ڃ��̐�  ���̎�������
    If pap8 > 0 Then 'er78���߂�B
        '�V�s��
        If StrConv(Left(bfshn.Cells(sr(7), a).Value, 1), 8) = "*" Then  '���̎��񃋁[�`��
            If Val(rvsrz3(mr(1, 7, bni), pap7 + 1, mr(2, 4, bni), 0)) = 0.1 Then '7�s�E���u�|�v
                soroeru = 1
                paq7 = (pap7 - 1) / 2
                '��ԉE�Ɂu�|�v�A����soroeru�r�b�g�𗧂ĂāA�[�͖����O��(�Fpap7-1)��paq7��ݒ�
            Else
                paq7 = pap7 / 2  'paq7�͔���(�F���̃O���[�s���O���@.5�����蓾��)
            End If
            
            qaap = 0
            For ii = 0 To Int(paq7) '���y�A�O���[�v���Ŏ���
                fma = Val(rvsrz3(mr(1, 7, bni), ii * 2 + 1, mr(2, 4, bni), 0)) 'from��
                er78(0, qaap) = fma  '�y�A�O���[�v���̈��
                qap = qaap  '�����ł�qap�͌�fma�̔z��ʒu(0,2,,,)
                pap7 = qaap
                If ii < Int(paq7) Or paq7 - Int(paq7) = 0.5 Then
                '(�E�́[�͖�������ł�)�ŏI���łȂ��A���邢�͍ŏI����to����
                '���E���[�̏����͉���soreoeru=1�@�̏��Ŏ��{�����B
                    tob = Val(rvsrz3(mr(1, 7, bni), ii * 2 + 2, mr(2, 4, bni), 0)) 'to��
                    
                    If qaap + 1 + Abs(fma - tob) - 1 > pap8 Then  '86_011����k
                        Call oshimai("", bfn, shn, 1, 0, "�]�ڗ񐔁F7�s��>8�s�ڂł��B�m�F���B")
                    End If
                    
                    If fma > tob Then sa7 = -1 Else sa7 = 1
                    For qap = qaap + 1 To qaap + 1 + Abs(fma - tob) - 1 '30��70fromto����
                        er78(0, qap) = er78(0, qaap) + (sa7) * (qap - qaap)
                    Next
                    pap7 = qap - 1
                    qaap = qap '�����ł�qap,qaap�́A���遖�O���[�v���Z���next��fma���ꍞ�ވʒu
                End If
            Next
            
            If pap7 > pap8 Then Call oshimai("", bfn, shn, 1, 0, "pap7>pap8�ł��B�m�F���B")
 
            If soroeru = 1 Then '7�s�E���u�|�v�E�Epap8�Ƒ�����
                For qaap = pap7 + 1 To pap8
                    er78(0, qaap) = er78(0, pap7) + (qaap - pap7)
                Next
                    If er78(0, qaap - 1) >= mghz Then
                        Call oshimai("", bfn, shn, 1, Int(er78(0, qaap - 1)), "mghz�͂ݏo�Ă܂��B" & vbCrLf _
                        & rvsrz3(bfshn.Cells(1, Int(er78(0, qaap - 1))).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0) & "��܂Ńf�[�^����")
                    End If
                pap7 = pap8
            End If
        Else '���ʂ̃��[�`��
            If pap7 > pap8 Then Call oshimai("", bfn, shn, 1, 0, "pap7>pap8�ł��B�m�F���B")
            
            For ii = 1 To pap8
                er78(0, ii) = Val(rvsrz3(mr(1, 7, bni), ii + 1, mr(2, 4, bni), 0))
            Next
        End If
    End If

    If (er(5, bni) < 0 And pap7 <> pap8) Then '30s79�ǉ�
        Call oshimai("", bfn, shn, sr(7), a, "�񐔂���v���܂���i�����̕������r�j")
    End If
    
    '4��ڂɗv�f�]�L(�W�s�ڕʓr) ���Q���߈ȍ~�����{,4��ڔ��f��
    For ii = 1 To 7
        hk1 = "" '��3���q�Ή�
        If mr(3, ii, bni) <> "" Then
        'MsgBox "��3���q:" & mr(3, ii, bni)
        hk1 = "�" & mr(3, ii, bni)
        End If
        If StrConv(Left(mr(0, ii, bni), 1), 8) = "*" Then  '30s75�i*�L�����ǉ��j�@'mr(1,��mr(4 ��
            bfshn.Cells(sr(ii), 4).Value = "*�" & mr(4, ii, bni) & "�" & mr(2, ii, bni) & hk1
        Else
            bfshn.Cells(sr(ii), 4).Value = "�" & mr(4, ii, bni) & "�" & mr(2, ii, bni) & hk1
        End If
    Next

    '8�s���e��4��ڂɗv�f�]�L�@��������ϊ��ς�
    If mr(2, 8, bni) = Chr(Val("&H" & "0A")) Then '�s�̖c��ݑj�~16s
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then  'mr(1,��mr(4 ��
            bfshn.Cells(sr(8), 4).Value = "*�" & mr(4, 8, bni) & "�(LF)"
        Else
            bfshn.Cells(sr(8), 4).Value = "�" & mr(4, 8, bni) & "�(LF)"
        End If
    Else
        If StrConv(Left(mr(0, 8, bni), 1), 8) = "*" Then
            bfshn.Cells(sr(8), 4).Value = "*�" & mr(4, 8, bni) & "�" & mr(2, 8, bni)
        Else
            bfshn.Cells(sr(8), 4).Value = "�" & mr(4, 8, bni) & "�" & mr(2, 8, bni)
        End If
    End If
    
    bfshn.Cells(sr(0), 4).Value = bni & "/" & bun & "����"   '26���@(12, 4)��(sr(0), 4)
    bfshn.Cells(sr(4), 5).Value = mr(2, 11, bni)    'koum�i�����Ƃ��j(14, 4)��(sr(4), 5)
    cnt = 0  '�����J�E���^���Z�b�g
    qq = er(11, bni) '30s81

    '�ΏۃV�[�g�́u1�v�T�� 30s81 �㑤��������z��
    If Left(mr(2, 11, bni), 2) = "����" And qq = 0 Then qq = 1 '�����̏ꍇ�i�G�j
    If Left(mr(2, 11, bni), 2) <> "����" Then '�H���ڍs����ꍇ�͂��̎��s�A�Ȃ��ꍇ�͂P�s�ڂ���T��
        Do Until Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(4, bni))).Value = "all1" Or Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(4, bni))).Value = 1
            qq = qq + 1
            If qq = 2000 Then Call oshimai("", bfn, shn, k, a, "�ΏۃV�[�g��all1��́u1�v��������Ȃ��悤�ł��B")
        Loop
        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(4, bni))).Value = "all1" Then qq = qq + 1
    Else '�����F���s����F�@����2�A����a��
        qq = qq + 1
    End If
    
    If pap8 > 0 Then
        '7�s�� ��U�������A�ۗ��@30s79 ������R�����g�}��
        If pap7 > 0 Then '���Q�ƃV�[�g�ɍ��ڍs�������ꍇ������
            Call tnsai(tst, ct3, er78(), a, sr(7), bni, 1, k - 1, -7, mr(), er(), pap7, mr8())
        End If
        '8�s�����A��U�ۗ�
        If er(11, bni) > 0 And pap8 > 0 Then '���Q�ƃV�[�g�ɍ��ڍs�������ꍇ������
            If mr(2, 11, bni) = "����" Then
                Call tnsai(tst, ct3, er78(), a, sr(8), bni, 1, qq - 1, -8, mr(), er(), 0, mr8())
            Else
                Call tnsai(tst, ct3, er78(), a, sr(8), bni, 1, Int(er(11, bni)), -8, mr(), er(), 0, mr8())
            End If
        End If
    End If
        
    '�������V�[�g�쐬30s82f
    
    If Round(kurai) = 1 Then '�㋉(1.x)
        If mr(1, 11, 1) = "������" Then
            Call kskst(pap7, h, er78(), er9(), mr9(), er3, mr3(), er5(), mr5(), c5, pap3, pap5, bni, qq, rrr, mr(), er(), a, cted())
            hirt = Range(twt.Cells(1, 1), twt.Cells(rrr + 1, 9)).Value '��rrr��rrr+1(�����΍�)
        End If '�����V�[�g�쐬�����܂�
    Else  '����(2.x)
        Call oshimai("", bfn, shn, sr(2), a, "�㋉�ȊO�ł͍����V�[�g�쐬�͂ł��܂���B")
    End If
    
    ii = qq 'rrr�͍����V�[�g�̃f�[�^�I���s(�܃��b�N���q)�������ցAii�������Ă���
    cnt = 0 '�J�E���^���Z�b�g
    pqp = 0 '���b�N�I���ۃ��Z�b�g

    If mr(1, 11, 1) = "������" Or mr(1, 11, 1) = "�ߎ�����" Then
        hiru = Range(bfshn.Cells(1, mghz), bfshn.Cells(h, mghz + 1)).Value
    End If
    '����������s��
    Do While ii <= cted(0) '628���
        ct3 = ""
        If Abs(er(4, bni)) >= 1 Then
            If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(ii, Abs(er(4, bni))).Value = "" Then
                Exit Do  '�ΏۃV�[�g��all_1��
            End If
        End If

        If mr(1, 11, 1) = "������" Then  '�������b�N�I������
            If hirt(ii, 2) = 0 Then  '���������͍����V�[�g�Q�Ƃց@30s86_012r
                UserForm4.StartUpPosition = 2 '1�@�G�N�Z���̒����A�@2�@��ʂ̒����A�@3�@��ʂ̍���
                UserForm4.Show vbModeless
                UserForm4.Repaint
                bfshn.Cells(sr(2), 5).Value = "����ۯ�" '20191118
                k0 = h0  '���ꂪ�|�C���g
                pqp = 1
                Unload UserForm4
                UserForm1.Repaint
            End If
            
            ct3 = hirt(ii + pqp, 6)  '������ct3����(���������Ŏg�p�A�]�ڃG�������g)
            
            '�����ꂷ�邽�߂ɁA�T�s�O�ł��A�����V�[�g�R��ڂɂP�𖄂ߍ��ށB�ցB
            If hirt(ii + pqp, 3) = "" Then
                c5 = 0
            Else
                c5 = 1
                qq = hirt(ii + pqp, 2) '�����V�[�g���ΏۃV�[�g�̍s�Ɋ��Z,�z��ϐ������Ac5=1�݂̂ɓK�p��(�ꉞ�o�O)
            End If
        Else  '�������ȊO�@ct3�͕K��""
            c5 = kaunta(mr(), ii, pap5, bni, er5(), mr5())  '�����ł�ii�͍����V�[�g��0�J�E���g�s
            If c5 = 1 Then qq = ii  'c5=1�݂̂ɓK�p��(�ꉞ�o�O)�@86_015b
        End If
        
        '�J�E���g�Ώۂł���Ύ��{�i�Ŗ�����Δ�΂��j
        If c5 = 1 Then
            '�������߂̕Ԃ�l��Ԃ�(�����������`��)�B�����߂Ŗ����Ă��펞�g����(am2)
            If mr(1, 11, 1) = "������" Then  '�������b�N�I������
                am2 = hirt(ii + pqp, 1)   '30s85_027����6
            Else
                am2 = wetaiou(mr(2, 1, bni), mr(1, 1, bni), qq, er3(), mr(2, 4, bni), mr(1, 11, 1), mr3(), 3) '30s51
            End If
            
            If am2 = "" Then am2 = "�[(���󔒍s)�["
            hk1 = ""

            '�ˍ��L�[unicode��?�ɂȂ�΍�J�n�@86_016q
            If InStr(am2, "?") = 0 And InStr(StrConv(am2, 24), "?") > 0 Then MsgBox "uni���蒍�Ӂi" & am2

            'p����--��--��--��--�@p��pkt�@86_014n

            If mr(1, 11, 1) = "������" Or mr(1, 11, 1) = "�ߎ�����" Then  '�������b�N�I������
               pkt = kskup(am1, am2, n1, n2, h, er(2, bni), er(6, bni), k0, h0, pap2, er2(), mr(1, 11, 1), pqp, er5(0), er3(), hiru)
            Else '�ᑬ
               pkt = tskup(am1, am2, n1, n2, h, er(2, bni), er(6, bni), k0, h0, pap2, er2(), mr(1, 11, 1), pqp, er5(0), er3())
            End If

            kahi = 0 '(���Z�۔��ʃt���O) zyou��kahi�ց@86_014j
            kasan = 0 '16s
            
            If pkt = -1 Then Exit Do
            If pkt = -2 Then '��629(85_001) pap2�̓[���ł���B�@�����x�^
                MsgBox "�����x�^"
                '���������ł̍����x�^�͂ǂ��Ȃ�̂ł��傤���H�����͋����I��
                If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0) Then Call oshimai("", bfn, shn, 1, 0, "���������A���v�H")
                
                If rrr - qq <= 1 Then '�ΏۃV�[�g�f�[�^2�s�ȉ��Ȃ獂���x�^�łȂ��ʏ���@�ɖ߂�(p=2)
                     pkt = 2
                Else
                    MsgBox "�\�󔒃x�^�\��]�ڊJ�n���܂��B" & vbCrLf & _
                    "�s���F(�\��)" & rrr - qq + 1 & vbCrLf & _
                    "�񐔁F(�\��)" & 1 _
                    , 64, "�x�^�\��u�]�ځv�J�n"
                    
                    '�����A�T��̃I�[�g�t�B���Ւn
                    Call betat4(bfn, shn, k, 0.4, k + rrr - qq, 0.4, bfn, shn, k, 1, "pp", "1")
                    
                    If Not StrConv(Left(mr(0, 2, bni), 1), 8) = "*" Then '�[�Q�ɂ�2�s���T�O���f��
                        Call betat4(bfn, shn, k, 0.4, k + rrr - qq, 0.4, bfn, shn, k, 2, "pp", "c")
                    End If
                    
                    Call cpp2("", Now(), 0, 0, 0, -1, bfn, shn, k, 5, k + rrr - qq, 5, -4163) '5��ځinow�j����ɋÏk ��
                    
                    '�ˍ���@'��������ł̒l���i���B�̂��߁j
                    Call betat4(twn, "�����V�[�g_" & syutoku(), qq, 1, rrr, 1, bfn, shn, k, Abs(er2(0)), "mm", "")
                    
                    '����
                    If mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then '���萔���Z�Ή�
                        Call betat4(bfn, shn, k, 0.4, k + rrr - qq, 0.4, bfn, shn, k, a, "pm", mr(2, 9, bni))
                    ElseIf (er(8, bni) = 0 Or er(7, bni) <> 0) And er(6, bni) > 0.5 Then  '���Z������(����]�ڂȂ���s���Ȃ�)
                        Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Int(er(6, bni)), rrr, Int(er(6, bni)), bfn, shn, k, a, "pm", "")
                    End If
                   
                    If pap7 <> pap8 Then MsgBox "p=-2��pap7 <> pap8�ł�.���ӂ�(�����͑��s)"
                        'MsgBox "�������x�^�\��ł�"
                    If er(8, bni) <> 0 Then
                        For jj = 0 To pap7 'pap7,pap8��0�̎��͈��|�b�L�����{ ii��jj
                            
                            qap = jj 'iji��jj��qap
                            If er78(1, jj) <> 0 Then
                                If er78(0, jj) = 0 Then
                                    axa = a
                                    If er(6, bni) > 0 And er(5, bni) >= 0 Then
                                        Call oshimai("", bfn, shn, 1, 0, "�������~�Ba���Z����œ���]�ڂ��悤�Ƃ��Ă��܂��B�m�F���B")
                                    End If
                                Else
                                    axa = er78(0, jj)
                                End If
                            End If
                    
                            If pap7 > 0 And jj <> pap7 Then
                                Do While er78(0, jj + 1) - er78(0, jj) = 1 And er78(1, jj + 1) - er78(1, jj) = 1
                                    jj = jj + 1   ' for����jj�𑝂₷�B
                                    If jj = pap7 Then Exit Do
                                Loop
                                'If qap <> jj Then MsgBox "�������x�^�\��ł�"
                            End If
                            DoEvents
                            Application.StatusBar = "���x�^���A" & Str(jj) & " / " & Str(pap7) & " �A " & Str(Abs(qap - jj) + 1) & "��"
                
                            If (er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then '�����񎞁@85_017 mm
                                er34 = "mm" '������w��(�[�[)
                            ElseIf er(3, bni) > 0.2 And (er(4, bni) < 0 Or er(4, bni) = 0.1) Then '�ʉݎ��@85_017 pm
                                er34 = "pm"  '�ʉݎw��i�{�[�j
                            ElseIf (er(3, bni) < 0 Or er(3, bni) = 0.1) Then
                                er34 = "mp"  '�Z�����P "ap"��"mp�i�[�{�j"�@�x�E�R�s�y�p�^�[��
                            Else
                                er34 = "pp"  '�i�{�{�j
                            End If
                            Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, er78(1, qap), rrr, er78(1, jj), bfn, shn, k, axa, er34, mr8(jj))
                        Next
                    End If
                    Exit Do
                End If
                Application.StatusBar = False
            End If  'p=-2�����܂�

            '�ˍ����AM���鎞(p=1)�̏����@n2�F���V�[�g�����s�@qq�F�ΏۃV�[�g�捞�s
            If pkt = 1 Then
                If Round(er(6, bni)) = -2 Then
                    'c=-2�̉��Z����(�ˁA��)
                    If er(7, bni) <> 0 And er(5, bni) >= 0 Then
                        Call oshimai("", bfn, shn, 1, 0, "�u-2�v�ő��񑀍�����悤�Ƃ��Ă��܂��B�C�����B")
                    End If
                    '�����K�p�p�~30s64
                    If mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then
                        kasan = tszn(er9(), bni, mr(), qq, pap9, mr9())
                        kahi = 1
                    ElseIf Round(er(9, bni)) <> 0 Then
                        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(9, bni))).Value <> "" Then
                            kasan = tszn(er9(), bni, mr(), qq, pap9, mr9())
                            kahi = 1
                        End If
                    End If
                    If kahi = 1 Then bfshn.Cells(n2, a).Value = bfshn.Cells(n2, a).Value + kasan
                ElseIf er(8, bni) <> 0 And Not (er(5, bni) < 0 And er(6, bni) > 0) Then
                    '�]�ڏ���(�A�ڌ^) ��������E(�㏑�^) ���A���@����(�㏑�^) ���A���A���A�� (��������)
                    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) > 0) Then  '30s86_012w '�������Z�^�̕�
                    '���������@������(���Z�^)�͂����̓X���[
                    'MsgBox "iei"
                    Else  '����܂ł͂����火�@������(�A�ڌ^)��������
                        If er(6, bni) = -5 Or er(6, bni) = -7 Then
                            bfshn.Cells(n2, a).Value = hunpan(mr(2, 1, bni), mr(1, 1, bni), qq, er(3, bni), er(4, bni), er(5, bni), er(6, bni), er(8, bni), mr(2, 6, bni), hk1)
                            If hk1 <> "" And er(7, bni) <> 0 Then bfshn.Cells(n2, a).Value = hk1
                        Else
                            '����������(�A�ڌ^)�ł́��������ʂ�B
                            Call tnsai(tst, ct3, er78(), a, n2, bni, pkt, qq, 0, mr(), er(), pap7, mr8()) '30s62�ꌳ��
                        End If
                    End If
                End If
                
                '���Z����(����]�ڂȂ���s���Ȃ�) ���A���A���A��
                If (er(10, bni) = 0 Or Abs(er(7, bni)) > 0.2) And (er(6, bni) > 0 Or Round(er(6, bni)) = -1) Then
                    If (er(5, bni) >= 0 And er(6, bni) < 0 And er(7, bni) <> 0) Or _
                    (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And Int(er(7, bni)) = 0 And er(8, bni) <> 0) Then
                        kasan = 1  '6�s-1�ő���]�ڎ� or ����������86_14s
                        kahi = 1
                    ElseIf mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then '���萔���Z�Ή�
                        '�����󋵂̉��Z�^�͂�����ʂ�P�[�X����i����p=1�Ȃ�j���ʂ�Ȃ���86_14s
                        kasan = tszn(er9(), bni, mr(), qq, pap9, mr9()) 'er()��ee->qq
                        kahi = 1
                    ElseIf Round(er(9, bni)) > 0 Then
                        '�����󋵂̉��Z�^�͂�����ʂ�P�[�X����i����p=1�Ȃ�j�����ʂ�Ȃ���86_14s
                        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = "" _
                            Or (mr(2, 9, bni) = "n0" And Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = 0) Then
                            '���Z�ΏۃZ������or0�Ȃ���Z�������{���Ȃ�
                        Else
                            '�����󋵂̉��Z�^�͂�����ʂ�P�[�X����i����p=1�Ȃ�j�����ʂ�Ȃ���86_14s
                            kasan = tszn(er9(), bni, mr(), qq, pap9, mr9())
                            kahi = 1
                        End If
                    End If
                    
                    If kahi = 1 Then
                        If er(2, 1) < 0 And n2 > h0 Then  'And pap2 = 0�@���P�p�@�������V�K�̊���
                            hirt(n2 - h0 + 1, 5) = hirt(n2 - h0 + 1, 5) + kasan    '5��ڏ���
                            '(�����V�[�g�T��ڎg�p�A�������V�K�̊����A�i�������������̃P�[�X����j
                        Else
                            bfshn.Cells(n2, a).Value = bfshn.Cells(n2, a).Value + kasan
                            '���������ł��̃P�[�X������@�ŏ���������̎�
                        End If
                    End If
                End If
                n1 = n2
            End If '(p=1�����܂�)
            
            '�ˍ����AM�������i�ǉ��j(p=2)�@n2�F���V�[�g�����s�@qq�F�ΏۃV�[�g�捞�s
            If pkt = 2 And er(6, bni) >= 0 Then
                '�ᑬ�V�K�͂��܂���1,2,5��ǉ��ցi�u�ǁv�W�v���̑Ή��d�l�j�ŏ��Ɂ@86_011q
                If mr(1, 11, 1) = "�m�[�}��" And bfshn.Cells(n2, 1).Value <> 1 Then
                    bfshn.Cells(n2, 1).Value = 1
                    If Not StrConv(Left(mr(0, 2, bni), 1), 8) = "*" Then '30s73
                        bfshn.Cells(n2, 2).Value = "c"
                    End If
                    bfshn.Cells(n2, 5).Value = Now() '�^�C���X�^���v�ǋL
                End If
            
                If bfshn.Cells(n2, Abs(er(2, bni))).Value <> "" Then    '1��abs (er(2,0) )��23��
                    Call oshimai("", bfn, shn, 1, 0, "�V�K�\��s�Ɋ��ɏ�񂪂���܂��B�m�F���B")
                End If
                    
                If bfshn.Cells(n2 + 1, a).Value <> "" Then Call oshimai("", bfn, shn, 1, 0, "�V�K���s�Ɋ��ɏ�񂪂���܂��B�m�F���B")

                If er(2, 1) < 0 Then  '85_007 �������ꍇ�����@�iAnd pap2 = 0���P�p�j
                        hirt(n2 - h0 + 1, 4) = am2      '4��ڏ���
                Else  '�m�[�}���i������E�P��j
                    If pap2 > 0 Then
                        '86_012j�ǉ��o�O�Ή�
                        With bfshn.Cells(n2, Abs(er(2, bni)))  '�L�[�ǉ��imghz)�m�[�}��������
                            .NumberFormatLocal = "@"
                            .Value = am2
                        End With
                    Else  '
                        With bfshn.Cells(n2, Abs(er2(0)))  '�L�[�ǉ��imghz�̃P�[�X�͔������Ȃ��Ber(2,x)�ł͂Ȃ�er(x)�Ȃ̂�
                            .NumberFormatLocal = "@"
                            .Value = am2
                        End With
                    End If
                End If
                
                If pap2 > 0 Then    '�Q�s�ڃ��Ή�(������̂Ƃ�) ���S���̕����񂪑Ώہ@mghz�ł͂Ȃ�
                    For jj = 0 To pap2  'qap��jj
                        If Round(Abs(er2(jj))) > 0 Then  '30s70 0.4(null,�[)�͖����Ƃ���(�G���[�h�~�̂���)
                            With bfshn.Cells(n2, Abs(er2(jj)))
                                .NumberFormatLocal = "@"
                                .Value = rvsrz3(am2, jj + 1, mr(2, 4, bni), 0)
                            End With
                        End If
                    Next
                End If
             
                '�]�ڏ����� �A���^ ���A���A���A���A�A��(�Ǔ_����)��������E�㏑�^ ���A�� �i����]�ځj�E�㏑�^ ���A�� ��������
                
                If er(8, bni) <> 0 And Not (er(5, bni) < 0 And (Round(er(6, bni)) = -2 Or er(6, bni) > 0)) Then  '30s69�o�O�C���i���Fer(6, bni)  0>�@)
                    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) > 0) Then  '�������Z�^�̕�
                        '���������ǉ�30s86_012s�@'�������]�ڂ͂��Ă͂����܂���ł��B�@'���������@������(���Z�^)�͂����̓X���[
                    Else  '����܂ł͂����火�@������(�A�ڌ^)�͂�����
                        '����������(�A�ڌ^)�ł́��������ʂ�B
                        Call tnsai(tst, ct3, er78(), a, n2, bni, pkt, qq, 0, mr(), er(), pap7, mr8()) '30s62�ꌳ��
                    End If
                End If

                '���Z������(����]�ڂȂ���s���Ȃ�) ���A���A��,��
                If (er(10, bni) = 0 Or er(7, bni) <> 0) And er(6, bni) > 0 Then  '���������Ή��ł�
                    '�����������Z�^�͂����ł����������B�@'�����K�p�p�~30s64
                    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And Int(er(7, bni)) = 0 And er(8, bni) <> 0) Then
                        '���������ꗥ����Ł@30s86s
                        kahi = 1
                        kasan = 1
                    ElseIf mr(2, 9, bni) <> "" And Round(er(9, bni)) = 0 Then
                        kahi = 1
                        kasan = tszn(er9(), bni, mr(), qq, pap9, mr9())
                    ElseIf Round(er(9, bni)) > 0 Then
                        If Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = "" Or _
                                (mr(2, 9, bni) = "n0" And _
                            Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, er(9, bni)).Value = 0) Then
                        Else
                            kasan = tszn(er9(), bni, mr(), qq, pap9, mr9())
                            kahi = 1
                        End If
                    End If
                    If kahi = 1 Then
                        If er(2, 1) < 0 Then  ' And pap2 = 0�@���P�p
                            hirt(n2 - h0 + 1, 5) = kasan
                            '�����������������ۂ��@�����Ȃ̂�
                        Else
                            bfshn.Cells(n2, a).Value = kasan
                        End If
                    End If
                End If
                    
                If er(2, 1) > 0 Then h0 = n2 'h0���X�V(�������ȊO) ���ꕶ�ߖڂŔ��f30s22
                '�P��Q��T��@���܂��܁����ߍŌ�ɂ܂Ƃ߂Ă�
                h = h + 1
            
            End If '(p=2�����܂�)
            
            If (pkt = 2 Or pkt = 1) Then  '-1-2�V�K�ȊO�̎� �V��
                am1 = am2
                n1 = n2
            End If
        
        End If '�J�E���g�Ώێ��s�������܂�

        ii = ii + 1
        
        If er(2, 1) < 0 Then
            Call hdrst2(ii, a, 10000, k0, h0)
        Else
            Call hdrst2(ii, a, 1000, k0, h0) '201904�@100��1000
        End If
    
    Loop '����������s�������܂�
  
    '�����ł�qq�E�E�ΏۃV�[�g�ōŏI�ŃJ�E���g�ΏۂƂ����s
  
    bfshn.Cells(2, 4).Value = k0 'Loop�I����s��30s86_002���
    bfshn.Cells(3, 4).Value = h0 'Loop�I����s��30s86_002���
   
    If mr(1, 11, 1) = "������" Or mr(1, 11, 1) = "�ߎ�����" Then Erase hiru
    
    If mr(1, 11, 1) = "������" Then '�����V�[�gA�[E����ꍞ�ݖ߂��A����A�ˍ���փx�^
        Range(twt.Cells(1, 1), twt.Cells(rrr + 1, 5)).Value = hirt  '�������΍�(rrr��rrr+1)
        Erase hirt

        '�m�[�}�����l�A���1,2,5�񏈗���
        bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
    
        '���Q��T��V�K�������̂܂Ƃ߂Ė��ߍ��݁@�������݂̂�(�ᑬ�́��Ŏ��{�ς�)
        If bfshn.Cells(1, 4).Value - 1 + k - 1 < h Then  '�u�ǁv�l���^
            '�Q��ځi���j�����炪��
            If Not StrConv(Left(mr(0, 2, bni), 1), 8) = "*" Then '�����[�Q�ɂ�2�s���T�O���f��
                Call betat4(bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 0.4, h, 0.4, bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 2, "pp", "c")
            End If
            '5��ځinow�j
            Call cpp2("", Now(), 0, 0, 0, -1, bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 5, h, 5, -4163) '5��ځinow�j����ɋÏk�@-1��-2
            '���ځi�P�j�͍Ō��
            Call betat4(bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 0.4, h, 0.4, bfn, shn, bfshn.Cells(1, 4).Value - 1 + k, 1, "pp", "1")
            DoEvents
        End If
        
        If h0 < h Then  '�V�K����Ƃ��̂� ��h1��h0�ցB������only�Ȃ̂Ł@And pap2�P�p
            '�ˍ���V�K�x�^(�P���񎞂̂�)
            If pap2 = 0 Then
                Call betat4(twn, "�����V�[�g_" & syutoku(), 2, 4, h - h0 + 1, 4, bfn, shn, h0 + 1, Abs(er2(0)), "mm", "")
            End If
            
            '����V�K�x�^�i���Z���̂݁j������]�ڂȂ���s���Ȃ�)
            '�������Ή�
            If (er(10, bni) = 0 Or er(7, bni) <> 0) And er(6, bni) > 0 Then
                '����T��
                Call betat4(twn, "�����V�[�g_" & syutoku(), 2, 5, h - h0 + 1, 5, bfn, shn, h0 + 1, a, "pm", "")
                '�����̌�A�����V�[�g�T��ڂ́A����strconv24�p�r�Ŏg�p
            End If
            
            'mghz��V�K�x�^ 30s86_011
            Application.Calculation = xlCalculationAutomatic    '�����v�Z���@�����Ɂ@'�V�`���@85_007
            
            twt.Cells(1, 5).Value = "��ASC(PHONETIC(D1))"  '�����̊֐��A�����ނɂȂ��Ă���Ȃ�
            twt.Cells(1, 5).Replace What:="��", Replacement:="=", LookAt:=xlPart, _
              SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
              ReplaceFormat:=False 'strconv24�Ɠ���
            
            Call cpp2(twn, "�����V�[�g_" & syutoku(), 1, 5, 0, 0, twn, "�����V�[�g_" & syutoku(), 2, 5, h - h0 + 1, 5, -4123) 'strconv24
            Call cpp2(twn, "�����V�[�g_" & syutoku(), 2, 5, h - h0 + 1, 5, twn, "�����V�[�g_" & syutoku(), 2, 5, 0, 0, -4163) '5���������������(�y�����邽��)
            '��������Z���͈̓R�s�y��betat4�͕s��(�]�ڌ����N���A����Ă��܂�����)
            Call betat4(twn, "�����V�[�g_" & syutoku(), 2, 5, h - h0 + 1, 5, bfn, shn, h0 + 1, mghz, "mm", "")  'mghz�ɕ�����Ŗ��߂Ȃ���΂Ȃ�Ȃ�
            
            '������
            Range(Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(2, 5), Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(h - h0 + 1, 5)).Replace What:="��", Replacement:="��", LookAt:=xlPart, _
                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                        ReplaceFormat:=False
            
            Application.Calculation = xlCalculationManual  '�Čv�Z�Ăю蓮�Ɂi�d���Ȃ邽�߁j30s66
            
            'mghz+1�� �s�ԍ����� 30s86_011
            Call cpp2("", "", 0, 0, h0 + 1, -2, bfn, shn, h0 + 1, mghz + 1, h, mghz + 1, -4163) 'mghz+1�A�s�ԍ��A�t�B��
        End If
    End If
    
    cnt = 0
    Call hdrst(ii, a)   'exitdo���l�����A�����ɂ�
    bfshn.Cells(1, a).Value = cted(0)
    '���s�v���������A�O�̂���
    bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1

    '�T�}���l�����i�]�ڕ��jp=-2�́A�]�ڂ����{
    If er(5, bni) >= 0 And er(6, bni) >= 0 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then
        '�Čv�Z����U�����ɖ߂�
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        If pap7 > pap8 Then ii = pap8 Else ii = pap7
        '���Ƃ���p=-2�̓]�ڏ���
        jj = ii  'jj�͍ŏI��
        
        bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
        
        For ii = 0 To jj 'pap7,pap8��0�̎��͈��|�b�L�����{
            Application.StatusBar = "�T�}�����A" & Str(ii) & " / " & Str(jj)
            If er78(0, ii) > 0.2 And er78(0, ii) <> a Then  '86_013f 7�s0.4�Ώ���0.1�Ώ�
                Call samaru(Int(er78(0, ii)), mr(1, 1, 1)) '�����񂳂܂�
            End If
        Next
        '�Čv�Z���Ăю蓮��
        Application.Calculation = xlCalculationManual
        Application.DisplayStatusBar = True
    End If
    
    Next
'���c�������߂ӂ��i���s�ҁj�����܂Ł�

'���d���P���߂ӂ��i����ҁj�������火
    
    bni = 1 '1���ߖڂ̒l�Ŕ��f�A���s
    mg2 = 0 'mg2�̕ϐ��͂Ƃ肠����public�œ���Ă���B������ق�

    '������ -1-2 �̋󔒖��ߍ��݃��[�`���@������Ή�30s61_4
    If er(6, bni) = -1 Or er(6, bni) = -2 Or er(6, bni) <= -3 Then
      
        If nifuku = 1 Then
            'MsgBox "2���ߖڂł̕����߂���B���ӂ��B"  '���ւց@86_016v
            Call oshimai("", bfn, shn, 1, 0, "2���ߖڕ����߂ł̎w��͋��e����Ă��܂���")
        End If
      
      'c<=-1����ύX(C<=-3 �̋󔒖��ߍ��݂͋ߎ��������݂̂��ɕ����ց@85_014
        DoEvents
        cnt = 0
        k0 = k  'k0��U���Z�b�g(�������A�܂��������瑝���Ă���)
        
        If mr(1, 11, bni) = "�ߎ�����" Then  '�ߎ��^
            If er(6, bni) = Round(er(6, bni), 0) Then '-15.1�����ʑΉ����Ȃ�
                mg2 = 1
                For ii = h To k Step -1  'i��ii
                    Do Until bfshn.Cells(ii, mghz).Value <> ""  '�g���Ă��Ȃ��H
                        ii = ii - 1
                    Loop
                    '���s(�O�s)��v��a�ɏ�񂠂�@'���������(strconv�u2�v�킹��)
                    If StrConv(bfshn.Cells(ii, mghz).Value, 2) = StrConv(bfshn.Cells(ii + 1, mghz).Value, 2) Then
                        If bfshn.Cells(bfshn.Cells(ii + 1, mghz + 1).Value, a).Value <> "" Then
                            Call tnsai(tst, ct3, er78(), a, bfshn.Cells(ii, mghz + 1).Value, bni, 1, 0, bfshn.Cells(ii + 1, mghz + 1).Value, mr(), er(), pap7, mr8())
                        End If
                    End If
                    Call hdrst(h - ii, a)  '�����X�e�[�^�X�\����
                Next
                cnt = 0
            End If
        ElseIf er(2, bni) < 0 Then
            Call oshimai("", bfn, shn, sr(2), a, "�������g�p�I���ł��B")
        Else '�m�[�}��
            If er(6, bni) = -1 Or er(6, bni) = -2 Then
                For ii = k + 1 To h  'i��ii
                    Do Until bfshn.Cells(ii, Abs(a)).Value = NullString
                        ii = ii + 1
                    Loop
                    If bfshn.Cells(ii, Abs(er(2, bni))).Value = bfshn.Cells(ii - 1, Abs(er(2, bni))).Value Then '�O�s�Ɠ����ꍇ
                        If bfshn.Cells(ii - 1, a).Value <> "" Then '-1����]�ڎ��A����܂ł̂悤��match�Ă��؂�ɕK����񂪂���Ƃ͌���Ȃ��Ȃ����̂ŁA���̑Ώ��j
                            Call tnsai(tst, ct3, er78(), a, ii, bni, 1, 0, ii - 1, mr(), er(), pap7, mr8()) '30s81qq������
                        End If
                    Else
                        '�������́��ȉ��ʂ�Ȃ�(�L�[�������̑O��ł��邽�߁B�K���G���[�ɂȂ�)�B
                        If Not IsError(Application.Match(bfshn.Cells(ii, Abs(er(2, bni))).Value, Range(bfshn.Cells(k0, Abs(er(2, bni))), bfshn.Cells(ii - 1, Abs(er(2, bni)))), 0)) Then  'match�g�p
                            m = Application.WorksheetFunction.Match(bfshn.Cells(ii, Abs(er(2, bni))).Value, Range(bfshn.Cells(k0, Abs(er(2, bni))), bfshn.Cells(ii - 1, Abs(er(2, bni)))), 0) 'h��ii-1�ɏC��30s24
                            If bfshn.Cells(k0 + m - 1, a).Value <> "" Then  '-1����]�ڎ��A����܂ł̂悤��match�Ă��؂�ɕK����񂪂���Ƃ͌���Ȃ��Ȃ����̂ŁA���̑Ώ�
                                Call tnsai(tst, ct3, er78(), a, ii, bni, 1, 0, k0 + m - 1, mr(), er(), pap7, mr8()) '30s63�W��
                            End If
                        End If
                    End If
                    Call hdrst(ii, a)    '�����X�e�[�^�X�\����
                    bfshn.Cells(2, 4).Value = k0 'Cells(13, 3)��Cells(2, 4)
                Next

                cnt = 0
            End If
        End If
    End If
    
    cnt = 0

    '�ʎ��
    '���A�ڌ^�����������ǉ��ց@86_013d
    If ((er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) < 0) Or _
        (er(9, bni) = 0 And er(10, bni) < 0)) And _
        Not ((er(3, bni) < 0 Or er(3, bni) = 0.1) And (er(4, bni) < 0 Or er(4, bni) = 0.1)) Then '(����)�Ō�́u�A�v�����B
        
        Application.Cursor = xlWait '85_026
        hirt = Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).Value  '���V�o
        For ii = k To h  '�`�`�V�o�[�W�����@i��ii
            'hirt�ł́A�����P�A�������|���{�P�@�Afor��ii�ɁA-k+1�����Ԃ���@ii��ii-k+1
            If Right(hirt(ii - k + 1, 1), 1) = mr(2, 10, bni) Then hirt(ii - k + 1, 1) = Left(hirt(ii - k + 1, 1), Len(hirt(ii - k + 1, 1)) - 1)
            If hirt(ii - k + 1, 1) = "" Then hirt(ii - k + 1, 1) = NullString '�@""���󔒂Ɂi�J�E���g�����Ȃ����߁j
            '�������̓P�K�̌����ɂȂ�Ȃ������B
            Call hdrst(ii, a)   '�����X�e�[�^�X�\����
        Next
        
        Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).Value = hirt '�V�o
        
        Erase hirt '�V�o
        Application.Cursor = xlDefault
    
    End If
    cnt = 0
    
    '����
    If (er(5, bni) < 0 And Round(kurai) = 1) And er(6, bni) <= 0 And Round(er(6, bni)) <> -2 And er(7, bni) < 0 And Not IsNumeric(bfshn.Cells(sr(7), a).Value) Then '2s
        For ii = k To h  'i��ii
            bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value & bfshn.Cells(ii, Abs(er(7, bni))).Value '��������
        Next
        Call hdrst(ii, a)    '�����X�e�[�^�X�\����
    End If
        
    '��������(er(5,0)<0)
    If (er(5, bni) < 0 And Round(kurai) = 1) Then '2s
        cnt = 0
        axa = 0
        dif = 1
        If IsNumeric(bfshn.Cells(sr(7), a).Value) Then
            axa = a + er(7, bni) '���l�F�ʒu�����Βl
        Else
            axa = Abs(er(7, bni)) '�����F�ʒu����Βl
            If er(7, bni) < 0 Then dif = -1 '�a�E�������{�t���O
        End If
        bfshn.Cells(sr(0) + 3, 3).Value = "" '30s64 sou���g�p�I����

        Application.Cursor = xlWait
        For ii = k To h
            saemp = 0
            If pap7 <> 0 Or (bfshn.Cells(ii, a).Value <> "" Or bfshn.Cells(ii, axa).Value <> "") Then '�����󗓂Ȃ���{���Ȃ�(�F�P��ȂȈȊO���{)
                If IsNumeric(bfshn.Cells(ii, axa).Value) Or IsDate(bfshn.Cells(ii, axa).Value) Then
                    saemp = bfshn.Cells(ii, axa).Value
                End If
                '������r(�V�Ή�)���]���^�͉���elseif��
                If er(6, bni) <= 0 And Round(er(6, bni)) <> -2 And dif = 1 Then
                    zzz = bfshn.Cells(ii, a).Value '�����̃f�[�^��
                    ReDim zxxz(pap7) 'pap7=8�̑O��i����������r���j
                    ReDim xxxx(pap7) 'pap7=8�̑O��i����������r���j
                    ReDim zyyz(pap7) 'pap7=8�̑O��i����������r���j20190614�A013q
                    
                    bfshn.Cells(ii, a).Value = bfshn.Cells(sr(0) + 2, 1).Value '��U�󔒂�
                        
                    zyyz() = Split(zzz, mr(2, 4, bni))  'Split�֐��őΏۃV�[�g���X�v�f����C�ɓ����
                    
                    For jj = 0 To pap7  '��r����X���Ă��� ii��jj��
                        xxxx(jj) = bfshn.Cells(ii, er78(0, jj)).Value '���񑤔�r�Ώۗ�
                        'kahi(�e�X�̃Z����r���ʂ̍���)
                        If zzz = "" Then
                            kahi = -1
                        '�������ŃG���[
                        ElseIf zyyz(jj) = "" And xxxx(jj) = "" Then 'zyz��zyyz(jj)
                            kahi = 2  '�Ȃ��Ȃ�
                        ElseIf zyyz(jj) = xxxx(jj) Then
                            kahi = 0
                        Else
                            kahi = 1
                        End If
                            
                        If mr(2, 6, bni) = "-1" Then 'op:1�@�����͋�̓I�����\�L�𗅗�p�^�[�� )1��-1�ɕύX
                            If kahi <> -1 Then  '-1�͗��񂷂炵�Ȃ�
                                If kahi = 1 Then '�������t�L
                                    zxxz(jj) = zyyz(jj)  '�Z���]�L�͍Ōに�Ajoin�֐��ɂ�
                                Else  '���f�ڂ����ikahi=0,2)
                                End If
                            End If
                        Else 'op:2�A���Ac�Acc�i�����͐��l�t���O�\�L�p�^�[���j
                            If mr(2, 6, bni) = "-2" Then  '���l�t���O�𗅗�ŕ\�L�@���Z���]�L�͍Ōに�Ajoin�֐��ɂā@2��-2�ɕύX
                                '���l�t���O�ݒ�
                                If kahi <> 2 Then
                                    zxxz(jj) = CStr(kahi) '=LTrim(Str(kahi)) -1,0,1�����܂��
                                    '�}���ӏ�(�q�d�P�Z��������̃R�����g�j
                                End If
                            Else 'op:���Ac�Acc�@ '���l�t���O�����Z�ŕ\�L
                                If kahi <> 2 Then  '��r���ʂQ�����Z���쎩�̂��s��Ȃ�
                                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value + kahi  '�Z���Ɍ��ʋL�ځi���v�l�j
                                        
                                    'different�Z���̐F�t��
                                    If kahi = 1 And Left(mr(2, 6, bni), 1) = "c" Then '�F�t��(c,cc)
                                        bfshn.Cells(ii, a).Interior.Color = 15662846 '������ɂ��F�t����13434622(��)
                                        bfshn.Cells(ii, er78(0, jj)).Interior.Color = 15662846
                                        bfshn.Cells(ii, er78(0, jj)).ClearComments
                                        If mr(2, 6, bni) = "cc" Then '����ɃR�����g�t���i��r�����j(cc)�����d��
                                            bfshn.Cells(ii, er78(0, jj)).AddComment
                                            bfshn.Cells(ii, er78(0, jj)).Comment.Text Text:=zyyz(jj) 'zyz��zyyz(jj)
                                            bfshn.Cells(ii, er78(0, jj)).Comment.Shape.TextFrame.AutoSize = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    '����^(�P���Q)�͂����ŋL�ځ�join�֐��g�p
                    If ((mr(2, 6, bni) = "-1" And zzz <> "") Or mr(2, 6, bni) = "-2") Then
                        bfshn.Cells(ii, a).Value = Join(zxxz, mr(2, 4, bni))
                    End If
                    '������r(�V�Ή�)201708�ǉ������܂�
                ElseIf er(6, bni) <= 0 And Round(er(6, bni)) <> -2 And dif = -1 Then
                    '�����������ڍs�ɂ�肱���ł͉������Ȃ����Ƃ�
                ElseIf dif = -1 And er(8, bni) >= 0 Then  '1m =0��>=0 30s69�o�N�C��
                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value + saemp '�a
                ElseIf dif = -1 And er(8, bni) < 0 Then  '<-0.5��<0�@30s69�o�N�C��
                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value * saemp '��
                ElseIf er(8, bni) >= 0 Then 'e=0��e>=0�C��16��
                    bfshn.Cells(ii, a).Value = bfshn.Cells(ii, a).Value - saemp '��
                ElseIf er(8, bni) < 0 Then  '1m�@<-0.5��<0�@30s69�o�N�C��
                    bfshn.Cells(ii, a).Value = saemp - bfshn.Cells(ii, a).Value '��(���])
                End If
            End If
            '��r�̈ڍs��͂������H
            Call hdrst(ii, a)              '�����X�e�[�^�X�\����
        Next  '�����̂����s�������[�`���͂����܂�
        
        Application.Cursor = xlDefault
        cnt = 0
    End If
  End If '��A�@-99or*,**�͒ʉ߁A�����܂�

  '�I�����̕�������(����or���܂��܁��s�v�A�ʏ�̉��Z��]��or�ˍ������̍ŏI�s�̏����ɍ��킹�鎞�̂ݗv�A)
  If er(6, bni) >= 0 Or (er(6, bni) < 0 And tst = 7 And trt <> -9) Then 'c<0��vlookup�^�����{���Ȃ���(tst7,-99�ȊO)�@86_013d
        If tst = -2 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "#,##0;[��]-#,##0"
        ElseIf tst = 0 Then
            Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = "G/�W��"
        ElseIf tst = 7 Then
            If trt = -2 Then  '���Z�^�̈ꊇ���P
                If Abs(er(9, bni)) = 0.4 Or Abs(er(9, bni)) = 0.1 Then
                    MsgBox "���Z�Œ�l(0.4�A0.1�[)�Ȃ̂ňꊇ���P�͍s���܂���B"
                Else
                    If qq > 0 Then
                    Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(9, bni))).NumberFormatLocal
                    Else
                    MsgBox "�ꊇ���P�̓X���[�ƂȂ�܂��B"  '86_016z
                    End If
                End If
            ElseIf trt = -1 Then   '�]�ڌ^�̈ꊇ���P
                If er(10, bni) < 0.5 Then
                    MsgBox "�A�ځA����(�F�]�ڏ�񏑎������݂��Ȃ�) �͑ΏۊO"
                Else
                    If qq > 0 Then
                    Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).NumberFormatLocal = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er(10, bni))).NumberFormatLocal
                    Else
                    MsgBox "�ꊇ���P�̓X���[�ƂȂ�܂��B"  '86_016z
                    End If
                End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "���������P�[�X������̂��낤���H")
            End If
        End If

        '�܂�Ԃ��đS�̂�\�����Ȃ� (LF�Ώ�)16s�@�ꗥ�K�p��18s
        If trt <> -2 Then Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).WrapText = False

  End If
  
'���d�������܂�(not98)
        
'�`�`-99�-98�p��������`�`
  If er(6, bni) <= -90 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then '��C�@20190214�@*����͎��{���Ȃ���
    
    '�Ɨ��W�v(���ύX)
    '������́u�ǁv�̑Ή��͂ǂ����邩�v����
    If InStr(1, mr(0, 2, bni), "�") > 0 And InStr(1, rvsrz3(mr(0, 2, bni), 1, "�", 0), "��") > 0 Then
        ii = h
        Do Until ii = k - 1  '�P�c���������փT�[�`
            If bfshn.Cells(ii, Abs(er(2, 1))) <> "" Then Exit Do  'bfshn.��킹��
            ii = ii - 1
        Loop
        h = ii
    End If
    '�����ŃR�s�y
    If mr(2, 6, bni) = "c" Then '85_024_�܂��g�p����Ă��Ȃ��B
        bfshn.Cells(sr(8), a).Copy  '�قڔ������Ȃ�
        Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).PasteSpecial Paste:=xlPasteAll '���ׂ� -4104
    Else '�]���^
        Call cpp2(bfn, shn, sr(8), a, 0, 0, bfn, shn, k, a, h, a, -4123)
    End If
    
    Application.Calculation = xlCalculationAutomatic    '�����v�Z���@������
    
    If er(6, bni) = -99 Then '�Z����l�ɕϊ�
        Call cpp2(bfn, shn, k, a, h, a, bfn, shn, k, a, 0, 0, -4163) '����""��Nullstring�����(�P�K�̌����A�򉻍�p)
    End If
    
    '�i����j�� �� ��
    bfshn.Cells(sr(8), a).Replace What:="=", Replacement:="��", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    '4��ڗv�f�]�L(�|�X�X�A�W�s�ڂ̂݁j���͏�Ŏ��{�ς�
        bfshn.Cells(sr(8), 4).Value = bfshn.Cells(sr(8), a).Value
    '�������ʁi�R�����g���Ɂjat ������
    c99 = bfshn.Cells(sr(8), a).Value
    c99 = Replace(c99, "��", "=")
    bfshn.Cells(sr(8), a).ClearComments
    
    bfshn.Cells(sr(8), a).AddComment
    bfshn.Cells(sr(8), a).Comment.Text Text:=c99
    bfshn.Cells(sr(8), a).Comment.Shape.TextFrame.AutoSize = True
        
    bfshn.Cells(sr(8), a).Replace What:="��", Replacement:="=", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    bfshn.Cells(1, a).Value = h  'h0��h 30s64
    
    DoEvents
    Application.Calculation = xlCalculationManual  '�Čv�Z�Ăю蓮�Ɂi�d���Ȃ邽�߁j30s66
    
    cnt = 0
    If er(5, bni) <> 1 Then  '�J�E���g��񂪋󔒍s�͂�������󔒂� 30s68
        Application.Cursor = xlWait
        For ii = k To h
            If kaunta(mr(), ii, pap5, bni, er5(), mr5()) = 0 Then '30s85_024����6
                bfshn.Cells(ii, a).Value = bfshn.Cells(sr(0) + 2, 1).Value
            End If
            Call hdrst(ii, a)            '�����X�e�[�^�X�\����
        Next
        Application.Cursor = xlDefault
    End If
    
    cnt = 0
    If er(7, bni) = 0 Then '30s84_617
        Application.Cursor = xlWait
        yayuyo = 0
        For ii = k To h  'h
            If IsError(bfshn.Cells(ii, a)) Then '30s68_2�o�O����
                yayuyo = yayuyo + 1 '85_005
                If yayuyo = 10 Then Call oshimai("", bfn, shn, k, a, "����10��")
            ElseIf bfshn.Cells(ii, a) = "" Then
                bfshn.Cells(ii, a).Value = NullString '�@""���󔒂Ɂi�J�E���g�����Ȃ����߁j�@85_027���؂�
            End If
            Call hdrst(ii, a)            '�����X�e�[�^�X�\����
        Next
    
    End If
    Application.Cursor = xlDefault
  '�`�`-99�p�����܂Ł`�`
  End If  '��C
    Range(bfshn.Cells(k, a), bfshn.Cells(h, a)).WrapText = False  '86_016t�i�܂�Ԃ����Ȃ�API��XML�c���h�~�j��������
    
  '�����X�e�[�^�X�\�����@������������o�[�W����
  DoEvents
  If flag = True Then Call oshimai("", bfn, shn, k, a, "���~���܂���")    '���~�{�^������
  Application.StatusBar = Str(cnt) & "�A" & Str(a - dd1 + 1) & " / " & Str(dd2 - dd1 + 1)
  cnt = 0
 
End If '��(not"*")
     
    '�Čv�Z�������ɖ߂�
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
     
    '�T�}���l�����i�]�ڕ��j�����ł�(-1,-2�̂�)
    If er(5, bni) >= 0 And er(6, bni) > -90 And er(6, bni) < 0 And StrConv(Left(bfshn.Cells(sr(1), a).Value, 1), 8) <> "*" Then
        If pap7 > pap8 Then jj = pap8 Else jj = pap7 'ii��jj
        
        bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1
        
        For ii = jj To 0 Step -1
            If er78(0, ii) > 0.2 And er78(0, ii) <> a Then Call samaru(Int(er78(0, ii)), mr(1, 1, 1))  '�����񂳂܂�
        Next
    End If
    
    Call samaru(a, mr(1, 1, 1))  '���񂳂܂�
    bfshn.Cells(sr(0) + 3, a).Value = Now() '���p��1�^�C���X�^���v�����
    bfshn.Cells(sr(0) + 2, 4).Value = Now() '���p��1�^�C���X�^���v����� sr(0) + 3,��sr(0) + 2�@�R�O���V�S
    '����s���ɃC���N�������g�iID�ς�����Ƃ��̂݃��Z�b�g�����j �B��bfn�Ashn���͍X�V����Ȃ��i���񕡎ʎ��̒l���ڂ��Ă邾���j�B
    twbsh.Cells(14, 3).Value = twbsh.Cells(14, 3).Value + 1 '7s (3,3)��(14,3)25s
Next '�I��͈͗񕪂̌J��Ԃ��@��
'�����ł́ua�v�́Ad+1�ł���B
   
Call giktzg(a, rog)  '���O���ʃv���V�[�W���[��_201905
   
End Sub '�O�����������܂�
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub kskst(pap7 As Long, h As Long, er78() As Currency, er9() As Currency, mr9() As String, er3() As Currency, mr3() As String, er5() As Long, mr5() As String, c5 As Long, pap3 As Long, pap5 As Long, bni As Long, qq As Long, rrr As Long, mr() As String, er() As Currency, a As Long, cted() As Long) '�����V�[�g�͏��X�ɂ������
    
    ThisWorkbook.Activate
    Sheets("�����V�[�g_" & syutoku()).Select
    DoEvents

    '�����V�[�g�����z���@�������
    Dim hirt As Variant, ii As Long, tempo As String, baba As String ', hiro As Variant
    'hiro��hirt�Ɉ�{���\���Ɓi���̂����jjj��ii�Ɉ�{��
        
    twt.Cells.Clear
    twt.Cells.Delete Shift:=xlUp
    DoEvents
    twt.Columns("A:A").NumberFormatLocal = "@"  '���ڕ������
    twt.Columns("G:G").NumberFormatLocal = "@"  '7���(��4���)������Ɂi�]�ڑO�L�[��p�r�j�@30s85_027
    twt.Columns("F:F").NumberFormatLocal = "@"  '6��ڂ�������Ɂi�]�ڑO�L�[��p�r�j�@30s86_014g
    twt.Columns("D:D").NumberFormatLocal = "@"  '�S��ڕ�����Ɂi�V�K�����a���p�r�̂݁j�ց@30s85_027
        
    rrr = qq

    If Abs(er(4, bni)) >= 1 Then  '�ŏI�s�Ɣ����L���`�F�b�N(�P��̂�)
        '�����ł�hirt�͑ΏۃV�[�g�́uALL���v��1��E�Eseek�̍Ō㌟�m�p
        hirt = Range(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(1, Abs(er(4, bni))), Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(cted(0) + 1, Abs(er(4, bni)))).Value
        '�����ł�hiru�͑ΏۃV�[�g�́u�ˍ���v��1��E�E�L�[����chk�p
        Do Until hirt(rrr, 1) = ""
            rrr = rrr + 1
            Call hdrst2(rrr, a, 10000, 0, 0)
        Loop
        Erase hirt
    Else
        '�������Ή��@30d85_018
        rrr = cted(0) + 1  '�ł���
    End If

    DoEvents
        
    rrr = rrr - 1 'rrr�͑ΏۃV�[�g�̍ŏI�s�@qq�͑ΏۃV�[�g�̃x�^�\��J�n�s
    cnt = 0
        
    '2���(�s�ԍ�)�̏���(�t�B�����p)�@�x�^�E���܋��p
    Call betat4(twn, "�����V�[�g_" & syutoku(), qq, 0.1, rrr, 0.1, twn, "�����V�[�g_" & syutoku(), qq, 2, "pp", "")
        
    '3���(�J�E���g��)�ɓ��ꍞ�ޏ������펞���邱�Ƃ�
    If Abs(er(5, bni)) <> 0 Then  '�J�E���g��̏��(�������͈�ԍ�)��3��ڂ�
        '86_012r�F20190515�������
        If pap5 = 0 And mr(2, 5, bni) = "" Then
            '�ȑO�͈ꗥ���̎d�l��
            Call cpp2(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er(5, bni)), rrr, Abs(er(5, bni)), twn, "�����V�[�g_" & syutoku(), qq, 3, 0, 0, -4163) '85_027����3
        Else '�ǉ��d�l��������@�@86_012r 20190515�������
            cnt = 0
            For ii = qq To rrr
                c5 = kaunta(mr(), ii, pap5, bni, er5(), mr5())
                If c5 = 1 Then  '�J�E���g�ΏۂȂ���{
                    If pap5 = 0 Then '�J�E���^�P����E�E�J�E���^�Z�����R�s�y
                        'Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(ii, 3).Value = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(ii, Abs(er5(0)))
                        '�����ǂ���ł��B����
                        Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(ii, 3).Value = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(ii, Abs(er(5, bni)))
                    Else  '�J�E���^������E�E�E1������d�l
                        Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(ii, 3).Value = 1
                    End If
                End If
                Call hdrst2(ii, a, 1000, 0, 0)
            Next
        End If
    Else   '�J�E���g��[�����E�E�E����3��ɓ��ꍞ�ގd�l��(����܂œ���Ė�������)
        '1����ꍞ��ł�B
        Call betat4(twn, "�����V�[�g_" & syutoku(), qq, 0.4, rrr, 0.4, twn, "�����V�[�g_" & syutoku(), qq, 3, "pp", "1")
    End If
        
    '���ڍ쐬(pap3�P�����E�������ꍇ����)
    If pap3 = 0 Then '�P���񎞁i pap3=0 �j�F�x�^�\��d�l
        '�������͍����V�[�g���Z�g���ĂȂ��B hiro��hirt
        hirt = Range(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er3(0))), Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(rrr + 1, Abs(er3(0))))
        For ii = 1 To 1 + rrr - qq
            If hirt(ii, 1) = "" Then
                hirt(ii, 1) = "�[(���󔒍s)�[" '013t�������
            Else
                hirt(ii, 1) = hirt(ii, 1) & ""      '���l��������Ƃ�����Z�@86_012����n yatto��ii
            End If
        Next
        Range(Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(qq, 1), Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(rrr, 1)).Value = hirt
        '��hirt�̈�ԉ�(�_�~�[�s)�͖�������邾��
        Erase hirt
                
        Application.Cursor = xlWait
        Application.Cursor = xlDefault
        
    Else  '�����񎞁i pap3>0 �j baba�̎b��ϐ��͌�X�P�v�[�u����(�����g���񂹂Ȃ����H)�B
            '�����F������^�Apm�F�ʉ݌^�A
        If er3(0) <> 0.1 Then '�i�U��11��ڂɓ]�ځj��1��ł́[(�F�s�ԍ��]��)�͎��{���Ȃ��B
            Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er3(0)), rrr, Abs(er3(0)), twn, "�����V�[�g_" & syutoku(), qq, 11, "mm", mr3(0))
        End If
            
        For ii = 1 To pap3 '�i�V���P�Q��ڈȍ~�ɓ]�ځj��2�񂩂�͂�����@�s�ԍ��]�ڂ�����΂���B
            Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, er3(ii), rrr, er3(ii), twn, "�����V�[�g_" & syutoku(), qq, 11 + ii, "mm", mr3(ii))
        Next

        tempo = "R[0]C[1]"
        baba = tempo
        For ii = 1 To pap3 '�����̑f����
            tempo = "&" & """" & mr(2, 4, bni) & """" & "&" & "R[0]C[" & LTrim(Str(ii + 1)) & "]"
            baba = baba & tempo
        Next
                
        Application.Calculation = xlCalculationAutomatic   '�����v�Z���@������(���s���Z�̈�)
        twt.Cells(qq, 10).FormulaR1C1 = "=" & baba  '10���(��5���)1�s�ڂɐ����̑f�𒍓�
        Range(twt.Cells(qq, 10), twt.Cells(qq, 10)).AutoFill Destination:=Range(twt.Cells(qq, 10), twt.Cells(rrr, 10))  '10���(��5���)�������t�B���A�S�s�ɁB
                
        '���悤�₭10���(��5���)���P��ڂɃR�s�[�i�������l���j
        Call betat4(twn, "�����V�[�g_" & syutoku(), qq, 10, rrr, 10, twn, "�����V�[�g_" & syutoku(), qq, 1, "mm", "")
            
        twt.Columns("J:K").ClearContents  '10��11��(��5��6��)�N���A(12��ȍ~�͓��ɃN���A�͂��ĂȂ�)
        Application.Calculation = xlCalculationManual  '�Čv�Z�Ăю蓮��
            
    End If  '���ڍ쐬�����܂�
        
'    Workbooks(bfn).Activate  '86_017h�@�������牺���ցi���퐫�m�F���j
'    Sheets(shn).Select
                    
    twt.Columns("J:J").NumberFormatLocal = "@"
        '���ڂ̔��p���@excel2019�΍�(�Ђ炪�ȃJ�^�J�i�����ꎋ����Ȃ��Ȃ�������)
        '���������łȂ����{
        '�䂭�䂭��asc,PHONETIC�g���Ĉ�C��(�E�E����Ȃ烔�ɒ���)
    For ii = qq To rrr
        twt.Cells(ii, 10).Value = StrConv(twt.Cells(ii, 1).Value, 24) '���Ή�
    Next
        
    '�����������@7��(��4��)�x�^�E���O�\�[�g�����B
    If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) <> 0) Then
            
        '8�s��(�]�ڗ�E�ŏ�����)��7��ڂɃx�^���Ɠ]��
        Call cpp2(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er(8, bni)), rrr, Abs(er(8, bni)), twn, "�����V�[�g_" & syutoku(), qq, 7, 0, 0, -4163)
        '8�s��(�]�ڗ�E�Ō㐔��)��6��ڂɃx�^���Ɠ]��
        Call cpp2(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er78(1, UBound(er78(), 2))), rrr, Abs(er78(1, UBound(er78(), 2))), twn, "�����V�[�g_" & syutoku(), qq, 6, 0, 0, -4163)
            
        '���Z��̓]��(9��)�@86_013d��
        If er9(0) <> 0.1 Then
            Call betat4(mr(2, 1, bni), mr(1, 1, bni), qq, Abs(er9(0)), rrr, Abs(er9(0)), twn, "�����V�[�g_" & syutoku(), qq, 9, "pp", mr9(0))
        Else '�U�s�[�Ȃ�A������
            Call betat4(twn, "�����V�[�g_" & syutoku(), qq, 0.4, rrr, 0.4, twn, "�����V�[�g_" & syutoku(), qq, 9, "pp", "1")
            'MsgBox "6�s�[�ł��B"
        End If
        Call saato(twn, "�����V�[�g_" & syutoku(), 3, qq, 1, rrr, 10, 99) '3��ڍ~�\�[�g�@99�͍~���̈� 1��`9��͈̔͂�3��ڂ��A
        Call saato(twn, "�����V�[�g_" & syutoku(), 7, qq, 1, rrr, 10, 1) '7���(��4���)���\�[�g
        Call saato(twn, "�����V�[�g_" & syutoku(), 10, qq, 1, rrr, 10, 1) '���\�[�g 1��10���(Excel2019�΍�E���E�Г��ꎋ����Ȃ�)
            
        ii = rrr
        cnt = 0
            
        '3��ڏ����v���O�����@�@�@'rrr��rrr+1(�󔒃R�s�y�p) 86_013v
        hirt = Range(Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(qq, 1), Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(rrr + 1, 10)).Value

        Do Until ii = qq  '  30s86_012s�@qq�s(�f�[�^�J�n�s)�͈ȉ��̑�����Ȃ��Bqq+1�s�܂ł��Ώ�
            '1��ڂ͏�ɏ�񂪂���B
            If hirt(-qq + 1 + ii, 3) <> "" Then '3��ڏ�񂠂�Ȃ�A�ȉ� hiro �� hirt ��
                If hirt(-qq + 1 + ii, 7) = "" Then '3����L,7���Ȃ�A3���(�@)�A8�񉽂����Ȃ�
                    hirt(-qq + 1 + ii, 3) = hirt(rrr - qq + 2, 3) '���Z���󔒉�
                Else  '3��4�񋤂ɏ�񂠂�@8�A9
                    'If hirt(-qq + 1 + ii, 8) = "" Then hirt(-qq + 1 + ii, 8) = hirt(-qq + 1 + ii, 9)
                    '��86_014g
                    If hirt(-qq + 1 + ii, 8) = "" And hirt(-qq + 1 + ii, 9) <> "" Then hirt(-qq + 1 + ii, 8) = hirt(-qq + 1 + ii, 9)
                            
                        '�������s�Ə�̍s��(1��10��7��)����v�Ȃ�Έȉ��i�́j7
                    If hirt(-qq + 1 + ii, 10) = hirt(-qq + 1 + ii - 1, 10) And hirt(-qq + 1 + ii, 7) = hirt(-qq + 1 + ii - 1, 7) Then  ',1)��,10)
                        hirt(-qq + 1 + ii, 3) = hirt(rrr - qq + 2, 3) '���󔒉�
                        hirt(-qq + 1 + ii, 7) = hirt(rrr - qq + 2, 3) '���󔒉�
                        hirt(-qq + 1 + ii, 6) = hirt(rrr - qq + 2, 3) '���󔒉�
                            
                        If hirt(-qq + 1 + ii, 8) <> "" Or hirt(-qq + 1 + ii - 1, 9) <> "" Then '86_014g if�����ǉ�
                            hirt(-qq + 1 + ii - 1, 8) = hirt(-qq + 1 + ii, 8) + hirt(-qq + 1 + ii - 1, 9) '8���s��8�񓖍s+9���s
                            hirt(-qq + 1 + ii, 8) = hirt(rrr - qq + 2, 3) '������A8�񓖍s�͋󔒉�
                        End If
                    End If
                End If
            ElseIf hirt(-qq + 1 + ii, 7) <> "" Then '3����i�V��7�񂠂�Ȃ炱����
                hirt(-qq + 1 + ii, 7) = hirt(rrr - qq + 2, 3) '7��󔒉�
                hirt(-qq + 1 + ii, 6) = hirt(rrr - qq + 2, 3) '6��󔒉�
            End If
            ii = ii - 1
            Call hdrst2(rrr - ii, a, 10000, 0, 0)
        Loop  '3��ڏ����v���O���������܂�
        '�������ł�ii��qq(�J�n�s)
            
        Range(Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(qq, 1), Workbooks(twn).Sheets("�����V�[�g_" & syutoku()).Cells(rrr + 1, 10)).Value = hirt
        Erase hirt
            
        Application.Cursor = xlWait
            
        '���J�n�s���ʏ��uA�@8��擪�s�@'���X��null�łW��[�����荞�ޑj�~�p
        If twt.Cells(qq, 3).Value <> "" And twt.Cells(qq, 7).Value <> "" And twt.Cells(qq, 8).Value = "" And twt.Cells(qq, 9).Value <> "" Then
            twt.Cells(qq, 8).Value = twt.Cells(qq, 9).Value  'qq=ii�ł���
        End If
            
        '���J�n�s���ʏ��uB�@6��A7��擪�s�@�i�����Ă��o�O��Ȃ���������Ȃ����j
        If twt.Cells(qq, 3).Value = "" And twt.Cells(qq, 7).Value <> "" Then
            twt.Cells(qq, 7).Value = twt.Cells(rrr + 10, 3).Value  '���󔒉�
            twt.Cells(qq, 6).Value = twt.Cells(rrr + 10, 3).Value  '���󔒉�
        End If
            
        '���W��3��ڂɃR�s�[�i�������l���j�p�^�[��A�`C�ǂ̏ꍇ���r���i�K�Ƃ���
        Call betat4(twn, "�����V�[�g_" & syutoku(), qq, 8, rrr, 8, twn, "�����V�[�g_" & syutoku(), qq, 3, "pp", "")  '8(��5)
            
        '���Z��̓]��(8��6��)�@86_013d��
        If er9(0) <> 0.1 And pap7 = 0 Then
            '���W��6��ڂɃR�s�[�i�������l���j ct3�͏펞�U��ڂ����Ă���B�]�ڗ�̖{����6��
            Call betat4(twn, "�����V�[�g_" & syutoku(), qq, 8, rrr, 8, twn, "�����V�[�g_" & syutoku(), qq, 6, "pp", "")  'from8(��5),to6(��3)
        End If
            
        Application.Cursor = xlDefault

    End If '�������������i�S��x�^�j�����܂�
        
    cnt = 0
    rrr = rrr + 1  'rrr�͍ŏI�s�̎��s(all1�I�ɂ͋󔒂ɂȂ����s)

    '���b�N���ʎq�}��
    If h >= k Then
        twt.Cells(rrr, 1).Value = bfshn.Cells(h, Abs(er(2, bni)))  'mghz(strconv24��)�̍ŉ��s
        twt.Cells(rrr, 10).Value = StrConv(bfshn.Cells(h, Abs(er(2, bni))), 24) 'mghz�̍ŉ��s�@1��10���(Excel2019�΍�
        twt.Cells(rrr, 2).Value = 0
        rrr = rrr + 1
    End If

    cnt = 0
    rrr = rrr - 1
    '���̎��_��rrr�͍����V�[�g�̃f�[�^�I���s(�܃��b�N���q)�Aqq�͑��ς�炸�f�[�^�J�n�s(�ΏۃV�[�g�y�э����V�[�g)
        
    '�����̍����V�[�g�����~���͂�����
    Call saato(twn, "�����V�[�g_" & syutoku(), 3, qq, 1, rrr, 10, 99)  '3��(���ė�)�~���@5�s�ڃ[���ł����{��86_012y
    Call saato(twn, "�����V�[�g_" & syutoku(), 10, qq, 1, rrr, 10, 1)   '���\�[�g 1��10���(Excel2019�΍�E���E�Г��ꎋ����Ȃ�)

    Workbooks(bfn).Activate  '86_017h�@�ォ�炱����ցi���퐫�m�F���j
    Sheets(shn).Select
    DoEvents

End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub giktzg(a As Long, rog As String)     '�O������
    
    Dim ii As Long, jj As Long
    '����{�^�����ɃC���N�������g�iID�ς�����Ƃ��̂݃��Z�b�g�����j �B��bfn�Ashn���͍X�V����Ȃ��i���񕡎ʎ��̒l���ڂ��Ă邾���j�B
    twbsh.Cells(13, 3).Value = twbsh.Cells(13, 3).Value + 1 '7s (2,3)��(13,3)25s
    
    '���O����������@30s76
    If rog <> "" Then     'if,rog��null�Ȃ�(�������̎��j�L�ڂ��Ȃ�
    'MsgBox rog
        jj = 1
        Do Until Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = ""
            jj = jj + 1
            If jj = 10000 Then
                MsgBox "�󔒍s��������Ȃ��悤�ł�"
                Exit Sub
            End If
        Loop

        Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = 1 '���ږ�
        Workbooks(twn).Sheets(shog).Cells(jj, 2).Value = jj '����

        'log
        Workbooks(twn).Sheets(shog).Cells(jj, 3).Value = "��.�@�O���A" & Mid(twn, 1, Len(twn) - 5) _
        & "�A�" & shn & "�" & bfn & "�A�" & shn & "�" & bfn & "�Afrom" & dd1 & "to" & dd2 _
        & "�A���ږ��Ab" & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "�A" & twbsh.Cells(2, 2).Value & "�A" _
        & Format(Now(), "yyyymmdd_hhmmss") & "�A" & bfshn.Cells(sr(8), 5).Value & "�A" & Application.WorksheetFunction.Sum(bfshn.Range("A:A")) & "�A" & dd2 - dd1 + 1
        '��������A�O��������(min:1)�Aall1(�s��)��(�{�P�ł�
        'twbsh.Cells(2, 3). (25�n)��twbsh.Cells(2, 2).�@(89�n)�ց@86_016e

        Workbooks(twn).Sheets(shog).Cells(jj, 4).Value = Format(Now(), "yyyymmdd")  'date
        Workbooks(twn).Sheets(shog).Cells(jj, 5).Value = Format(Now(), "yyyymmdd_hhmmss")  'timestamp
        Workbooks(twn).Sheets(shog).Cells(jj, 7).Value = bfn & "\" & shn  'to

        frmx = 9 'from�̊J�n��
        ii = 1
        Do Until rvsrz3(rog, ii, "��", 0) = ""
            Workbooks(twn).Sheets(shog).Cells(jj, ii + frmx - 1).Value = rvsrz3(rog, ii, "��", 0) 'from
            ii = ii + 1
            If ii = 200 Then
                Call oshimai("", bfn, shn, k, a, "���܂������ĂȂ�2�B")
            End If
        Loop
        Workbooks(twn).Sheets(shog).Cells(jj, 8).Value = ii + frmx - 2 '�ŉE��̗�ɓ����l
    End If
    '���O�������܂�
    
    Application.CutCopyMode = False

    bfshn.Cells(1, 4).Value = Application.WorksheetFunction.Sum(bfshn.Range("A:A")) + 1

    DoEvents
    'sheets(shn).Select
    'Worksheets(shn).Select    '��Excel2019�ɂȂ�A�G���[���悭�o��ӏ�
    Call oshimai("", bfn, shn, k, dd2, "")
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub teisai()
    '�V�K�H���ӏ�
    Call oshimai("", bfn, shn, 1, 0, "�H�����ł��B������������B")
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub saato(fbk As String, fsh As String, sas As Long, fmg1 As Long, fmr1 As Long, fmg2 As Long, fmr2 As Long, sot As Long)
    '�\�[�g��Afmg1�Afmg2�@�܂��͏����A�P�[�S������
    'MsgBox "ha"
    'https://excelwork.info/excel/cellsortcollection/
    If sot = 1 Then '����
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Clear
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Add Key:=Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas), Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas)), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With Workbooks(fbk).Worksheets(fsh).Sort 'Sort�I�u�W�F�N�g�ɑ΂��� '���בւ���͈͂��w�肵��
            .SetRange Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Worksheets(fsh).Cells(fmg2, fmr2))
            .Header = xlNo '1�s�ڂ��^�C�g���s���ǂ������w�肵�i�K��l�FxlNo�j
            .MatchCase = False '�啶���Ə���������ʂ��邩�ǂ������w�肵
            .Orientation = xlTopToBottom '���בւ��̕���(�s/��)���w�肵  (�K��l�FxlTopToBottom)
            .SortMethod = xlPinYin '�ӂ肪�Ȃ��g�����ǂ������w�肵  (�K��l�FxlPinYin)
            .Apply '���בւ������s���܂� �@�ȗ��͂��Ȃ���������A�O��̂������p���炵���̂�
        End With
    ElseIf sot = 99 Then '�~��
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Clear
        Workbooks(fbk).Worksheets(fsh).Sort.SortFields.Add Key:=Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas), Workbooks(fbk).Worksheets(fsh).Cells(fmg1, sas)), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With Workbooks(fbk).Worksheets(fsh).Sort
            .SetRange Range(Workbooks(fbk).Worksheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Worksheets(fsh).Cells(fmg2, fmr2))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Else
        Call oshimai("", bfn, shn, 1, 0, "sot�̈������ςł�")
    End If
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function kaunta(mr() As String, qq As Long, pap5 As Long, bni As Long, er5() As Long, mr5() As String) As Long
    kaunta = 1 '���_����
    For qap = 0 To pap5  '������̌��_����
        If Abs(er5(qap)) < 1 Then
            'A,�O��O�D�S�F�J�E���g�Ώ�
        ElseIf mr5(qap) = "��#N/A" Then
            If WorksheetFunction.IsNA(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) Then
                '���G���[�ɂȂ邗�����H
                kaunta = 0 '�J�E���g��Ώ�
                Exit For '���������Ȃ��ƕ����񎞃G���[�ɂȂ�B
            ElseIf Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value = "" Then
                kaunta = 0 '�J�E���g��Ώ�
            End If
            
        '�ȉ��̓Z����N/A����ƃG���[�ɂȂ�B20200207
        ElseIf Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value = "" Then
            'B,�Z���󔒂̓J�E���g��Ώ�
                kaunta = 0
        '�ȉ��Z���ɏ�񂠂�
        ElseIf mr5(qap) = "" Then
            'C,�J�E���g�Ώ�
        ElseIf Left(mr5(qap), 1) = "��" Then
            'D1,85_014
            If Val(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) < Val(Mid(mr5(qap), 2)) Then
                kaunta = 0 '�J�E���g��Ώ�
            End If
        ElseIf Left(mr5(qap), 1) = "��" Then
            'D1,85_014
            If Val(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value) > Val(Mid(mr5(qap), 2)) Then
                kaunta = 0 '�J�E���g��Ώ�
            End If
        ElseIf Left(mr5(qap), 1) = "�[" Or Left(mr5(qap), 1) = "��" Or mr5(qap) = "n0" Then  '�u�[�v�ǉ�86_012y
            'E�@��D��n0��E�ŏ����i�������0�ɂ��Ή��\
            '30s75 strcomp ������(�����́��ƃZ������r�����܂��s���Ȃ�����)
            If StrComp(Mid(mr5(qap), 2), Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value, vbBinaryCompare) = 0 Then
                kaunta = 0 '�J�E���g��Ώ�
            End If
        ElseIf mr5(qap) = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er5(qap))).Value Then
            'F,n0�F�J�E���g�Ώ�
        Else
            'Z,��L�ȊO�F�J�E���g��Ώہ@�@��Z����@�ŕs��v�̃P�[�X���z�肳���B
            kaunta = 0
        End If
    Next
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function tszn(er9() As Currency, bni As Long, mr() As String, qq As Long, pap9 As Long, mr9() As String) As Variant  '���Z(�����Z)
    '���Ή�6or8�s�A����a,�������ݍshx,����bni,P�l(AM�L��)�A����0����1(�g�p�I��)�A-1-2����,mr,er,6�s�ڂW�s�ڃ��̐�(pap9),mr9
    Dim qap(2) As Long
    'Dim nuez As Double
    qap(1) = 0
    qap(2) = pap9
   
    For qap(0) = qap(1) To qap(2)  '������戵���A���[�v����B
        If Abs(er9(qap(0))) = 0.4 Or Abs(er9(qap(0))) = 0.1 Then
            nuex = Val(mr9(qap(0)))
        Else '�ʏ펞
            nuex = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er9(qap(0)))).Value 'er(11, bni)��ee��qq
        End If
        If qap(0) = 0 Then
            '�����O�̂Ƃ����������Ȃ�
        Else
            nuex = nuey * nuex '������̂Q��ڈȍ~����Z���{
        End If
        nuey = nuex
    Next
    tszn = nuex
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub samaru(aa As Long, mr101 As String)  '�T�}���l����30��59���
    '���܂��Amr(1, 1, 1)���V�[�g��
    Dim gx As Long, gy As Long, ii As Long

    For ii = sr(0) + 4 To sr(0) + 5
        uu = 0
        gx = 0
        gy = 0
        If IsError(bfshn.Cells(ii, aa)) Then '30s66_2�o�O����
            uu = 1  '�Z�����G���[���l�������Ă���B
        ElseIf bfshn.Cells(ii, aa) <> "" Then
            uu = 1
        End If
        
        If uu = 1 Then '��
            '30s71_3�ꏊ�኱�ύX�A�R�O���V�S����ɕύX
            If bfshn.Cells(ii, aa).Font.Color = 2499878 Or bfshn.Cells(ii, aa).Font.Color = 255 Then '��[��]sum
                gy = sr(8) + 5
                gx = 4
            ElseIf bfshn.Cells(ii, aa).Font.Color = 16737793 Then '���l���O
                gy = sr(8) + 5 'sr(8)�͑�:�]�ڗ�̍s
                gx = 2  '4
            ElseIf bfshn.Cells(ii, aa).Font.Color = 1137350 Then '��������
                gy = sr(8) + 6 '17
                gx = 4 '2
            ElseIf bfshn.Cells(ii, aa).Font.Color = 5287937 Then '��nonzero
                gy = sr(8) + 6 '16
                gx = 2
            End If
        End If  '��
             
        If gx > 0 Then  '�R�s�y��
            Call cpp2(bfn, shn, gy, gx, gy, gx, bfn, shn, ii, aa, 0, 0, -4123) 'xlPasteFormulas) �x
            '�R�s�y�����Z����l�ɕϊ�(���s�h�̂�)
            If kurai = 1.1 And StrConv(Left(mr101, 1), 8) <> "*" Then
                Call cpp2(bfn, shn, ii, aa, ii, aa, bfn, shn, ii, aa, 0, 0, -4163) '-4163�͒l���R�s�[�@��
            Else
                '�Q�s���������̎��A�����ʉ߂��Ă�B
                'MsgBox "������ʂ�P�[�X�������ɂ���B"
            End If
            bfshn.Cells(sr(0) + 3, aa).Value = Now() '���p��1�^�C���X�^���v�����
        End If
    Next
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function tnele(ax As Long, qap0 As Long, ct3 As String, er78() As Currency, a As Long, hx As Long, bni As Long, qq As Long, fuk12 As Long, mr() As String, er() As Currency, mr8() As String) As String  '�]�ڃG�������g����@86_013j
    '�]�ڃG�������g����@30s86_013j function��
    If fuk12 > 0 Then '30s62 ���F-1-2���ʎ�
        tnele = bfshn.Cells(fuk12, ax).Value
    ElseIf fuk12 = -7 Then '��������ȉ��Afuk12��0�ȉ�
        If er78(0, qap0) = 0 Then   'bx��0��
            tnele = Trim$(CStr(Workbooks(mr(2, 0, bni)).Sheets(mr(1, 0, bni)).Cells(qq, Abs(a)).Value))
        Else
            tnele = Trim$(CStr(Workbooks(mr(2, 0, bni)).Sheets(mr(1, 0, bni)).Cells(qq, Abs(er78(0, qap0))).Value))
        End If
    '�ȍ~�Abx:1
    ElseIf er78(1, qap0) = 0.4 Then 'And mr(2, 8, bni) <> "" �̏����P�p30s59
        tnele = Trim$(mr8(qap0))     '30s75������Ή���
    ElseIf Round(er(6, bni), 0) = -15 Then    '-15 ���镶�߃G�������g���o
        tnele = rvsrz3(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), Val(mr(2, 6, bni)), mr(2, 8, bni), 0)
    ElseIf Round(er(6, bni), 0) = -14 Then '��؂萔�@30s86_016m
        tnele = kgcnt(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), mr(2, 8, bni)) + Val(mr(2, 6, bni))
    ElseIf Round(er(6, bni), 0) = -10 Then 'naka2
        tnele = Mid(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), Val(mr(2, 6, bni)))
    ElseIf Round(er(6, bni), 0) = -9 Then 'hiduke -9
        tnele = Format(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), mr(2, 6, bni))
    ElseIf Round(er(6, bni), 0) = -8 Then 'mojihen    -13��-8
        tnele = StrConv(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))), mr(2, 6, bni))
    ElseIf Abs(er78(1, qap0)) = 0.1 Then
        tnele = Format(qq, "0000000")  '30s85_014 �s�ԍ��]�ڑΉ��̏C��
    ElseIf er78(1, qap0) > 0.2 And mr8(qap0) <> "" And bfshn.Cells(hx, a).Value <> "" Then '20180813�V��Ή�
        '�������Ȃ��itnele=""�̂܂܁j
        '�����q�����p
        'MsgBox "6448"
    Else '�ʏ펞�@fuk12=-8�͂����ʂ�B
        If (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) < 0) Then
            tnele = ct3  '��������(�A�ڌ^�g�p��)
            'MsgBox "�ł��ł�"
        Else '�]���̒ʏ�p�^�[��
            tnele = Trim$(CStr(Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap0))).Value))
        End If
    End If
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub tnsai(tst As Long, ct3 As String, er78() As Currency, a As Long, hx As Long, bni As Long, p As Long, qq As Long, fuk12 As Long, mr() As String, er() As Currency, pap7 As Long, mr8() As String)
    '���Ή�7�s8�s�A����a,�����shx(0�̓R�����g�p),����bni,P�l(AM�L��)�A�Q�ƍs(�R�����g���̃P�[�X(���ڍs���Q�ƍs)������)�A-1-2����,mr,er,7�s�ڃ��̐�(pap7)
    Dim ax As Long '������i����Aa�͓���j
    Dim tenx As String, teny As String, tenz As String
    Dim qap(3) As Long, apa7 As Long    'qap3�͏����p
    Dim bx As Long, mm As Long
    
    'bx����
    If fuk12 = -7 Then bx = 0 Else bx = 1  '-7�͂V�s�R�����g�p,fuk12:-1-2���ʃt���O�̂���
    
    If er(5, bni) < 0 Or fuk12 = -7 Then apa7 = 0 Else apa7 = pap7
    '���������邢��7�s�̃R�����g����pap7����������(�ȍ~�����[�`��apa7�g�p�Apap7�s�g�p(-7�ȊO))
    
    If fuk12 > 0 Then '-1-2���ʎ� fuk12 �͓]�ڌ��̍s
        qap(1) = -1
        qap(2) = apa7
    Else
        qap(1) = 0
        If fuk12 = -7 Then  '7�s�̃R�����g�̓��ꎞ
            If pap7 <= UBound(er78(), 2) Then '���ʂȎ�
                qap(2) = 0    'pap7��0��
            Else '���ʂłȂ��Ƃ��@30s85_021�V��
                'MsgBox "�����͂������Ȃ��̂ł�"    '�V�s���̐���8�s���̐��@�͂ǂ�����oshimai�����������悤�ȁB�B
                Call oshimai("", bfn, shn, sr(7), a, "�����͂������Ȃ��̂ł�")
                qap(2) = UBound(er78(), 2)
            End If
        ElseIf fuk12 = -8 Then
            qap(2) = 0    'UBound(er78(), 2)��0��
        Else '�ʏ펞
            qap(2) = apa7  '�i�ʏ�jUBound(er78(), 2)[�W�s�ڂ̃��̐�]��apa7[7�s�ڂ̃��̐�]��
        End If
    End If
    
    For qap(0) = qap(1) To qap(2)  '������]�ڎ����[�v����B�P����̓��[�v����1��̂݁B-1-2���ʎ��A-1���烋�[�v����Bqap(2)�܂Ń��[�v����B
                                   'qap(2)�͒ʏ��8�s�ڃ��̐��B7�s�R�����g����7�s���̐��B8�s������7�s�����̎��ӎ�����B
        '���V�s�[�̎��̓X���[�ց@86_013
        If er78(0, qap(0)) = 0.1 Then   '86_013f 7�s0.1�����ɔ����Ώ�(0.4�͎g��Ȃ��d�l�ցB
            'MsgBox "7�s0.1(�[)�A�ȗ����ڂ���"
        ElseIf er78(0, qap(0)) = 0.4 Then
            Call oshimai("", bfn, shn, sr(7), a, "7�s��0.4�͍��̏����蓾�Ȃ����ƁB")
        Else
            '���񂩑��񂩁H�iax����j�]�ڌ��E�]�ڐ�Ŏg�p
            If qap(0) <= apa7 Then '�P���񁕕�����@apa7�ŏ��l�[���i�F7�s�������j
                'ax�X�V(���́A�͂ݏo�����͍X�V���Ȃ�)
                If er78(0, qap(0)) > 0 And er(5, bni) >= 0 Then '�������łȂ��A7�s���l�̎��̂ݑ��񋖗e�A(���ʂ̓]��)
                    ax = er78(0, qap(0))  '����
                Else
                    '��30s86_012w ���������Ή���
                    If er(9, bni) > 0 And er(10, bni) > 0 And er(5, bni) >= 0 And Not (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And Int(er(7, bni)) = 0 And er(8, bni) <> 0) Then
                    If er(20, bni) = 1 Then MsgBox "yaa"
                        Call oshimai("", bfn, shn, 1, 0, "b���Z����œ���]�ڂ��悤�Ƃ��Ă��܂��B�m�F���B")
                    Else
                        ax = a  '�������ł��ꔭ�ڂ����͗L���@-1�̎�
                    End If
                End If
            End If
    
            tenx = ""   '�s�v(tnele�Ń��Z�b�g�����)�����A�ꉞ�O�̂���
            tenx = tnele(ax, qap(0), ct3, er78(), a, hx, bni, qq, fuk12, mr(), er(), mr8())
            '��30s86_013j function��
            
            '�A�ڏ����k�a�l�E�E�k�`�l��肱���炪���
            If fuk12 <= 0 Then  '����-1�A-2�ȊO 30s82=��<= ��(�R�����g-7-8�Ή�)
                If qap(0) = qap(2) Then '�ŏI���݈̂ȉ�
                    If UBound(er78(), 2) > qap(2) Then '�͂ݏo�Ă���ꍇ�̂݁A�ȉ��͂ݏo�����A�������k�a�l
                        For mm = qap(2) + 1 To UBound(er78(), 2)
                            If er78(0, mm) <> 0.1 Then
                                tenx = tenx & mr(2, 4, bni) & tnele(ax, mm, ct3, er78(), a, hx, bni, qq, fuk12, mr(), er(), mr8())
                            End If
                        Next
                    End If
                End If
            End If
       
            '�A�ڏ����k�`�l(tenx�X�V)
            If fuk12 <= 0 Then  '����-1�A-2�ȊO 30s82=��<= ��(�R�����g-7-8�Ή�)
                    '�k�`�l�����̘A�ڏ���(����t��)
                    '���Ȃ��A����A�U�s�[��or-1��A�W�}�C�i�X�������A�����񋖗e��
                            
                            '���m�[�}���A�ڂ�-2�ȉ��͋��e���ĂȂ��B
                    If (er(6, bni) = 0 Or Round(er(6, bni)) = -1) Or _
                        (er(2, bni) < 0 And er(5, bni) >= 0 And er(6, bni) > 0 And (er(7, bni) = 0 Or er(7, bni) = 0.1) And er(8, bni) < 0) Then '<0��<0.3 85s023
                        '�����������A�ڌ^��or�Œǉ�
                        If mr8(qap(0)) <> "" And (er78(bx, qap(0)) < 0 Or er78(bx, qap(0)) = 0.1) Then
                            tenx = bfshn.Cells(hx, ax).Value & tenx & mr8(qap(0)) '�� ���������������� mr(2,8,bni)��mr8(qap(0))
                        ElseIf er78(bx, qap(0)) < -0.5 Then
                            tenx = bfshn.Cells(hx, ax).Value & tenx & "�A"  '��
                        Else
                            '�s�ԍ��㏑���^(�����ł͉������Ȃ�)
                        End If
                    End If
            End If
        
            '������̍ŏI���ڈȊO�������ŏ������݁A-7-8�R�����g���͂����Ŏ��{���Ȃ�(�Ō��)�B�1�2���ʏ����͂����ł͍s���Ȃ��B
            '������]�ڍs�ׂ̍ŏI����(�P����͂��̈��)�́A�����ł͎��{����fornext�̌��Ŏ��{(�ŏI�܂Ƃߕ���tenx���܂Ƃ܂��ĂȂ����߁B)
            If tenx <> "" And hx <> 0 And fuk12 >= 0 Then  'And qap(0) < apa7 �̏����P�p�@86_013q
                'tst�Atrt���g�������B-7-8�R�����g�͂����g���Ă��Ȃ����B
                If p = 2 And tst = 1 Then
                    With bfshn.Cells(hx, ax)  '������V�K���܂���(-3<0&-4<0�@���A�ڐV�K),��fuk12�͂����e�����B
                        .NumberFormatLocal = "@"
                        .Value = tenx
                    End With
                ElseIf tst = 8 And fuk12 = 0 Then  '���V�[�g�Z�����܂���
                    If Abs(er78(1, qap(0))) = 0.4 Or Abs(er78(1, qap(0))) = 0.1 Then Call oshimai("", bfn, shn, sr(8), a, "���P�^���܂��܂łͦ����͎w��ł��܂���a")
                    With bfshn.Cells(hx, ax)
                        .NumberFormatLocal = Workbooks(mr(2, 1, bni)).Sheets(mr(1, 1, bni)).Cells(qq, Abs(er78(1, qap(0)))).NumberFormatLocal
                        .Value = tenx
                    End With
                ElseIf tst = 8 And er(6, bni) <> -2 And fuk12 > 0 Then   '�Z�����ʂ��܂���(fuk12)86_012k
                    With bfshn.Cells(hx, ax)
                        .NumberFormatLocal = bfshn.Cells(fuk12, ax).NumberFormatLocal
                        .Value = tenx
                    End With
                Else 'normal(���m���Ȃ��A���̂܂ܒl�œ\��t���A�W���ł��Ȃ��i��ŕW���Ȃ�ʉ݂Ȃ菈���j)
                    bfshn.Cells(hx, ax).Value = tenx
                End If
            End If
        End If
    Next '������]�ڎ��A���[�v����B-1-2���ʎ��A-1���烋�[�v����B�́A�����܂�
    
    qap(0) = qap(0) - 1 '1�߂�(Next��̃C���N���߂�)
    '�����ł�tenx�́A�ŏI�܂Ƃߕ���tenx���܂Ƃ܂��Ă�����
    '�P����]�ڂ͕K�������ōs����Bfor���ł͍s���Ȃ��B�1�2���ʂ͂����ōs����B
    If hx = 0 Then MsgBox "hx=0�Ƃ������Ƃ����蓾�邾�낤���H"

    If (fuk12 = -8 Or fuk12 = -7) And tenx <> "" Then '30s79 8�s�R�����g�p
        bfshn.Cells(hx, a).ClearComments
        bfshn.Cells(hx, a).AddComment
        bfshn.Cells(hx, a).Comment.Text Text:=tenx
        bfshn.Cells(hx, a).Comment.Shape.TextFrame.AutoSize = True
    End If
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function wetaiou(fn As String, f As String, ii As Long, erx() As Currency, kg2 As String, spd As String, mrx() As String, gyu As Long) As String
    '�t�@�C�����A�V�[�g���A�s���A��(����)�A��؂蕶���A�����ہA������q(����)�A�Q�s�ڂ��R�s�ڂ�
    Dim am2 As String, am3 As String, qap As Long, jj As Long
    '��������̕Ԃ�l��Ԃ�(�����������`��)�B������Ŗ����Ă��펞�g����(am2)
    If Abs(erx(0)) < 1 Then
        If erx(0) > 0.3 Then
            If mrx(0) = "" Then '�����q����΂�����̋L�ڗD��Ł@85_027����8
                am2 = "" '�P����͂����͂��蓾�Ȃ��B�����߂ł���w�薳����(0.4)���Y���B
            Else
                am2 = mrx(0)  '85_027����8�@�R�s�ڑ����q
            End If
        End If
    ElseIf spd = "������" Or spd = "�ߎ�����" Then  '�P���ߎ�or������ꔭ��(0�Ԗ�)����
        am2 = CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(0))).Value)  '��������trim���{���Ȃ��ő�����ց@85�Q027����6
    Else '�P���ߎ�or������ꔭ��(0�Ԗ�)�ᑬ ���Ԃ�lnull(�Z�����󔒕Ԃ�l)�̏ꍇ������B
        am2 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(0))).Value))
    End If
    
    If UBound(erx()) > 0 Then '��������Ƃ��̂݁i�񔭖ڈȍ~�j�B�P���߂͒ʂ�Ȃ��B
        jj = UBound(erx())   '2.11(�]���ʂ�j�@�R������H
        For qap = 1 To jj
            If Abs(erx(qap)) < 1 Then
                If erx(qap) < 0.3 Then
                    am3 = Format(ii, "0000000") '�u�[�v�̎�(0.1)�A�s�ԍ����L�[�Ƃ���B
                ElseIf mrx(qap) = "" Then '(0.4)�@New�����@85_027����8
                    am3 = "" '�]��(0.4)���R�s�ڑ����q�Ȃ��@���]���^
                Else
                    am3 = mrx(qap)  '���R�s�ڑ����q����@��New�@85_027����8
                End If
            ElseIf spd = "������" Or spd = "�ߎ�����" Then    '��������trim���{���Ȃ��ő�����ց@85�Q027����6
                am3 = CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(qap))).Value)
            Else  '�ᑬ�͏]���ʂ�
                am3 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(erx(qap))).Value))
            End If
            am2 = am2 & kg2 & am3
        Next
    End If
    wetaiou = am2
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub hdrst(ii As Long, a As Long)  '21���ȑf��
    '�����X�e�[�^�X�\����  100��1000��201904
    If cnt < Int(ii / 1000) * 1000 Then  'cnt �̓p�u���b�N�ϐ�
        cnt = Int(ii / 1000) * 1000
        DoEvents
        If flag = True Then Call oshimai("", bfn, shn, 1, 0, "���~���܂����ł�") '���~�{�^������
        Application.StatusBar = Str(cnt) & "�A" & Str(a - dd1 + 1) & " / " & Str(dd2 - dd1 + 1)
    End If
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub hdrst2(ii As Long, a As Long, ak As Long, kkk As Long, hhh As Long)
    '�����X�e�[�^�X�\����
    If ak <= 0 Then ak = 100 '�ُ�l�̂Ƃ���100�̃f�t�H�l(�]���ς�炸)��
    If cnt < Int(ii / ak) * ak Then  'cnt �̓p�u���b�N�ϐ�
        If kkk <> 0 Then
            bfshn.Cells(2, 4).Value = kkk
            bfshn.Cells(3, 4).Value = hhh
        End If
        cnt = Int(ii / ak) * ak
        DoEvents
        If flag = True Then Call oshimai("", bfn, shn, 1, 0, "���~���܂�����") '���~�{�^������
        Application.StatusBar = Str(cnt) & "�A" & Str(a - dd1 + 1) & " / " & Str(dd2 - dd1 + 1)
    End If
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function hunpan(fn As String, f As String, ii As Long, g As Currency, e7 As Currency, e5 As Currency, c As Currency, e As Currency, mc As String, hk1 As String) As String
    Dim et5 As String, et As String
    '�Í����̕�����
    If mc = "1" Then 'mc�E�E�U�sop�A���̎��
        et5 = "1"
    ElseIf mc = "e5" Then  '�g���Ă���@5�s�ڃJ�E���g���ɂ��镶���񂪌�(a234567�@�Ƃ�)
        et5 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(e5))))
    ElseIf mc = "e" Then  '8�s�ړ]�ڗ�ɂ��镶���񂪌�(a234567�@�Ƃ�)�@�g���Ă��Ȃ�
        et5 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(e))))
    Else     '3�s�ړ]�ڗ�ɂ��镶���񂪌�(a234567�@�Ƃ�)�@�g���Ă��Ȃ�
        et5 = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(g))))  '
    End If
    et = Trim$(CStr(Workbooks(fn).Sheets(f).Cells(ii, Abs(e))))  '�Í����m��
    hunpan = hunk2(c, et, et5, hk1)    '�Í���
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function hunk2(cc As Currency, et As String, et5 As String, hk1 As String) As String
    '�Í������A�Ώە����A���A�G���[�r�b�g �ϐ��񕶎����@86_014k
    Dim mjs As Long '�]�ڕ����̕�����
    Dim gjs As Long '�ˍ������̕�����
    Dim pp As Long, vv As Long
    Dim hh As String '�ϊ��p�ꕶ��
    Dim uu As String '�ϊ��㕶��
    Dim tt As Long '�ϊ�unicode�V�t�g�萔
    Dim jj As Long '�͈̓`�F�b�N
    Dim qq As Long '�]��
    Dim mb As Long '�܂Ԃ��W��
    Dim ww As Long, tn As Long, ss As Long

    uu = ""
    tn = 0
    gjs = Len(et5) '�Í����̕�����
    mjs = Len(et) '�Í��Ώە����̕�����
    vv = 0
    For pp = 1 To gjs
        ss = AscW(Mid(et5, pp, 1))
        vv = vv + ss
    Next
    ww = vv Mod 10  '�Í����̂܂Ԃ��o�C�A�X�l
    vv = 0
    
    For pp = 1 To mjs 'pp�͓]�ڕ����́A���镶����
        qq = (pp - tn) Mod gjs + 1  '�Í����̕������z��
        ss = AscW(Mid(et5, qq, 1))
        mb = (ss + pp - tn + ww) Mod 50 '�܂Ԃ��W��(�O�`�{�|�S�X)
        tt = 25000 - 30 + 99 * mb '����tt
        If cc = -5 Then
            jj = 0
            tt = tt
        ElseIf cc = -7 Then
            jj = tt
            tt = (-1) * tt
        End If
        If Mid(et, pp, 1) = "�A" Then
            hh = "�A"
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
        hunk2 = "�K��O��������F" & et
        hk1 = "1"
    End If
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function yhwat1(fn As String, f As String, ee As Currency, yhs As Long, a As Long, bni As Long, kg1 As String, nkg As Long, kg2 As String, bugyo As String) As String
    '��_�t�@�C�����A��_�V�[�g���A��_���ڍs�A��_�s�ԍ��A��_��ԍ��A���ߐ��A��؂蕶��(��)�A������؂�r�b�g�A��؂蕶��(��)�A���̕��߂̌��{���i�����܂ށj
    '���z�L�[�@�������������Ή�30s43(doloop��fornext�Ɂj
    Dim ymj1 As String, ymj2 As String, mh As Long, yap As Long '�����̌�
    Dim pm As Long, ii As Long, wenk As Long  '�[���L�@�A���̏��Ԗ�(kgwe��ii)�@�A���̋�؂��
    Dim chwat1 As String, chwat2 As String '�Ԃ�l�̋�؂薈�������@�A�Ԃ�l�̑f�i�~�ό^�j
    
    If kg2 = "" Then wenk = 1 'wenk=1�F����؂薳���Awenk=0 �F����؂�L��
    pm = 1
    If Right(fn, 4) = ".xls" Then mh = 256 Else mh = 2000 '85_009�@.xls�ɂ��Ή�
    ymj1 = rvsrz3(bugyo, 2, "�", 2) 'nkg�F�Q�A�F��ȍs�Ł��������@�`���i�����胒�s�g�p�j�����e�i30s73�L�j
    If bni >= 2 And bugyo = "" Then  '30s75 ymj1��bugyo ��(��񕶐߈ȍ~�����̎��A�O�ߏ�񓥏P�ɂȂ�o�O�Ώ�)
        yhwat1 = ""   '��yhwat1�������֐��ɂȂ炴��𓾂Ȃ�����
    Else 'bni=1 �̂Ƃ��A�Ⴕ���́Aymj1��񂠂�̂Ƃ�
        yap = kgcnt(ymj1, kg2)  '�����̌�
        For ii = 1 To yap + 1
            pm = 1  '30s81_7�ǉ��i�o�O�Apm��for���Ƀ��Z�b�g���Ȃ���΂Ȃ�Ȃ��j
            ymj2 = rvsrz3(ymj1, ii, kg2, wenk)
            If Mid(ymj2, 1, 1) = "�[" Then
                pm = -1
                If ymj2 = "�[" Then ymj2 = "" Else ymj2 = Mid(ymj2, 2)
            End If
            chwat1 = ""
            If IsNumeric(ymj2) Or IsDate(ymj2) Then
                chwat1 = ymj2
            ElseIf ymj2 = "" Then
                chwat1 = pm * 0.4  '�����߂�l0.4�A��[����߂�l-0.4������
                If chwat1 = -0.4 Then chwat1 = 0.1 '85_024 -0.4��0.1
            ElseIf ee > 0 Then
                If IsError(Application.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)) Then
                    Call oshimai("", bfn, shn, yhs, a, "�i�������~�j" & vbCrLf & "�V�[�g���F" & f & " ���" & vbCrLf & "���ږ��u" & ymj2 & "�v��������܂���")
                Else
                    chwat1 = pm * Application.WorksheetFunction.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)
                End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "�u���ږ��v���u�Ȃ��v�ŕ����������Ă��܂�")
            End If
            If ii = 1 Then chwat2 = chwat1 Else chwat2 = chwat2 & kg2 & chwat1
        Next
        yhwat1 = chwat2
    End If
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function zhwat1(fn As String, f As String, ee As Currency, yhs As Long, a As Long, bni As Long, kg1 As String, nkg As Long, kg2 As String, bugyo As String) As String
    '��_�t�@�C�����A��_�V�[�g���A��_���ڍs�A��_�s�ԍ��A��_��ԍ��A���ߐ��A��؂蕶��(��)�A������؂�r�b�g�A��؂蕶��(��)�A���̕��߂̌��{���i�����܂ށj
    '���z�L�[�@�������������Ή�30s43(doloop��fornext�Ɂj
    Dim ymj1 As String, ymj2 As String, mh As Long, yap As Long '�����̌�
    Dim pm As Long, ii As Long, wenk As Long  '�[���L�@�A���̏��Ԗ�(kgwe��ii)�@�A���̋�؂��
    Dim chwat1 As String, chwat2 As String '�Ԃ�l�̋�؂薈�������@�A�Ԃ�l�̑f�i�~�ό^�j
    
    If kg2 = "" Then wenk = 1 'wenk=1�F����؂薳���Awenk=0 �F����؂�L��
    pm = 1
    If Right(fn, 4) = ".xls" Then mh = 256 Else mh = 2000 '85_009�@.xls�ɂ��Ή�
    ymj1 = rvsrz3(bugyo, 2, "�", 2) 'nkg�F�Q�A�F��ȍs�Ł��������@�`���i�����胒�s�g�p�j�����e�i30s73�L�j
    If bni >= 2 And bugyo = "" Then  '30s75 ymj1��bugyo ��(��񕶐߈ȍ~�����̎��A�O�ߏ�񓥏P�ɂȂ�o�O�Ώ�)
        zhwat1 = ""   '��zhwat1�������֐��ɂȂ炴��𓾂Ȃ�����
    Else 'bni=1 �̂Ƃ��A�Ⴕ���́Aymj1��񂠂�̂Ƃ�
        yap = kgcnt(ymj1, kg2)  '�����̌�
        For ii = 1 To yap + 1
            pm = 1  '30s81_7�ǉ��i�o�O�Apm��for���Ƀ��Z�b�g���Ȃ���΂Ȃ�Ȃ��j
            ymj2 = rvsrz3(ymj1, ii, kg2, wenk)
            If Mid(ymj2, 1, 1) = "�[" Then
                pm = -1
                If ymj2 = "�[" Then ymj2 = "" Else ymj2 = Mid(ymj2, 2)
            End If
            chwat1 = ""
            If IsNumeric(ymj2) Or IsDate(ymj2) Then
                chwat1 = ymj2
            ElseIf ymj2 = "" Then
                chwat1 = pm * 0.4  '�����߂�l0.4�A��[����߂�l-0.4������
                If chwat1 = -0.4 Then chwat1 = 0.1 '85_024 -0.4��0.1
            ElseIf ee > 0 Then
                If IsError(Application.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)) Then
                    Call oshimai("", bfn, shn, yhs, a, "�i�������~�j" & vbCrLf & "�V�[�g���F" & f & " ���" & vbCrLf & "���ږ��u" & ymj2 & "�v��������܂���")
                Else
                    chwat1 = pm * Application.WorksheetFunction.Match(ymj2, Range(Workbooks(fn).Sheets(f).Cells(ee, 1), Workbooks(fn).Sheets(f).Cells(ee, mh)), 0)
                End If
                '86_014e
                    If Val(chwat1) > 0 Then
                        chwat1 = rvsrz3(Workbooks(fn).Sheets(f).Cells(1, Abs(Val(chwat1))).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0)
                    Else
                        chwat1 = "�" & rvsrz3(Workbooks(fn).Sheets(f).Cells(1, Abs(Val(chwat1))).Address(RowAbsolute:=True, ColumnAbsolute:=False), 1, "$", 0)
                    End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "�u���ږ��v���u�Ȃ��v�ŕ����������Ă��܂�")
            End If
            If ii = 1 Then chwat2 = chwat1 Else chwat2 = chwat2 & kg2 & chwat1
        Next
        zhwat1 = chwat2
    End If
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function rvsrz3(cef As String, bni As Long, kgr As String, nkg As Long) As String '30s03
    '�Ώە�����(A:*������)�A���ߐ��A��؂蕶���A��؂肵�Ȃ��t���O �ϐ��񕶎����@86_014k
    'bni=0 (���߃[��)�����̂܂܃t���ŕԂ�
    Dim ipc As Long, bb As Long, cc As Long, prd As String
    If kgr = "" Then
        'Call oshimai("", bfn, shn, 1, 0, "���̃P�[�X�A����̂��ȁH")�@'������܂���
        prd = "�A"
    Else
        prd = kgr
    End If
    If nkg = 1 Then
        rvsrz3 = cef  '��؂�Ȃ��t���O�L�������̂܂ܕԂ�
    'nkg=2����Ȃ�����(bni)2����(����)�Ώ����[�`���B����̉�͂̎�(sr)�����g���Ȃ��B
    ElseIf bni = 2 And nkg = 2 And StrConv(Left(cef, 1), 8) <> "*" And InStr(1, cef, "�") = 0 Then 'nkg:2����Ή��@���i�������j�@bni2�����,�F��ȍs�Ŏg�p 30s81�ړ��ǑΉ�
        rvsrz3 = cef 'bni=1��2�@30��61
    ElseIf bni = 2 And nkg = 2 And StrConv(Left(cef, 1), 8) = "*" And InStr(1, cef, "�") = 0 Then  '�����i�������j�A���@bni2�����,�F��ȍs�Ŏg�p 30s81 �ړ��ǑΉ�
        If Len(cef) > 1 Then
            If StrConv(Mid(cef, 2, 1), 8) = "*" Then '**
                If Len(cef) = 2 Then
                    Call oshimai("", bfn, shn, 1, 0, "�u**�v�͎g�p����Ȃ��ł��B" & vbCrLf & "�u**���v�`���ł�낵��")
                Else 'New ���ݒ��@������
                    rvsrz3 = Mid(cef, 3) '�������@�����v��Ԃ�
                End If
            Else '�]���^
                rvsrz3 = Mid(cef, 2) '�����@�����v��Ԃ�
            End If
        Else
            rvsrz3 = cef '��
        End If
    Else 'nkg=0, nkg=2�F���L��i�擪���܂ށj
        If bni > 0 Then
            Do
                ipc = ipc + 1
                bb = cc
                cc = InStr(bb + 1, cef, prd)
            Loop Until cc = 0 Or ipc = bni '����ȍ~�Y���Ȃ�or�K�蕶�ߓ��B�Ŕ�����B
        End If
        If cc = 0 And ipc = bni Then  '�W���X�g�K�蕶�߂ŋ�ؖ����Ɂ@�O���߂�������ɓ���B
            rvsrz3 = Mid(cef, bb + 1)
        ElseIf cc = 0 And ipc < bni Then  '�K�蕶�ߖ����B(����ĊY�����߂�"")
            rvsrz3 = ""
        Else '�K�蕶�߂ŋ�؂蕶�����܂�����B
            rvsrz3 = Mid(cef, bb + 1, cc - bb - 1)
        End If
    End If
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function ctdg(rtyu As String, tyui As String, qwer As Currency, wert As Long) As Long '�@�ŏI�s��Ԃ��B'������͍s�̕�
    '�u�b�N���A�V�[�g���Aer4�n�A����
    '��������ɋÏk�Apubikou��
    ctdg = ctreg(rtyu, tyui)

    '����b(����)�ł̐���������
    If ctdg > 10000 And Abs(qwer) < 1 Then  '1000��10000�@86_014r
        Call oshimai("", bfn, shn, sr(1), wert, "�ΏۃV�[�g���ꖜ�s����(" & ctdg & ")�ł�")
    End If

End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function ctdr(rtyu As String, tyui As String, qwer As Currency, wert As Long) As Long '�@�ŉE���Ԃ��B'������͗�̕�
    '�u�b�N���A�V�[�g���Aer4�n�A����
    
    '������ ���ΏۃV�[�g�̍ŉE��@628
    ctdr = Workbooks(rtyu).Sheets(tyui).Range("A1").SpecialCells(xlLastCell).Column()

    If ctdr > 300 And Abs(qwer) < 1 Then
        Call oshimai("", bfn, shn, sr(1), wert, "�ΏۃV�[�g��300�񒴂�(" & ctdr & ")�ł�")
    End If

    ctdr = ctdr + 1
    Do Until Workbooks(rtyu).Sheets(tyui).Cells(1, ctdr).EntireColumn.Hidden = False
        ctdr = ctdr + 1
    Loop  'ctrl+end�̎���hidden�������ꍇ�̑Ώ��i85_020)
    ctdr = ctdr - 1
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function passwordGet(knt As Integer) As String
    Dim ii As Integer, aa As Integer, bb As String
    'PW�������͂Q�����ȏ�Ȃ��ƃG���[�ɂȂ�B �ϐ��񕶎����@86_014k
    For ii = 4 To knt
        aa = 0
        Do Until aa = 1 'Randomize
'            bb = rndchr(0, 9, "nsuu")
            bb = rndchr("nsuu")
            If InStr(1, passwordGet, bb) = 0 Then aa = 1
        Loop
        passwordGet = passwordGet & bb
    Next ii
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function rndchr(suei As String) As String
    Dim aa As Long, bb As Integer, san As Integer
    '����(�����p��)�擾
    Randomize
    san = Int(4 * Rnd + 1)  '1~4���擾�͈�
    If suei = "suu" Then  '�����w��
        rndchr = LTrim(Str(Int(7 * Rnd + 3)))  '0+3~6+3��3~9���擾�͈�(0~2�͎擾���O)
    ElseIf suei = "syou" Then '�p�������w��
        aa = 0
        Do Until aa = 1
            bb = Int((26 - 1 + 1) * Rnd + 1)
            If InStr(1, "aeiucklosvwxz", Chr(bb + 97 - 1)) = 0 Then aa = 1 '�擾���O���X�g
        Loop
        rndchr = Chr(bb + 97 - 1)
    ElseIf suei = "dai" Then '�p�啶���w��
        aa = 0
        Do Until aa = 1
            bb = Int((26 - 1 + 1) * Rnd + 1)
            If InStr(1, "ABCEIKOSUVWXZ", Chr(bb + 65 - 1)) = 0 Then aa = 1 '�擾���O���X�g
        Loop
        rndchr = Chr(bb + 65 - 1)
    '�������疳�w��
    ElseIf san = 1 Then '����
        rndchr = LTrim(Str(Int(7 * Rnd + 3)))  '0+3~6+3��3~9���擾�͈�(0~2�͎擾���O)
    ElseIf san >= 2 And san <= 3 Then '�p������
        aa = 0
        Do Until aa = 1
            bb = Int((26 - 1 + 1) * Rnd + 1)
            If InStr(1, "aeiucklosvwxz", Chr(bb + 97 - 1)) = 0 Then aa = 1 '�擾���O���X�g
        Loop
        rndchr = Chr(bb + 97 - 1)
    Else  '�p�啶��
        aa = 0
        Do Until aa = 1
            bb = Int((26 - 1 + 1) * Rnd + 1)
            If InStr(1, "ABCEIKOSUVWXZ", Chr(bb + 65 - 1)) = 0 Then aa = 1 '�擾���O���X�g
        Loop
        rndchr = Chr(bb + 65 - 1)
    End If
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub oshimai(msx As String, fn As String, ff As String, ii As Long, aa As Long, msbv As String)
         '��msx�͎g���ĂȂ��ł��ˁ��Ďg�p�ց@201912
    Unload UserForm1
    Unload UserForm3
    Workbooks(fn).Activate '30s76�ǉ��i�O�����O�΍�j
    DoEvents
    Worksheets(ff).Select  '30s76�ǉ��i�O�����O�΍�j
    If aa > 0 And ii > 0 Then
        Workbooks(fn).Sheets(ff).Cells(ii, aa).Select
        If msx = "" Then
            If msbv <> "" Then MsgBox msbv & vbCrLf & "(�I���Z��)"
        Else
            If msbv <> "" Then MsgBox msbv & vbCrLf & "(�I���Z��)", 289, msx
        End If
    Else
        If msx = "" Then
            If msbv <> "" Then MsgBox msbv
        Else
            If msbv <> "" Then MsgBox msbv, 289, msx
        End If
    End If
    '���O�̒�`�̍폜
    Dim nnn As Name
    For Each nnn In ActiveWorkbook.Names
        On Error Resume Next  ' �G���[�𖳎��B
        nnn.Delete
    Next
    '�t�B���^��
    If k > 1 Then bfshn.Rows(k - 1).AutoFilter         '��x���āA
    
    Application.Calculation = xlCalculationAutomatic  '�Čv�Z�����ɖ߂�
    Application.StatusBar = False
    If aa > 0 And ii > 0 Then Workbooks(fn).Sheets(ff).Cells(ii, aa).Select
    Application.Cursor = xlDefault
    End
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub ����()  ' �I��͈͂�ʃV�[�g�ɃR�s�[
    Dim a, j, x As Integer   'g1��gg1,g��gg2(���[�J���P�p,�O���[�o����)�@30s74
    Dim i As Long, jj(2) As Long, kk
    Dim shemei, wd, dg1, dg2 As String, hk1 As String, pasu As String
    Dim se_name As String, fimei As String, c99 As String
    
    '���܂��Ȃ�(�u�R�[�h�̎��s�����f����܂����v�Ώ�)
    Application.EnableCancelKey = xlDisabled
    pasu = ActiveWorkbook.Path
    
    kyosydou
    If dd1 = 0 Then Call oshimai("", bfn, shn, 1, 0, "dd1���[���ł�")
    
    '���̓��̏���`�F�b�N
    j = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(j, 1).Value = ""
        j = j + 1
        If j = 50000 Then
            MsgBox "�󔒍s��������Ȃ��悤�ł�"
            Exit Sub
        End If
    Loop
    
    If Workbooks(twn).Sheets(shog).Cells(j - 1, 4).Value <> Val(Format(Now(), "yyyymmdd")) Then
        Call oshimai("", bfn, shn, 1, 0, "���̓��̏���́A�ŏ���[����]�{�^���������ĉ������B")
    End If
    
    If bfshn.Cells(sr(0) - 1, 5) = "" Then
        Call oshimai("", bfn, shn, sr(0) - 1, 5, "�W�v������͂��ĉ�����(�ΐF�Z��)")
    End If
    

    Call iechc(hk1)  '��igchc(hk1)
    
    a = Len(bfn)
    zikan = Format(Now(), "yymmdd_hhmmss")
    shemei = bfshn.Cells(1, 3).Value & "_" & zikan
    se_name = bfshn.Cells(1, 3).Value
    If bfshn.Cells(1, 5).Value = "��" Then '30s81
        fimei = shemei
    ElseIf bfshn.Cells(1, 5).Value <> "" Then
        fimei = bfshn.Cells(1, 5).Value & "_" & zikan
    End If
    
    '�Čv�Z��������
    Application.Calculation = xlCalculationAutomatic
    
    j = 6
    i = 19967
    Do Until j >= mghz
        If bfshn.Cells(2, j).Value <> "" Then
            If Not IsError(Application.Match(bfshn.Cells(2, j).Value, Range(bfshn.Cells(2, j + 1), bfshn.Cells(2, mghz)), 0)) Then
                Range(bfshn.Cells(2, Application.Match(bfshn.Cells(2, j).Value, Range(bfshn.Cells(2, j + 1), bfshn.Cells(2, mghz)), 0) + j), bfshn.Cells(2, Application.Match(bfshn.Cells(2, j).Value, Range(bfshn.Cells(2, j + 1), bfshn.Cells(2, mghz)), 0) + j)).Select
                Call oshimai("", bfn, shn, 2, Int(j), "�����d���Z������i" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & "�Z���j")
            End If
        End If
        If bfshn.Cells(2, j).Value <> "" And IsNumeric(bfshn.Cells(3, j).Value) Then '30s76 �኱����(�R�s�ڕ����̎��͖���)
            If bfshn.Cells(3, j).Value > i Then i = bfshn.Cells(3, j).Value
        End If
        j = j + 1
        Application.StatusBar = "�d���m�F�A" & Str(j) & " / " & Str(mghz) '9s
    Loop

    i = i + 1
    Application.StatusBar = False
  
    If gg2 <= 3 Then '---Unicode����������---
        '�Čv�Z���蓮��
        Application.Calculation = xlCalculationManual
    
        For j = dd1 To dd2
            Application.StatusBar = "���ߍ��ݒ��A" & Str(j - dd1 + 1) & " / " & Str(dd2 - dd1 + 1) '9s
            If gg2 = 3 And bfshn.Cells(2, j).Value = "" And bfshn.Cells(3, j).Value = "" Then '30s76
                bfshn.Cells(3, j).Value = i
                i = i + 1
            End If
            If gg1 <= 2 And bfshn.Cells(2, j).Value = "" And bfshn.Cells(3, j).Value <> "" Then
                If Right(twbsh.Cells(2, 3).Value, 1) = "r" _
                Or twbsh.Cells(12, 3).Value = "����" Then
                    If IsError(Application.Match(ChrW(bfshn.Cells(3, j).Value), Range(bfshn.Cells(2, 6), bfshn.Cells(2, mghz)), 0)) Then
                        bfshn.Cells(2, j).Value = ChrW(bfshn.Cells(3, j).Value)
                    Else
                        Range(bfshn.Cells(2, Application.Match(ChrW(bfshn.Cells(3, j).Value), Range(bfshn.Cells(2, 6), bfshn.Cells(2, mghz)), 0) + 5), bfshn.Cells(2, Application.Match(ChrW(bfshn.Cells(3, j).Value), Range(bfshn.Cells(2, 6), bfshn.Cells(2, mghz)), 0) + 5)).Select
                        Call oshimai("", bfn, shn, 2, Int(j), "�}���\�蕶���u" & ChrW(bfshn.Cells(3, j).Value) & "�v�F" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & "�Z���Əd��")
                    End If
                Else
                    bfshn.Cells(2, j).Value = "��" & LTrim(Str(bfshn.Cells(3, j).Value)) 'entry�d�l
                End If
            ElseIf gg1 = 3 And bfshn.Cells(2, j).Value <> "" And bfshn.Cells(3, j).Value = "" Then
                bfshn.Cells(3, j).Value = AscW(bfshn.Cells(2, j).Value)
                If bfshn.Cells(3, j).Value < 0 Then bfshn.Cells(3, j).Value = bfshn.Cells(3, j).Value + 65536
                bfshn.Cells(3, j).Value = "��" & bfshn.Cells(3, j).Value
            End If
        Next
        '�Čv�Z��������
        Application.Calculation = xlCalculationAutomatic
        bfshn.Cells(3, 3).Value = i
        Application.StatusBar = False   '---Unicode�������܂�---
    ElseIf gg2 = sr(8) And gg1 = sr(8) Then
       '8�s�R�����g�R�s�[ 30s86_012_a
        If dd1 = dd2 And bfshn.Cells(sr(6), dd1).Value = -99 Then
            If TypeName(ActiveCell.Comment) = "Comment" Then '�R�����g�L��̏ꍇ
                c99 = ActiveCell.Comment.Text  '�R�����g���e
                If Left(c99, 1) = "=" Then
                    If MsgBox("�Z�����R�����g�Ɍf�ڂ̎�:" & vbCrLf & c99 & vbCrLf & "�ɁA�u�������Ă����ł����H", 289, "���O�����ɂ�") = vbOK Then 'ok��
                        c99 = Replace(c99, "��", "=")
                        ActiveCell.Value = c99
                        ActiveCell.Replace What:="��", Replacement:="=", LookAt:=xlPart, _
                            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                            ReplaceFormat:=False
                    End If
                Else
                    MsgBox "���{�ΏۊO"
                End If
            End If
        End If
    Else  '�{��
        '�㑤���炱����Ɉ����z���@86_014e
        If bfshn.Cells(1, 3) = "" Then
            Call oshimai("", bfn, shn, 1, 3, "���ʖ�����͂��ĉ�����(���F�Z��)")
        ElseIf Len(bfshn.Cells(1, 3)) > 14 Then
            Call oshimai("", bfn, shn, 1, 3, "���ʖ���14�����ȓ��Ɏ��߂ĉ�����(" & Len(bfshn.Cells(1, 3)) & ")")
        End If
                
        'Aa�F���V�[�g�A��������i���C�����j
        Selection.Copy
        'Aa�F�f�[�^�J�n�s�̎�i��̃t�B���^�����O�A�E�B���h�E�̌Œ�̂��߁j
        j = 1
        Do Until bfshn.Cells(j, 1).Value = 1 Or bfshn.Cells(j, 1).Value = "all1"
            j = j + 1
            If j = 200 Then
                MsgBox "�u1�v��������Ȃ��悤�ł�"
                Exit Sub
            End If
        Loop
        If bfshn.Cells(j, 1).Value = 1 Then k = j Else k = j + 1 'k�̓f�[�^�J�n�s(�T���v���s�ł͂Ȃ��Ȃ���)
        'Aa�F���V�[�g�A�����܂�
        'Ab�F�V�V�[�g�A��������
        Worksheets.Add  '�V�V�[�g����
        ActiveSheet.Name = shemei
    
        Range(Cells(2, 3), Cells(2, 3)).Select
        'Ab�F�\��t���i�l�Ɛ��l�̏����œ\�t�j
        Selection.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats '85_24����3����12�������Ɛ��l�̏����œ\�t ��
        
        'Ab�F�\��t���i�R�����g��\�t�j30s79
        Selection.PasteSpecial Paste:=xlPasteComments  '-4144
    
        'Ab�F������\��t��
        Selection.PasteSpecial Paste:=xlPasteFormats '-4122
        
        'Ab�F�r���ҏW
        Selection.Borders.Color = -4210753 '30s86_012i (191,191,191)
            
        '85_024���S���V�[�g
        If gg1 < k Then
            If gg1 > sr(8) Then x = gg1 Else x = sr(8) + 1 '�܂�������x�g�p(�l�̃R�s�[�J�n�s)
            bfshn.Select
            Range(Cells(x, dd1), Cells(k - 1, dd2)).Copy
            
            'Bbop�F�V�V�[�g
            Workbooks(bfn).Sheets(shemei).Select
            Range(Cells(x - gg1 + 2, 3), Cells(x - gg1 + 2, 3)).PasteSpecial Paste:=xlPasteValuesAndNumberFormats  '12 'Bb1op�F�\��t���i�l�Ɛ��l�̏����œ\�t�j
        End If
            
        Range(Cells(1, 1), Cells(1, dd2 - dd1 + 3)).Select  '��s�ڃZ�������΂�
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 15204275   '������
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
        Range(Cells(1, 2), Cells(1, dd2 - dd1 + 3)).Select  '��s��unicode�����F�قړ�����
        With Selection.Font
            .Color = -2428958
            .TintAndShade = 0
        End With
    
        Cells.FormatConditions.Delete      '�����t����������
    
        x = k - gg1 + 1 'Ab�Fx�F�]�ڃV�[�g�̃t�B���^�����O��s

        If gg2 < k - 1 Then x = gg2 - gg1 + 1 + 1 '����p�^�[��
        
        If x > 0 Then  'x>0�łȂ��p�^�[���͂Ȃ����ƁB���͂Q�ȏ�
            Range(Cells(2, 1), Cells(x, 2)).Select  'Ab�F1��2��^�e�㕔�������΂�
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 15204275   '������
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
        
        Selection.ColumnWidth = 2  'Ab�F1��2�񕝒���
        Range(Cells(1, 2), Cells(1, 2)).Select
        Cells(1, 2).Value = "."   'if���P�p�A30s74
    
        x = k - gg1 + 1 'Ab�Fx�F�]�ڃV�[�g�̃t�B���^�����O��s(��)
    
        'Ab�F�V�V�[�g�A�����܂�
        If bfshn.Cells(1, 5).Value <> "" Then '�s�v����625
            'wd = InputBox("PW����͂��ĉ������B", , passwordGet(10))�@'624�܂�
            dw = ""  '625�ȍ~�V�o�[�W����
            UserForm2.Show vbModal
            wd = dw
            If Application.Version < 16 And fmt = "csv" Then
                Call oshimai("", bfn, shn, 1, 0, "excel2013�ȑO��csv(utf-8)�ŕۑ��͂ł��܂���)")
            End If
        End If
        'Ba�F���V�[�g�A���[�^�e�n��
        bfshn.Select
    
        If x > 1 And k - gg2 < 1 Then  'k - gg2 < 2��1 (���ڍs���[�͑ΏۊO��)
    
            Range(Cells(gg1, 1), Cells(gg2, 2)).Copy     '1��2��^�e��G�c�R�s�[(all1�ƃJ�E���^c�̏�)
    
            'Bbop�F�V�V�[�g
            Workbooks(bfn).Sheets(shemei).Select '�P��Q���G�c�\��t��(�����n���Ƃ�)
            Range(Cells(2, 1), Cells(2, 1)).Select
            'Bb1op�F�\��t���i�l�Ɛ��l�̏����œ\�t�j
            Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats  '12
            With Selection.Font
                .Color = 8421504   '�˂��ݐF��
                .TintAndShade = 0
            End With
    
            'Bb2op�F�I�[�g�t�B���^�ݒ�
            Rows(x).AutoFilter
            
            'Bb2op�F�E�B���h�E�g�̌Œ�
            Range(Cells(x + 1, 3), Cells(x + 1, 3)).Select
            ActiveWindow.FreezePanes = True
        
            Range(Cells(2, 1), Cells(x, 2)).ClearContents   '30s74�@Bb2op�F1��2��^�e�㕔�������΂�
            
            Cells(x, 1).Value = "����"
            Cells(x, 2).Value = "���F�W�v�Ώہ@�@"
        
            Range(Cells(x, 1), Cells(x, 2)).Select
            With Selection 'Bb2op�F1��2��̍��ڍs�͏��
                .VerticalAlignment = xlTop
                .Orientation = -90
            End With
            Selection.Font.Size = 9 'Bb2op�F1��2��̍��ڍs�t�H���g����
       
            If x > 3 Then
                Range(Cells(2, 1), Cells(x - 1, 1)).Select 'Bb2op�F1��2��̏㕔�t�H���g�قړ�����
                With Selection.Font
                    .Color = -1572941
                    .TintAndShade = 0
                End With
            End If
    
            If gg1 < sr(0) + 6 Then  'subtotal�R�s�y�@N/A�Ή����R�s�y��
                
                'subtotal���O
                'i(mghz�R�s�y�J�n�p)��`(gg1�F�I��͈͊J�n�s�ɍ��E�����)
                
                If gg1 = sr(0) + 5 Then 'Bb2op 30s77����
                    i = sr(0) + 4 '3�s����
                    jj(1) = 1
                ElseIf gg1 < sr(0) + 5 Then
                    i = sr(0) + 3  '�S�s�W��
                    jj(1) = i - gg1 + 2 + 1
                End If
        
                'subtotal�{�ԁABb2op_opC
                
                For j = mghz + 1 To mghz Step -1  '85_024����9
                    'Bb2op_opCa�F��������A���V�[�g(subtotal�R�s�y)
                    bfshn.Select     'B�F�E�[subtotal�R�s�[
                    Range(bfshn.Cells(i, j), bfshn.Cells(sr(0) + 6, j)).Copy
                
                    'Bb2op_opCb�F��������A�V�V�[�g�isubtotal�R�s�y�j
                    Workbooks(bfn).Sheets(shemei).Select

                    If gg1 > sr(0) + 3 Then
                        Range(Cells(1, 2), Cells(1, 2)).Select '�͂ݏo���y�[�X�g
                    Else
                        Range(Cells(i - gg1 + 2, 2), Cells(i - gg1 + 2, 2)).Select '�W���y�[�X�g
                    End If
                    ActiveSheet.Paste
                
                    Range(Cells(jj(1), 2), Cells(jj(1) + 1, 2)).Columns.AutoFit
                
                    For jj(0) = jj(1) To jj(1) + 1
                        jj(2) = Range(Cells(jj(0), 2), Cells(jj(0), 2)).Font.Color
                        
                        Range(Cells(jj(0), 2), Cells(jj(0), 2)).Copy
                                                                
                        For ii = jj(1) To jj(1) + 1
                            For kk = 3 To dd2 - dd1 + 3
                                If WorksheetFunction.IsNA(Cells(ii, kk).Value) Then 'N/A�����ǉ�(85_024����9
                                    If Cells(ii, kk).Font.Color = jj(2) Then
                                        Range(Cells(ii, kk), Cells(ii, kk)).Select
                                        ActiveSheet.Paste
                                        '�������H
'                                        Range(Cells(ii, kk), Cells(ii, kk)).Paste
                                    End If
                                ElseIf Cells(ii, kk).Font.Color = jj(2) And Cells(ii, kk) <> "" Then '30s77null�����ǉ�(�]��)
                                        Range(Cells(ii, kk), Cells(ii, kk)).Select
                                        ActiveSheet.Paste
                                        '�������H
'                                        Range(Cells(ii, kk), Cells(ii, kk)).Paste
                                End If
                            Next
                        Next
                    Next
                Next
            'Bb2op_opC�����܂�
            End If
        'Bb2op�����܂�
        End If
        
        Range(Cells(2, 1), Cells(2, 1)).Select
        'Bb�F�V�V�[�g�A�����܂�
    
        'D�F��������A���V�[�g(�ْ̍���)
        'Da
        bfshn.Select
    
        Range(Cells(1, 1), Cells(sr(0) + 5, mghz - 1)).Select  'Da1�F1-18�s(�W��)�𒲐�
        With Selection.Font  'Da�F�t�H���g����
            .Name = "�l�r �o�S�V�b�N"  'Da�F�������Win10�e��(���C���I�Ȃ�)�e���󂯂Ȃ��Ȃ�B
            .Size = 9
        End With
    
        Range(Cells(sr(0) + 3, 5), Cells(sr(0) + 3, mghz - 1)).Select 'Da1�F�W15�s����
        With Selection.Font  'Da�F�t�H���g�A�F����
            .Name = "Haettenschweiler"
            .Size = 10
        End With
        Selection.NumberFormatLocal = "m/d h:mm"
    
        Range(Cells(sr(0), 5), Cells(sr(0), mghz - 1)).Select  'Da1�F�W12�s����
        With Selection.Font  '�t�H���g�A�F����
            .Name = "Haettenschweiler"
            .Size = 10
        End With
        Selection.NumberFormatLocal = "m/d h:mm"
        
        'Da1�F�O���W�v�l��                                              '�Ԃ��o�鎖���蒍�Ӂ�
        Range(Cells(sr(0) + 1, 5), Cells(sr(0) + 2, mghz - 1)).NumberFormatLocal = "#,##0;[��]-#,##0"
        'Da1�F�O���W�v�l��                                              '�Ԃ��o�鎖���蒍�Ӂ�
        Range(Cells(sr(0) + 4, 5), Cells(sr(0) + 5, mghz - 1)).NumberFormatLocal = "#,##0;[��]-#,##0"
    
        '��86_014m
        Range(Cells(sr(0), 3), Cells(sr(0) + 1, 5)).Select 'Da1�F�ԃZ�����t�H���g�Ē���1��2
        With Selection.Font
            .Name = "�l�r �o�S�V�b�N"
            .Size = 9
        End With
        Selection.NumberFormatLocal = "G/�W��" '��with�ɓ����ƃG���[�ɂȂ�B
    
        'Da2�F���V�[�g�񕝂��R�s�[
        Range(Cells(2, dd1), Cells(2, dd2)).Copy '�񕝃R�s��i2�s�ڂɂāj
        
        Range(Cells(gg1, dd1), Cells(gg1, dd1)).Select  '�J�n�Z���i�F����j
        dg1 = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) '�J�n�Z���]�L
        Range(Cells(gg2, dd2), Cells(gg2, dd2)).Select  '�I���Z���i�F�E���j
        dg2 = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) '�I���Z���]�L
        Range(Cells(gg1, dd1), Cells(gg2, dd2)).Select '�I��͈͖͂߂�
        'Da�F���V�[�g�A�����܂�
        
        'Db�F��������A�V�V�[�g�i�񕝂̂݃R�s�y�j
        Workbooks(bfn).Sheets(shemei).Select
        
        Range(Cells(1, 3), Cells(1, 3)).PasteSpecial Paste:=xlPasteColumnWidths '8'�񕝂�\��t��
    
        If x > 1 And k - gg2 < 2 Then  'Dbop_Unicode�̓\��t��(���ڍs(k-1)���܂܂�Ă��邱�Ƃ��\��t���̏���)
        'Dbop�Fx�F�t�B���^�����O��s
            Range(Cells(1, 3), Cells(1, 3)).PasteSpecial Paste:=xlPasteValues '-4163
        End If
        
        'Db
        Cells(1, 1).Value = "��.�@���ʁA" & twn & "�A" & shemei & "�A�" & shn _
        & "�" & bfn & "�A" & dg1 & dg2 & "�A��ŗL��"
        Range(Cells(2, 2), Cells(2, 2)).Select
    
        '���O���������灨kyosydou�ֈڐ�(Activate�����P�p)
        j = 1
        Do Until Workbooks(twn).Sheets(shog).Cells(j, 1).Value = ""
            j = j + 1
            If j = 10000 Then
                MsgBox "�󔒍s��������Ȃ��悤�ł�"
                Exit Sub
            End If
        Loop

        Workbooks(twn).Sheets(shog).Cells(j, 1).Value = 1 '���ږ�
        Workbooks(twn).Sheets(shog).Cells(j, 2).Value = j  '����
        
        'log 'twbsh.Cells(2, 3). (25�n)��twbsh.Cells(2, 2).�@(89�n)�ց@86_016e
        Workbooks(twn).Sheets(shog).Cells(j, 3).Value = Workbooks(bfn).Sheets(shemei).Cells(1, 1).Value & "�Az" & wd & "z" & "b" _
        & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "�A" & twbsh.Cells(2, 2).Value & "�A" & _
        Format(Now(), "yyyymmdd_hhmmss") & "�A" & bfshn.Cells(sr(8), 5).Value & "�A" & gg2 - gg1 + 1 & "�A" & dd2 - dd1 + 1
        
        Workbooks(twn).Sheets(shog).Cells(j, 4).Value = Format(Now(), "yyyymmdd") 'date
        Workbooks(twn).Sheets(shog).Cells(j, 5).Value = Format(Now(), "yyyymmdd_hhmmss") 'timestamp
        If fimei = "" Then
            Workbooks(twn).Sheets(shog).Cells(j, 7).Value = shemei 'to
        Else
            Workbooks(twn).Sheets(shog).Cells(j, 7).Value = fimei & "." & fmt & "\" & shemei 'to
        End If
        Workbooks(twn).Sheets(shog).Cells(j, 8).Value = 9  '�ŉE��(���ʂ͌Œ�l)
        Workbooks(twn).Sheets(shog).Cells(j, 9).Value = bfn & "\" & shn  'from
        
        '���O�������܂�
 
        '�V�[�g��ʃt�@�C���Ƃ��Ă�����(E1�Z�����L�̎�)30s81
        If bfshn.Cells(1, 5).Value <> "" Then
            'Eb
            Workbooks(bfn).Activate
            Sheets(shemei).Select
            Sheets(shemei).Copy
    
            '30s79�ǉ��A�V�t�@�C���̃t�H���g����S�V�b�N�ł͂Ȃ��AMSP�S�V�b�N�d�l��
            If Application.Version < 16 Then
                'MsgBox "excel2013�ȑO�ł�"
            Else                  'MsgBox "excel2016�ȍ~�ł�"
                If IsNumeric(syutoku()) Then
                    '�����m�{��PC
                    ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                    "C:\Program Files (x86)\Microsoft Office\Root\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml")
                Else
                    '��fmv��PC
                    'ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                        "C:\Program Files\WindowsApps\Microsoft.Office.Desktop_16051.12228.20364.0_x86__��������\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml" _
                        )
                    '20200115��Rev�オ�������ƁB�܂����̍��́Astore��
                    'ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                        "C:\Program Files\WindowsApps\Microsoft.Office.Desktop_16051.12325.20288.0_x86__��������\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml" _
                        )
                    '202004��Rev�@(store�ł���DL�ł�)
                    ActiveWorkbook.Theme.ThemeFontScheme.Load ( _
                        "C:\Program Files (x86)\Microsoft Office\Root\Document Themes 16\Theme Fonts\Office 2007 - 2010.xml" _
                        )
                End If
            End If
            
            If fmt = "xlsx" Then
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlOpenXMLWorkbook, Password:=wd, CreateBackup:=False    'xlsx�t�H�[�}�b�g
            ElseIf fmt = "csv" Then
                Rows("1:1").Select
                Selection.Delete Shift:=xlUp
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlCSVUTF8, CreateBackup:=False                          'csv(utf-8)�t�H�[�}�b�g
            ElseIf fmt = "csvsjis" Then
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlCSV, CreateBackup:=False                              'csv�t�H�[�}�b�g
            Else
                ActiveWorkbook.SaveAs Filename:=pasu & "\" & fimei, _
                FileFormat:=xlExcel12, Password:=wd, CreateBackup:=False            'xlsb�t�H�[�}�b�g
            End If
            
            ActiveWorkbook.Save      '�㏑�ۑ�
            
            '627��������PW��Excel�V�V�[�g��
            If wd <> "" Then
                ThisWorkbook.Activate
                Sheets("�����V�[�g_" & syutoku()).Select
                Sheets("�����V�[�g_" & syutoku()).Copy
                Sheets("�����V�[�g_" & syutoku()).Name = shemei & "��PW"
                
                Cells(2, 1).Value = "�V�[�g���F" & shemei
                
                Cells(4, 1).Value = fimei & "." & fmt
                Cells(5, 1).Value = "�o�v�F" & wd
                Cells(7, 1).Value = "�����p"
                Range(Cells(1, 1), Cells(1, 1)).Select
                Shell "c:\windows\system32\notepad.exe", vbNormalFocus 'PW�p�����������グ
            End If
        End If
    End If
    Application.CutCopyMode = False
End Sub  '���ʂ����܂�
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub ����̂�()
    Dim hk1 As String, mghx As Long
    
    kyosydou  '���ʏ���
 
    If Not (shn = "���W�v_���`" And bfn = twn) Then
        bfshn.Cells(sr(0) + 1, 3).Value = ""    '13��sr(0)+1
        bfshn.Cells(sr(0) + 2, 3).Value = ""     '14��sr(0)+2
    End If

    Call iechc(hk1)

    mghx = Application.Match("�B", Range(twbsh.Cells(1, 1), twbsh.Cells(1, 5000)), 0) 'mghx�̓}�N���t�@�C���́u�B�v�̗�,mghz�͓��V�[�g�́A
    
    jj = 6 'j��jj 30s82
    ii = 19967 'i��ii 30s82
    Do Until jj >= mghz
        If bfshn.Cells(2, jj).Value <> "" Then
            If Not IsError(Application.Match(bfshn.Cells(2, jj).Value, Range(bfshn.Cells(2, jj + 1), bfshn.Cells(2, mghz)), 0)) Then
                Range(bfshn.Cells(2, Application.Match(bfshn.Cells(2, jj).Value, Range(bfshn.Cells(2, jj + 1), bfshn.Cells(2, mghz)), 0) + jj), bfshn.Cells(2, Application.Match(bfshn.Cells(2, jj).Value, Range(bfshn.Cells(2, jj + 1), bfshn.Cells(2, mghz)), 0) + jj)).Select
                Call oshimai("", bfn, shn, 2, Int(jj), "�����d���Z������i" & ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False) & "�Z���j")
            End If
        End If
        If bfshn.Cells(2, jj).Value <> "" Then '17s
            If bfshn.Cells(3, jj).Value > ii Then ii = bfshn.Cells(3, jj).Value
        End If
        jj = jj + 1
        Application.StatusBar = "�d���m�F�A" & Str(jj) & " / " & Str(mghz) '9s
    Loop

    '�Čv�Z����U������
    Application.Calculation = xlCalculationAutomatic
    Application.ExtendList = False '�f�[�^�͈͊g��:�I�t�i�F�E�׃Z��������ɏ����ς����̂�j�~�j"
    
    '�I�[�g�t�B���^���ݒ肳��Ă��邩�ǂ������f������
    If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    Application.EnableAutoComplete = False  '�I�[�g�R���v���[�g
       
    ThisWorkbook.Activate
    
    jj = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = ""
        jj = jj + 1
        If jj = 10000 Then
            MsgBox "�󔒍s��������Ȃ��悤�ł�"
            Exit Sub
        End If
    Loop
    
    Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = 1 '���ږ�
    Workbooks(twn).Sheets(shog).Cells(jj, 2).Value = jj  '����
    
    Workbooks(twn).Sheets(shog).Cells(jj, 3).Value = _
    "��.�@����A" & _
    twn _
    & "�A�" & shn & "�" & bfn _
    & "�A����W�v_���`" & "�" & twn _
    & "�Afrom" & dd1 & "to" & dd2 _
    & "�A���ږ��Ab" & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "�A" & twbsh.Cells(2, 2).Value & "�A" _
    & Format(Now(), "yyyymmdd_hhmmss") & "�A" & bfshn.Cells(sr(8), 5).Value & "�A0�A0"
    Workbooks(twn).Sheets(shog).Cells(jj, 4).Value = Format(Now(), "yyyymmdd") 'date
    Workbooks(twn).Sheets(shog).Cells(jj, 5).Value = Format(Now(), "yyyymmdd_hhmmss") 'timestamp
    Workbooks(twn).Sheets(shog).Cells(jj, 7).Value = bfn & "\" & shn 'to
    Workbooks(twn).Sheets(shog).Cells(jj, 8).Value = 9  '�ŉE��(���ʂ͌Œ�l)
    Workbooks(twn).Sheets(shog).Cells(jj, 9).Value = twn & "\���W�v_���`"   'from
    '���O�������܂�
    
    Call cpp2(twn, "���W�v_���`", 14, 1, 18, 5, bfn, shn, sr(0) - 1 + 5 - 2, 1, 0, 0, -4104) '�T�}���֐����ӊۂ��ƃR�s�y�Ȃ̂łS�P�O�S
    Call cpp2(twn, "���W�v_���`", 15, mghx, 18, mghx + 2, bfn, shn, sr(0) - 1 + 5 - 1, mghz, 0, 0, -4104) '���A�T�}���֐����ӊۂ��ƃR�s�y(mghz��)
    Call cpp2(twn, "���W�v_���`", 11, 5, 15, 6, bfn, shn, sr(0) - 1, 5, 0, 0, -4122)  '���A�T�}���֐����ӊۂ��ƃR�s�y(mghz��)
    
    Workbooks(bfn).Activate '����.copy��A����������ɓ����ƁA�Z�����������I������Ă��閭�ȉf��͉����������ۂ�
    Sheets(shn).Select
    
    '30s86_012i (191,191,191)
    Range(bfshn.Cells(2, mghz - 2), bfshn.Cells(21, mghz)).Borders.Color = -4210753
    
    bfshn.Cells(sr(0) - 1, 5).Select  '�ΐF�Z��

    '�Čv�Z���蓮��
    Application.Calculation = xlCalculationManual

    With Application.AutoCorrect      '�I�[�g�R���N�g�����Ȃ��@�R�O���T�Q
        .TwoInitialCapitals = False
        .CorrectSentenceCap = False
        .CapitalizeNamesOfDays = False
        .CorrectCapsLock = False
        .ReplaceText = False
        .DisplayAutoCorrectOptions = True
    End With
    
    Range(bfshn.Cells(gg1, dd1), bfshn.Cells(gg2, dd2)).Select '�I��͈͖͂߂� bfshn�킹��
    
    Call oshimai(syutoku(), bfn, shn, 1, 0, "���񏈗�����")
    '����݂̂����܂�
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub betat4(fbk As String, fsh As String, fmg1 As Long, fmr1 As Currency, fmg2 As Long, fmr2 As Currency, tbk As String, tsh As String, tog As Long, tor As Long, er34 As String, mr_8 As String)
    'e ver
    Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).ClearContents  'betatn��cpp�ƈقȂ�A�d�l�Ƃ��āA�N���A���邱�ƂƂ���B
    If er34 = "pp" Then '�W���^����
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).NumberFormatLocal = "G/�W��"
    ElseIf er34 = "mm" Then '������^����
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).NumberFormatLocal = "@"
    ElseIf er34 = "pm" Then '�ʉ݌^����
        Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor + Int(fmr2) - Int(fmr1))).NumberFormatLocal = "#,##0;[��]-#,##0"
    Else
        'mp
    End If
    
    If Abs(fmr1) = 0.4 Or Abs(fmr1) = 0.1 Then '���蕶���Ή��A��[��s�ԍ��Ή��@�i�t�B���^�j
        '������ɂ�(���A0.4����)
        If Abs(fmr1) = 0.1 Then '��[��s�ԍ��Ή��@���b��^�p
            
            Workbooks(tbk).Sheets(tsh).Cells(tog, tor).Value = Format(fmg1, "0000000")
                
            If fmg2 > fmg1 Then '�͈͂�1�s��2�s���������ꍇ�̑Ώ�(�ȉ�����)
                Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor).Value = Format(fmg1 + 1, "0000000")
            '�����A�����V�[�g��2��ڂ��R�s�[���邱�Ƃ�����������B
            End If
            If fmg2 > fmg1 + 1 Then  '�t�B���͂R�s�ȏ゠��ꍇ�̂ݎ��{
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor)).AutoFill Destination:=Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor))
            End If
        Else '���蕶���Ή�(�ʏ�) (0.4)
            Workbooks(tbk).Sheets(tsh).Cells(tog, tor).Value = Trim$(mr_8)
            If fmg2 > fmg1 Then
                Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor).Value = Trim$(mr_8)
            End If
            If fmg2 > fmg1 + 1 Then  '�t�B���͂R�s�ȏ゠��ꍇ�̂ݎ��{
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + 1, tor)).AutoFill Destination:=Range(Workbooks(tbk).Sheets(tsh).Cells(tog, tor), Workbooks(tbk).Sheets(tsh).Cells(tog + fmg2 - fmg1, tor))
            End If
        End If
    
    Else '��ʑΉ����A�ܕ�����x�^�^
        If er34 = "mp" Then  '�Z�����P
            Call cpp2(fbk, fsh, fmg1, Int(fmr1), fmg2, Int(fmr2), tbk, tsh, tog, tor, 0, 0, 12) '12:�l�Ɛ��l�̏��� '�x�E�R�s�y�p�^�[��
        Else 'mm,pp,pm
            Call cpp2(fbk, fsh, fmg1, Int(fmr1), fmg2, Int(fmr2), tbk, tsh, tog, tor, 0, 0, -4163) '-4163:�l
        End If
    End If
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub cpp2(fbk As String, fsh As String, fmg1 As Long, fmr1 As Long, fmg2 As Long, fmr2 As Long, tbk As String, tsh As String, tog1 As Long, tor1 As Long, tog2 As Long, tor2 As Long, mdo As Long)
    If tog2 = 0 And tor2 = 0 Then '�]���p�^�[��
        '�R�s�y���[�`���@30s85_004
        If mdo = -4163 Then  '�V�^���x�����B
            Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1 + fmg2 - fmg1, tor1 + fmr2 - fmr1)) = _
              Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Value
    
        Else '�]���^(�����ǂ�) �R�s�y�Ȃ̂Œx��(�����R�s�y�͂���A�������Ȃ�)�B
            UserForm3.StartUpPosition = 3 '1�@�G�N�Z���̒����A�@2�@��ʂ̒����A�@3�@��ʂ̍���
            UserForm3.Show vbModeless
            UserForm3.Repaint
            Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Copy
            Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1).PasteSpecial Paste:=mdo
            '(�Q�l)��.copy�̃R�s�[���\�b�h�́Aselect�`�b�N�ɔ͈͎͂w�肳��鋓���ł���B
            Unload UserForm3
            UserForm1.Repaint
        End If
    ElseIf fmr2 <= 0 Then  '�V�p�^�[���i��s�𕡐��s�ɃR�s�y�j-99��ߎ������̦�ASC(PHONETIC())�Ƃ��Ŏg����B
        If fmr2 = 0 Then
            If mdo = -4163 Then  '�V�^���x�����B
                Call oshimai("", bfn, shn, 1, 0, "�܂�������a")
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)) = _
                Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg2, fmr2)).Value
            Else '�]���^(�����ǂ�) �R�s�y�Ȃ̂Œx��(�����R�s�y�͂���A�������Ȃ�)�B
                UserForm3.StartUpPosition = 3 '1�@�G�N�Z���̒����A�@2�@��ʂ̒����A�@3�@��ʂ̍���
                UserForm3.Show vbModeless
                UserForm3.Repaint
                Range(Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1), Workbooks(fbk).Sheets(fsh).Cells(fmg1, fmr1)).Copy
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor2)).PasteSpecial Paste:=mdo
                Unload UserForm3
                UserForm1.Repaint
            End If
        ElseIf fmr2 = -1 Then
            If tog1 > tog2 Then Call oshimai("", bfn, shn, 1, 0, "tog1��tog2���ł���a")
            UserForm3.StartUpPosition = 3 '1�@�G�N�Z���̒����A�@2�@��ʂ̒����A�@3�@��ʂ̍���
            UserForm3.Show vbModeless
            UserForm3.Repaint
            Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1)) = fsh
            If tog2 > tog1 Then
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1)).Copy
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog2, tor1)).PasteSpecial Paste:=mdo 'tor2�͎g���ĂȂ�
            End If
            Unload UserForm3
            UserForm1.Repaint
        '4163����Ȃ�
        ElseIf fmr2 = -2 Then  '�t�B��(�A�Ԃ̂݁A�Œ�t�B���͂��Ȃ��Ɂj
          '4163����Ȃ�
            If tog1 > tog2 Then Call oshimai("", bfn, shn, 1, 0, "tog1��tog2���ł���b")
                Range(Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1, tor1)) = fmg2
                If tog2 > tog1 Then
                    Range(Workbooks(tbk).Sheets(tsh).Cells(tog1 + 1, tor1), Workbooks(tbk).Sheets(tsh).Cells(tog1 + 1, tor1)) = fmg2 + 1
                    If tog2 > tog1 + 1 Then  '�t�B��
                        Range(bfshn.Cells(tog1, tor1), bfshn.Cells(tog1 + 1, tor1)).AutoFill Destination:=Range(bfshn.Cells(tog1, tor1), bfshn.Cells(tog2, tor1))   'tor2�͎g���ĂȂ�
                    End If
                End If
            Else
                Call oshimai("", bfn, shn, 1, 0, "�ǂ��Ȃ邩����")
            End If
        Else
        Call oshimai("", bfn, shn, 1, 0, "�܂�������b")
    End If
    '��86_014c�ǉ��iexcel2019 �΍����test�j
    Application.CutCopyMode = False
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function kskup(am1 As String, am2 As String, n1 As Long, n2 As Long, h As Long, b As Currency, c As Currency, k0 As Long, h0 As Long, pap2 As Long, er2() As Currency, spd As String, pqp As Long, e5 As Long, er3() As Currency, hiru As Variant) As Long  '����
    'p(kskup)����V�d�l�An3�͋ߎ��p�@,(����)m2��n1�� ���t�@���Fn3�An2,h0,k0, pqp�ǉ�624,e5&er3()�ǉ�629
    Dim m As Long, n3 As Long
    Dim tskosk As Long 'strconv�p(�葬�F2(�]���ʂ�)�A�����F26(New85_006))
    Dim n3a As Long
    '��������(���E�ߎ�)��p�B�ᑬ����tskup��
    tskosk = 26 '(���E�Г��ꎋ�p���ցjexcel2019�΍�

    kskup = 0 '���Z�b�g(�s�v����)
    n3 = 0 '�[���X�^�[�g
    krpm2 = 0 '�������b�N��p=-2����t���O

    If h < k Then '���V�[�g�ɉ��������ꍇ(����̂ݒʉ߃]�[��) k��
        If c < 0 Then
            kskup = -1    '-1-2���̎��_��exitdo(p:-1�Ƃ���B)
            MsgBox "�\�󔒂�-1-2�ł�(p=-1,exitdo�A���̂܂܏I������܂�)�B"
        Else  '��c>=0���O��ƂȂ�B
            kskup = 2
            n2 = h + 1
            
            If er2(0) < 0 Then  '�������b�N�I��(�\��)�ł��B"
                UserForm4.StartUpPosition = 2 '1�@�G�N�Z���̒����A�@2�@��ʂ̒����A�@3�@��ʂ̍���
                UserForm4.Show vbModeless
                UserForm4.Repaint
                bfshn.Cells(sr(2), 5).Value = "����ۯ�(�\��)" '���̂���
                If pap2 = 0 And Abs(e5) < 1 And UBound(er3()) > 0 Then 'p=-2�̔���@�U�Q�X
                    For jj = 1 To UBound(er3())
                        If er3(jj) = 0.1 Then krpm2 = 1  'er3(jj)<0��=0.1 ��
                    Next
                    If krpm2 = 1 Then kskup = -2 'p=-2�̊m��@�U�Q�X
                End If
                Unload UserForm4
                UserForm1.Repaint
            End If
        End If
    ElseIf Round(kurai) = 2 Then '�x�[�V�b�N��
        'p=0
    ElseIf StrConv(am2, tskosk) = StrConv(am1, tskosk) Then '(�A�h�o���s�h)�O���v
        kskup = 1
        n2 = n1
    ElseIf spd = "������" And (k0 > h0 Or pqp = 1) Then  '(�A�h�o���s�h)�������b�N�I����ԁ@pqp�ǉ�624
        If c < 0 Then
            kskup = -1    '-1-2���̎��_��exitdo(p:-1�Ƃ���B)
            MsgBox "�����͂����ʂ�Ȃ��͂�(-1-2)�Ȍ㓖�V�[�g�����Ȃ��ł�(p=-1)�B"
        Else
            kskup = 2
            n2 = h + 1
        End If
    '(�A�h�o���s�h)���s��v c = Round(c, 0)�ǉ�85�Q026�i�F-15.�P���{���Ȃ��A-15���{����j
    ElseIf Round(c) <> -1 And Round(c) <> -2 And c = Round(c, 0) And (spd = "������" Or spd = "�m�[�}��") And n1 < h0 And _
            StrConv(am2, tskosk) = StrConv(bfshn.Cells(n1 + 1, Abs(b)).Value, tskosk) Then
        kskup = 1
        If b < 0 Then '��������
            n3 = n1 + 1 'n2�͍���ˍ������Ώۍs(��)
            n2 = hiru(n3, 2) '�ߎ��̓ǂݑւ� �W�T�Q027���؂�
            k0 = n3
        Else '�m�[�}��
            n2 = n1 + 1
        End If
    Else
        'p=0
    End If
    
    If kskup = 0 Then '�܂����܂炸(p=0)���}�b�`���O���{
        If k0 > h0 Then Call oshimai("", bfn, shn, 1, 0, "k0>h0��match�ɂ����悤�Ȃ��Ƃ͂����Ă͂Ȃ�Ȃ��B")

        If IsError(Application.Match(StrConv(am2, tskosk), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 1)) Then '�ߎ��ł��G���[�͔�������B
            m = 0
        Else
            m = Application.WorksheetFunction.Match(StrConv(am2, tskosk), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 1) '������
            If StrConv(am2, tskosk) <> StrConv(bfshn.Cells(k0 + m - 1, Abs(b)), tskosk) Then '��v���ĂȂ���Γ�i�K��
                m = 0
            Else 'ok�ߎ����{
                kskup = 1
                n3 = k0 + m - 1 'n2�͍���ˍ������Ώۍs(��)
                n2 = hiru(n3, 2) '�ߎ��̓ǂݑւ� �W�T�Q027���؂�
                If spd = "������" Then k0 = n3
            End If
        End If

        If spd = "������" And m = 0 Then '�V�K�ŏ������͓������s��
            kskup = 2
            n2 = h + 1
        End If
    End If
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function tskup(am1 As String, am2 As String, n1 As Long, n2 As Long, h As Long, b As Currency, c As Currency, k0 As Long, h0 As Long, pap2 As Long, er2() As Currency, spd As String, pqp As Long, e5 As Long, er3() As Currency) As Long  '����
    'p(tskup)����V�d�l�An3�͋ߎ��p�@,(����)m2��n1�� ���t�@���Fn3�An2,h0,k0, pqp�ǉ�624,e5&er3()�ǉ�629
    Dim m As Long, n3 As Long
    Dim tskosk As Long 'strconv�p(�葬�F2(�]���ʂ�)�A�����F26(New85_006))
    
    '���ᑬ����p
    tskosk = 2  '�啶������������
    tskup = 0 '���Z�b�g(�s�v����)
    n3 = 0 '�[���X�^�[�g
    krpm2 = 0 '�������b�N��p=-2����t���O
    If h < k Then  '���V�[�g�ɉ��������ꍇ(����̂ݒʉ߃]�[��) k��
        If c < 0 Then
            tskup = -1    '-1-2���̎��_��exitdo(p:-1�Ƃ���B)
            MsgBox "�\�󔒂�-1-2�ł�(p=-1,exitdo�A���̂܂܏I������܂�)�B"
        Else  '��c>=0���O��ƂȂ�B
            tskup = 2
            n2 = h + 1
            h0 = n2 '(=k0,���������b�N�I��) h0�����͊�{tskup����A������������(hk�t�]�C���M�����[�����[�u
            If er2(0) < 0 Then
                MsgBox "������͒ᑬ��p�ɂȂ�܂����B�����ł�����ʂ�̂͂��������B"
            End If
        End If
    ElseIf Round(kurai) = 2 Then '�x�[�V�b�N��
    
    ElseIf LCase(am2) = LCase(am1) Then '(�A�h�o���s�h)�O���v�@'��86_016q(uni�΍�)�@StrConv(am2��LCase(am2)
        tskup = 1
        n2 = n1
    ElseIf spd = "������" And k0 = h0 And b < 0 And pqp = 1 Then '(�A�h�o���s�h)�������b�N�I����ԁ@pqp�ǉ�624
        MsgBox "������͒ᑬ��p�ɂȂ�܂����B�������ł�����ʂ�̂͂��������B"
    
    '(�A�h�o���s�h)���s��v c = Round(c, 0)�ǉ�85�Q026�i�F-15.�P���{���Ȃ��A-15���{����j
    ElseIf Round(c) <> -1 And Round(c) <> -2 And c = Round(c, 0) And (spd = "������" Or spd = "�m�[�}��") And n1 < h0 And _
            LCase(am2) = LCase(bfshn.Cells(n1 + 1, Abs(b)).Value) Then    '��86_016q(uni�΍�)�@StrConv(am2��LCase(am2)
        tskup = 1
        If b < 0 Then '��������
            MsgBox "������͒ᑬ��p�ɂȂ�܂����B�����ł�����ʂ�̂͂��������B"
        Else '�m�[�}��
            n2 = n1 + 1
        End If
    Else
        'p=0
    End If
                
    If tskup = 0 Then '�܂����܂炸(p=0)���}�b�`���O���{
        If c < 0 And (spd = "�ߎ�����") Then  'spd = "���ߎ�����" Or spd = "���ߎ��m�[�}��" Or�@�͏��O 85_006
            MsgBox "������͒ᑬ��p�ɂȂ�܂����B�ߎ������ł�����ʂ�̂͂��������B"
        Else '�m�[�}��or���@ am2��strcnv���@85_008
            If IsError(Application.Match(LCase(am2), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 0)) Then  '�V�K2
                If c >= 0 Then '-1-2�͂��Ȃ��Ɂip=0�̂܂܏I���ց@'��86_016q(uni�΍�)�@StrConv(am2��LCase(am2)
                    tskup = 2
                    n2 = h + 1
                End If
            Else
                tskup = 1
            '                                           ��86_016q(uni�΍�)�@StrConv(am2��LCase(am2)
                m = Application.WorksheetFunction.Match(LCase(am2), Range(bfshn.Cells(k0, Abs(b)), bfshn.Cells(h0, Abs(b))), 0)  '���S��v
                If spd = "������" Then '������
                    MsgBox "������͒ᑬ��p�ɂȂ�܂����B�������ł�����ʂ�̂͂��������B"
                Else 'not������
                    n2 = k0 + m - 1
                End If
            End If
        End If
    End If
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function iptfg(jmj As String, czc As Long, ww As String) As String
    Do Until jj >= czc '��>�����Ă�͖̂������[�v�h�~
        fzf = tzt
        tzt = InStr(fzf + 1, jmj, ww)
        jj = jj + 1
    Loop
    If tzt = 0 Then tzt = Len(jmj) + 1 '4�I�N�Ή�
    iptfg = Mid(jmj, fzf + 1, tzt - fzf - 1)
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Function koudicd(fn As String, ff As String, er0 As Currency, mr04 As String) As String
    'er0�Ԃ�l�͍��ڍs�̍s���i�Ԃ�l0���L�F�����̏ꍇ�j
    '�@�ΏۃV�[�g�́u���ږ��v�T��
    Dim ywe As String
    For er0 = 1 To 2000  '�@
        If Workbooks(fn).Sheets(ff).Cells(er0, 1).Value = "���ږ�" Then
            koudicd = "���L"
            Exit For
        End If
        If Right(Workbooks(fn).Sheets(ff).Cells(er0, 1).Value, 4) = "��ŗL��" Then  '30s82���
            koudicd = "����"
            Exit For
        End If
    Next
    
    If koudicd = "" Then koudicd = "����"  '��
    
    If Round(kurai) = 1 Then        '�u���v�i�A�h�o���R�[�X�j
        If InStr(1, mr04, "�") > 0 Then '�����[��]
            'ywe�͕���
            ywe = iptfg(mr04, 3, "�") '��Ɂu���v�c�� iptfg�E�ENewVersion[��]
            '30s84_3 �����C���i�o�O�j��
            If ywe <> "" And InStr(1, iptfg(mr04, 2, "�"), ywe) > 0 Then '����������
                ywmoji = iptfg(iptfg(mr04, 2, "�"), 1, ywe) '[��]
                yw10 = iptfg(iptfg(mr04, 2, "�"), 2, ywe)  '[��]
            Else '���Ȃ�
                ywmoji = iptfg(mr04, 2, "�") 'yw10��null[��]
            End If
        Else  '��Ȃ�
            ywmoji = mr04  'yw10��null
        End If
        If Mid(ywmoji, 1, 1) = "�[" Then ywmoji = Mid(ywmoji, 2)
        
        '�A�C�ΏۃV�[�g�́u���ږ��v���w���L��(����������)or����
        If ywmoji = "" Or ywmoji = "0" Then  '����628
            If yw10 = "" Then
                er0 = 0
            ElseIf IsNumeric(yw10) Then
                er0 = Val(yw10) - 1
            Else
                Call oshimai("", bfn, shn, 1, 0, "����b�����{�ł��Ȃ��悤�ł��B")
            End If
            koudicd = "����b"  '��X����ɕς������B
            'MsgBox "����b�i����j"
        ElseIf yw10 <> "" Then   '��(����������)
            er0 = 1
            If Not IsNumeric(ywmoji) Then '�ꍇ����622
                '�]���p�^�[��
                Do Until Workbooks(fn).Sheets(ff).Cells(er0, 1).Value = yw10
                    If er0 = 20000 Then '2000��20000
                        Call oshimai("", bfn, shn, 1, 0, "���w�̍��ڍs��������Ȃ��悤�ł�")
                    End If
                    er0 = er0 + 1  '�C
                Loop
                koudicd = "����2"  '�A�@30s64�@���w������2(�����w)�ɕύX
            Else        'New�p�^�[���@����a ������yw10��all1�̑��
                Do Until Workbooks(fn).Sheets(ff).Cells(er0, Abs(Val(ywmoji))).Value = yw10
                    If er0 = 20000 Then '2000��20000
                        Call oshimai("", bfn, shn, 1, 0, "���w�̍��ڍs��������Ȃ��悤�ł�")
                    End If
                    er0 = er0 + 1  '�C
                Loop
                koudicd = "����a"  '�A�@30s64�@���w������2�ɕύX
            End If
        End If
        
        '�B�D�ΏۃV�[�g�́u���ږ��v�y�э��w���Ȃ����Aall1�����l�L�ڂłȂ��ꍇ(�����ږ�) �I���[�u628
        If koudicd = "����" And Not IsNumeric(ywmoji) Then
            er0 = 1
            Do Until Workbooks(fn).Sheets(ff).Cells(er0, 1).Value = ywmoji
                If er0 = 20000 Then '2000��20000
                    Call oshimai("", bfn, shn, 1, 0, "�����̍��ڍs��������Ȃ��悤�ł�")
                End If
                er0 = er0 + 1  '�D
            Loop
            koudicd = "����" 'yw10��""�@�B
            Call oshimai("", bfn, shn, 1, 0, "�����͏I�����܂����B����a�Ɉڍs���ĉ������B") '�I���[�u628
        End If
    End If '���i�A�h�o���R�[�X�j�����܂�
    
    '�ȍ~�x�[�V�b�N�A�h�o������
    If koudicd = "����" Then er0 = 0
End Function
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub iechc(hk1 As String)  '�ȑO��igchc
    If twbsh.Cells(3, 2).Value = "" Then  '86_017g ����N�ł��g����悤�ɁB
        twbsh.Cells(3, 2).Value = hunk2(-5, syutoku(), "1", hk1)
        If hk1 = "1" Then
            Call oshimai("", twn, "���W�v_���`", 3, 2, "(�����I��)ID�L�[�����܂���" & vbCrLf & "��ID:" & syutoku())
        End If
    ElseIf syutoku() = hunk2(-7, twbsh.Cells(3, 2).Value, "1", hk1) Then
        'MsgBox "�����ł�"
    Else  '"�s�����ł�"
        ThisWorkbook.Activate
        Sheets("���W�v_���`").Select
        twbsh.Cells(1, 1).Select  '�ΐF�Z��
        Call oshimai("", twn, "���W�v_���`", 3, 2, "(�����I��)ID�L�[�s��v" & vbCrLf & "��ID:" & syutoku())
    End If
    
    '(�ȉ��A���@igchc�d�l�𓥏P�@)
    kurai = 1.1  '���s�h�Œ�
    hk1 = Left(twn, 7) & "r"   '���s�h�Œ�
    twbsh.Cells(2, 3).Value = Left(twn, 7) & "r"
    twbsh.Cells(2, 2).Value = syutoku() & "r" '�V��86_016e
End Sub
'�[���[���[���[���[���[���[���[���[���[���[���[���[���[
Sub �G���[�Ή��p()  '��������̓}�N���{�^���Ƃ��āA�o�^����Ă���B
    Dim hk1 As String ', mghx As Long
    
    kyosydou  '���ʏ���

    Call iechc(hk1)
    
    '�Čv�Z����U������
    Application.Calculation = xlCalculationAutomatic
    Application.ExtendList = False '�f�[�^�͈͊g��:�I�t�i�F�E�׃Z��������ɏ����ς����̂�j�~�j"
    
    '�I�[�g�t�B���^���ݒ肳��Ă��邩�ǂ������f������
    If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    Application.EnableAutoComplete = False  '�I�[�g�R���v���[�g
       
    ThisWorkbook.Activate
    
    jj = 1
    Do Until Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = ""
        jj = jj + 1
        If jj = 10000 Then
            MsgBox "�󔒍s��������Ȃ��悤�ł�"
            Exit Sub
        End If
    Loop
    
    Workbooks(twn).Sheets(shog).Cells(jj, 1).Value = 1 '���ږ�
    Workbooks(twn).Sheets(shog).Cells(jj, 2).Value = jj  '����
    
    Workbooks(twn).Sheets(shog).Cells(jj, 3).Value = _
    "��.�@�G���A" & _
    twn _
    & "�A�" & shn & "�" & bfn _
    & "�A����W�v_���`" & "�" & twn _
    & "�Afrom" & dd1 & "to" & dd2 _
    & "�A���ږ��Ab" & twbsh.Cells(13, 3).Value & "R" & twbsh.Cells(14, 3).Value & "�A" & twbsh.Cells(2, 2).Value & "�A" _
    & Format(Now(), "yyyymmdd_hhmmss") & "�A" & bfshn.Cells(sr(8), 5).Value & "�A0�A0"

    Workbooks(twn).Sheets(shog).Cells(jj, 4).Value = Format(Now(), "yyyymmdd") 'date
    Workbooks(twn).Sheets(shog).Cells(jj, 5).Value = Format(Now(), "yyyymmdd_hhmmss") 'timestamp
    Workbooks(twn).Sheets(shog).Cells(jj, 7).Value = bfn & "\" & shn 'to
    Workbooks(twn).Sheets(shog).Cells(jj, 8).Value = 9  '�ŉE��(���ʂ͌Œ�l)
    Workbooks(twn).Sheets(shog).Cells(jj, 9).Value = twn & "\���W�v_���`"   'from
    '���O�������܂�
    
    Workbooks(bfn).Activate '����.copy��A����������ɓ����ƁA�Z�����������I������Ă��閭�ȉf��͉����������ۂ�
    DoEvents  '�����ʂ��邩���ؒ�
    Application.CutCopyMode = False '�����ʂ��邩���ؒ�
    MsgBox shn
    
    Worksheets(shn).Select
    bfshn.Cells(sr(0) - 1, 5).Select  '�ΐF�Z��

    '�Čv�Z���蓮��
    Application.Calculation = xlCalculationManual

    With Application.AutoCorrect      '�I�[�g�R���N�g�����Ȃ��@�R�O���T�Q
        .TwoInitialCapitals = False
        .CorrectSentenceCap = False
        .CapitalizeNamesOfDays = False
        .CorrectCapsLock = False
        .ReplaceText = False
        .DisplayAutoCorrectOptions = True
    End With
    
    Range(bfshn.Cells(gg1, dd1), bfshn.Cells(gg2, dd2)).Select '�I��͈͖͂߂� bfshn�킹��
    
    Call oshimai("", bfn, shn, 1, 0, "�G���[��������(" & syutoku() & ")")
    '����݂̂����܂�
End Sub
