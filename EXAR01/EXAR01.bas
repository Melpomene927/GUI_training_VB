Attribute VB_Name = "mod_EXAR01"
Option Explicit             'Do not allow ambiguous declaration
Option Compare Text         'Set all compare method as text-compare

'�b���B�w�q���Ҧ��ܼ�, �@�ߥHG�}�Y, �pG_AAA$, G_BBB#, G_CCC&
'�B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
' $: ��r
' #: �Ҧ��Ʀr�B��(���B�μƶq)
' &: �{���j���ܼ�
' %: ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
' �ť�: �N��VARIENT, �ʺA�ܼ�

Global G_FormFrom$    '�ťեN��������

'========================================================================
'   �w�q�U Form ���D��r�ܼ�
'========================================================================
Global G_Form_EXAR01$
Global G_Form_EXAR01q$

'========================================================================
'   �w�q�U�����D��r
'========================================================================
Global G_Pnl_A1601$
Global G_Pnl_A1602$
Global G_Pnl_A1605$
Global G_Pnl_A1606$
Global G_Pnl_A1609$
Global G_Pnl_A1614$
Global G_Pnl_A1617$
Global G_Pnl_A1620$
Global G_Pnl_A1621$
Global G_Pnl_A1643$
Global G_Pnl_Sum$
Global G_Pnl_Total$
Global G_Pnl_Credit$

Global G_Pnl_A0801$
Global G_Pnl_A0802$

Global G_Pnl_Dash$
Global G_Pnl_PrtType$
Global G_Pnl_Printer$
Global G_Pnl_Screen$
Global G_Pnl_File$
Global G_Pnl_Excel$

'========================================================================
'   Def �{���@���ܼ�
'========================================================================
Global G_PathNotFound$
Global G_Report_Heading$

Global G_A1620_Total#
Global G_A1621_Total#
Global G_A1643_Total#
Global G_Credit_Total#

Global G_A1601s$
Global G_A1601e$
Global G_A1617s$
Global G_A1617e$
Global G_A1609s$
Global G_A1609e$


'========================================================================
'??? �b���ŧi���{�����Ҧ���Spread�ۭq���A�ܼ�,�C�Ӵ���User�ۭq��쪺vaSpread,
'    �����ŧi�@��Spread�ۭq���A�ܼ�,�R�W�p�U:
'    vaSread Name : Spd_EXAR01   Spread Type Name: tSpd_EXAR01
'========================================================================
Global tSpd_Help As Spread
Global tSpd_EXAR01 As Spread

'========================================================================
'   Def ����榡
'========================================================================
'Global Const H0$ = "....5...10....5...20....5...30....5...40....5...50....5...60....5...70....5...80....5...90....5..100....5..110....5..120....5..130....5..140....5..150....5..160....5..170....5..180....5..190....5..."
'Global Const H1$ = " "
'Global Const H2$ = "  <SCR01>                                                     ***  �ϥΤ�x�C�L  ***"
'Global Const H3$ = "  �_�l���/�ɶ� : 89/02/15   / 10:01:01"
'Global Const H4$ = "  �I����/�ɶ� : 89/02/15   / 11:44:47"
'Global Const H5$ = "  �t�ΥN��:"
'Global Const H6$ = "  �{���N�X      :                                                                                                    �����G1"
'Global Const H7$ = "  �s�եN��      :                                                                                                    ����G89/02/15"
'Global Const H8$ = "  User ID       :            -                                                                                       �ɶ��G11:44:47"
'Global Const H9$ = "  ================================================================================================================================="
'Global Const FC$ = "  �t�ΥN��  ���       �ɶ�     �n��   �{���W��                                  �Ƶ�                                              "
'Global Const B1$ = "  �ϥΪ� : "
'Global Const B2$ = "  ---------------------------------------------------------------------------------------------------------------------------------"
'Global Const B3$ = "  ��ئX�p   : 2   Start : 1  Exit : 1"
'Global Const B3$ = "  ��ؤp�p : 2   Start : 1  Exit : 1"
'Global Const B3$ = "  �ϥΪ̦X�p : 2   Start : 1  Exit : 1"
'Global Const N1$ = "                                                                 ... �� �U �� ...                          �L��H :                "
'Global Const N2$ = "                                                                                                           �L��H :                "

'Global Const H0$ = "....5...10....5...20....5...30....5...40....5...50....5...60....5...70....5...80....5...90....5..100....5..110....5..120....5..130....5..140....5..150....5..160....5..170....5..180....5..190....5..."
'Global Const H1$ = " "
'Global Const H2$ = "  <SCR01> ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
'Global Const H3$ = "  ##############: ########## / ########"
'Global Const H4$ = "  ##############: ########## / ########"
'Global Const H5$ = "  ##############: ########## ########################################"
'Global Const H6$ = "  ##############: ########## ########################################                                        ######## : #####"
'Global Const H7$ = "  ##############: ### #########################################                                              ######## : ##########"
'Global Const H8$ = "  ##############: ########## - ##########                                                                    ######## : ##########"
'Global Const H9$ = "  ================================================================================================================================"
'Global Const FC$ = "  ########## ######## ##### ######################################## ########## ##################################################"
'Global Const FD$ = "  ########## ######## ##### ######################################## ########## ##################################################"
'Global Const N1$ = "                         ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^######## : ############"
'Global Const N2$ = "                                                                                                           ######## : ############"
'Global Const B1$ = "  ############ : "
'Global Const B2$ = "  --------------------------------------------------------------------------------------------------------------------------------"
'Global Const B3$ = "  ��ئX�p     : #######   Start : #######  Exit :#######"
'Global Const B3$ = "  ��ؤp�p   : #######   Start : #######  Exit :#######"
'Global Const B3$ = "  �ϥΪ̦X�p   : #######   Start : #######  Exit :#######"

'========================================================================
'??? �ŧi����榡�ܼ�
'========================================================================
Global Const H0$ = "....5...10....5...20....5...30....5...40....5...50....5...60....5...70....5...80....5...90....5..100....5..110....5..120....5..130....5..140....5..150....5..160....5..170....5..180....5..190....5..."
Global Const H1$ = " "
Global H2$
Global H3$
Global H4$
Global H5$
Global H6$
Global H7$
Global H8$
Global H3l$
Global H4l$
Global H5l$
Global H6l$
Global H7l$
Global H8l$
Global HDate$
Global HPerson$
Global H9$              'Line: =========================
Global B1$              'Break Header Line
Global B11$             'Break value Format of 1 column
Global B2$              'Line: -------------------------
Global B3$              'Line: #########################
Global B31$             'Break value Format of 3 columns
Global FC$              'Column Header Description
Global fd$              'Column Header Data
Global N1$
Global N2$

Sub GetPanelCaption()
'��FORM���D��r
    G_Form_EXAR01$ = GetCaption("FormTitle", "", "�Ȥ��ƦC�L")
    G_Form_EXAR01q$ = GetCaption("FormTitle", "", "�Ȥ��ƦC�L")
    
'�������D��r
    G_Pnl_A1609$ = GetCaption("PanelDescpt", "unifyno", "�Τ@�s��")
    G_Pnl_A1617$ = GetCaption("", "", "�t�d�~��")
    G_Pnl_A1601$ = GetCaption("PanelDescpt", "buyerid", "�Ȥ�s��")
    G_Pnl_A1602$ = GetCaption("PanelDescpt", "custmer", "�Ȥ�²��")
    G_Pnl_A1614$ = GetCaption("PanelDescpt", "liaison", "�p���H")
    G_Pnl_A1605$ = GetCaption("PanelDescpt", "telno", "�q�ܸ��X")
    G_Pnl_A1606$ = GetCaption("PanelDescpt", "faxno", "�ǯu���X")
    G_Pnl_A1620$ = GetCaption("PanelDescpt", "credit_limit", "�«H�B��")
    G_Pnl_A1621$ = GetCaption("PanelDescpt", "current_credit", "�����b��")
    G_Pnl_A1643$ = GetCaption("PanelDescpt", "nr", "��������")
    G_Pnl_Credit$ = GetCaption("PanelDescpt", "credit use", "�i���B��")
    
    G_Pnl_Sum$ = GetCaption("EXAR01", "sum", "�p�p")
    G_Pnl_Total$ = GetCaption("EXAR01", "total", "�X�p")
    
    G_Pnl_A0801$ = GetCaption("PanelDescpt", "", "�~�ȭ��s��")
    G_Pnl_A0802$ = GetCaption("PanelDescpt", "p_name_c", "�i���B��")
    
    G_Pnl_Dash$ = GetCaption("PanelDescpt", "dash", "��")
    G_Pnl_PrtType$ = GetCaption("PanelDescpt", "printtype", "�C�L�覡")
    G_Pnl_Printer$ = GetCaption("PanelDescpt", "printer", "�L���")
    G_Pnl_Screen$ = GetCaption("PanelDescpt", "screen", "�ù����")
    G_Pnl_File$ = GetCaption("PanelDescpt", "file", "�ɮ�")
    G_Pnl_Excel$ = GetCaption("PanelDescpt", "excel", "Excel")

'���C�L���N��r
'    G_SlipAttrib_1$ = Reference_SINI("SlipAttrib", "1")
'    G_SlipAttrib_2$ = Reference_SINI("SlipAttrib", "2")
'    G_AccountUse_1$ = Reference_SINI("AccountUse", "1")
'    G_AccountUse_2$ = Reference_SINI("AccountUse", "2")
'    G_AccountUse_3$ = Reference_SINI("AccountUse", "3")
'    G_SlipType_1$ = Reference_SINI("SlipType", "1")
'    G_SlipType_2$ = Reference_SINI("SlipType", "2")
'    G_SlipType_3$ = Reference_SINI("SlipType", "3")
'    G_SlipType_4$ = Reference_SINI("SlipType", "4")
'    G_SlipType_5$ = Reference_SINI("SlipType", "5")
'    G_SlipType_6$ = Reference_SINI("SlipType", "6")
'    G_SlipType_7$ = Reference_SINI("SlipType", "7")
'    G_SlipType_8$ = Reference_SINI("SlipType", "8")
    
'����L�ܼƤ��t��
    G_PathNotFound$ = GetCaption("PgmMsg", "path_not_found", "�ɮ׸��|���~!")
    G_Report_Heading$ = GetCaption("ReportHeading", "EXAR01", "��ئC�L")
End Sub

Sub Main()
' ���Ҳդ�, �������ӤU�C���ǰ���, �p�G���S���p���N�Y�ǼҲ�������,
' �Цb�ӼҲիe�W ' �Y�i, ���o�R��.

    Screen.MousePointer = HOURGLASS
    IsAppropriateCheck        ' �ˬd���{���O�_��MENU���I�s����
    DoubleRunCheck            ' �ˬd���{�����o���а���
    GetSystemINIString        ' ������t�ΨϥΤ���Ʈw�����w�ܼ�,
                              ' CHECK (C:\WINDOWS) LOCAL INI.
    OpenDB                    ' �}�_���t�ΩҦ��{���|�ϥΤ���Ʈw
    GetSystemDefault          ' ������t�Φ@�P���������ҰѼƳ]�w,
                              ' CHECK LXXX.MDB����INI TABLE.
    GetSvrDefault             ' ������t�ΨϥΤW, �S�w�������, �p���b��,
                              ' �����ɦW, ����榡, ...
    GetPanelCaption           ' ������{���w�]�w�@���ܼƤ����t��
    SetReportCols             ' �]�w�����Ҧ�����Spread Type��
    
'??? �N�Ҧ����ӵe����Load�iMemory,�Эק�Form Name
    Load frm_EXAR01           ' ���bQ�e�����]�w��Ĳ�o��,����V�e��Spread�W
                              ' ��Caption,�G��{������ɥ�Load V�e��
                              
'??? �Эק令�Ĥ@�ӵe����Form Name
    frm_EXAR01q.Show       ' �����e�����
    Screen.MousePointer = Default
End Sub

Sub PageCheck(Spd As vaSpread, Optional Break As Boolean = False)
'   under 2 circumstances do jump as below
'   1. reach maximum line
'   2. change to next break column
'   !!! "Excel" & "Screen Spread" does not need to jump to next page

    If G_PrintSelect = G_Print2Excel And Not Break Then Exit Sub
    If G_PrintSelect = G_Print2Screen Then Exit Sub
    '�����B�z
    If G_LineNo > G_OverFlow Or Break Then          '@Alter R2: Adding "Break" mechanism
        If G_PageNo > 0 Then
           If G_PrintSelect <> G_Print2Excel Then   '@Alter R2:
              PrintOut3 Spd, H1$, ""
              PrintOut3 Spd, H1$, ""
              PrintOut3 Spd, N1$, ""
           End If
           If G_PrintSelect = G_Print2Printer Then
              Printer.NewPage
           ElseIf G_PrintSelect = G_Print2Excel Then
              SetExcelNewPage
           Else
              Print #1, G_G1
           End If
        End If
        If G_PrintSelect <> G_Print2Excel Then ReportHeader Spd
    End If
End Sub

Sub PrePare_Data(Frm As Form, Prb As ProgressBar, Spd As vaSpread, A_Exit%)
On Local Error GoTo MY_Error
    
    '??? �]�wProgressBar�̤j��
    If G_ReportDataFrom = G_FromRecordSet Then
       Spd.MaxRows = 0
       DY_A16.MoveLast
       Prb.MAX = DY_A16.RecordCount
       DY_A16.MoveFirst
    Else
       Prb.MAX = Spd.MaxRows
    End If
    
    '�}�Ҥ�r��
    If G_PrintSelect = G_Print2File Then
        Open G_OutFile For Output As #1
    ElseIf G_PrintSelect = G_Print2Excel Then
        If Not OpenExcelFile(G_OutFile) Then
           Frm!Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
           Exit Sub
        End If
        '??? Excel�]�w��l��
        Set_Excel_Property Spd, tSpd_EXAR01
    End If

    '��l��tSpd��������
    InitialtSpdTextValue tSpd_EXAR01

    '�]�w�ʺA������榡
    SetPrintFormatStr
    
    '�]�w����r��,�r���ΦL����]�w
    If Not ReportSet() Then Frm!Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE): Exit Sub

    '??? �Y��Break����,�����s�վ������e
    AdjustColWidth Spd, tSpd_EXAR01, "A1617", B31$
    
    
    '��ƦC�L�B�z
    If G_ReportDataFrom = G_FromRecordSet Then
       Print2Spread Prb, Spd, A_Exit%
    Else
       PrintSub Prb, Spd, True, A_Exit%
    End If
    
    '��Esc��QĲ�o,�����C�L�ʧ@
    If A_Exit% Then Exit Sub
    
    Frm!Sts_MsgLine.Panels(1) = G_PrintOk
    Exit Sub

MY_Error:
    Beep
    Select Case Err
      Case 55
           Resume Next
      Case Else
           MsgBox Error$(Err)
    End Select
    Err = 0
End Sub

Sub Print2Spread(Prb As ProgressBar, Spd As vaSpread, A_Exit%)
Dim A_FmtStr$                                   'Format String
Dim A_A1601$, A_A1602$, A_A1614$                'Output Column Values
Dim A_A1605$, A_A1606$, A_A1620$
Dim A_A1621$, A_A1643$, A_credit$
Dim A_A1617$, A_A1617_Brk$                      'Break Column(A1607��ؤj��) & Previous Value of it
Dim A_Row#, A_Index#                            'Statical Counter
Dim A_Break_Value#                              '�«H�B�פp�p of A1620
Dim A_Break_Value2#                             '�����b�ڤp�p of A1621
Dim A_Break_Value3#                             '�������ڤp�p of A1643
Dim A_Break_Value4#                             '�i���B�פp�p of credit

    'Initialize
    Prb.Visible = True
    Prb.Value = 0
    A_Row# = 0
    A_Index# = 0
    Spd.MaxRows = 0
    
    'Initialize Summary Counter
    A_Break_Value# = 0
    A_Break_Value2# = 0
    A_Break_Value3# = 0
    A_Break_Value4# = 0
    G_A1620_Total# = 0
    G_A1621_Total# = 0
    G_A1643_Total# = 0
    G_Credit_Total# = 0
    
    
    '�C�L���Y
    ReportHeader Spd
    
    'Keep Break Value
    A_A1617_Brk$ = Trim$(DY_A16.Fields("A0802") & "")
    A_A1617$ = A_A1617_Brk$
    
    'Setup Output format
    A_FmtStr$ = "FD$"   'Format: [Break Header] + [-------] + [Data]
'    A_FmtStr$ = "B1$;B2$;FD$"   'Format: [Break Header] + [-------] + [Data]

    'Loop to Dump Report Values
    Do While Not DY_A16.EOF And Not A_Exit%
       
        '�֥[�ثe�B�z����Ƶ���
        A_Index# = A_Index# + 1
    
        'If change to another break
        If StrComp(A_A1617_Brk$, Trim$(DY_A16.Fields("A0802") & ""), _
            vbTextCompare) <> 0 Then
                 
            '�C�L�p�p
            PrintBreak Spd, A_Row#, G_Pnl_Sum$, A_Break_Value#, A_Break_Value2#, _
                A_Break_Value3#, A_Break_Value4#, "B2$;B3$;B2$"
          
            '�ܼ��k�s,�H�K���s�֭p
            A_Break_Value# = 0
            A_Break_Value2# = 0
            A_Break_Value3# = 0
            A_Break_Value4# = 0
          
            'Keep Break Value
            A_A1617_Brk$ = Trim$(DY_A16.Fields("A0802") & "")
            A_A1617$ = A_A1617_Brk$
            
            'Setup Output format
            A_FmtStr$ = "NP;FD$"    'Format: [NewPage] + [Data]
'            A_FmtStr$ = "NP;B1$;B2$;FD$"    'Format: [NewPage] + [Break Header] + [-------] + [Data]
        End If
              
        'Keep�C�L��Ʀ��ܼ�
        'col2 �Ȥ�s�� 10
        A_A1601$ = Trim$(DY_A16.Fields("A1601") & "")
        'col3 �Ȥ�²�� 12
        A_A1602$ = Trim$(DY_A16.Fields("A1602") & "")
        'col4 �p���H 20
        A_A1614$ = Trim$(DY_A16.Fields("A1614") & "")
        'col5 �q�ܸ��X 15
        A_A1605$ = Trim$(DY_A16.Fields("A1605") & "")
        'col6 �ǯu���X 15
        A_A1606$ = Trim$(DY_A16.Fields("A1606") & "")
        'col7 �«H�B�� 8
        A_A1620$ = Trim$(DY_A16.Fields("A1620") & "")
        'col8 �����b�� 8
        A_A1621$ = Trim$(DY_A16.Fields("A1621") & "")
        'col9 �������� 8
        A_A1643$ = Trim$(DY_A16.Fields("A1643") & "")
        'col10 �i���B�� 8
        A_credit$ = Str(Val(A_A1620$) - Val(A_A1621$) - Val(A_A1643$))
        
        'sum up col7~10 to break value
        G_A1620_Total# = G_A1620_Total# + CDbl(A_A1620$)
        G_A1621_Total# = G_A1621_Total# + CDbl(A_A1621$)
        G_A1643_Total# = G_A1643_Total# + CDbl(A_A1643$)
        G_Credit_Total# = G_Credit_Total# + CDbl(A_credit$)

        A_Break_Value# = A_Break_Value# + CDbl(A_A1620$)
        A_Break_Value2# = A_Break_Value2# + CDbl(A_A1621$)
        A_Break_Value3# = A_Break_Value3# + CDbl(A_A1643$)
        A_Break_Value4# = A_Break_Value4# + CDbl(A_credit$)
        
       
        '�NSpread�W��MaxRows�[�@
        AddSpreadMaxRows Spd, A_Row#
        
        '========================================================================
        ' [Mechanism Desciption]:
        '??? �H���W�ٳ]�w���Ȧ�vaSpread
        '    �ѼƤ@ : Spread Name           �ѼƤG : �ѼƤ@���ݪ�Spead Type Name
        '    �ѼƤT : �ۭq�����W��        �Ѽƥ| : ��ƦC
        '    �ѼƤ� : ��J��
        '========================================================================
        SetSpdText Spd, tSpd_EXAR01, "A1617", A_Row#, A_A1617$   '�t�d�~��
        SetSpdText Spd, tSpd_EXAR01, "A1601", A_Row#, A_A1601$   '�Ȥ�s��
        SetSpdText Spd, tSpd_EXAR01, "A1602", A_Row#, A_A1602$   '�Ȥ�²��
        SetSpdText Spd, tSpd_EXAR01, "A1614", A_Row#, A_A1614$   '�p���H
        SetSpdText Spd, tSpd_EXAR01, "A1605", A_Row#, A_A1605$   '�q�ܸ��X
        SetSpdText Spd, tSpd_EXAR01, "A1606", A_Row#, A_A1606$   '�ǯu���X
        SetSpdText Spd, tSpd_EXAR01, "A1620", A_Row#, A_A1620$   '�«H�B��
        SetSpdText Spd, tSpd_EXAR01, "A1621", A_Row#, A_A1621$   '�����b��
        SetSpdText Spd, tSpd_EXAR01, "A1643", A_Row#, A_A1643$   '��������
        SetSpdText Spd, tSpd_EXAR01, "credit", A_Row#, A_credit$ '�i���B��
        SetSpdText Spd, tSpd_EXAR01, "Flag", A_Row#, A_FmtStr$
'        SetSpdText Spd, tSpd_EXAR01, "TEST", A_Row#, "TEST"
        
        
       
        '�]�wSpread�Ĥ@�C���C��
        If G_PrintSelect = G_Print2Screen Then Spd.TopRow = SetSpreadTopRow(Spd)
       
        '========================================================================
        ' [Mechanism Desciption]:
        '   �Y��Q�e����� "�D�ù����" ���C�L�覡
        '   ����N���Prepare��V Screen��Spread�W.
        '   �YSpread��MaxRows�j�󵥩�100��,�h������PrintSub�NSpread�W����ƦL�X,
        '   �ñNMaxRows�k�s,�A�~��Prepare��Ʀ�V Screen.
        '========================================================================
        If (G_ReportDataFrom = G_FromRecordSet And G_PrintSelect <> _
            G_Print2Screen) And A_Row# >= 100 Then
            GoSub Print2SpreadA
        End If
       
        '�M��,Break�H��,���C�L����쪺���
        A_A1617$ = ""
       
        '�]�w��ƦC���M�ή榡
        A_FmtStr$ = "FD$"   'Format: [ReportData]
       
        '��ܥثe�B�z�i��
        Prb.Value = A_Index#
       
        DoEvents
       
        '��Esc��QĲ�o,�����C�L�ʧ@
        If A_Exit% Then Exit Do
       
        DY_A16.MoveNext
       
    Loop
    
    '�wĲ�o�������, ���X���{��
    If A_Exit% Then Exit Sub

    '�C�L���
    '�C�L��ئX�p��Break
    PrintBreak Spd, A_Row#, G_Pnl_Sum$, A_Break_Value#, A_Break_Value2#, _
        A_Break_Value3#, A_Break_Value4#, "B2$;B3$;H9$"
          
    '�C�L��ؤp�p��Break
    PrintBreak Spd, A_Row#, G_Pnl_Total$, G_A1620_Total#, G_A1621_Total#, G_A1643_Total#, G_Credit_Total#, "B3$;H9$"
          
    '�Y��Q�e����ܫD�ù���ܪ��C�L�覡,���ƳB�z����,���A�NSpread�W����ƦL�X.
    If (G_ReportDataFrom = G_FromRecordSet And G_PrintSelect <> G_Print2Screen) _
        And Spd.MaxRows > 0 Then
       GoSub Print2SpreadA
    End If
    
    '�B�z��ƦC�L�����᪺�����ʧ@
    PrintBottom Prb, Spd
    Exit Sub
    
Print2SpreadA:
    '�N��ƥ�SpreadŪ���C�L�ܤ�r��,�L�����Excel
    PrintSub Prb, Spd, False, A_Exit%
    ClearSpreadText Spd
    Spd.MaxRows = 0
    Return
End Sub

Sub PrintBottom(Prb As ProgressBar, Spd As vaSpread)
    
    '??? �C�L�L��H
    PrintOut3 Spd, H1$, "", -1
    PrintOut3 Spd, H1$, "", -1
    PrintOut3 Spd, N2$, "", -1


    '??? �N�_�l��줤�����,�HG_G1�r���N��Ƥ��Φ��h������
    SetExcelTextToColumns G_XlsStartCol%, 1, G_XlsHRows% + G_ExcelIndex#, _
        SetXlsFldDataType(tSpd_EXAR01)
    
    '�]�wExcel������榡
    SetExcelFormat

    '??? �B�z�U�ئC�L�覡�������ʧ@
    PrintEnd4 Spd, tSpd_EXAR01
    Prb.Visible = False
End Sub

Sub PrintBreak(Spd As vaSpread, A_Row#, ByVal A_Desc$, ByVal A_Break_Value#, _
    ByVal A_Break_Value2#, ByVal A_Break_Value3#, ByVal A_Break_Value4#, ByVal _
    A_FmtStr$)
Dim A_STR$, A_PrtStr$, A_Col&, A_Len&, A_Len2&

    'Keep�C�L��Ʀ��ܼ�
    A_STR$ = A_Desc$ & G_G1 & Format(A_Break_Value#, "#,##0.00")
    A_STR$ = A_STR$ & G_G1 & Format(A_Break_Value2#, "#,##0.00")
    A_STR$ = A_STR$ & G_G1 & Format(A_Break_Value3#, "#,##0.00")
    A_STR$ = A_STR$ & G_G1 & Format(A_Break_Value4#, "#,##0.00")
    A_STR$ = PrintUse(B31$, A_STR$)
    
    '�NSpread�W��MaxRows�[�@
    AddSpreadMaxRows Spd, A_Row#
    
    '�H���W�ٳ]�w���Ȧ�Spread
    SetSpdText Spd, tSpd_EXAR01, "A1617", A_Row#, A_STR$
    
    '�]�w�ӦC�C�L�ɩҮM�Ϊ��榡�r��
    SetSpdText Spd, tSpd_EXAR01, "Flag", A_Row#, A_FmtStr$
    
    '�]�wBreak��ƦC���C��
    SetSpreadColor Spd, A_Row#, -1, CStr(COLOR_YELLOW), G_TextGotFore_Color
    
    '�]�wSpread�Ĥ@�C���C��
    If G_PrintSelect = G_Print2Screen Then Spd.TopRow = SetSpreadTopRow(Spd)
End Sub

Sub PrintSub(Prb As ProgressBar, Spd As vaSpread, ByVal ShowProgress%, A_Exit%)
Dim A_PrtStr$, A_A1617$, A_FmtStr$()
Dim A_Row#, I#

    '��V Screen���檺�C�L�ʧ@,�~���B�z���@�~
    If ShowProgress% Then
        Prb.Visible = True
        Prb.Value = 0
        ReportHeader Spd
    End If
    
    '�]�w��l��
    A_Row# = 0
    
    '�B�z�C�C��ƪ��C�L
    Do While A_Row# < Spd.MaxRows And Not A_Exit%
        
        '�֥[�ثe�B�z����Ƶ���
        A_Row# = A_Row# + 1
        '========================================================================
        '??? �H���W�٨��o����,�m�JColumns Type��Text�ݩʤ�
        '    �ѼƤ@ : Spread Name           �ѼƤG : �ѼƤ@���ݪ�Spead Type Name
        '    �ѼƤT : �ۭq�����W��        �Ѽƥ| : ��ƦC
        '========================================================================
        A_A1617$ = GetSpdText(Spd, tSpd_EXAR01, "A1617", A_Row#)
        GetSpdText Spd, tSpd_EXAR01, "A1601", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1602", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1614", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1605", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1606", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1620", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1621", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1643", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "credit", A_Row#
        A_FmtStr$ = Split(GetSpdText(Spd, tSpd_EXAR01, "Flag", A_Row#), ";")
        
        For I# = 0 To UBound(A_FmtStr$)
       
            '�֭p�ثe�C�L���,�Y�W�L�@���h����
            G_LineNo = G_LineNo + 1
            PageCheck Spd
           
            '??? �N�r��ǵ�PrintOut3�B�z�C�L�ʧ@
            Select Case UCase$(A_FmtStr$(I#))
                Case "H1$"      'Single white space
                    PrintOut3 Spd, H1$, "", -1
                    
                Case "H9$"      'Line: =========================
                    PrintOut3 Spd, H9$, "", -1
                    
                Case "B2$"      'Line: -------------------------
                    PrintOut3 Spd, B2$, "", -1
                    
                Case "B1$"      'Break Header
                    G_ExcelIndex# = G_ExcelIndex# + 1
                    If G_PrintSelect = G_Print2Excel Then
                       A_PrtStr$ = PrintUse(B1$, G_Pnl_A1617$ & G_G1 & A_A1617$)
                    Else
                       A_PrtStr$ = G_Pnl_A1617$ & G_G1 & A_A1617$
                    End If
                    PrintOut3 Spd, B1$, A_PrtStr$, G_ExcelIndex#
                    '�Y�C�L��Excel��,�X��Break��쪺�x�s��
                    SetCellAlignment GetMergeCols(1, G_ExcelIndex# + _
                        G_XlsHRows%, G_ExcelMaxCols%, G_ExcelMaxCols%, 0), xlLeft, _
                        xlCenter, True
                        
                Case "B3$"      'Break Value
                    G_ExcelIndex# = G_ExcelIndex# + 1
                    PrintOut3 Spd, B3$, A_A1617$, G_ExcelIndex#
                    
                    '�]�wExcel Cells Range���I���C��
                    SetExcelRangeColor G_XlsHRows% + G_ExcelIndex#, G_XlsHRows% _
                        + G_ExcelIndex#, G_XlsStartCol%, G_ExcelMaxCols%, _
                        COLOR_YELLOW
                        
                    '�Y�C�L��Excel��,�X��Break��쪺�x�s��
                    SetCellAlignment GetMergeCols(1, G_ExcelIndex# + _
                        G_XlsHRows%, G_ExcelMaxCols%, G_ExcelMaxCols%, 0), xlRight, _
                        xlCenter, True
                        
                Case "FD$"      'Contents
                    G_ExcelIndex# = G_ExcelIndex# + 1
                    PrintOut3 Spd, fd$, PrintStrConnect(tSpd_EXAR01, 2), _
                        G_ExcelIndex#
                        
                Case "NP"       'New Page
                    PageCheck Spd, True
                    
            End Select
        Next I#
       
        '��Esc��QĲ�o,�����C�L�ʧ@
        If A_Exit% Then Exit Do
       
        '��V Screen���檺�C�L�ʧ@,����ܥثe�B�z�i��
        If ShowProgress% Then Prb.Value = A_Row#
        DoEvents
    Loop
    
    '�wĲ�o�������, ���X���{��
    If A_Exit% Then Exit Sub
    
    '��V Screen���檺�C�L�ʧ@������,�����檺�����ʧ@
    If ShowProgress% Then PrintBottom Prb, Spd
End Sub

Sub ReDefineHeaderAlign()
'�w��P�w�]�Ȥ��P�����,���s�]�w������Y��쪺����覡

    ChangeReportHeaderAlign tSpd_EXAR01, "A1617", SS_CELL_H_ALIGN_LEFT
'    ChangeReportHeaderAlign tSpd_EXAR01, "A0902", SS_CELL_H_ALIGN_CENTER
'    ChangeReportHeaderAlign tSpd_EXAR01, "A0906", SS_CELL_H_ALIGN_CENTER
'    ChangeReportHeaderAlign tSpd_EXAR01, "A0907", SS_CELL_H_ALIGN_RIGHT
'    ChangeReportHeaderAlign tSpd_EXAR01, "A0909", SS_CELL_H_ALIGN_RIGHT
'                                   :
'                                   :
End Sub

Sub ReDefineReportHeader()
'�w��P�w�]�Ȥ��P�����,���s�]�w������Y��쪺Caption

'    ChangeReportHeader tSpd_EXAR01, "A0901", "Test"
'    ChangeReportHeader tSpd_EXAR01, "A0902", "Test"
'    ChangeReportHeader tSpd_EXAR01, "A0906", "Test"
'    ChangeReportHeader tSpd_EXAR01, "A0907", "Test"
'    ChangeReportHeader tSpd_EXAR01, "A0909", "Test"
'                                   :
'                                   :
End Sub

Private Function Reference_SINI(ByVal A_Section$, ByVal A_Topic$) As String
On Local Error GoTo MyError
Dim A_Sql$

    Reference_SINI = ""
    A_Sql$ = "Select TOPICVALUE From SINI"
    A_Sql$ = A_Sql$ & " where SECTION='" & A_Section$ & "'"
    A_Sql$ = A_Sql$ & " and TOPIC='type" & A_Topic$ & "'"
    A_Sql$ = A_Sql$ & " order by SECTION,TOPIC"
    CreateDynasetODBC DB_ARTHGUI, DY_INI, A_Sql$, "DY_INI", True
    If Not (DY_INI.BOF And DY_INI.EOF) Then
       Reference_SINI = Trim$(DY_INI.Fields("TOPICVALUE") & "")
    End If
    Exit Function
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Sub ReportHeader(Spd As vaSpread)
Dim A_H2$, A_H3$, A_H4$, A_H5$, A_FC$
Dim A_FirstColName$, A_LastColName$

    '��ܿù��C�L��,���C�L���Y
    If G_PrintSelect = G_Print2Screen Then Exit Sub

    '??? ��l�ȭ��ثe�C��=���Y�`�C��
    G_LineNo = 8
    
    '�C�C�L�@�����Y,���Ʋ֥[�@
    G_PageNo = G_PageNo + 1

    '�걵���Y��Ʀ��ܼ�
    If G_PrintSelect = G_Print2Excel Then
        G_ExcelWkb.Visible = True
        '===========================================
        '???���oExcel����γ̫�@�檺�ۭq���W��
        '===========================================
        A_FirstColName$ = GetRptColName(tSpd_EXAR01, 1)
        A_LastColName$ = GetRptColName(tSpd_EXAR01, GetReportCols(tSpd_EXAR01))
       
        '===========================================
        '???�NExcel Header�����,�Ȧs��Spread Type
        '   �A�Q��PrintStrConnect�걵�C�L�C��Ʀ��ܼ�.
        '   ��PrintStrConnect Function�N�ѼƤG�]��2
        '   �H���oHeader��Ʀr��
        '===========================================
        SetTSpdText tSpd_EXAR01, A_FirstColName$, G_Report_Heading$
        A_H2$ = PrintStrConnect(tSpd_EXAR01, 2)
        SetTSpdText tSpd_EXAR01, A_FirstColName$, H3l$
        A_H3$ = PrintStrConnect(tSpd_EXAR01, 2)
        SetTSpdText tSpd_EXAR01, A_FirstColName$, H4l$
        SetTSpdText tSpd_EXAR01, A_LastColName$, PrintUse(HDate$, G_Print_Date & G_G1 & GetCurrentDay(1))
        A_H4$ = PrintStrConnect(tSpd_EXAR01, 2)
        SetTSpdText tSpd_EXAR01, A_FirstColName$, H5l$
        SetTSpdText tSpd_EXAR01, A_LastColName$, PrintUse(HDate$, G_Print_Time & G_G1 & Format(Now, "HH:MM:SS"))
        A_H5$ = PrintStrConnect(tSpd_EXAR01, 2)
        
'        SetTSpdText tSpd_EXAR01, A_FirstColName$, H6l$
'        A_H6$ = PrintStrConnect(tSpd_EXAR01, 2)
'        SetTSpdText tSpd_EXAR01, A_FirstColName$, H7l$
'        SetTSpdText tSpd_EXAR01, A_LastColName$, PrintUse(HDate$, G_Print_Date & G_G1 & GetCurrentDay(1))
'        A_H7$ = PrintStrConnect(tSpd_EXAR01, 2)
'        SetTSpdText tSpd_EXAR01, A_FirstColName$, H8l$
'        SetTSpdText tSpd_EXAR01, A_LastColName$, PrintUse(HDate$, G_Print_Time & G_G1 & Format(Now, "HH:MM:SS"))
'        A_H8$ = PrintStrConnect(tSpd_EXAR01, 2)
    Else
        '??? �]�w�C�L�ܦL����Τ�r��Header���r���ܼ�,��ƶ��HG_G1���Ϲj
        A_H2$ = G_Report_Heading$
        A_H3$ = G_Print_Page & G_G1 & Format(G_PageNo, "###0")
        A_H4$ = G_Print_Date & G_G1 & GetCurrentDay(1)
        A_H5$ = G_Print_Time & G_G1 & Format(Now, "HH:MM:SS")
    End If
    
    '??? �C�L������Y
    PrintOut3 Spd, H1$, "", 1         '�̫�@�ӰѼ�: �C�L�����
    PrintOut3 Spd, H2$, A_H2$, 2
    PrintOut3 Spd, H3$, A_H3$, 3
    PrintOut3 Spd, H4$, A_H4$, 4
    PrintOut3 Spd, H5$, A_H5$, 5
'    PrintOut3 Spd, H6$, A_H6$, 6
'    PrintOut3 Spd, H7$, A_H7$, 7
'    PrintOut3 Spd, H8$, A_H8$, 8
    PrintOut3 Spd, H9$, "", -1
    PrintOut3 Spd, FC$, FC$, 6
    PrintOut3 Spd, H9$, "", -1

    '??? �]�w�C�L��Excel��,���Y�ҥΪ��C��
    If G_PrintSelect = G_Print2Excel Then G_XlsHRows% = 6
End Sub

Function ReportSet() As Boolean
    ReportSet = True
    
    'Initial����i�ƤΨC������ܼƭ�
    G_PageNo = 0
    G_LineNo = 0
    
    'Initial����O�_����Ƥw�C�L���ܼƭ�
    G_HaveDataPrint% = False
    
    '���Excel or Screen�C�L,�����]�w�L��]�w
    If G_PrintSelect = G_Print2Screen Then Exit Function
    If G_PrintSelect = G_Print2Excel Then Exit Function
    
    '�]�w�����ݩʪ�l��
    G_FontName = GetIniStr("Font", "Name", "GUI.INI")
    G_FontSize = GetGLRptFont("Font3")
    G_PageSize = GetGLRptPageLine("Font3")
    G_OverFlow = GetGLRptPageSize("Font3")
    
    '�Y��ܦL����C�L,�h��ܦL��]�w����
    If G_PrintSelect = G_Print2Printer Then
        Printer.FontName = G_FontName
        Printer.FontSize = G_FontSize
        rptset.Show MODAL
        G_OverFlow = G_PageSize - 6
        If G_PageSize <= 0 Then G_PageSize = 60
        If G_OverFlow <= 0 Then G_OverFlow = 60
        ReportSet = G_RptSet
    End If
End Function

Sub Set_Excel_Property(Spd As vaSpread, tSPD As Spread)
'�]�w�_�l��쬰1,�ñNHeaders���`�C���k�s
    G_XlsStartCol% = 1: G_XlsHRows% = 0

'�N�ثe�C���k�s
    G_ExcelIndex# = 0
    
'���oExcel�����`���
    G_ExcelMaxCols% = GetReportCols(tSPD)
    
 '�]�wExcel�C����쪺��ƫ��A�ι����m
    SetExcelDataType Spd, tSPD
End Sub

Sub SetExcelFormat()
Dim A_MaxCol$, A_Row#

    If G_PrintSelect <> G_Print2Excel Then Exit Sub

    '========================================================================
    ' Excel style Setting
    '========================================================================
    '??? �ثeEXCEL�L��ĴX�C
    A_Row# = G_ExcelIndex# + G_XlsHRows%
    
    '??? �ӳ̤j���Ƥ��W��
    A_MaxCol$ = Chr(Asc("A") + G_ExcelMaxCols% - 1)

    '??? �b����|�P�[�W�u��
    xlsDrawLine G_ExcelWkb, "A" & Trim(Str(G_XlsHRows%)) & ":" & A_MaxCol$ & _
        Trim(Str(A_Row#))

    '??? ���D����m��(�Y���ݭn,�Эק�Rang�����d��Y�i�ϥ�) FALSE-���X���x�s��
    SetCellAlignment "A2:" & A_MaxCol$ & "2", xlCenter, xlCenter, True

    '========================================================================
    ' Header Left Part Setting
    '========================================================================
    '??? N/A(�x�s��X��)
    SetCellAlignment GetMergeCols(1, 3, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
        1), xlLeft, xlCenter, True

    '??? ���q�O(�x�s��X��)
    SetCellAlignment GetMergeCols(1, 4, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
        1), xlLeft, xlCenter, True

    '??? ��ؽd��(�x�s��X��)
    SetCellAlignment GetMergeCols(1, 5, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
        1), xlLeft, xlCenter, True

'    '??? �{���N�X(�x�s��X��)
'    SetCellAlignment GetMergeCols(1, 6, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
'        1), xlLeft, xlCenter, True
'
'    '??? �s�եN��(�x�s��X��)
'    SetCellAlignment GetMergeCols(1, 7, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
'        1), xlLeft, xlCenter, True
'
'    '??? User ID(�x�s��X��)
'    SetCellAlignment GetMergeCols(1, 8, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
'        1), xlLeft, xlCenter, True
'
    
    '========================================================================
    ' Header Right Part Setting
    '========================================================================
    '??? �C�L����m�k
    SetCellAlignment GetExcelColName(G_ExcelMaxCols% + G_XlsStartCol% - 1) & _
        "4", xlRight, xlCenter, True

    '??? �C�L�ɶ��m�k
    SetCellAlignment GetExcelColName(G_ExcelMaxCols% + G_XlsStartCol% - 1) & _
        "5", xlRight, xlCenter, True
    
    '??? ���D�C�m��
    SetCellAlignment Trim(Str(G_XlsHRows%)) + ":" + Trim(Str(G_XlsHRows%)), _
        xlCenter, xlCenter, False
    
    
    
    '========================================================================
    ' Other Setting
    '========================================================================
    '??? �]�wExcel��ΦC���w�]�j�p,�ýվ������e�ܳ̾A�e��
    SetExcelSize "A:" & A_MaxCol$
    
    '�����]�w (Orientation%�Ѽƭ� - xlPortrait:���L  xlLandscape:��L)
    SetExcelAllocate "$1:$" & Trim(Str(G_XlsHRows%))
    
    '�]�w�@���x�s��
    SelectExcelCells "A1"
End Sub

Sub SetPrintFormatStr()

    '??? �������Y�榡�i���ܼƪ�l�Ȫ��ʧ@
    H3l$ = "########## : ########## - ##########"
    H4l$ = "########## : ############ - ############"
    H5l$ = "########## : ############### - ###############"
    HDate$ = "######## : ##########"
    HPerson$ = "######## : ############"
    B31$ = "######## : ~~~~~~~~~~~~~~~ ~~~~~~~~~~~~~~~ ~~~~~~~~~~~~~~~ ~~~~~~~~~~~~~~~  "
    B11$ = "######## : ###############"

    '�ù���ܤ����]�w����榡
    If G_PrintSelect = G_Print2Screen Then Exit Sub

    '??? �]�w�����k���Ŷ����涡�j,�Y�ϥιw�]�ȥi����J
    SetRptAllocate
    
    '??? ���o�����̤p�e��
    GetRptMinWidth H3l$ & Space(1) & HDate$
    
' �@�C�H�WHeader���榡�]�w =====================================================================
'??? ���o���D�θ�ƪ��r��榡(�ѼƤG��Ǧ^���榡���A -- 1:���D�榡 2:�����ܪ��榡)
'??? Multi Line �ɨϥ�
'    ' �w��P�w�]�Ȥ��P�����,���s�]�w������Y��쪺����覡
'    ReDefineHeaderAlign
'    ' �]�w�Ĥ@�CHeader��Caption
'    ReDefineReportHeader
'    '���o�Ĥ@�CHeader��Format
    'FC$ = GetRptFormatStr(tSpd_EXAR01, 3)
'    ' �]�w�ĤG�CHeader��Caption
'    ReDefineReportHeader
'    '���o�ĤG�CHeader��Format
'    FC$ = GetRptFormatStr(tSpd_EXAR01, 3)
'    fd$ = GetRptFormatStr(tSpd_EXAR01, 2)
' ==============================================================================================
   
    '??? ���o���D�θ�ƪ��r��榡(�ѼƤG��Ǧ^���榡���A -- 1:���D�榡 2:�����ܪ��榡)
    '   �w��P�w�]�Ȥ��P�����,���s�]�w������Y��쪺����覡
    ReDefineHeaderAlign

    '??? ���Y��Single Line �ɨϥ�
    FC$ = GetRptFormatStr(tSpd_EXAR01, 1)
    fd$ = GetRptFormatStr(tSpd_EXAR01, 2)

    '??? ���o������Y���榡
    H2$ = GetRptTitleFormat()
    
    '??? ���o������Y��ƪ��榡
    H3l$ = PrintUse(H3l$, G_Pnl_A1601$ & G_G1 & G_A1601s$ & G_G1 & G_A1601e$)
    H4l$ = PrintUse(H4l$, G_Pnl_A1617$ & G_G1 & G_A1617s$ & G_G1 & G_A1617e$)
    H5l$ = PrintUse(H5l$, G_Pnl_A1609$ & G_G1 & G_A1609s$ & G_G1 & G_A1609e$)
    H3$ = GetRptHeaderFormat(H3l$, HDate$)
    H4$ = GetRptHeaderFormat(H4l$, HDate$)
    H5$ = GetRptHeaderFormat(H5l$, HDate$)
    B31$ = GetRptHeaderFormat("", B31$)

    '??? ���o����Break��쪺�榡
    B1$ = GetRptHeaderFormat(B11$)
   
    '??? ���o��U���ΦL��H���榡
    N1$ = GetRptFootFormat(HPerson$)
    N2$ = PrintUse(GetRptLineFormat("~"), HPerson$)
    
    '??? ���o�Ϲj�C���榡
    B2$ = GetRptLineFormat("-")
    B3$ = GetRptLineFormat("~")
    H9$ = GetRptLineFormat("=")
End Sub

Sub SetReportCols()
    '========================================================================
    '*** �]�wQ Screen����Spd_Help vaSpread **********************************
    '??? �ŧiSpread���A��Columns��Sorts���}�C�Ӽ�,
    '    �ѼƤ@ : Spread Type Name
    '    �ѼƤG : vaSpread�W������`��
    '    �ѼƤT : �O�_���\User�ۭq�Ƨ����Ψ䶶��
    '========================================================================
    InitialCols tSpd_Help, 2, False
    
    '========================================================================
    '??? �]�wvaSpread�W���Ҧ����αƧ�����Spread Type��
    '    �ѼƤ@ : Spread Type Name
    '    �ѼƤG : �]�w�ΨӦs��vaSpread�W��쪺���W��
    '    �ѼƤT : Optional - �]�w�������(0:���  1:�Ȯ�����,�w�]��  2:�ä[����)
    '    �Ѽƥ| : Optional - �]�w�{���w�]�Ƨ���쪺����
    '    �ѼƤ� : Optional - �]�w�{���w�]�Ƨ���쪺��V(1:���W,�w�]��  2:����)
    '    �ѼƤ� : Optional - �]�wBreak��쪺����
    '    �ѼƤC : Optional - �]�wBreak���O�_�P��L�����ܩ�P�@�C�W(True,�w�]�� / False)
    '========================================================================
    AddReportCol tSpd_Help, "A0801", , 1
    AddReportCol tSpd_Help, "A0802", , 2
    
    '========================================================================
    '??? ���User�ۭq���������ܶ��ǤαƧ����
    '    �ѼƤ@ : Spread Type Name
    '    �ѼƤG : vaSpread�Ҧb��Form Name
    '    �ѼƤT : vaSpread Name
    '========================================================================
    GetSpreadDefault tSpd_Help, "frm_EXAR01q", "Spd_Help"

    '========================================================================
    '*** �]�wV Screen����Spd_EXAR01 vaSpread *********************************
    '??? �ŧiSpread���A��Columns��Sorts���}�C�Ӽ�,
    '    �ѼƤ@ : Spread Type Name
    '    �ѼƤG : vaSpread�W������`��
    '    �ѼƤT : �O�_���\User�ۭq�Ƨ����Ψ䶶��
    '========================================================================
    InitialCols tSpd_EXAR01, 11, False
    
    '========================================================================
    '??? �]�wvaSpread�W���Ҧ����αƧ�����Spread Type��
    '    �ѼƤ@ : Spread Type Name
    '    �ѼƤG : �]�w�ΨӦs��vaSpread�W��쪺���W��
    '    �ѼƤT : Optional - �]�w�������(0:���  1:�Ȯ�����,�w�]��  2:�ä[����)
    '    �Ѽƥ| : Optional - �]�w�{���w�]�Ƨ���쪺����
    '    �ѼƤ� : Optional - �]�w�{���w�]�Ƨ���쪺��V(1:���W,�w�]��  2:����)
    '    �ѼƤ� : Optional - �]�wBreak��쪺����
    '    �ѼƤC : Optional - �]�wBreak���O�_�P��L�����ܩ�P�@�C�W(True,�w�]�� / False)
    '========================================================================
    AddReportCol tSpd_EXAR01, "A1617", , 1, , 1
    AddReportCol tSpd_EXAR01, "A1601", , 2
    AddReportCol tSpd_EXAR01, "A1602"
    AddReportCol tSpd_EXAR01, "A1614"
    AddReportCol tSpd_EXAR01, "A1605"
    AddReportCol tSpd_EXAR01, "A1606"
    AddReportCol tSpd_EXAR01, "A1620"
    AddReportCol tSpd_EXAR01, "A1621"
    AddReportCol tSpd_EXAR01, "A1643"
    AddReportCol tSpd_EXAR01, "credit"
    AddReportCol tSpd_EXAR01, "Flag", 2
    
    '========================================================================
    '??? ���User�ۭq���������ܶ��ǤαƧ����
    '    �ѼƤ@ : Spread Type Name
    '    �ѼƤG : vaSpread�Ҧb��Form Name
    '    �ѼƤT : vaSpread Name
    '========================================================================
    GetSpreadDefault tSpd_EXAR01, "frm_EXAR01", "Spd_EXAR01"
End Sub

