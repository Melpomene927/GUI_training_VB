Attribute VB_Name = "mod_TSR03"
Option Explicit
Option Compare Text

'�b���B�w�q���Ҧ��ܼ�, �@�ߥHG�}�Y, �pG_AAA$, G_BBB#, G_CCC&
'�B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
' $: ��r
' #: �Ҧ��Ʀr�B��(���B�μƶq)
' &: �{���j���ܼ�
' %: ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
' �ť�: �N��VARIENT, �ʺA�ܼ�


'���n�ܼ�
Global G_FormFrom$    '�ťեN��������

'Def �{���@���ܼ�
Global G_A0901s$
Global G_A0901e$
Global G_A0902s$
Global G_A0902e$
Global G_A0904s$
Global G_A0904e$
Global G_A0905$
Global G_A0905o$
Global G_A0906$
Global G_A0906o$
Global G_A0911$
Global G_A0911o$
''SAMPLE
'Global G_BB#
'Global G_CC!

'??? �b���ŧi���{�����Ҧ���Spread�ۭq���A�ܼ�,�C�Ӵ���User�ۭq��쪺vaSpread,
'    �����ŧi�@��Spread�ۭq���A�ܼ�,�R�W�p�U:
'    vaSread Name : Spd_PatternR2   Spread Type Name: tSpd_PatternR2
Global tSpd_Help As Spread
Global tSpd_PATTERNR2 As Spread


'Def ����榡
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
'Global Const B3$ = "  �t�Τp�p   : 2   Start : 1  Exit : 1"
'Global Const B3$ = "  �ϥΪ̤p�p : 2   Start : 1  Exit : 1"
'Global Const B3$ = "  �ϥΪ̦X�p : 2   Start : 1  Exit : 1"
'Global Const N1$ = "                                                                 ... �� �U �� ...                          �L��H :                "
'Global Const N2$ = "                                                                                                           �L��H :                "

Global Const H0$ = "....5...10....5...20....5...30....5...40....5...50....5...60....5...70....5...80....5...90....5..100....5..110....5..120....5..130....5..140....5..150....5..160....5..170....5..180....5..190....5..."
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
'Global Const B3$ = "  �t�Τp�p     : #######   Start : #######  Exit :#######"
'Global Const B3$ = "  �ϥΪ̤p�p   : #######   Start : #######  Exit :#######"
'Global Const B3$ = "  �ϥΪ̦X�p   : #######   Start : #######  Exit :#######"

'??? �ŧi����榡�ܼ�
Global Const H1$ = " "
Global H2$
Global H3$
Global H3l$
Global H4$
Global H4l$
Global H5$
Global H5l$
Global H6$
Global H6l$
Global H7$
Global H7l$
Global H8$
Global H8l$
Global HDate$
Global HPerson$
Global H9$
Global B1$
Global B11$
Global B2$
Global B3$
Global B31$
Global FC$
Global fd$
Global N1$
Global N2$

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
    SetReportCols             ' �]�w�����Ҧ�����Spread Type��
    
'??? �N�Ҧ����ӵe����Load�iMemory,�Эק�Form Name
    Load frm_PATTERNR2        ' ���bQ�e�����]�w��Ĳ�o��,����V�e��Spread�W
                              ' ��Caption,�G��{������ɥ�Load V�e��
                              
'??? �Эק令�Ĥ@�ӵe����Form Name
    frm_PATTERNR2q.Show       ' �����e�����
    Screen.MousePointer = Default
End Sub


Sub PageCheck(Spd As vaSpread, Optional Break As Boolean = False)
    If G_PrintSelect = G_Print2Excel And Not Break Then Exit Sub
    If G_PrintSelect = G_Print2Screen Then Exit Sub
'�����B�z
    If G_LineNo > G_OverFlow Or Break Then
        If G_PageNo > 0 Then
           If G_PrintSelect <> G_Print2Excel Then
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
       DY_A09.MoveLast
       Prb.MAX = DY_A09.RecordCount
       DY_A09.MoveFirst
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
        Set_Excel_Property Spd, tSpd_PATTERNR2
    End If

'��l��tSpd��������
    InitialtSpdTextValue tSpd_PATTERNR2

'�]�w�ʺA������榡
    SetPrintFormatStr
    
'�]�w����r��,�r���ΦL����]�w
    If Not ReportSet() Then Frm!Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE): Exit Sub

'??? �Y��Break����,�����s�վ������e
    AdjustColWidth Spd, tSpd_PATTERNR2, "A0909", B31$
    
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

Sub PrintBreak(Spd As vaSpread, A_Row#, ByVal A_Desc$, ByVal A_Start_Break1#, ByVal A_Exit_Break2#, ByVal A_FmtStr$)
'�B�zBreak���C�L
Dim A_STR$, A_PrtStr$, A_Col&, A_Len&, A_Len2&

'Keep�C�L��Ʀ��ܼ�
    A_STR$ = A_Desc$ & G_G1 & Format(A_Start_Break1# + A_Exit_Break2#, "#,##0")
    A_STR$ = A_STR$ & G_G1 & "Start"
    A_STR$ = A_STR$ & G_G1 & Format(A_Start_Break1#, "#,##0")
    A_STR$ = A_STR$ & G_G1 & "Exit"
    A_STR$ = A_STR$ & G_G1 & Format(A_Exit_Break2#, "#,##0")
    A_STR$ = PrintUse(B31$, A_STR$)
    
'�NSpread�W��MaxRows�[�@
    AddSpreadMaxRows Spd, A_Row#
    
'�H���W�ٳ]�w���Ȧ�Spread
    SetSpdText Spd, tSpd_PATTERNR2, "A0909", A_Row#, A_STR$
    
'�]�w�ӦC�C�L�ɩҮM�Ϊ��榡�r��
    SetSpdText Spd, tSpd_PATTERNR2, "Flag", A_Row#, A_FmtStr$
    
'�]�wBreak��ƦC���C��
    SetSpreadColor Spd, A_Row#, -1, CStr(COLOR_YELLOW), G_TextGotFore_Color
    
'�]�wSpread�Ĥ@�C���C��
    If G_PrintSelect = G_Print2Screen Then Spd.TopRow = SetSpreadTopRow(Spd)
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

Sub ReportHeader(Spd As vaSpread)
'�C�L������Y
Dim A_H2$, A_H3$, A_H4$, A_H5$, A_H6$, A_H7$, A_H8$, A_FC$
Dim A_FirstColName$, A_LastColName$

'��ܿù��C�L��,���C�L���Y
    If G_PrintSelect = G_Print2Screen Then Exit Sub

'??? ��l�ȭ��ثe�C��=���Y�`�C��
    G_LineNo = 11
    
'�C�C�L�@�����Y,���Ʋ֥[�@
    G_PageNo = G_PageNo + 1

'�걵���Y��Ʀ��ܼ�
    If G_PrintSelect = G_Print2Excel Then
       
       G_ExcelWkb.Visible = True
       
'??? ���oExcel����γ̫�@�檺�ۭq���W��
       A_FirstColName$ = GetRptColName(tSpd_PATTERNR2, 1)
       A_LastColName$ = GetRptColName(tSpd_PATTERNR2, GetReportCols(tSpd_PATTERNR2))
       
'??? �NExcel Header�����,�Ȧs��Spread Type,�A�Q��PrintStrConnect�걵�C�L�C��Ʀ��ܼ�.
'    ��PrintStrConnect Function�N�ѼƤG�]��2,�H���oHeader��Ʀr��
       SetTSpdText tSpd_PATTERNR2, A_FirstColName$, GetCaption("ReportHeading", "PATTERNR", "�ϥΤ�x�C�L")
       A_H2$ = PrintStrConnect(tSpd_PATTERNR2, 2)
       SetTSpdText tSpd_PATTERNR2, A_FirstColName$, H3l$
       A_H3$ = PrintStrConnect(tSpd_PATTERNR2, 2)
       SetTSpdText tSpd_PATTERNR2, A_FirstColName$, H4l$
       A_H4$ = PrintStrConnect(tSpd_PATTERNR2, 2)
       SetTSpdText tSpd_PATTERNR2, A_FirstColName$, H5l$
       A_H5$ = PrintStrConnect(tSpd_PATTERNR2, 2)
       SetTSpdText tSpd_PATTERNR2, A_FirstColName$, H6l$
       A_H6$ = PrintStrConnect(tSpd_PATTERNR2, 2)
       SetTSpdText tSpd_PATTERNR2, A_FirstColName$, H7l$
       SetTSpdText tSpd_PATTERNR2, A_LastColName$, PrintUse(HDate$, G_Print_Date & G_G1 & GetCurrentDay(1))
       A_H7$ = PrintStrConnect(tSpd_PATTERNR2, 2)
       SetTSpdText tSpd_PATTERNR2, A_FirstColName$, H8l$
       SetTSpdText tSpd_PATTERNR2, A_LastColName$, PrintUse(HDate$, G_Print_Time & G_G1 & Format(Now, "HH:MM:SS"))
       A_H8$ = PrintStrConnect(tSpd_PATTERNR2, 2)
       
    Else
       
'??? �]�w�C�L�ܦL����Τ�r��Header���r���ܼ�,��ƶ��HG_G1���Ϲj
       A_H2$ = GetCaption("ReportHeading", "PATTERNR", "�ϥΤ�x�C�L")
       A_H6$ = G_Print_Page & G_G1 & Format(G_PageNo, "###0")
       A_H7$ = G_Print_Date & G_G1 & GetCurrentDay(1)
       A_H8$ = G_Print_Time & G_G1 & Format(Now, "HH:MM:SS")
       
    End If
    
'??? �C�L������Y
    PrintOut3 Spd, H1$, "", 1         '�̫�@�ӰѼ�: �C�L�����
    PrintOut3 Spd, H2$, A_H2$, 2
    PrintOut3 Spd, H3$, A_H3$, 3
    PrintOut3 Spd, H4$, A_H4$, 4
    PrintOut3 Spd, H5$, A_H5$, 5
    PrintOut3 Spd, H6$, A_H6$, 6
    PrintOut3 Spd, H7$, A_H7$, 7
    PrintOut3 Spd, H8$, A_H8$, 8
    PrintOut3 Spd, H9$, "", -1
    PrintOut3 Spd, FC$, FC$, 9
    PrintOut3 Spd, H9$, "", -1

'??? �]�w�C�L��Excel��,���Y�ҥΪ��C��
    If G_PrintSelect = G_Print2Excel Then G_XlsHRows% = 9
End Sub


Sub PrintSub(Prb As ProgressBar, Spd As vaSpread, ByVal ShowProgress%, A_Exit%)
'�N��ƥ�SpreadŪ���C�L�ܤ�r��,�L�����Excel
Dim A_PrtStr$, A_A0909$, A_FmtStr$()
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
       
'??? �H���W�٨��o����,�m�JColumns Type��Text�ݩʤ�
'    �ѼƤ@ : Spread Name                               �ѼƤG : �ѼƤ@���ݪ�Spead Type Name
'    �ѼƤT : �ۭq�����W��                            �Ѽƥ| : ��ƦC
       GetSpdText Spd, tSpd_PATTERNR2, "A0901", A_Row#, , , , , True
       GetSpdText Spd, tSpd_PATTERNR2, "A0902", A_Row#
       GetSpdText Spd, tSpd_PATTERNR2, "A0907", A_Row#
       A_A0909$ = GetSpdText(Spd, tSpd_PATTERNR2, "A0909", A_Row#)
       GetSpdText Spd, tSpd_PATTERNR2, "A0906", A_Row#
       GetSpdText Spd, tSpd_PATTERNR2, "A0911", A_Row#
       GetSpdText Spd, tSpd_PATTERNR2, "A0912", A_Row#
       A_FmtStr$ = Split(GetSpdText(Spd, tSpd_PATTERNR2, "Flag", A_Row#), ";", , vbTextCompare)
       
       For I# = 0 To UBound(A_FmtStr$)
       
'�֭p�ثe�C�L���,�Y�W�L�@���h����
           G_LineNo = G_LineNo + 1
           PageCheck Spd
           
'??? �N�r��ǵ�PrintOut3�B�z�C�L�ʧ@
           Select Case UCase$(A_FmtStr$(I#))
             Case "H1$"
                  PrintOut3 Spd, H1$, "", -1
             Case "H9$"
                  PrintOut3 Spd, H9$, "", -1
             Case "B2$"
                  PrintOut3 Spd, B2$, "", -1
             Case "B1$"
                  G_ExcelIndex# = G_ExcelIndex# + 1
                  If G_PrintSelect = G_Print2Excel Then
                     A_PrtStr$ = PrintUse(B1$, GetCaption("PATTERNR", "username", "�ϥΪ�") & G_G1 & A_A0909$)
                  Else
                     A_PrtStr$ = GetCaption("PATTERNR", "username", "�ϥΪ�") & G_G1 & A_A0909$
                  End If
                  PrintOut3 Spd, B1$, A_PrtStr$, G_ExcelIndex#
                  '�Y�C�L��Excel��,�X��Break��쪺�x�s��
                  SetCellAlignment GetMergeCols(1, G_ExcelIndex# + G_XlsHRows%, G_ExcelMaxCols%, G_ExcelMaxCols%, 0), xlLeft, xlCenter, True
             Case "B3$"
                  G_ExcelIndex# = G_ExcelIndex# + 1
                  PrintOut3 Spd, B3$, A_A0909$, G_ExcelIndex#
                  '�]�wExcel Cells Range���I���C��
                  SetExcelRangeColor G_XlsHRows% + G_ExcelIndex#, G_XlsHRows% + G_ExcelIndex#, G_XlsStartCol%, G_ExcelMaxCols%, COLOR_YELLOW
                  '�Y�C�L��Excel��,�X��Break��쪺�x�s��
                  SetCellAlignment GetMergeCols(1, G_ExcelIndex# + G_XlsHRows%, G_ExcelMaxCols%, G_ExcelMaxCols%, 0), xlLeft, xlCenter, True
             Case "FD$"
                  G_ExcelIndex# = G_ExcelIndex# + 1
                  PrintOut3 Spd, fd$, PrintStrConnect(tSpd_PATTERNR2, 2), G_ExcelIndex#
             Case "NP"
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


Sub PrintBottom(Prb As ProgressBar, Spd As vaSpread)
'�C�L������
    
'??? �C�L�L��H
    PrintOut3 Spd, H1$, "", -1
    PrintOut3 Spd, H1$, "", -1
    PrintOut3 Spd, N2$, "", -1


'??? �N�_�l��줤�����,�HG_G1�r���N��Ƥ��Φ��h������
    SetExcelTextToColumns G_XlsStartCol%, 1, G_XlsHRows% + G_ExcelIndex#, SetXlsFldDataType(tSpd_PATTERNR2)
    
'�]�wExcel������榡
    SetExcelFormat

'??? �B�z�U�ئC�L�覡�������ʧ@
    PrintEnd4 Spd, tSpd_PATTERNR2
    Prb.Visible = False
End Sub


Sub Print2Spread(Prb As ProgressBar, Spd As vaSpread, A_Exit%)
'�N��ƦC�L��Spread�W
Dim A_FmtStr$, A_A0901$, A_A0902$, A_A0906$, A_A0907$, A_A0909$
Dim A_A0911$, A_A0912$, A_A0909_Brk$, A_A0911_Brk$
Dim A_Row#, A_Index#, A_Flag%
Dim A_Start_Break1#, A_Exit_Break1#             '�t�Τp�p
Dim A_Start_Break2#, A_Exit_Break2#             '�ϥΪ̤p�p
Dim A_Start_Break3#, A_Exit_Break3#             '�ϥΪ̦X�p

    Prb.Visible = True
    Prb.Value = 0
    
'�]�w��l��
    A_Row# = 0: A_Index# = 0
    A_Start_Break1 = 0: A_Exit_Break1# = 0
    A_Start_Break2 = 0: A_Exit_Break2# = 0
    A_Start_Break3 = 0: A_Exit_Break3# = 0
    Spd.MaxRows = 0
    
'�C�L���Y
    ReportHeader Spd
    
'�N�ϥΪ̤Ψt�ΥN��Keep��Break�ܼƤ�
    A_A0909_Brk$ = Trim$(DY_A09.Fields("A0909") & "")
    A_A0911_Brk$ = Trim$(DY_A09.Fields("A0911") & "")
          
'�]�w�ϥΪ̤Ψt�ΥN���ܦU�۪��C�L�ܼ�,�]�Ĥ@�������Break��쪺���
    A_A0909$ = A_A0909_Brk$
    A_A0911$ = A_A0911_Brk$
    
'�]�w�ӦC�C�L�M�Ϊ��榡������,�榡�����H�����Ϲj
    A_FmtStr$ = "B1$;B2$;FD$"

'�B�z�C�C��ƪ��C�L
    Do While Not DY_A09.EOF And Not A_Exit%
       
'�֥[�ثe�B�z����Ƶ���
       A_Index# = A_Index# + 1
    
'�ϥΪ̤��P���B�z�ʧ@
       If StrComp(A_A0909_Brk$, Trim$(DY_A09.Fields("A0909") & ""), vbTextCompare) <> 0 Then
       
'�C�L�t�Τp�p��Break
          PrintBreak Spd, A_Row#, GetCaption("PATTERNR", "systemsubtotal", "�t�Τp�p"), A_Start_Break1#, A_Exit_Break1#, "B2$;B3$;B2$"
          
'�C�L�ϥΪ̤p�p��Break
          PrintBreak Spd, A_Row#, GetCaption("PATTERNR", "usersubtotal", "�ϥΪ̤p�p"), A_Start_Break2#, A_Exit_Break2#, "B3$;H9$;H1$"
          
'�N�t�Τp�p�ΨϥΪ̤p�p���ܼ��k�s,�H�K���s�֭p
          A_Start_Break1# = 0
          A_Exit_Break1# = 0
          A_Start_Break2# = 0
          A_Exit_Break2# = 0
          
'Keep�ثe�ϥΪ̤Ψt�ΥN����Break�ܼƤ�
          A_A0909_Brk$ = Trim$(DY_A09.Fields("A0909") & "")
          A_A0911_Brk$ = Trim$(DY_A09.Fields("A0911") & "")
          
'�]�w�ϥΪ̤Ψt�ΥN���ܦU�۪��C�L�ܼ�,�]�Ĥ@�������Break��쪺���
          A_A0909$ = A_A0909_Brk$
          A_A0911$ = A_A0911_Brk$
          
'�]�w�ӦC�C�L�M�Ϊ��榡������,�榡�����H�����Ϲj
          A_FmtStr$ = "NP;B1$;B2$;FD$"
          
'�t�ΥN�����P���B�z�ʧ@
       ElseIf StrComp(A_A0911_Brk$, Trim$(DY_A09.Fields("A0911") & ""), vbTextCompare) <> 0 Then
       
'�C�L�t�Τp�p��Break
          PrintBreak Spd, A_Row#, GetCaption("PATTERNR", "systemsubtotal", "�t�Τp�p"), A_Start_Break1#, A_Exit_Break1#, "B2$;B3$;B2$"
          
'�N�t�Τp�p���ܼ��k�s,�H�K���s�֭p
          A_Start_Break1# = 0
          A_Exit_Break1# = 0

'Keep�ثe�t�ΥN����Break�ܼƤ�
          A_A0911_Brk$ = Trim$(DY_A09.Fields("A0911") & "")
          
'�]�w�t�ΥN���ܦC�L�ܼ�,�]�Ĥ@�������Break��쪺���
          A_A0911$ = A_A0911_Brk$
       End If
              
'Keep�C�L��Ʀ��ܼ�
       A_A0901$ = DateFormat2(DateOut(DY_A09.Fields("A0901") & ""))
       A_A0902$ = Format$(Mid$(DY_A09.Fields("A0902") & "", 1, 6), "00:00:00")
       A_Flag% = False
       Select Case Trim$(DY_A09.Fields("A0907") & "")
         Case "1"
              A_A0907$ = "Start"
'�֥[�n�����O=Start���t�Τp�p,�ϥΪ̤p�p�ΨϥΪ̦X�p
              A_Start_Break1# = A_Start_Break1# + 1
              A_Start_Break2# = A_Start_Break2# + 1
              A_Start_Break3# = A_Start_Break3# + 1
         Case "2"
              A_A0907$ = "Exit"
'�֥[�n�����O=Exit���t�Τp�p,�ϥΪ̤p�p�ΨϥΪ̦X�p
              A_Exit_Break1# = A_Exit_Break1# + 1
              A_Exit_Break2# = A_Exit_Break2# + 1
              A_Exit_Break3# = A_Exit_Break3# + 1
         Case "3"
              A_Flag% = True
              A_A0907$ = "Add"
         Case "4"
              A_Flag% = True
              A_A0907$ = "Delete"
         Case "5"
              A_Flag% = True
              A_A0907$ = "Edit"
       End Select
       A_A0906$ = ""
       If Trim$(DY_A09.Fields("A0906") & "") <> "" Then
          GetProgramName Trim$(DY_A09.Fields("A0906") & "")
          A_A0906$ = G_ProgramName
       End If
       A_A0912$ = ""
       If A_Flag% Then A_A0912$ = Trim$(DY_A09.Fields("A0912") & "")
       
'�NSpread�W��MaxRows�[�@
       AddSpreadMaxRows Spd, A_Row#

'??? �H���W�ٳ]�w���Ȧ�vaSpread
'    �ѼƤ@ : Spread Name                               �ѼƤG : �ѼƤ@���ݪ�Spead Type Name
'    �ѼƤT : �ۭq�����W��                            �Ѽƥ| : ��ƦC
'    �ѼƤ� : ��J��
       SetSpdText Spd, tSpd_PATTERNR2, "A0901", A_Row#, A_A0901$
       SetSpdText Spd, tSpd_PATTERNR2, "A0902", A_Row#, A_A0902$
       SetSpdText Spd, tSpd_PATTERNR2, "A0907", A_Row#, A_A0907$
       SetSpdText Spd, tSpd_PATTERNR2, "A0909", A_Row#, A_A0909$
       SetSpdText Spd, tSpd_PATTERNR2, "A0906", A_Row#, A_A0906$
       SetSpdText Spd, tSpd_PATTERNR2, "A0911", A_Row#, A_A0911$
       SetSpdText Spd, tSpd_PATTERNR2, "A0912", A_Row#, A_A0912$
       SetSpdText Spd, tSpd_PATTERNR2, "Flag", A_Row#, A_FmtStr$
       SetSpdText Spd, tSpd_PATTERNR2, "TEST", A_Row#, "TEST"
       
'�]�wSpread�Ĥ@�C���C��
       If G_PrintSelect = G_Print2Screen Then Spd.TopRow = SetSpreadTopRow(Spd)
       
'�Y��Q�e����ܫD�ù���ܪ��C�L�覡,����N���Prepare��V Screen��Spread�W.
'�YSpread��MaxRows�j�󵥩�100��,�h������PrintSub�NSpread�W����ƦL�X,
'�ñNMaxRows�k�s,�A�~��Prepare��Ʀ�V Screen.
       If (G_ReportDataFrom = G_FromRecordSet And G_PrintSelect <> G_Print2Screen) And A_Row# >= 100 Then
          GoSub Print2SpreadA
       End If
       
'�M�ŨϥΪ̤Ψt�ΥN�����C�L�ܼ�,Break�H��,���C�L����쪺���
       A_A0909$ = ""
       A_A0911$ = ""
       
'�]�w��ƦC���M�ή榡
       A_FmtStr$ = "FD$"
       
'��ܥثe�B�z�i��
       Prb.Value = A_Index#
       
       DoEvents
       
'��Esc��QĲ�o,�����C�L�ʧ@
       If A_Exit% Then Exit Do
       
       DY_A09.MoveNext
       
    Loop
    
'�wĲ�o�������, ���X���{��
    If A_Exit% Then Exit Sub

'�C�L���
'�C�L�t�Τp�p��Break
    PrintBreak Spd, A_Row#, GetCaption("PATTERNR", "systemsubtotal", "�t�Τp�p"), A_Start_Break1#, A_Exit_Break1#, "B2$;B3$;B2$"
          
'�C�L�ϥΪ̤p�p��Break
    PrintBreak Spd, A_Row#, GetCaption("PATTERNR", "usersubtotal", "�ϥΪ̤p�p"), A_Start_Break2#, A_Exit_Break2#, "B3$;H9$"
          
'�C�L�ϥΪ̦X�p
    PrintBreak Spd, A_Row#, GetCaption("PATTERNR", "usertotal", "�ϥΪ̦X�p"), A_Start_Break3#, A_Exit_Break3#, "B3$;H9$"

'�Y��Q�e����ܫD�ù���ܪ��C�L�覡,���ƳB�z����,���A�NSpread�W����ƦL�X.
    If (G_ReportDataFrom = G_FromRecordSet And G_PrintSelect <> G_Print2Screen) And Spd.MaxRows > 0 Then
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

Sub SetReportCols()
'*** �]�wQ Screen����Spd_Help vaSpread *******************************************************
'??? �ŧiSpread���A��Columns��Sorts���}�C�Ӽ�,
'    �ѼƤ@ : Spread Type Name
'    �ѼƤG : vaSpread�W������`��
'    �ѼƤT : �O�_���\User�ۭq�Ƨ����Ψ䶶��
    InitialCols tSpd_Help, 2, True
    
'??? �]�wvaSpread�W���Ҧ����αƧ�����Spread Type��
'    �ѼƤ@ : Spread Type Name
'    �ѼƤG : �]�w�ΨӦs��vaSpread�W��쪺���W��
'    �ѼƤT : Optional - �]�w�������(0:���  1:�Ȯ�����,�w�]��  2:�ä[����)
'    �Ѽƥ| : Optional - �]�w�{���w�]�Ƨ���쪺����
'    �ѼƤ� : Optional - �]�w�{���w�]�Ƨ���쪺��V(1:���W,�w�]��  2:����)
'    �ѼƤ� : Optional - �]�wBreak��쪺����
'    �ѼƤC : Optional - �]�wBreak���O�_�P��L�����ܩ�P�@�C�W(True,�w�]�� / False)
    AddReportCol tSpd_Help, "A0826", , 1
    AddReportCol tSpd_Help, "A0802", , 2
    
'??? ���User�ۭq���������ܶ��ǤαƧ����
'    �ѼƤ@ : Spread Type Name
'    �ѼƤG : vaSpread�Ҧb��Form Name
'    �ѼƤT : vaSpread Name
    GetSpreadDefault tSpd_Help, "frm_PATTERNR2q", "Spd_Help"

'*** �]�wV Screen����Spd_PatternR1 vaSpread ***************************************************
'??? �ŧiSpread���A��Columns��Sorts���}�C�Ӽ�,
'    �ѼƤ@ : Spread Type Name
'    �ѼƤG : vaSpread�W������`��
'    �ѼƤT : �O�_���\User�ۭq�Ƨ����Ψ䶶��
    InitialCols tSpd_PATTERNR2, 8, False
    
'??? �]�wvaSpread�W���Ҧ����αƧ�����Spread Type��
'    �ѼƤ@ : Spread Type Name
'    �ѼƤG : �]�w�ΨӦs��vaSpread�W��쪺���W��
'    �ѼƤT : Optional - �]�w�������(0:���  1:�Ȯ�����,�w�]��  2:�ä[����)
'    �Ѽƥ| : Optional - �]�w�{���w�]�Ƨ���쪺����
'    �ѼƤ� : Optional - �]�w�{���w�]�Ƨ���쪺��V(1:���W,�w�]��  2:����)
'    �ѼƤ� : Optional - �]�wBreak��쪺����
'    �ѼƤC : Optional - �]�wBreak���O�_�P��L�����ܩ�P�@�C�W(True,�w�]�� / False)
    AddReportCol tSpd_PATTERNR2, "A0909", , 1, , 1, False
    AddReportCol tSpd_PATTERNR2, "A0911", , 2, , 2
    AddReportCol tSpd_PATTERNR2, "A0901", , 3
    AddReportCol tSpd_PATTERNR2, "A0902", , 4
    AddReportCol tSpd_PATTERNR2, "A0907"
    AddReportCol tSpd_PATTERNR2, "A0906"
    AddReportCol tSpd_PATTERNR2, "A0912"
    AddReportCol tSpd_PATTERNR2, "Flag", 2
    
'??? ���User�ۭq���������ܶ��ǤαƧ����
'    �ѼƤ@ : Spread Type Name
'    �ѼƤG : vaSpread�Ҧb��Form Name
'    �ѼƤT : vaSpread Name
    GetSpreadDefault tSpd_PATTERNR2, "frm_PATTERNR2", "Spd_PATTERNR2"
End Sub

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
'�]�wExcel����榡,���ƦC�L����~�]�w
Dim A_MaxCol$, A_Row#

    If G_PrintSelect <> G_Print2Excel Then Exit Sub

'??? �ثeEXCEL�L��ĴX�C
    A_Row# = G_ExcelIndex# + G_XlsHRows%
    
'??? �ӳ̤j���Ƥ��W��
    A_MaxCol$ = Chr(Asc("A") + G_ExcelMaxCols% - 1)

'??? �b����|�P�[�W�u��
    xlsDrawLine G_ExcelWkb, "A" & Trim(Str(G_XlsHRows%)) & ":" & A_MaxCol$ & Trim(Str(A_Row#))

'??? ���D����m��(�Y���ݭn,�Эק�Rang�����d��Y�i�ϥ�) FALSE-���X���x�s��
    SetCellAlignment "A2:" & A_MaxCol$ & "2", xlCenter, xlCenter, True

'??? �_�l���/�ɶ�(�x�s��X��)
    SetCellAlignment GetMergeCols(1, 3, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, 1), xlLeft, xlCenter, True

'??? �I����/�ɶ�(�x�s��X��)
    SetCellAlignment GetMergeCols(1, 4, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, 1), xlLeft, xlCenter, True

'??? �t�ΥN��(�x�s��X��)
    SetCellAlignment GetMergeCols(1, 5, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, 1), xlLeft, xlCenter, True

'??? �{���N�X(�x�s��X��)
    SetCellAlignment GetMergeCols(1, 6, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, 1), xlLeft, xlCenter, True

'??? �s�եN��(�x�s��X��)
    SetCellAlignment GetMergeCols(1, 7, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, 1), xlLeft, xlCenter, True

'??? User ID(�x�s��X��)
    SetCellAlignment GetMergeCols(1, 8, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, 1), xlLeft, xlCenter, True

'??? �C�L����m�k
    SetCellAlignment GetExcelColName(G_ExcelMaxCols% + G_XlsStartCol% - 1) & "7", xlRight, xlCenter, True

'??? �C�L�ɶ��m�k
    SetCellAlignment GetExcelColName(G_ExcelMaxCols% + G_XlsStartCol% - 1) & "8", xlRight, xlCenter, True
    
'??? ���D�C�m��
    SetCellAlignment Trim(Str(G_XlsHRows%)) + ":" + Trim(Str(G_XlsHRows%)), xlCenter, xlCenter, False

'??? �]�wExcel��ΦC���w�]�j�p,�ýվ������e�ܳ̾A�e��
    SetExcelSize "A:" & A_MaxCol$
    
'�����]�w (Orientation%�Ѽƭ� - xlPortrait:���L  xlLandscape:��L)
    SetExcelAllocate "$1:$" & Trim(Str(G_XlsHRows%))
    
'�]�w�@���x�s��
    SelectExcelCells "A1"
End Sub

Sub SetPrintFormatStr()
'Run Time�]�w�����榡

'??? �������Y�榡�i���ܼƪ�l�Ȫ��ʧ@
    H3l$ = "############## : ########## / ########"
    H4l$ = "############## : ########## / ########"
    H5l$ = "############## : ########## ########################################"
    H6l$ = "############## : ########## ########################################"
    H7l$ = "############## : ### #########################################"
    H8l$ = "############## : ########## - ##########"
    HDate$ = "######## : ##########"
    HPerson$ = "######## : ############"
    B31$ = "########## : ~~~~~~~     ##### : ~~~~~~~     #### :~~~~~~~"
    B11$ = "############ : ############"

'�ù���ܤ����]�w����榡
    If G_PrintSelect = G_Print2Screen Then Exit Sub

'??? �]�w�����k���Ŷ����涡�j,�Y�ϥιw�]�ȥi����J
    SetRptAllocate
    
'??? ���o�����̤p�e��
    GetRptMinWidth H5l$ & Space(1) & HDate$
    
' �@�C�H�WHeader���榡�]�w =====================================================================
'??? ���o���D�θ�ƪ��r��榡(�ѼƤG��Ǧ^���榡���A -- 1:���D�榡 2:�����ܪ��榡)
'??? Multi Line �ɨϥ�
'    ' �w��P�w�]�Ȥ��P�����,���s�]�w������Y��쪺����覡
'    ReDefineHeaderAlign
'    ' �]�w�Ĥ@�CHeader��Caption
'    ReDefineReportHeader
'    '���o�Ĥ@�CHeader��Format
    'FC$ = GetRptFormatStr(tSpd_PATTERNR2, 3)
'    ' �]�w�ĤG�CHeader��Caption
'    ReDefineReportHeader
'    '���o�ĤG�CHeader��Format
'    FC$ = GetRptFormatStr(tSpd_PATTERNR2, 3)
'    fd$ = GetRptFormatStr(tSpd_PATTERNR2, 2)
' ==============================================================================================
   
'??? ���o���D�θ�ƪ��r��榡(�ѼƤG��Ǧ^���榡���A -- 1:���D�榡 2:�����ܪ��榡)
    ' �w��P�w�]�Ȥ��P�����,���s�]�w������Y��쪺����覡
    ReDefineHeaderAlign

'??? ���Y��Single Line �ɨϥ�
    FC$ = GetRptFormatStr(tSpd_PATTERNR2, 1)
    fd$ = GetRptFormatStr(tSpd_PATTERNR2, 2)

'??? ���o������Y���榡
    H2$ = GetRptTitleFormat()
    
'??? ���o������Y��ƪ��榡
    H3l$ = PrintUse(H3l$, GetCaption("PATTERNR", "startdate", "�_�l���/�ɶ�") & G_G1 & DateFormat(G_A0901s$) & G_G1 & Format(Left(G_A0902s$, 6), "00:00:00"))
    H4l$ = PrintUse(H4l$, GetCaption("PATTERNR", "enddate", "�I����/�ɶ�") & G_G1 & DateFormat(G_A0901e$) & G_G1 & Format(Left(G_A0902e$, 6), "00:00:00"))
    H5l$ = PrintUse(H5l$, GetCaption("PATTERNR", "systemid", "�t�ΥN��") & G_G1 & G_A0911$ & G_G1 & G_A0911o$)
    H6l$ = PrintUse(H6l$, GetCaption("PATTERNR", "programid", "�{���N�X") & G_G1 & G_A0906$ & G_G1 & G_A0906o$)
    H7l$ = PrintUse(H7l$, GetCaption("PanelDescpt", "groupid", "�s�եN�X") & G_G1 & G_A0905$ & G_G1 & G_A0905o$)
    H8l$ = PrintUse(H8l$, GetCaption("PATTERNR", "userid", "User ID") & G_G1 & G_A0904s$ & G_G1 & G_A0904e$)
    H3$ = GetRptHeaderFormat(H3l$)
    H4$ = GetRptHeaderFormat(H4l$)
    H5$ = GetRptHeaderFormat(H5l$)
    H6$ = GetRptHeaderFormat(H6l$, HDate$)
    H7$ = GetRptHeaderFormat(H7l$, HDate$)
    H8$ = GetRptHeaderFormat(H8l$, HDate$)

'??? ���o����Rreak��쪺�榡
    B1$ = GetRptHeaderFormat(B11$)
   
'??? ���o��U���ΦL��H���榡
    N1$ = GetRptFootFormat(HPerson$)
    N2$ = PrintUse(GetRptLineFormat("~"), HPerson$)
    
'??? ���o�Ϲj�C���榡
    B2$ = GetRptLineFormat("-")
    B3$ = GetRptLineFormat("#")
    H9$ = GetRptLineFormat("=")
End Sub


Sub ReDefineHeaderAlign()
'�w��P�w�]�Ȥ��P�����,���s�]�w������Y��쪺����覡

    ChangeReportHeaderAlign tSpd_PATTERNR2, "A0901", SS_CELL_H_ALIGN_LEFT
'    ChangeReportHeaderAlign tSpd_PATTERNR2, "A0902", SS_CELL_H_ALIGN_CENTER
'    ChangeReportHeaderAlign tSpd_PATTERNR2, "A0906", SS_CELL_H_ALIGN_CENTER
'    ChangeReportHeaderAlign tSpd_PATTERNR2, "A0907", SS_CELL_H_ALIGN_RIGHT
'    ChangeReportHeaderAlign tSpd_PATTERNR2, "A0909", SS_CELL_H_ALIGN_RIGHT
'                                   :
'                                   :
End Sub

Sub ReDefineReportHeader()
'�w��P�w�]�Ȥ��P�����,���s�]�w������Y��쪺Caption

'    ChangeReportHeader tSpd_PATTERNR2, "A0901", "Test"
'    ChangeReportHeader tSpd_PATTERNR2, "A0902", "Test"
'    ChangeReportHeader tSpd_PATTERNR2, "A0906", "Test"
'    ChangeReportHeader tSpd_PATTERNR2, "A0907", "Test"
'    ChangeReportHeader tSpd_PATTERNR2, "A0909", "Test"
'                                   :
'                                   :
End Sub

