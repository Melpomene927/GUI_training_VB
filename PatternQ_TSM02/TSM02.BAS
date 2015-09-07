Attribute VB_Name = "mod_TSM02"
'========================================================================
' Module    : mod_TSM02
' Author    : Mike_chang
' Date      : 2015/8/27
' Purpose   :
'========================================================================
Option Explicit
Option Compare Text

'========================================================================
'   Coding Rule
'========================================================================
'�b���B�w�q���Ҧ��ܼ�, �@�ߥHG�}�Y, �pG_AAA$, G_BBB#, G_CCC&
'�B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
' $: ��r
' #: �Ҧ��Ʀr�B��(���B�μƶq)
' &: �{���j���ܼ�
' %: ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
' �ť�: �N��VARIENT, �ʺA�ܼ�
'========================================================================

'���n�ܼ�
Global G_FormFrom$    '�ťեN����������

'�w�q�U Form ���D��r�ܼ�
Global G_Form_TSM02 As String
Global G_Form_TSM02v As String
Global G_Form_TSM02q As String

'�w�q�U�����D��r
Global G_Pnl_Dash$
Global G_Pnl_A02001$
Global G_Pnl_A02005$
Global G_Pnl_A0201$
Global G_Pnl_A0202$
Global G_Pnl_A0203$
Global G_Pnl_A0204$
Global G_Pnl_A0205$
Global G_Pnl_A0206$
Global G_Pnl_A0207$
Global G_Pnl_A0208$
Global G_Pnl_A0209$
Global G_Pnl_A0213$
Global G_Pnl_A0214$
Global G_Pnl_A0215$
Global G_Pnl_A0216$
Global G_Pnl_A0217$
Global G_Pnl_A0218$
Global G_Pnl_A0219$
Global G_Pnl_Menu$
Global G_Pnl_Selected$

'Global G_Pnl_A15001$
'Global G_Pnl_A15005$
'Global G_Pnl_A1501$
'Global G_Pnl_A1502$
'Global G_Pnl_A1504$
'Global G_Pnl_A1505$
'Global G_Pnl_A1506$
'Global G_Pnl_A1507$
'Global G_Pnl_A1508$
'Global G_Pnl_A1509$
'Global G_Pnl_A1510$
'Global G_Pnl_A1511$
'Global G_Pnl_A1512$
'Global G_Pnl_A1513$
'Global G_Pnl_A1514$
'Global G_Pnl_A1519$
'Global G_Pnl_A1521$
'Global G_Pnl_A1523$
'Global G_Pnl_A1601$
'Global G_Pnl_A1602$


'Def �{���@���ܼ�
Global G_A0201$                      'Keep�����O
Global G_A0201o$                     'Keep Pre �����O

Global G_ActiveRow#                  'Keep��ƥثe�Ҧb�C
Global G_MaxRows#                    'Keep����`����
'SAMPLE
Global G_AA$
Global G_BB#
Global G_CC!

'========================================================================
' Module    : mod_TSM02
' Procedure : GetPanelCaption
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Get Panel Caption from "Sini" table
' Details   : Using GUI_common_module function:
'             "GetSIniStr" to get topicvalue from sini
'========================================================================
Sub GetPanelCaption()
'��FORM���D��r
    G_Form_TSM02 = GetSIniStr("FormTitle", "TSM02")
    G_Form_TSM02v = GetSIniStr("FormTitle", "TSM02V")
    G_Form_TSM02q = GetSIniStr("FormTitle", "TSM02Q")
'�������D��r
    G_Pnl_A02001$ = GetSIniStr("PanelDescpt", "Date_Add")
    G_Pnl_A02005$ = GetSIniStr("PATTERNQ", "date_update")
    G_Pnl_A0201$ = GetSIniStr("TSM02", "A0201")
    G_Pnl_A0202$ = GetSIniStr("TSM02", "A0202")
    G_Pnl_A0203$ = GetSIniStr("TSM02", "A0203")
    G_Pnl_A0204$ = GetSIniStr("TSM02", "A0204")
    G_Pnl_A0205$ = GetSIniStr("TSM02", "A0205")
    G_Pnl_A0206$ = GetSIniStr("TSM02", "A0206")
    G_Pnl_A0207$ = GetSIniStr("TSM02", "A0207")
    G_Pnl_A0208$ = GetSIniStr("TSM02", "A0208")
    G_Pnl_A0209$ = GetSIniStr("TSM02", "A0209")
    
    G_Pnl_A0213$ = GetSIniStr("TSM02", "A0213")
    G_Pnl_A0214$ = GetSIniStr("TSM02", "A0214")
    G_Pnl_A0215$ = GetSIniStr("TSM02", "A0215")
    G_Pnl_A0216$ = GetSIniStr("TSM02", "A0216")
    G_Pnl_A0217$ = GetSIniStr("TSM02", "A0217")
    G_Pnl_A0218$ = GetSIniStr("TSM02", "A0218")
    G_Pnl_A0219$ = GetSIniStr("TSM02", "A0219")
    
'����L�ܼƤ��t��
'    G_Pnl_A1602$ = GetSIniStr("PATTERNQ", "bankname")
    G_Pnl_Dash$ = GetSIniStr("PanelDescpt", "dash")
    G_Pnl_Menu$ = GetSIniStr("PATTERNQ", "listmenu")
    G_Pnl_Selected$ = GetSIniStr("PATTERNQ", "listselected")
End Sub

'========================================================================
' Module    : mod_TSM02
' Procedure : main
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   :
' Details   :
'========================================================================
Sub main()
' ���Ҳդ�, �������ӤU�C���ǰ���, �p�G���S�����p���N�Y�ǼҲ�������,
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
    Load frm_TSM02            ' ���NDetail�e��Load��Memory
    frm_TSM02q.Show           ' �����e�����
    Screen.MousePointer = Default
End Sub