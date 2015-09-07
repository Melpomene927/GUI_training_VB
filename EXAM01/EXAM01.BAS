Attribute VB_Name = "mod_EXAM01"
'========================================================================
' Module    : mod_EXAM01
' Author    : Mike_chang
' Date      : 2015/8/27
' Purpose   :
'========================================================================
Option Explicit
Option Compare Text

'========================================================================n
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
Global G_Form_EXAM01 As String
Global G_Form_EXAM01v As String
Global G_Form_EXAM01q As String

'�w�q�U�����D��r
Global G_Pnl_Dash$
Global G_Pnl_A1601$
Global G_Pnl_A1602$
Global G_Pnl_A1605$
Global G_Pnl_A1606$
Global G_Pnl_A1609$
Global G_Pnl_A1628$
Global G_Pnl_A1612$

Global G_Pnl_A1902$
Global G_Pnl_A1902n$
Global G_Pnl_A1903$


'SAMPLE
'Global G_AA$
'Global G_BB#
'Global G_CC!

'Def �{���@���ܼ�
Global G_A1601$                      'Keep�����O
Global G_A1601o$                     'Keep Pre �����O
Global G_A1902$
Global G_Hlp_Return                  'Keep return value from form_GD

Global G_ActiveRow#                  'Keep��ƥثe�Ҧb�C
Global G_MaxRows#                    'Keep����`����


'========================================================================
' Procedure : GetPanelCaption (mod_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Get Panel Caption from "Sini" table
' Details   : Using GUI_common_module function:
'             "GetCaption" to get topicvalue from sini
'========================================================================
Sub GetPanelCaption()
    '��FORM���D��r
    G_Form_EXAM01 = GetCaption("FormTitle", "MCFGD", "�Ȥ�򥻸�ƺ��@")
    G_Form_EXAM01v = GetCaption("FormTitle", "", "�Ȥ�򥻸�ƤT�B�s��")
    G_Form_EXAM01q = GetCaption("FormTitle", "MCFGDV", "�Ȥ�򥻸�Ƭd��")
    '�������D��r

    G_Pnl_A1601$ = GetCaption("MCFGD", "order1", "�Ȥ�s��")
    G_Pnl_A1602$ = GetCaption("paneldescpt", "custmer", "�Ȥ�²��")
    G_Pnl_A1605$ = GetCaption("PanelDescpt", "10", "�p���q��")
    G_Pnl_A1606$ = GetCaption("KBM01", "fax", "�ǯu�q��")
    G_Pnl_A1612$ = GetCaption("MCFGD", "address0", "�p���a�}")
    G_Pnl_A1609$ = GetCaption("MCFGD", "", "������/�νs")
    G_Pnl_A1628$ = GetCaption("MCFGD", "", "�ͤ�/���ߤ�")
    
    
    G_Pnl_A1902$ = GetCaption("MCFGD", "", "���Y�H�s��")
    G_Pnl_A1902n$ = GetCaption("MCFGD", "", "���Y�H²��")
    G_Pnl_A1903$ = GetCaption("PanelDescpt", "relationship", "���Y")

    
    '����L�ܼƤ��t��
'    G_Pnl_A1602$ = GetCaption("EXAM01", "bankname")
    G_Pnl_Dash$ = GetCaption("PanelDescpt", "dash", "��")
End Sub

'========================================================================
' Module    : mod_EXAM01
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
    Load frm_EXAM01q          ' ���NDetail�e��Load��Memory
    frm_EXAM01q.Show          ' �����e�����
    Screen.MousePointer = Default
End Sub