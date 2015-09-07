Attribute VB_Name = "mod_TSM01"
Option Explicit
Option Compare Text

'========================================================================
' Coding Rules
'========================================================================
'�b���B�w�q���Ҧ��ܼ�, �@�ߥHG�}�Y
'       Ex: G_AAA$, G_BBB#, G_CCC&
'
'�B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
' $: ��r
' #: �Ҧ��Ʀr�B��(���B�μƶq)
' &: �{���j���ܼ�
' %: ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
' �ť�: �N��VARIENT, �ʺA�ܼ�
'========================================================================

'-- Fixed Variables (���n�ܼ�) :
Global G_FormFrom$    '�ťեN����������


'�w�q�U Form ���D��r�ܼ�
Global G_Form_PATTERN$

'�w�q�U�����D��r

Global G_Pnl_A0101$
Global G_Pnl_A0102$
Global G_Pnl_A0104$
Global G_Pnl_A0111$

'Def �{���@���ܼ�
''SAMPLE
Global G_BB#
Global G_CC!


Sub GetPanelCaption()
'��FORM���D��r
    G_Form_PATTERN$ = GetSIniStr("FormTitle", "TSM01")
'�������D��r
     G_Pnl_A0101$ = GetCaption("TSM01", "ik01", "���q�N��")
     G_Pnl_A0102$ = GetCaption("TSM01", "ik02", "���q²��")
     G_Pnl_A0104$ = GetCaption("TSM01", "ik03", "���q���W")
     G_Pnl_A0111$ = GetCaption("TSM01", "ik04", "�Τ@�s��")
'����L�ܼƤ��t��
End Sub

Sub Main()
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
    
    Frm_TSM01.Show            ' �����e�����
    Screen.MousePointer = Default
End Sub
