Attribute VB_Name = "mod_EXAM02"
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

'�w�q�U Form ���D��r�ܼ�
Global G_Form_EXAM02 As String
Global G_Form_EXAM02v As String
Global G_Form_EXAM02q As String

'�w�q�U�����D��r
Global G_Pnl_Dash$
Global G_Pnl_A0801$
Global G_Pnl_A0802$
Global G_Pnl_A0803$
Global G_Pnl_A0804$
Global G_Pnl_A0805$
Global G_Pnl_A0806$
Global G_Pnl_A0807$
Global G_Pnl_A0808$
Global G_Pnl_A0809$
Global G_Pnl_A0810$
Global G_Pnl_A0811$
Global G_Pnl_A0812$
Global G_Pnl_A0813$
Global G_Pnl_A0814$
Global G_Pnl_A0815$
Global G_Pnl_A0816$
Global G_Pnl_A0817$
Global G_Pnl_A0818$
Global G_Pnl_A0819$
Global G_Pnl_A0820$
Global G_Pnl_A0821$
Global G_Pnl_A0822$
Global G_Pnl_A0823$
Global G_Pnl_A0824$
Global G_Pnl_A0825$
Global G_Pnl_A0826$
Global G_Pnl_A0201$
Global G_Pnl_A0202$
Global G_Pnl_A0601$
Global G_Pnl_A0602$
Global G_RecordNotExist$


'Def �{���@���ܼ�
Global G_A0801$                      'Keep�����O
Global G_A0801o$                     'Keep Pre �����O
Global G_A0824n$

Global G_ActiveRow#                  'Keep��ƥثe�Ҧb�C
Global G_MaxRows#                    'Keep����`����

Sub GetPanelCaption()
'��FORM���D��r
    G_Form_EXAM02 = GetCaption("FormTitle", "EXAM02", "���u�򥻸�ƺ޲z")
    G_Form_EXAM02v = GetCaption("FormTitle", "EXAM02V", "���u�򥻸�ƥؿ�")
    G_Form_EXAM02q = GetCaption("FormTitle", "EXAM02Q", "���u�򥻸�Ƭd��")
'�������D��r
    G_Pnl_A0801$ = GetCaption("EXAM02", "A0801", "���u�s��")
    G_Pnl_A0802$ = GetCaption("EXAM02", "A0802", "����m�W")
    G_Pnl_A0803$ = GetCaption("EXAM02", "A0803", "�^��m�W")
    G_Pnl_A0804$ = GetCaption("EXAM02", "A0804", "�����N��")
    G_Pnl_A0805$ = GetCaption("EXAM02", "A0805", "��¾���")
    G_Pnl_A0806$ = GetCaption("EXAM02", "A0806", "��¾���")
    G_Pnl_A0807$ = GetCaption("EXAM02", "A0807", "�K�X")
    G_Pnl_A0808$ = GetCaption("EXAM02", "A0808", "�X�ͤ��")
    G_Pnl_A0809$ = GetCaption("EXAM02", "A0809", "�����Ҹ��X")
    G_Pnl_A0810$ = GetCaption("EXAM02", "A0810", "����a�}")
    G_Pnl_A0811$ = GetCaption("EXAM02", "A0811", "�^��a�}")
    G_Pnl_A0812$ = GetCaption("EXAM02", "A0812", "����")
    G_Pnl_A0813$ = GetCaption("EXAM02", "A0813", "�l���ϸ�")
    G_Pnl_A0814$ = GetCaption("EXAM02", "A0814", "��a")
    G_Pnl_A0815$ = GetCaption("EXAM02", "A0815", "�s���q��")
    G_Pnl_A0816$ = GetCaption("EXAM02", "A0816", "�s���ǯu")
    G_Pnl_A0817$ = GetCaption("EXAM02", "A0817", "BB Call")
    G_Pnl_A0818$ = GetCaption("EXAM02", "A0818", "��ʹq��")
    G_Pnl_A0819$ = GetCaption("EXAM02", "A0819", "E-Mail Address")
    G_Pnl_A0820$ = GetCaption("EXAM02", "A0820", "���Ĥ��")
    G_Pnl_A0821$ = GetCaption("EXAM02", "A0821", "�ʧO")
    G_Pnl_A0822$ = GetCaption("EXAM02", "A0822", "�B�ê��p")
    G_Pnl_A0823$ = GetCaption("EXAM02", "A0823", "¾��")
    G_Pnl_A0824$ = GetCaption("EXAM02", "A0824", "���q�O�N�X")
    G_Pnl_A0825$ = GetCaption("EXAM02", "A0825", "�s�եN��")
    G_Pnl_A0826$ = GetCaption("EXAM02", "A0826", "User ID")
    
    G_Pnl_A0201$ = GetCaption("EXAM02", "A0201", "���u�s��")
    G_Pnl_A0202$ = GetCaption("EXAM02", "A0202", "���u����W��")
    
    G_Pnl_A0601$ = GetCaption("EXAM02", "A0601", "�s�եN��")
    G_Pnl_A0602$ = GetCaption("EXAM02", "A0602", "�s�ջ���")
    
    G_RecordNotExist$ = GetCaption("PgmMsg", "g_record_no_exist", "��Ƥ��s�b! �ЦA�d��!")
    
'����L�ܼƤ��t��
'    G_Pnl_A1602$ = GetCaption("EXAM02", "bankname")
    G_Pnl_Dash$ = GetCaption("PanelDescpt", "dash", "��")
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
    Load frm_EXAM02            ' ���NDetail�e��Load��Memory
    Frm_EXAM02q.Show           ' �����e�����
    Screen.MousePointer = Default
End Sub
