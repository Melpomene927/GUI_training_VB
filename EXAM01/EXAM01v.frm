VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_EXAM01v 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0FFFF&
   Caption         =   "�|�p��إؿ�"
   ClientHeight    =   6735
   ClientLeft      =   3090
   ClientTop       =   1200
   ClientWidth     =   9390
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "EXAM01v.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6735
   ScaleWidth      =   9390
   Begin VsOcxLib.VideoSoftElastic Vse_background 
      Height          =   6360
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   9390
      _Version        =   327680
      _ExtentX        =   16563
      _ExtentY        =   11218
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Align           =   5
      BevelOuter      =   6
      Picture         =   "EXAM01v.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "EXAM01v.frx":0326
      Begin FPSpread.vaSpread Spd_EXAM01v 
         Height          =   6165
         Left            =   60
         OleObjectBlob   =   "EXAM01v.frx":0342
         TabIndex        =   7
         Top             =   90
         Width           =   7755
      End
      Begin Threed.SSCommand cmd_delete 
         Height          =   405
         Left            =   7900
         TabIndex        =   1
         Top             =   540
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�R��F3"
      End
      Begin Threed.SSCommand cmd_add 
         Height          =   405
         Left            =   7900
         TabIndex        =   2
         Top             =   990
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�s�WF4"
      End
      Begin Threed.SSCommand cmd_previous 
         Height          =   405
         Left            =   7905
         TabIndex        =   3
         Top             =   1440
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�e��F7"
      End
      Begin Threed.SSCommand cmd_next 
         Height          =   405
         Left            =   7905
         TabIndex        =   4
         Top             =   1890
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "����F8"
      End
      Begin Threed.SSCommand cmd_exit 
         Height          =   405
         Left            =   7905
         TabIndex        =   5
         Top             =   5850
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "����Esc"
      End
      Begin Threed.SSCommand cmd_help 
         Height          =   405
         Left            =   7900
         TabIndex        =   0
         Top             =   90
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "����F1"
      End
      Begin ComctlLib.ProgressBar Prb_Percent 
         Height          =   405
         Left            =   8580
         TabIndex        =   9
         Top             =   3810
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   714
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   6360
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_EXAM01v"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'   �b���B�w�q���Ҧ��ܼ�, �@�ߥHM�}�Y, �pM_AAA$, M_BBB#, M_CCC&
'   �B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
'   $: ��r
'   #: �Ҧ��Ʀr�B��(���B�μƶq)
'   &: �{���j���ܼ�
'   %: ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
'   �ť�: �N��VARIENT, �ʺA�ܼ�

'�۩w�ܼ�
'Dim m_A4101Flag%
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

'���n�ܼ�
Dim m_FieldError%    '���ܼƦb�P�_���O�_���~, �����^�����줧�ʧ@
Dim m_ExitTrigger%   '���ܼƦb�P�_������O�_�QĲ�o, �N����ثe���b�B�z���@�~

Private Function MoveDB2Spread() As Boolean
On Local Error GoTo MY_Error
Dim A_Row&
Dim A_Records&
        
    Me.MousePointer = HOURGLASS
    MoveDB2Spread = True
    
    '�NSpread�W���`���Ƴ]��0
    Spd_EXAM01v.MaxRows = 0

    '���o�`����
    DY_A16.MoveLast
    A_Records& = DY_A16.RecordCount
    DY_A16.MoveFirst

    '��ƬO�_���
    '=========================================================
    '   Function: "DisplayOverMaxLines" will show a Dialog to
    '   ask whether user want to show the data
    '   if "cancel" clicked: exit
    '   else continue the process
    '=========================================================
    If Not DisplayOverMaxLines(A_Records&) Then
       Me.MousePointer = Default
       MoveDB2Spread = False
       Exit Function
    End If
    
    ProgressBoxShow Me, Spd_EXAM01v     'Show Progress Box
    Prb_Percent.MAX = A_Records&        '�]�wProgress Box���̤j��

    '�N��ƥ��Spread�W
    With Spd_EXAM01v
         Do While Not DY_A16.EOF And Not m_ExitTrigger%
            A_Row& = A_Row& + 1
            .MaxRows = A_Row&
            .Row = A_Row&
            .Col = 1
            .text = Trim$(DY_A16.Fields("A1601") & "")
            .Col = 2
            .text = Trim$(DY_A16.Fields("A1602") & "")
            .Col = 3
            .text = Trim$(DY_A16.Fields("A1609") & "")
            .Col = 4
            .text = DateFormat2(DateOut(DY_A16.Fields("A1628") & ""))
            .Col = 5
            .text = Trim$(DY_A16.Fields("A1605") & "")
            .Col = 6
            .text = Trim$(DY_A16.Fields("A16121") & "")
            .text = .text & Trim$(DY_A16.Fields("A16122") & "")
            .text = .text & Trim$(DY_A16.Fields("A16123") & "")
            
            '�]�wSpread�W�Ĥ@�����d�b�ĴX�C
            .TopRow = SetSpreadTopRow(Spd_EXAM01v)
            
            '��ܥثe�i��
            Prb_Percent.Value = A_Row&
            
            DoEvents
            DY_A16.MoveNext
         Loop
    
        '��ƥ����᧹��,�]�wSpread�W�Ĥ@�����d�b�Ĥ@�C
        .TopRow = 1
    End With

    'Hide Progress Box
    ProgressBoxHide Me, Spd_EXAM01v
    Me.MousePointer = Default
    Exit Function
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Private Function Reference_INI(ByVal A_Section$, ByVal A_Topic$) As String
On Local Error GoTo MyError
Dim A_Sql$

    Reference_INI = ""
    A_Sql$ = "Select TOPICVALUE From SINI"
    A_Sql$ = A_Sql$ & " where SECTION='" & A_Section$ & "'"
    A_Sql$ = A_Sql$ & " and TOPIC='type" & A_Topic$ & "'"
    A_Sql$ = A_Sql$ & " order by SECTION,TOPIC"
    
    CreateDynasetODBC DB_ARTHGUI, DY_INI, A_Sql$, "DY_INI", True
    
    If Not (DY_INI.BOF And DY_INI.EOF) Then
       Reference_INI = Trim$(DY_INI.Fields("TOPICVALUE") & "")
    End If
    Exit Function
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Private Sub Set_Property()
'�]�w��Form�����D,�r�ΤΦ�t
    Form_Property frm_EXAM01v, G_Form_EXAM01v, G_Font_Name
    
    '�]Form���Ҧ�Command�����D�Φr��
    Command_Property cmd_help, G_CmdHelp, G_Font_Name
    Command_Property cmd_delete, G_CmdDel, G_Font_Name
    Command_Property cmd_add, G_CmdAdd, G_Font_Name
    Command_Property cmd_previous, G_CmdPrvPage, G_Font_Name
    Command_Property cmd_next, G_CmdNxtPage, G_Font_Name
    Command_Property cmd_exit, G_CmdExit, G_Font_Name
    
    'Disable the Pre & Next Buttom
    cmd_previous.Enabled = False
    cmd_previous.Visible = False
    cmd_next.Enabled = False
    cmd_next.Visible = False
    
    '�]Form��Spread���ݩ�
    Set_Spread_Property
    
    '�]Form��Progress Bar ���ݩ�
    ProgressBar_Property Prb_Percent
    
    '�H�U���зǫ��O, ���o�ק�
    VSElastic_Property Vse_background
    StatusBar_ProPerty Sts_MsgLine
End Sub

Private Sub Set_Spread_Property()
    Spd_EXAM01v.UnitType = 2      '<--- @!!! Fixed Property, DO NOT CHANGE. !!!
    
    '�]�w��Spread�����Ƥ�����
    Spread_Property Spd_EXAM01v, 0, 6, WHITE, G_Font_Size, G_Font_Name

    '�]�w��Spread���U����D����ܼe��, 0�N�����줣���
    Spread_Col_Property Spd_EXAM01v, 1, TextWidth("X") * 10, G_Pnl_A1601$
    Spread_Col_Property Spd_EXAM01v, 2, TextWidth("X") * 12, G_Pnl_A1602$
    Spread_Col_Property Spd_EXAM01v, 3, TextWidth("X") * 15, G_Pnl_A1609$
    Spread_Col_Property Spd_EXAM01v, 4, TextWidth("9") * 10, G_Pnl_A1628$
    Spread_Col_Property Spd_EXAM01v, 5, TextWidth("X") * 15, G_Pnl_A1605$
    Spread_Col_Property Spd_EXAM01v, 6, TextWidth("X") * 50, G_Pnl_A1612$
    
    '========================================
    '   �]�w��Spread���U���ݩʤ���ܦr��
    '   SS_CELL_TYPE_EDIT        = ��r�i��J
    '   SS_CELL_TYPE_FLOAT       = �Ʀr�i��J
    '   SS_CELL_TYPE_STATIC_TEXT = �����
    '   SS_CELL_TYPE_CHECKBOX    = �I�ﶵ��
    '========================================
    Spread_DataType_Property Spd_EXAM01v, 1, SS_CELL_TYPE_EDIT, "", "", 10
    Spread_DataType_Property Spd_EXAM01v, 2, SS_CELL_TYPE_EDIT, "", "", 12
    Spread_DataType_Property Spd_EXAM01v, 3, SS_CELL_TYPE_EDIT, "", "", 15
    Spread_DataType_Property Spd_EXAM01v, 4, SS_CELL_TYPE_EDIT, "", "", 10
    Spread_DataType_Property Spd_EXAM01v, 5, SS_CELL_TYPE_EDIT, "", "", 15
    Spread_DataType_Property Spd_EXAM01v, 6, SS_CELL_TYPE_EDIT, "", "", 120
    
    Spd_EXAM01v.EditEnterAction = SS_CELL_EDITMODE_EXIT_NONE
    
    '�T�w�V�k���ʮ�, �ҭ�����
    Spd_EXAM01v.ColsFrozen = 1
    
    '�w�q�Y����m����m���]�w 0:���a  1:�k�a  2:�m��
    Spd_EXAM01v.Row = -1
'    Spd_EXAM01v.Col = 7: Spd_EXAM01v.TypeHAlign = 2
    
    '�w�q�Y�������,���i�ק���
    Spd_EXAM01v.Col = -1
    Spd_EXAM01v.Lock = True
End Sub

Private Sub cmd_add_Click()
    Me.MousePointer = HOURGLASS
    
    'set status to "ADD"
    G_AP_STATE = G_AP_STATE_ADD
    
    '=====================================================
    '   clean up rows since "ADD" procedure need to show
    '   those records which are added by D-form
    '=====================================================
    Spd_EXAM01v.MaxRows = 0
    
    DoEvents
    
    Me.Hide
    frm_EXAM01.Show
    Me.MousePointer = Default
End Sub

Private Sub cmd_delete_Click()
    If Spd_EXAM01v.MaxRows = 0 Then Exit Sub

    Me.MousePointer = HOURGLASS
    
    With Spd_EXAM01v
        '���oV�e�����`���ƤΥثe���d�C
        G_AP_STATE = G_AP_STATE_DELETE
        G_MaxRows# = .DataRowCnt
        G_ActiveRow# = .ActiveRow  'keep the record to delete

        'Keep P-Key, ��Detail�e���R�����
        .Row = G_ActiveRow#
        .Col = 1
        StrCut .text, Space(1), G_A1601$, ""   'fetch Pkey to Global var
'       .Col = 2
'       G_A1502$ = Trim$(.text)
    End With
    
    '����V�e��, ������Detail�e��
    DoEvents
    Me.Hide
    frm_EXAM01.Show
    
    Me.MousePointer = Default
End Sub

Private Sub Cmd_Exit_Click()
    '�����ثe����,���X��L�B�z�{��
    m_ExitTrigger% = True
    
    '���åثe�e��, ���Q�e��
    DoEvents
    Me.Hide
    frm_EXAM01q.Show
End Sub

Private Sub Cmd_Help_Click()
Dim a$
    a$ = "notepad " + G_Help_Path + "EXAM01v.HLP"
    retcode = Shell(a$, 4)
End Sub

Private Sub Cmd_Next_Click()
'    cmd_next.Enabled = False
'    Spd_EXAM01v.SetFocus
'    SendKeys "{PgDn}"
'    DoEvents
'    cmd_next.Enabled = True
End Sub

Private Sub Cmd_Previous_Click()
'    cmd_previous.Enabled = False
'    Spd_EXAM01v.SetFocus
'    SendKeys "{PgUp}"
'    DoEvents
'    cmd_previous.Enabled = True
End Sub

Private Sub Form_Activate()
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    Me.Refresh
    
    'Initial Form�������n�ܼ�
    m_FieldError% = -1
    m_ExitTrigger% = False
    
    If G_AP_STATE = G_AP_STATE_QUERY Then
        Sts_MsgLine.Panels(1) = G_Process
        Sts_MsgLine.Refresh
        '�N�d�߸�ƥ��Spread�W,�Y���ƹL�h�����,�h�^Q�e��
        If Not MoveDB2Spread() Then
            DoEvents
            Me.Hide
            frm_EXAM01q.Show
            Exit Sub
        End If
        Sts_MsgLine.Panels(1) = G_Query_Ok
    Else
        G_AP_STATE = G_AP_STATE_NORMAL
        Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    End If
    
    '�NForm��m��ù������h
    frm_EXAM01v.ZOrder 0
    If frm_EXAM01v.Visible Then Spd_EXAM01v.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case KEY_F1
                KeyCode = 0
                If cmd_help.Visible = True And cmd_help.Enabled = True Then
                   cmd_help.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    
           Case KEY_F3
                KeyCode = 0
                If cmd_delete.Visible = True And cmd_delete.Enabled = True Then
                   cmd_delete.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    
           Case KEY_F4
                KeyCode = 0
                If cmd_add.Visible = True And cmd_add.Enabled = True Then
                   cmd_add.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    
           Case KEY_F7
                KeyCode = 0
                If cmd_previous.Visible = True And cmd_previous.Enabled = True Then
                   cmd_previous.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    
           Case KEY_F8
                KeyCode = 0
                If cmd_next.Visible = True And cmd_next.Enabled = True Then
                   cmd_next.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    
           Case KEY_ESCAPE
                KeyCode = 0
                If cmd_exit.Visible = True And cmd_exit.Enabled = True Then
                   cmd_exit.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    KeyPress KeyAscii
End Sub

Private Sub Form_Load()
    FormCenter Me
    Set_Property
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    If cmd_exit.Enabled Then Cmd_Exit_Click
End Sub

Private Sub Spd_EXAM01v_Click(ByVal Col As Long, ByVal Row As Long)
'��Column Header Click��, �̸����Ƨ�
    If Row = 0 Then SpreadSort Spd_EXAM01v, Col
End Sub

Private Sub Spd_EXAM01v_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    Me.MousePointer = HOURGLASS
    
    With Spd_EXAM01v
'���oV�e�����`���ƤΥثe���d�C
         G_AP_STATE = G_AP_STATE_UPDATE
         G_MaxRows# = .DataRowCnt
         G_ActiveRow# = Row

'Keep P-Key, ��Detail�e���ק���
         .Row = G_ActiveRow#
'         .Col = 1
'         StrCut .text, Space(1), G_A1601$, ""
         .Col = 1
         G_A1601$ = Trim$(.text)
    End With
    
'����V�e��, ������Detail�e��
    DoEvents
    Me.Hide
    frm_EXAM01.Show
    
    Me.MousePointer = Default
End Sub

Private Sub Spd_EXAM01v_GotFocus()
    SpreadGotFocus -1, Spd_EXAM01v.ActiveRow
End Sub

Private Sub Spd_EXAM01v_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_RETURN Then
       Spd_EXAM01v_DblClick CLng(Spd_EXAM01v.ActiveCol), CLng(Spd_EXAM01v.ActiveRow)
    End If
End Sub

Private Sub Spd_EXAM01v_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'��_�e�@��쪺�C��
    SpreadLostFocus -1, Row

'���ܷs��쪺�C��
    If NewCol > 0 Then
       SpreadGotFocus -1, NewRow
    End If
End Sub

