VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_TSR01 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "�ϥΰO���C�L"
   ClientHeight    =   6525
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "TSR01.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6525
   ScaleWidth      =   9480
   Begin VsOcxLib.VideoSoftElastic Vse_Background 
      Height          =   6150
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   9480
      _Version        =   327680
      _ExtentX        =   16722
      _ExtentY        =   10848
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Align           =   5
      BevelOuter      =   6
      Picture         =   "TSR01.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "TSR01.frx":0326
      Begin FPSpread.vaSpread Spd_TSR01 
         Height          =   4635
         Left            =   60
         OleObjectBlob   =   "TSR01.frx":0342
         TabIndex        =   0
         Top             =   630
         Width           =   7875
      End
      Begin ComctlLib.ProgressBar Prb_Percent 
         Height          =   210
         Left            =   1290
         TabIndex        =   12
         Top             =   4800
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Frame Fra_PrintType 
         Caption         =   "�C�L�覡"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   13
         Top             =   5310
         Width           =   7875
         Begin VB.TextBox Txt_FileName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3330
            TabIndex        =   4
            Text            =   " "
            Top             =   270
            Width           =   4440
         End
         Begin Threed.SSOption Opt_File 
            Height          =   360
            Left            =   1260
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   270
            Width           =   1068
            _Version        =   65536
            _ExtentX        =   1884
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "�ɮ�"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption Opt_Printer 
            Height          =   360
            Left            =   156
            TabIndex        =   1
            Top             =   276
            Width           =   1068
            _Version        =   65536
            _ExtentX        =   1884
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "�L���"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption Opt_Excel 
            Height          =   360
            Left            =   2340
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   270
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "Excel "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSCommand Cmd_Help 
         Height          =   405
         Left            =   8010
         TabIndex        =   5
         Top             =   630
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "���U F1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Cmd_Print 
         Height          =   405
         Left            =   8010
         TabIndex        =   6
         Top             =   1080
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�C�LF6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Cmd_Next 
         Height          =   405
         Left            =   8010
         TabIndex        =   8
         Top             =   1980
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "���� F8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Cmd_Previous 
         Height          =   405
         Left            =   8010
         TabIndex        =   7
         Top             =   1530
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�e�� F7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Cmd_exit 
         Height          =   405
         Left            =   8010
         TabIndex        =   9
         Top             =   5640
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "����Esc"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel Pnl_A1501 
         Height          =   390
         Left            =   1125
         TabIndex        =   15
         Top             =   135
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   15790320
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel Pnl_A1501n 
         Height          =   390
         Left            =   1575
         TabIndex        =   16
         Top             =   135
         Width           =   1860
         _Version        =   65536
         _ExtentX        =   3281
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   15790320
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Lbl_A1501 
         Caption         =   "���q�O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   14
         Top             =   180
         Width           =   1635
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6150
      Width           =   9480
      _ExtentX        =   16722
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
Attribute VB_Name = "frm_TSR01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'�b���B�w�q���Ҧ��ܼ�, �@�ߥHM�}�Y, �pM_AAA$, M_BBB#, M_CCC&
'�B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
' $: ��r
' #: �Ҧ��Ʀr�B��(���B�μƶq)
' &: �{���j���ܼ�
' %: ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
' �ť�: �N��VARIENT, �ʺA�ܼ�


'�۩w�ܼ�
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

'���n�ܼ�
Dim m_FieldError%    '���ܼƦb�P�_���O�_���~, �����^�����줧�ʧ@
Dim m_ExitTrigger%   '���ܼƦb�P�_������O�_�QĲ�o, �N����ثe���b�B�z���@�~

Private Sub Set_Property()
'�]�w��Form�����D,�r�ΤΦ�t
    Form_Property frm_TSR01, G_Form_TSR01$, G_Font_Name

'�]Form���Ҧ�Panel, Label�����D, �r�ΤΦ�t
    Label_Property Lbl_A1501, G_Pnl_A1501$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Pnl_A1501, G_A1501$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Pnl_A1501n, G_A1501n$, G_Label_Color, G_Font_Size, G_Font_Name
    
'�]Form���Ҧ�TextBox���r�ΤΥi��J����
    Text_Property Txt_FileName, 60, G_Font_Name
    Txt_FileName.Visible = False
        
'�]Form���Ҧ�Frame�����D�Φr��
    Frame_Property Fra_PrintType, G_Pnl_PrtType$, G_Font_Size, G_Font_Name
    
'�]Form���Ҧ�Option�����D�Φr��
    Option_Property Opt_Printer, G_Pnl_Printer$, G_Font_Size, G_Font_Name
    Option_Property Opt_File, G_Pnl_File$, G_Font_Size, G_Font_Name
    Option_Property Opt_Excel, G_Pnl_Excel$, G_Font_Size, G_Font_Name
    
'�]Form���Ҧ�Command�����D�Φr��
    Command_Property Cmd_Help, G_CmdHelp, G_Font_Name
    Command_Property Cmd_Print, G_CmdPrint, G_Font_Name
    Command_Property Cmd_exit, G_CmdExit, G_Font_Name
    Command_Property Cmd_Previous, G_CmdPrvPage, G_Font_Name
    Command_Property Cmd_Next, G_CmdNxtPage, G_Font_Name
    
'�]Form��Spread���ݩ�
    Set_Spread_Property

'�H�U���зǫ��O, ���o�ק�
    ProgressBar_Property Prb_Percent
    VSElastic_Property Vse_Background
    StatusBar_ProPerty Sts_MsgLine
End Sub

Private Sub Set_Spread_Property()
    Spd_TSR01.UnitType = 2

'�]�w��Spread�����Ƥ�����
    Spread_Property Spd_TSR01, 0, 7, WHITE, G_Font_Size, G_Font_Name

'�]�w��Spread���U����D����ܼe��, 0�N�����줣���
    Spread_Col_Property Spd_TSR01, 1, TextWidth("X") * 8, G_Pnl_A1502$
    Spread_Col_Property Spd_TSR01, 2, TextWidth("X") * 12, G_Pnl_A1505$
    Spread_Col_Property Spd_TSR01, 3, TextWidth("X") * 8, G_Pnl_A1504$
    Spread_Col_Property Spd_TSR01, 4, TextWidth("X") * 12, G_Pnl_A1507$
    Spread_Col_Property Spd_TSR01, 5, TextWidth("X") * 8, G_Pnl_A1510$
    Spread_Col_Property Spd_TSR01, 6, TextWidth("X") * 8, G_Pnl_A1512$
    Spread_Col_Property Spd_TSR01, 7, TextWidth("X") * 15, G_Pnl_A1508$
    
'�]�w��Spread���U���ݩʤ���ܦr��
  'SS_CELL_TYPE_EDIT        = ��r�i��J
  'SS_CELL_TYPE_FLOAT       = �Ʀr�i��J
  'SS_CELL_TYPE_STATIC_TEXT = �����
  'SS_CELL_TYPE_CHECKBOX    = �I�ﶵ��
    Spread_DataType_Property Spd_TSR01, 1, SS_CELL_TYPE_EDIT, "", "", 6
    Spread_DataType_Property Spd_TSR01, 2, SS_CELL_TYPE_EDIT, "", "", 40
    Spread_DataType_Property Spd_TSR01, 3, SS_CELL_TYPE_EDIT, "", "", 20
    Spread_DataType_Property Spd_TSR01, 4, SS_CELL_TYPE_EDIT, "", "", 20
    Spread_DataType_Property Spd_TSR01, 5, SS_CELL_TYPE_EDIT, "", "", 20
    Spread_DataType_Property Spd_TSR01, 6, SS_CELL_TYPE_EDIT, "", "", 20
    Spread_DataType_Property Spd_TSR01, 7, SS_CELL_TYPE_FLOAT, "-999999999.99", "999999999.99", 2
    Spd_TSR01.EditEnterAction = SS_CELL_EDITMODE_EXIT_NONE

'�T�w�V�k���ʮ�, �ҭ�����
    Spd_TSR01.ColsFrozen = 2

'���Spread���i�ק�
    Spd_TSR01.Row = -1
    Spd_TSR01.Col = -1
    Spd_TSR01.Lock = True

'�w�q�Y����m����m���]�w 0:���a  1:�k�a  2:�m��
    Spd_TSR01.Col = 1: Spd_TSR01.TypeHAlign = 2
    Spd_TSR01.Col = 2: Spd_TSR01.TypeHAlign = 2
    Spd_TSR01.Col = 3: Spd_TSR01.TypeHAlign = 2
    Spd_TSR01.Col = 4: Spd_TSR01.TypeHAlign = 2
    Spd_TSR01.Col = 5: Spd_TSR01.TypeHAlign = 2
    Spd_TSR01.Col = 6: Spd_TSR01.TypeHAlign = 2
    Spd_TSR01.Col = 7: Spd_TSR01.TypeHAlign = 1
End Sub

Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    If Not CheckRoutine_FileName() Then Exit Function
    IsAllFieldsCheck = True
End Function

Private Function CheckRoutine_FileName() As Boolean
    CheckRoutine_FileName = True
    
    If Opt_Printer.Value = True Then Exit Function
    
'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�Y����ɮצC�L,���Y�ť�,�h�a�X Default Value
    If Opt_File.Value Then
        SetDefaultFileName Txt_FileName, G_Print2File
    ElseIf Opt_Excel.Value Then
        SetDefaultFileName Txt_FileName, G_Print2Excel
    End If
    DoEvents
    
'�ˮָ��|�O�_�s�b
    Dim a$
    a$ = Trim(Txt_FileName)
    If Not CheckDirectoryExist(a$) Then
       CheckRoutine_FileName = False
       Sts_MsgLine.Panels(1) = G_PathNotFound$
       m_FieldError% = Txt_FileName.TabIndex
       Txt_FileName.SetFocus
    End If
End Function

Sub KeepFieldsValue()
    G_OutFile = Trim$(Txt_FileName)
    If Opt_Printer.Value Then G_PrintSelect = G_Print2Printer
    If Opt_File.Value Then G_PrintSelect = G_Print2File
    If Opt_Excel.Value Then G_PrintSelect = G_Print2Excel
End Sub

Private Sub Cmd_Exit_Click()
    Unload Me
End Sub

Private Sub Cmd_Help_Click()
Dim a$

'�бNTSR01�אּ��Form�W�r�Y�i, ��l���зǫ��O, ���o�ק�
    a$ = "notepad " + G_Help_Path + "TSR01.HLP"
    retcode = Shell(a$, 4)
End Sub

Private Sub Cmd_Next_Click()
    Cmd_Next.Enabled = False
    Spd_TSR01.SetFocus
    SendKeys "{PgDn}"
    DoEvents
    Cmd_Next.Enabled = True
End Sub

Private Sub Cmd_Print_Click()
    Me.MousePointer = HOURGLASS
    Cmd_Print.Enabled = False

'�ˮ���쥿�T��
    If IsAllFieldsCheck() = False Then
       Me.MousePointer = Default
       Cmd_Print.Enabled = True
       Exit Sub
    End If

'Keep�@���ܼƨѦL���
    KeepFieldsValue
    
'�B�z�C�L�ʧ@
    Sts_MsgLine.Panels(1) = G_Process

'����RepSet Form������,���|Ĳ�oForm_Activate
    If G_PrintSelect = G_Print2Printer Then
       G_FormFrom$ = "RptSet"
    End If
       
'�}�l�C�L����
    PrePare_Data frm_TSR01, Prb_Percent, Prb_Percent, m_ExitTrigger%
    
    Cmd_Print.Enabled = True
    Me.MousePointer = Default
End Sub

Private Sub Cmd_Previous_Click()
    Cmd_Previous.Enabled = False
    Spd_TSR01.SetFocus
    SendKeys "{PgUp}"
    DoEvents
    Cmd_Previous.Enabled = True
End Sub

Private Sub Form_Activate()
    Me.MousePointer = HOURGLASS
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    
'Initial Form�������n�ܼ�
    m_FieldError% = -1
    m_ExitTrigger% = False
         
'�P�_�O�_�Ѩ�L���U�e���^��, �ӫD��������
    If Trim(G_FormFrom$) <> "" Then
       G_FormFrom$ = ""
       '.....                '�[�J�ҭn�]�w���ʧ@
       '.....
       Exit Sub
    Else
       '.....                '�Ĥ@������ɤ��ǳưʧ@
       '.....
       Sts_MsgLine.Panels(1) = G_Process
       Cmd_Print.Enabled = False
       PrePare_Data frm_TSR01, Prb_Percent, Spd_TSR01, m_ExitTrigger%
       If m_ExitTrigger% Then Exit Sub
       Cmd_Print.Enabled = True
    End If
    
    '�NForm��m��ù������h
    frm_TSR01.ZOrder 0
    If frm_TSR01.Visible Then Spd_TSR01.SetFocus
    Me.MousePointer = Default
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case KEY_F1
           KeyCode = 0
           If Cmd_Help.Visible And Cmd_Help.Enabled Then
              Cmd_Help.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
        
      Case KEY_F6
           KeyCode = 0
           If Cmd_Print.Visible And Cmd_Print.Enabled Then
              Cmd_Print.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
           
      Case KEY_F7
           KeyCode = 0
           If Cmd_Previous.Visible And Cmd_Previous.Enabled Then
              Cmd_Previous.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
           
      Case KEY_F8
           KeyCode = 0
           If Cmd_Next.Visible And Cmd_Next.Enabled Then
              Cmd_Next.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
        
      Case KEY_ESCAPE
           KeyCode = 0
           If Cmd_exit.Visible And Cmd_exit.Enabled Then
              Cmd_exit.SetFocus
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
'�����ثe����,���X��L�B�z�{��
    m_ExitTrigger% = True
    G_FormFrom$ = "TSR01"
    frm_TSR01q.Show
End Sub

Private Sub Opt_Excel_Click(Value As Integer)
    SetDefaultFileName Txt_FileName, G_Print2Excel
End Sub

Private Sub Opt_File_Click(Value As Integer)
    SetDefaultFileName Txt_FileName, G_Print2File
End Sub

Private Sub Opt_Printer_Click(Value As Integer)
    Txt_FileName.Visible = False
End Sub

Private Sub Spd_TSR01_Click(ByVal Col As Long, ByVal Row As Long)
'��Column Heading Click��, �̸����Ƨ�
    If Row = 0 Then SpreadSort Spd_TSR01, Col
End Sub

Private Sub Spd_TSR01_GotFocus()
    SpreadGotFocus -1, CLng(Spd_TSR01.ActiveRow)
End Sub

Private Sub Spd_TSR01_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'��_�e�@��쪺�C��
    SpreadLostFocus -1, Row

'���ܷs��쪺�C��
    If NewCol > 0 Then
       SpreadGotFocus -1, NewRow
    End If
End Sub

Private Sub Txt_FileName_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_FileName_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If TypeOf ActiveControl Is SSCommand Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_FileName.TabIndex Then Exit Sub

    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_FileName()
End Sub

