VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_TSR01q 
   Caption         =   "�ϥΰO���C�L"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TSR01q.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VsOcxLib.VideoSoftElastic Vse_Background 
      Height          =   2340
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   7155
      _Version        =   327680
      _ExtentX        =   12621
      _ExtentY        =   4128
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
      Picture         =   "TSR01q.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "TSR01q.frx":0326
      Begin ComctlLib.ProgressBar Prb_Percent 
         Height          =   210
         Left            =   1260
         TabIndex        =   15
         Top             =   1140
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Frame Fra_Help 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   4830
         TabIndex        =   19
         Top             =   585
         Visible         =   0   'False
         Width           =   825
         Begin FPSpread.vaSpread Spd_Help 
            Height          =   495
            Left            =   90
            OleObjectBlob   =   "TSR01q.frx":0342
            TabIndex        =   3
            Top             =   210
            Width           =   615
         End
      End
      Begin VB.ComboBox Cbo_A1501 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IntegralHeight  =   0   'False
         ItemData        =   "TSR01q.frx":0572
         Left            =   1620
         List            =   "TSR01q.frx":0574
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   135
         Width           =   3945
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
         Height          =   1125
         Left            =   60
         TabIndex        =   18
         Top             =   1110
         Width           =   5505
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
            Left            =   150
            TabIndex        =   8
            Text            =   " "
            Top             =   660
            Width           =   5235
         End
         Begin Threed.SSOption Opt_File 
            Height          =   360
            Left            =   3060
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   270
            Width           =   1140
            _Version        =   65536
            _ExtentX        =   2011
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "�ɮ�"
            ForeColor       =   0
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
         Begin Threed.SSOption Opt_Scrn 
            Height          =   360
            Left            =   1500
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   270
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "�ù����"
            ForeColor       =   0
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
            Left            =   150
            TabIndex        =   4
            Top             =   270
            Width           =   1200
            _Version        =   65536
            _ExtentX        =   2117
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "�L���"
            ForeColor       =   0
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
            Left            =   4200
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   270
            Width           =   1185
            _Version        =   65536
            _ExtentX        =   2090
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "Excel "
            ForeColor       =   0
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
      Begin VB.TextBox Txt_A1502e 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4050
         MaxLength       =   6
         TabIndex        =   2
         Top             =   615
         Width           =   1515
      End
      Begin VB.TextBox Txt_A1502s 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   1
         Top             =   615
         Width           =   1515
      End
      Begin Threed.SSCommand Cmd_Help 
         Height          =   405
         Left            =   5670
         TabIndex        =   9
         Top             =   120
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "���U F1"
         ForeColor       =   0
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
      Begin Threed.SSCommand Cmd_Exit 
         Height          =   405
         Left            =   5670
         TabIndex        =   11
         Top             =   1830
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "����Esc"
         ForeColor       =   0
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
      Begin Threed.SSCommand Cmd_Print 
         Height          =   405
         Left            =   5670
         TabIndex        =   10
         Top             =   570
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�C�LF6"
         ForeColor       =   0
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
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   17
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Lbl_A1502 
         Caption         =   "��ؽd��"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   16
         Top             =   645
         Width           =   1560
      End
      Begin VB.Label Lbl_A1501 
         Caption         =   "���q�O"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   14
         Top             =   225
         Width           =   1560
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   2340
      Width           =   7155
      _ExtentX        =   12621
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
Attribute VB_Name = "frm_TSR01q"
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

'���n�ܼ�
Dim m_FieldError%    '���ܼƦb�P�_���O�_���~, �����^�����줧�ʧ@
Dim m_ExitTrigger%   '���ܼƦb�P�_������O�_�QĲ�o, �N����ثe���b�B�z���@�~

'�۩w�ܼ�
Dim m_A1501Flag%
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

Private Sub Set_Property()
'�]�w��Form�����D,�r�ΤΦ�t
    Form_Property frm_TSR01q, G_Form_TSR01q$, G_Font_Name
    
'�]Form���Ҧ�Panel, Label�����D, �r�ΤΦ�t
    Label_Property Lbl_A1502, G_Pnl_A15023$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1501, G_Pnl_A1501$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Sign(2), G_Pnl_Dash$, G_Label_Color, G_Font_Size, G_Font_Name
    
'�]Form��Help Frame�����D, �r�ΤΦ�t
    Label_Property Fra_Help, "", COLOR_SKY, G_Font_Size, G_Font_Name
    Fra_Help.Visible = False
    
'�]Form���Ҧ�Text Box ���r�ΤΥi��H����
    Text_Property Txt_A1502s, 6, G_Font_Name
    Text_Property Txt_A1502e, 6, G_Font_Name
    Text_Property Txt_FileName, 60, G_Font_Name
    Txt_FileName.Visible = False
    
'�]Form���Ҧ�Combo Box ���r��
    ComboBox_Property Cbo_A1501, G_Font_Size, G_Font_Name
        
'�]Form���Ҧ�Frame�����D�Φr��
    Frame_Property Fra_PrintType, G_Pnl_PrtType$, G_Font_Size, G_Font_Name
    
'�]Form���Ҧ�Option�����D�Φr��
    Option_Property Opt_Printer, G_Pnl_Printer$, G_Font_Size, G_Font_Name
    Option_Property Opt_Scrn, G_Pnl_Screen$, G_Font_Size, G_Font_Name
    Option_Property Opt_File, G_Pnl_File$, G_Font_Size, G_Font_Name
    Option_Property Opt_Excel, G_Pnl_Excel$, G_Font_Size, G_Font_Name
        
'�]Form���Ҧ�Command�����D�Φr��
    Command_Property Cmd_Help, G_CmdHelp, G_Font_Name
    Command_Property Cmd_Print, G_CmdPrint, G_Font_Name
    Command_Property Cmd_exit, G_CmdExit, G_Font_Name
    
'�H�U���зǫ��O, ���o�ק�
    ProgressBar_Property Prb_Percent
    VSElastic_Property Vse_Background
    StatusBar_ProPerty Sts_MsgLine
End Sub

Sub PrePare_ComboBox()
    CBO_A1501_Prepare
End Sub

Private Function CheckRoutine_A1502() As Boolean
     CheckRoutine_A1502 = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim$(Txt_A1502e) = "" Then Txt_A1502e = Txt_A1502s
    
    If Not CheckDataRange(Sts_MsgLine, Trim$(Txt_A1502s), Trim$(Txt_A1502e)) Then
       If ActiveControl.TabIndex = Txt_A1502e.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A1502e.TabIndex
       Else
          m_FieldError% = Txt_A1502s.TabIndex
          Txt_A1502s.SetFocus
       End If
       Exit Function
    End If
       
    CheckRoutine_A1502 = True
End Function

Private Function CheckRoutine_FileName() As Boolean
    CheckRoutine_FileName = True
    
    If Opt_Printer.Value = True Then Exit Function
    If Opt_Scrn.Value = True Then Exit Function
    
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



Private Sub DataPrepare_A15(Txt As TextBox)
Dim A_Sql$
Dim A_A1501$
    
    
    Me.MousePointer = HOURGLASS
    StrCut Cbo_A1501.text, Space(1), A_A1501$, ""

'�}�_�ɮ�
    A_Sql$ = "Select A1502,A1503 from A15"
    A_Sql$ = A_Sql$ & " where A1501 = '" & A_A1501$ & "'"
    A_Sql$ = A_Sql$ & " order by A1501"
    
    CreateDynasetODBC DB_ARTHGUI, DY_A15, A_Sql$, "DY_A15", True
    
    If DY_A15.BOF And DY_A15.EOF Then
       Me.MousePointer = Default
       Sts_MsgLine.Panels(1) = G_NoReference
       Exit Sub
    End If
    
    With Spd_Help

'�]�w���U����������ݩ�
         .UnitType = 2
         Spread_Property Spd_Help, 0, 1, WHITE, G_Font_Size, G_Font_Name
         Spread_Col_Property Spd_Help, 1, TextWidth("X") * 8, G_Pnl_A1502$
         Spread_DataType_Property Spd_Help, 1, SS_CELL_TYPE_EDIT, "", "", 6
         .Row = -1
         .Col = -1: .Lock = True
         .Col = 1: .TypeHAlign = 2
    
'�N����\�JSpread��
         Do While Not DY_A15.EOF
            .MaxRows = .MaxRows + 1
            .Row = Spd_Help.MaxRows
            .Col = 1
            .text = Trim(DY_A15.Fields("A1502") & "") & _
                    Trim(DY_A15.Fields("A1503") & "")
            DY_A15.MoveNext
         Loop
    
'�]�w���U��������ܦ�m
         SetHelpWindowPos Fra_Help, Spd_Help, 1300, 120, 4265, 2085
         .Tag = Txt.TabIndex
         .SetFocus
    End With
    
    Me.MousePointer = Default
End Sub

Private Sub CBO_A1501_Prepare()
On Local Error GoTo MyError
Dim A_Sql$

'���M��Combo Box���e
    Cbo_A1501.Clear
    
'�}�_�ɮ�
    A_Sql$ = "Select A0101,A0102 From A01 ORDER BY A0101"
    CreateDynasetODBC DB_ARTHGUI, DY_A01, A_Sql$, "DY_A01", True

'�N����\�JCombo Box��
    Do While Not DY_A01.EOF
       Cbo_A1501.AddItem Format(Trim$(DY_A01.Fields("A0101") & ""), "!@@@") & Trim$(DY_A01.Fields("A0102") & "")
       DY_A01.MoveNext
    Loop

'�YCombo Box�������, ���b�Ĥ@��
    If Cbo_A1501.ListCount > 0 Then Cbo_A1501.ListIndex = 0
    Exit Sub
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    If Not CheckRoutine_A1502() Then Exit Function
    If Not CheckRoutine_FileName() Then Exit Function
    DoEvents
    IsAllFieldsCheck = True
End Function

Private Sub KeepFieldsValue()
    G_A1502s$ = Trim$(Txt_A1502s)
    G_A1502e$ = Trim$(Txt_A1502e)
    StrCut Cbo_A1501.text, Space(1), G_A1501$, G_A1501n$
    G_OutFile = Trim$(Txt_FileName)
    If Opt_Printer.Value Then G_PrintSelect = G_Print2Printer
    If Opt_Scrn.Value Then G_PrintSelect = G_Print2Screen
    If Opt_File.Value Then G_PrintSelect = G_Print2File
    If Opt_Excel.Value Then G_PrintSelect = G_Print2Excel
End Sub


Private Sub OpenMainFile()
On Local Error GoTo MY_Error
Dim A_Sql$
Dim A_A1502e$

    A_Sql$ = "SELECT A1502,A1503,A1504,A1505,A1507,A1508,A1510,A1512,A1302 FROM A15"
    A_Sql$ = A_Sql$ & " INNER JOIN A13"
    A_Sql$ = A_Sql$ & " ON A13.A1301 = A15.A1507"
    A_Sql$ = A_Sql$ & " WHERE A1501='" & G_A1501$ & "'"
    
    If G_A1502s$ <> "" Then
        A_Sql$ = A_Sql$ & " and A1502+A1503>='" & G_A1502s$ & "'"
'        A_Sql$ = A_Sql$ & " and not (A1502<'" & Mid$(G_A1502s$, 1, 4) & "'"
'        A_Sql$ = A_Sql$ & " or A1502='" & Mid$(G_A1502s$, 1, 4) & "'"
'        A_Sql$ = A_Sql$ & " and A1503<'" & Mid$(G_A1502s$, 5) & "')"
    End If
    If G_A1502e$ <> "" Then
        A_A1502e$ = G_A1502e$ & "Z"
        A_Sql$ = A_Sql$ & " and A1502+A1503<='" & A_A1502e$ & "'"
'        A_Sql$ = A_Sql$ & " and not (A1502>'" & Mid$(A_A1502e$, 1, 4) & "'"
'        A_Sql$ = A_Sql$ & " or A1502='" & Mid$(A_A1502e$, 1, 4) & "'"
'        A_Sql$ = A_Sql$ & " and A1503>'" & Mid$(A_A1502e$, 5) & "')"
    End If
    
    A_Sql$ = A_Sql$ & " order by A1501,A1502,A1503"
    CreateDynasetODBC DB_ARTHGUI, DY_A15, A_Sql$, "DY_A15", True
'    A_Sql$ = A_Sql$ & " order by A0901,A0902,A0903"
'    CreateDynasetODBC DB_ARTHGUI, DY_A09, A_Sql$, "DY_A09", True

'!!!@bad one
'Open A13 on A1301 = A1507
'    A_Sql$ = "Select A1302 From A13 As a"
'    A_Sql$ = A_Sql$ & " INNER JOIN A15 As b"
'    A_Sql$ = A_Sql$ & " On a.A1301 = b.A1507"
'    A_Sql$ = A_Sql$ & " Order by A1302"
'    CreateDynasetODBC DB_ARTHGUI, DY_A13, A_Sql$, "DY_A13", True
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Sub Cbo_A1501_Click()

    If m_A1501Flag% Then Exit Sub
    
'�Y����Ƥ��e���I��B�ܰʮ�, �Ҷ��B�z���ƶ�
    If Trim(Cbo_A1501.text) <> Trim(Cbo_A1501.Tag) Then
       Cbo_A1501.Tag = Cbo_A1501.text
    End If
End Sub

Private Sub Cbo_A1501_DropDown()
    DoEvents
    
    m_A1501Flag% = True

'�N�ثeCombo Box�W���N�XKeep�U��
    Dim A_A1501$
    StrCut Cbo_A1501.text, Space(1), A_A1501$, ""
    
'���s�ǳƦ�Combo Box�����e
    CBO_A1501_Prepare
    
'�NCombo Box�W��ListIndex���VKeep�U�Ӫ����
    CboStrCut Cbo_A1501, A_A1501$, Space(1)
    
    m_A1501Flag% = False
End Sub

Private Sub Cbo_A1501_GotFocus()
    TextGotFocus
End Sub

Private Sub Cbo_A1501_LostFocus()
    TextLostFocus
End Sub

Private Sub Cmd_Exit_Click()
'�����ثe����,���X��L�B�z�{��
    m_ExitTrigger% = True
    CloseFileDB
    End
End Sub

Private Sub Cmd_Help_Click()
Dim a$

'�бNTSR01q�אּ��Form�W�r�Y�i, ��l���зǫ��O, ���o�ק�
    a$ = "notepad " + G_Help_Path + "TSR01q.HLP"
    retcode = Shell(a$, 4)
End Sub

Private Sub Cmd_Print_Click()
    Me.MousePointer = HOURGLASS
    Cmd_Print.Enabled = False

'�ˮ���쥿�T��
    If Not IsAllFieldsCheck() Then
       Me.MousePointer = Default
       Cmd_Print.Enabled = True
       Exit Sub
    End If

'Keep�@���ܼƨѦL���
    KeepFieldsValue
    
'�B�z�C�L�ʧ@
    Sts_MsgLine.Panels(1) = G_Process
    OpenMainFile
    If DY_A15.BOF And DY_A15.EOF Then

'�L��Ƥ����C�L
       Sts_MsgLine.Panels(1) = G_NoQueryData
    Else

'����RepSet Form������,���|Ĳ�oForm_Activate
       If G_PrintSelect = G_Print2Printer Then
          G_FormFrom$ = "RptSet"
       End If
       
'�}�l�C�L����
      If Not Opt_Scrn.Value Then
         PrePare_Data frm_TSR01q, Prb_Percent, Prb_Percent, m_ExitTrigger%
      Else
         DoEvents
         Me.Hide
         frm_TSR01.Show
         Sts_MsgLine.Panels(1) = G_PrintOk
      End If
    End If
    Cmd_Print.Enabled = True
    Me.MousePointer = Default
End Sub

Private Sub Form_Activate()
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    Me.Refresh
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
       PrePare_ComboBox
       G_AP_STATE = G_AP_STATE_NORMAL   '�]�w�@�~���A
       Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    End If
    
    '�NForm��m��ù������h
    frm_TSR01q.ZOrder 0
    If frm_TSR01q.Visible Then Txt_A1502s.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case KEY_DELETE
                If TypeOf ActiveControl Is ComboBox Then
                   ActiveControl.ListIndex = -1
                End If
           Case KEY_F1
                If ActiveControl.TabIndex = Txt_A1502s.TabIndex Then Exit Sub
                If ActiveControl.TabIndex = Txt_A1502e.TabIndex Then Exit Sub
                KeyCode = 0
                If Cmd_Help.Visible = True And Cmd_Help.Enabled = True Then
                   Cmd_Help.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F6
                KeyCode = 0
                If Cmd_Print.Visible = True And Cmd_Print.Enabled = True Then
                   Cmd_Print.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_ESCAPE
                KeyCode = 0
                If Cmd_exit.Visible = True And Cmd_exit.Enabled = True Then
                   Cmd_exit.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
'�D�ʱN��ƿ�J�Ѥp�g�ର�j�g
'  �Y���Y����줣�ݭn�ഫ��, �����H���L
   If ActiveControl.TabIndex <> Txt_A1502s.TabIndex And _
   ActiveControl.TabIndex <> Txt_A1502e.TabIndex Then _
   GoTo Form_KeyPress_A
   'If ActiveControl.TabIndex = txt_yyy.TabIndex Then GoTo Form_KeyPress_A
   'If ActiveControl.TabIndex = txt_zzz.TabIndex Then GoTo Form_KeyPress_A
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Form_KeyPress_A:
    KeyPress KeyAscii           'Enter�ɦ۰ʸ���U�@���
End Sub

Private Sub Form_Load()
    FormCenter Me                     '�e���m���B�z
    Set_Property                      '�]�w���e��������ݩ�
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'�YUser��������Windows, �ӥ��{�����b����, ���{���|���߰ݬO�_�n�������ۤv?
'�H�U���зǫ��O, ���o�ק�
    
    Dim MSG

    If UnloadMode > 0 Then
       MSG = GetSIniStr("PgmMsg", "g_gui_run")   ' If exiting the application.
    Else
       CloseFileDB
       End
    End If
    
' If user clicks the 'No' button, stop QueryUnload.
    If MsgBox(MSG, 36, Me.Caption) = 7 Then
       Cancel = True
    Else
       CloseFileDB
       End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cmd_Exit_Click
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

Private Sub Opt_Scrn_Click(Value As Integer)
    Txt_FileName.Visible = False
End Sub

Private Sub Spd_Help_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim A_Code$

    Me.MousePointer = HOURGLASS
    
'KEEP�ۻ��U�����I�諸���
    With Spd_Help
         .Row = .ActiveRow
         .Col = 1
         A_Code$ = Trim(.text)
    
'�NKEEP����Ʊa�J�e��
         Select Case Val(.Tag)
           Case Txt_A1502s.TabIndex
                Txt_A1502s = A_Code$
           Case Txt_A1502e.TabIndex
                Txt_A1502e = A_Code$
         End Select
    End With
    
'���û��U����
    Fra_Help.Visible = False
    
    Me.MousePointer = Default
End Sub

Private Sub Spd_Help_GotFocus()
    SpreadGotFocus Spd_Help.ActiveCol, Spd_Help.ActiveRow
End Sub

Private Sub Spd_Help_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_RETURN Then
       Spd_Help_DblClick Spd_Help.ActiveCol, Spd_Help.ActiveRow
    End If
End Sub

Private Sub Spd_Help_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'�зǫ��O,���o�ק�
    SpreadLostFocus Col, Row
    If NewCol > 0 Then SpreadGotFocus NewCol, NewRow
End Sub

Private Sub Spd_Help_LostFocus()
    Fra_Help.Visible = False
    Select Case Val(Spd_Help.Tag)
      Case Txt_A1502s.TabIndex
           Txt_A1502s.SetFocus
      Case Txt_A1502e.TabIndex
           Txt_A1502e.SetFocus
    End Select
End Sub

Private Sub Txt_A1502e_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A1502e_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A1502e_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A1502e_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If KeyCode = KEY_F1 Then DataPrepare_A15 Txt_A1502e
End Sub

Private Sub Txt_A1502e_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If Fra_Help.Visible = True Then Exit Sub
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A1502e.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A1502()
End Sub

Private Sub Txt_A1502s_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A1502s_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A1502s_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A1502s_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If KeyCode = KEY_F1 Then DataPrepare_A15 Txt_A1502s
End Sub

Private Sub Txt_A1502s_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If Fra_Help.Visible = True Then Exit Sub
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A1502s.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A1502()
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

