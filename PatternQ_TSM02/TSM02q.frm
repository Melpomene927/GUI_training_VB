VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_TSM02q 
   Caption         =   "�|�p��ظ�Ƭd��"
   ClientHeight    =   2580
   ClientLeft      =   3270
   ClientTop       =   2685
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TSM02q.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   6225
   Begin VsOcxLib.VideoSoftElastic Vse_Background 
      Height          =   2205
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   6225
      _Version        =   327680
      _ExtentX        =   10980
      _ExtentY        =   3889
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
      Picture         =   "TSM02q.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "TSM02q.frx":0326
      Begin VB.TextBox Txt_A0201e 
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
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   1
         Top             =   90
         Width           =   1395
      End
      Begin VB.TextBox Txt_A0201s 
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
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   0
         Top             =   90
         Width           =   1395
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
         Left            =   3825
         TabIndex        =   12
         Top             =   1665
         Visible         =   0   'False
         Width           =   855
         Begin FPSpread.vaSpread Spd_Help 
            Height          =   495
            Left            =   90
            OleObjectBlob   =   "TSM02q.frx":0342
            TabIndex        =   13
            Top             =   210
            Width           =   615
         End
      End
      Begin VB.TextBox Txt_A02005e 
         DataField       =   "z"
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
         Left            =   3240
         TabIndex        =   7
         Top             =   1380
         Width           =   1395
      End
      Begin VB.TextBox Txt_A02005s 
         DataField       =   "z"
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
         Left            =   1410
         TabIndex        =   6
         Top             =   1380
         Width           =   1395
      End
      Begin VB.TextBox Txt_A02001e 
         DataField       =   "z"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox Txt_A02001s 
         DataField       =   "z"
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
         Left            =   1395
         TabIndex        =   4
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox Txt_A0208e 
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
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   3
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox Txt_A0208s 
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
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   2
         Top             =   540
         Width           =   1395
      End
      Begin Threed.SSCommand Cmd_Help 
         Height          =   405
         Left            =   4740
         TabIndex        =   8
         Top             =   90
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "���U F1"
         ForeColor       =   0
      End
      Begin Threed.SSCommand Cmd_Add 
         Height          =   405
         Left            =   4740
         TabIndex        =   10
         Top             =   990
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�s�WF4"
         ForeColor       =   0
      End
      Begin Threed.SSCommand Cmd_Exit 
         Height          =   405
         Left            =   4740
         TabIndex        =   11
         Top             =   1680
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "����Esc"
         ForeColor       =   0
      End
      Begin Threed.SSCommand Cmd_Ok 
         Height          =   405
         Left            =   4740
         TabIndex        =   9
         Top             =   540
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�d��F2"
         ForeColor       =   0
      End
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "��"
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   3
         Left            =   2850
         TabIndex        =   23
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Lbl_A02005 
         Caption         =   "�ק���"
         DataField       =   "z"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   45
         TabIndex        =   22
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "��"
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   2
         Left            =   2850
         TabIndex        =   21
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Lbl_A02001 
         Caption         =   "���ɤ��"
         DataField       =   "z"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   60
         TabIndex        =   20
         Top             =   1020
         Width           =   1380
      End
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "��"
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   1
         Left            =   2850
         TabIndex        =   19
         Top             =   570
         Width           =   375
      End
      Begin VB.Label Lbl_A0201 
         Caption         =   "�����O"
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
         Left            =   45
         TabIndex        =   18
         Top             =   135
         Width           =   1380
      End
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "��"
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
         Index           =   0
         Left            =   2880
         TabIndex        =   17
         Top             =   150
         Width           =   300
      End
      Begin VB.Label Lbl_A0208 
         Caption         =   "���ߤ��"
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
         Left            =   45
         TabIndex        =   16
         Top             =   585
         Width           =   1380
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   2205
      Width           =   6225
      _ExtentX        =   10980
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
Attribute VB_Name = "frm_TSM02q"
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
'Dim m_A4101Flag%
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

'���n�ܼ�
Dim m_FieldError%    '���ܼƦb�P�_���O�_���~, �����^�����줧�ʧ@
Dim m_ExitTrigger%   '���ܼƦb�P�_������O�_�QĲ�o, �N����ثe���b�B�z���@�~


Private Sub Set_Property()
'�]�w��Form�����D,�r�ΤΦ�t
    Form_Property Me, G_Form_TSM02q, G_Font_Name
    
'�]Form���Ҧ�Panel, Label�����D, �r�ΤΦ�t
    Label_Property Lbl_A0201, G_Pnl_A0201$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0208, G_Pnl_A0208$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A02001, G_Pnl_A02001$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A02005, G_Pnl_A02005$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Sign(0), G_Pnl_Dash$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Sign(1), G_Pnl_Dash$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Sign(2), G_Pnl_Dash$, G_Label_Color, G_Font_Size, G_Font_Name
    
'�]Form��Help Frame�����D, �r�ΤΦ�t
    Label_Property Fra_Help, "", COLOR_SKY, G_Font_Size, G_Font_Name
    Fra_Help.Visible = False
   
'�]Form���Ҧ�Text Box ���r�ΤΥi��H����
    Text_Property Txt_A0201s, 6, G_Font_Name
    Text_Property Txt_A0201e, 6, G_Font_Name
    Text_Property Txt_A0208s, 8, G_Font_Name
    Text_Property Txt_A0208e, 8, G_Font_Name
    Text_Property Txt_A02001s, 8, G_Font_Name
    Text_Property Txt_A02001e, 8, G_Font_Name
    Text_Property Txt_A02005s, 8, G_Font_Name
    Text_Property Txt_A02005e, 8, G_Font_Name
    
'�]Form���Ҧ�Combo Box ���r��
'    ComboBox_Property Cbo_A1501, G_Font_Size, G_Font_Name
    
'�]Form���Ҧ�Command�����D�Φr��
    Command_Property Cmd_Help, G_CmdHelp, G_Font_Name
    Command_Property Cmd_Ok, G_CmdSearch, G_Font_Name
    Command_Property Cmd_Add, G_CmdAdd, G_Font_Name
    Command_Property Cmd_Exit, G_CmdExit, G_Font_Name
    
'�H�U���зǫ��O, ���o�ק�
    VSElastic_Property Vse_Background
    StatusBar_ProPerty Sts_MsgLine
End Sub

Private Sub DataPrepare_A02(Txt As TextBox)
Dim A_Sql$      'SQL Message
Dim A_A0201$    'PKey of A02 (�����N�X)
    
    
    Me.MousePointer = HOURGLASS
    
    A_A0201$ = Trim(Txt)    'parameter is

    '�}�_�ɮ�
    'concate SQL Message
    A_Sql$ = "Select A0201, A0202 A From A02"
    
    'generate wildcard compare SQL Statement
    A_Sql$ = A_Sql$ & " Where A0201 Like '" & A_A0201 & GetLikeStr(DB_ARTHGUI, True) & "'"
    A_Sql$ = A_Sql$ & " Order by A0201"
    
    'Old statements that belongs to PATTERNQ(A15)
'    If Len(A_A1502$) > 4 Then
'       A_Sql$ = A_Sql$ & " and A1502='" & Mid$(A_A1502$, 1, 4) & "'"
'       A_Sql$ = A_Sql$ & " and A1503 Like '" & Mid$(A_A1502$, 5) & GetLikeStr(DB_ARTHGUI, True) & "'"
'    Else
'       A_Sql$ = A_Sql$ & " and A1502 Like '" & A_A1502$ & GetLikeStr(DB_ARTHGUI, True) & "'"
'    End If
    
    'open dynaset of A02
    CreateDynasetODBC DB_ARTHGUI, DY_A02, A_Sql$, "DY_A02", True
    If DY_A02.BOF And DY_A02.EOF Then
       Me.MousePointer = Default
       Sts_MsgLine.Panels(1) = G_NoReference
       Exit Sub
    End If
    
    
    With Spd_Help

    '�]�w���U����(Spd_Help)������ݩ�
         .UnitType = 2          '<---- @!!! Fix property, DO NOT CHANGE IT. !!!
         
         Spread_Property Spd_Help, 0, 2, WHITE, G_Font_Size, G_Font_Name    'row: 0, col: 2
         Spread_Col_Property Spd_Help, 1, TextWidth("X") * 7, G_Pnl_A0201$  'col1 header: A0201
         Spread_Col_Property Spd_Help, 2, TextWidth("X") * 16, G_Pnl_A0201$ 'col2 header: A0202
         Spread_DataType_Property Spd_Help, 1, SS_CELL_TYPE_EDIT, "", "", 6
         Spread_DataType_Property Spd_Help, 2, SS_CELL_TYPE_EDIT, "", "", 12
         
         .Row = -1
         .Col = -1: .Lock = True
         .Col = 1: .TypeHAlign = 2
    
    '�N����\�JSpread��
         Do Until DY_A02.EOF
            .MaxRows = .MaxRows + 1
            .Row = Spd_Help.MaxRows
            .Col = 1
            .text = Trim(DY_A02.Fields("A0201") & "")
            .Col = 2
            .text = Trim(DY_A02.Fields("A0202") & "")
            DY_A02.MoveNext
         Loop
    
'�]�w���U��������ܦ�m
         SetHelpWindowPos Fra_Help, Spd_Help, 330, 90, 4305, 2025
         .Tag = Txt.TabIndex
         .SetFocus
    End With
    
    Me.MousePointer = Default
End Sub

Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    
'����d�ߩΦs�ɫe���N�Ҧ��ˮ����A���@��

    If Not CheckRoutine_A0201 Then Exit Function
    If Not CheckRoutine_A0208s() Then Exit Function
    If Not CheckRoutine_A0208e() Then Exit Function
    
    If Not CheckRoutine_A02001s() Then Exit Function
    If Not CheckRoutine_A02001e() Then Exit Function
    If Not CheckRoutine_A02005s() Then Exit Function
    If Not CheckRoutine_A02005e() Then Exit Function
    DoEvents
    
    IsAllFieldsCheck = True
End Function

Private Function CheckRoutine_A02001s() As Boolean
    CheckRoutine_A02001s = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim(Txt_A02001s) <> "" Then
       If Not IsDateValidate(Txt_A02001s) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A02001$ & G_DateError
          m_FieldError% = Txt_A02001s.TabIndex
          Txt_A02001s.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A02001s), Trim$(Txt_A02001e)) Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
       If ActiveControl.TabIndex = Txt_A02001e.TabIndex Then
          m_FieldError% = Txt_A02001e.TabIndex
       Else
          m_FieldError% = Txt_A02001s.TabIndex
          Txt_A02001s.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A02001s = True
End Function

Private Function CheckRoutine_A02005s() As Boolean
    CheckRoutine_A02005s = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim(Txt_A02005s) <> "" Then
       If Not IsDateValidate(Txt_A02005s) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A02005$ & G_DateError
          m_FieldError% = Txt_A02005s.TabIndex
          Txt_A02005s.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A02005s), Trim$(Txt_A02005e)) Then
       If ActiveControl.TabIndex = Txt_A02005e.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A02005e.TabIndex
       Else
          m_FieldError% = Txt_A02005s.TabIndex
          Txt_A02005s.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A02005s = True
End Function

Private Function CheckRoutine_A0201() As Boolean
    CheckRoutine_A0201 = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim$(Txt_A0201e) = "" Then Txt_A0201e = Txt_A0201s
    
    If Not CheckDataRange(Sts_MsgLine, Trim$(Txt_A0201s), Trim$(Txt_A0201e)) Then
       '==================
       'if from s to e
       'do not focus back (since it's correct to entering from s to e)
       '==================
       If ActiveControl.TabIndex = Txt_A0201e.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A0201e.TabIndex
       Else
          m_FieldError% = Txt_A0201s.TabIndex
          Txt_A0201s.SetFocus
       End If
       Exit Function
    End If
       
    CheckRoutine_A0201 = True
End Function

Private Function CheckRoutine_A0208s() As Boolean
    CheckRoutine_A0208s = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
'    If Trim(Txt_A0208s) = "" Then
'       Txt_A0208s = GetCurrentDay(0)
'    Else
    If Not Trim(Txt_A0208s) = "" Then
       If Not IsDateValidate(Txt_A0208s) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A0208$ & G_DateError
          m_FieldError% = Txt_A0208s.TabIndex
          Txt_A0208s.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A0208s), Trim$(Txt_A0208e)) Then
       If ActiveControl.TabIndex = Txt_A0208e.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A0208s.TabIndex
       Else
          m_FieldError% = Txt_A0208s.TabIndex
          Txt_A0208s.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A0208s = True
End Function

Private Function CheckRoutine_A02001e() As Boolean
    CheckRoutine_A02001e = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim(Txt_A02001e) = "" Then
       Txt_A02001e = GetCurrentDay(0)
    Else
       If Not IsDateValidate(Txt_A02001e) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A02001$ & G_DateError
          m_FieldError% = Txt_A02001e.TabIndex
          Txt_A02001e.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A02001s), Trim$(Txt_A02001e)) Then
       If ActiveControl.TabIndex = Txt_A02001s.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A02001s.TabIndex
       Else
          m_FieldError% = Txt_A02001e.TabIndex
          Txt_A02001e.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A02001e = True
End Function

Private Function CheckRoutine_A02005e() As Boolean
    CheckRoutine_A02005e = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim(Txt_A02005e) = "" Then
       Txt_A02005e = GetCurrentDay(0)
    Else
       If Not IsDateValidate(Txt_A02005e) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A02005$ & G_DateError
          m_FieldError% = Txt_A02005e.TabIndex
          Txt_A02005e.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A02005s), Trim$(Txt_A02005e)) Then
       If ActiveControl.TabIndex = Txt_A02005s.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A02005s.TabIndex
       Else
          m_FieldError% = Txt_A02005e.TabIndex
          Txt_A02005e.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A02005e = True
End Function

Private Function CheckRoutine_A0208e() As Boolean
    CheckRoutine_A0208e = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
'    If Trim(Txt_A0208e) = "" Then
'       Txt_A0208e = GetCurrentDay(0)
'    Else
    If Not Trim(Txt_A0208e) = "" Then
       If Not IsDateValidate(Txt_A0208e) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A0208$ & G_DateError
          m_FieldError% = Txt_A0208e.TabIndex
          Txt_A0208e.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A0208s), Trim$(Txt_A0208e)) Then
       If ActiveControl.TabIndex = Txt_A0208s.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A0208s.TabIndex
       Else
          m_FieldError% = Txt_A0208e.TabIndex
          Txt_A0208e.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A0208e = True
End Function

Private Sub OpenMainFile()
On Local Error GoTo MyError
Dim A_Sql$
Dim A_A02001s$
Dim A_A02001e$
Dim A_A02005s$
Dim A_A02005e$
Dim A_A0201s$
Dim A_A0201e$
Dim A_A0208s$
Dim A_A0208e$
Dim A_WhereAnd$

    'initialize
    A_WhereAnd = "Where"
    
'Keep TextBox ��Ʀ��ܼ�
    A_A0201s$ = Trim(Txt_A0201s)
    A_A0201e$ = Trim(Txt_A0201e)
    A_A0208s$ = Trim(DateIn(Trim(Txt_A0208s)))
    A_A0208e$ = Trim(DateIn(Trim(Txt_A0208e)))
    A_A02001s$ = Trim(DateIn(Trim(Txt_A02001s)))
    A_A02001e$ = Trim(DateIn(Trim(Txt_A02001e)))
    A_A02005s$ = Trim(DateIn(Trim(Txt_A02005s)))
    A_A02005e$ = Trim(DateIn(Trim(Txt_A02005e)))
    
    
'�}�Ҹ��
    'get the required Columns as SPEC
    A_Sql$ = "Select A0201,A0202,A0206,A0208,A0218 From A02 "
    
    'where clause: A0201
    If A_A0201s$ <> "" Then
       A_Sql$ = A_Sql$ & A_WhereAnd & " A0201>='" & A_A0201s$ & "' "
       If A_WhereAnd = "Where" Then A_WhereAnd = "and"
    End If
    If A_A0201e$ <> "" Then
       A_Sql$ = A_Sql$ & A_WhereAnd & " A0201<='" & A_A0201e$ & "' "
       If A_WhereAnd = "Where" Then A_WhereAnd = "and"
    End If
    
    'where clause A0208
    If A_A0208s$ <> "" Then
       A_Sql$ = A_Sql$ & A_WhereAnd & " A0208>='" & A_A0208s$ & "' "
       If A_WhereAnd = "Where" Then A_WhereAnd = "and"
    End If
    If A_A0208e$ <> "" Then
       A_Sql$ = A_Sql$ & A_WhereAnd & " A0208<='" & A_A0208e$ & "' "
       If A_WhereAnd = "Where" Then A_WhereAnd = "and"
    End If
    
    'where clause A02001
    If A_A02001s$ <> "" Then
       A_Sql$ = A_Sql$ & A_WhereAnd & " A02001>='" & A_A02001s$ & "' "
       If A_WhereAnd = "Where" Then A_WhereAnd = "and"
    End If
    If A_A02001e$ <> "" Then
       A_Sql$ = A_Sql$ & A_WhereAnd & " A02001<='" & A_A02001e$ & "' "
       If A_WhereAnd = "Where" Then A_WhereAnd = "and"
    End If
    
    'where clause A02005
    If A_A02005s$ <> "" Then
       A_Sql$ = A_Sql$ & A_WhereAnd & " A02005>='" & A_A02005s$ & "' "
       If A_WhereAnd = "Where" Then A_WhereAnd = "and"
    End If
    If A_A02005e$ <> "" Then
       A_Sql$ = A_Sql$ & A_WhereAnd & " A02005<='" & A_A02005e$ & "' "
       If A_WhereAnd = "Where" Then A_WhereAnd = "and"
    End If
    
    
    A_Sql$ = A_Sql$ & "Order by A0201"
    
    CreateDynasetODBC DB_ARTHGUI, DY_A02, A_Sql$, "DY_A02", True
    Exit Sub
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Sub cmd_add_Click()
'�N�@�~���A�]�w���s�W���A
    G_AP_STATE = G_AP_STATE_ADD
    
'����Q�e��, Show�XDetail�e��
    DoEvents
    Me.Hide
    frm_TSM02.Show
End Sub

Private Sub Cmd_Ok_Click()
    Me.MousePointer = HOURGLASS
    
    Sts_MsgLine.Panels(1) = G_Process
    Sts_MsgLine.Refresh
    
'�w�惡�e���������ˮ���찵PageCheck
    If Not IsAllFieldsCheck() Then
       Me.MousePointer = Default
       Exit Sub
    End If

'�}�Ҭd�߸��
    OpenMainFile
    
'�N�����ܨ�V�e��
    If Not (DY_A02.BOF And DY_A02.EOF) Then
       DoEvents
       Me.Hide
       Frm_TSM02v.Show
    Else
       Sts_MsgLine.Panels(1) = G_NoQueryData
    End If
    
    Me.MousePointer = Default
End Sub

Private Sub Cmd_Exit_Click()
'�����ثe����,���X��L�B�z�{��
    m_ExitTrigger% = True
    CloseFileDB
    End
End Sub

Private Sub Cmd_Help_Click()
Dim a$

    a$ = "notepad " + G_Help_Path + "TSM02q.HLP"
    retcode = Shell(a$, 4)
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
       'Do Something Here��
       
    End If
    G_AP_STATE = G_AP_STATE_QUERY  '�]�w�@�~���A
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    
    '�NForm��m��ù������h
    frm_TSM02q.ZOrder 0
    If frm_TSM02q.Visible Then Txt_A0201s.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case KEY_F1
                If ActiveControl.TabIndex = Txt_A0201s.TabIndex Then Exit Sub
                If ActiveControl.TabIndex = Txt_A0201e.TabIndex Then Exit Sub
                KeyCode = 0
                If Cmd_Help.Visible = True And Cmd_Help.Enabled = True Then
                   Cmd_Help.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F2
                KeyCode = 0
                If Cmd_Ok.Visible = True And Cmd_Ok.Enabled = True Then
                   Cmd_Ok.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F4
                KeyCode = 0
                If Cmd_Add.Visible = True And Cmd_Add.Enabled = True Then
                   Cmd_Add.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_ESCAPE
                KeyCode = 0
                If Cmd_Exit.Visible = True And Cmd_Exit.Enabled = True Then
                   Cmd_Exit.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    
'�D�ʱN��ƿ�J�Ѥp�g�ର�j�g
'  �Y���Y����줣�ݭn�ഫ��, �����H���L
   'If ActiveControl.TabIndex = txt_xxx.TabIndex Then GoTo Form_KeyPress_A
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Form_KeyPress_A:
'��J���N�r��(ENTER���~), �N��Ʋ����ܼƳ]��TRUE
    'If ActiveControl.TabIndex <> Spd_PATTERNM.TabIndex Then
       KeyPress KeyAscii           'Enter�ɦ۰ʸ���U�@���, spread���~
    'End If
End Sub

Private Sub Form_Load()
    FormCenter Me
    Set_Property
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim MSG ' Declare variable.

    If UnloadMode > 0 Then
       ' If exiting the application.
       MSG = GetSIniStr("PgmMsg", "g_gui_run")
    Else
       ' If just closing the form.
       Cmd_Exit_Click
    End If
    ' If user clicks the 'No' button, stop QueryUnload.
    If MsgBox(MSG, 36, Me.Caption) = 7 Then
       Cancel = True
    Else
       Cmd_Exit_Click
    End If
    
End Sub

Private Sub Spd_Help_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim A_Code$

    Me.MousePointer = HOURGLASS
    
'KEEP�ۻ��U�����I�諸���
    With Spd_Help
         'redirect to Pkey
         .Row = .ActiveRow
         .Col = 1
         A_Code$ = Trim(.text)
    
'�NKEEP����Ʊa�J�e��
         Select Case Val(.Tag)
           Case Txt_A0201s.TabIndex
                Txt_A0201s = A_Code$
           Case Txt_A0201e.TabIndex
                Txt_A0201e = A_Code$
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
      Case Txt_A0208s.TabIndex
           Txt_A0208s.SetFocus
      Case Txt_A0208e.TabIndex
           Txt_A0208e.SetFocus
    End Select
End Sub

Private Sub Txt_A02001e_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A02001e_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A02001e.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A02001e()
End Sub

Private Sub Txt_A02001s_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A02001s_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A02001s.TabIndex Then Exit Sub
    
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A02001s()
End Sub

Private Sub Txt_A02005e_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A02005e_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A02005e.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A02005e()
End Sub

Private Sub Txt_A02005s_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A02005s_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A02005s.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A02005s()
End Sub

Private Sub Txt_A0201e_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A0201e_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A0201e_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If KeyCode = KEY_F1 Then DataPrepare_A02 Txt_A0201e
End Sub

Private Sub Txt_A0201e_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A0201e_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If Fra_Help.Visible = True Then Exit Sub
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0201e.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0201()
End Sub

Private Sub Txt_A0201s_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A0201s_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A0201s_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If KeyCode = KEY_F1 Then DataPrepare_A02 Txt_A0201s
End Sub

Private Sub Txt_A0201s_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A0201s_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If Fra_Help.Visible = True Then Exit Sub
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0201s.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0201()
End Sub

Private Sub Txt_A0208e_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0208e_LostFocus()
     TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0208e.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0208e()
End Sub

Private Sub Txt_A0208s_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0208s_LostFocus()
     TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0208s.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0208s()
End Sub
