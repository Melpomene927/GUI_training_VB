VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_PATTERNR2q 
   Caption         =   "�ϥΰO���C�L"
   ClientHeight    =   4155
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
   Icon            =   "TSR03q.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4155
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VsOcxLib.VideoSoftElastic Vse_Background 
      Height          =   3780
      Left            =   0
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   7155
      _Version        =   327680
      _ExtentX        =   12621
      _ExtentY        =   6667
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
      Picture         =   "TSR03q.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "TSR03q.frx":0326
      Begin ComctlLib.ProgressBar Prb_Percent 
         Height          =   210
         Left            =   1260
         TabIndex        =   24
         Top             =   2580
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
         Left            =   5820
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   825
         Begin FPSpread.vaSpread Spd_Help 
            Height          =   495
            Left            =   90
            OleObjectBlob   =   "TSR03q.frx":0342
            TabIndex        =   9
            Top             =   210
            Width           =   615
         End
      End
      Begin VB.ComboBox Cbo_A0906 
         Appearance      =   0  'Flat
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
         ItemData        =   "TSR03q.frx":0572
         Left            =   1620
         List            =   "TSR03q.frx":0574
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   3945
      End
      Begin VB.ComboBox Cbo_A0905 
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
         ItemData        =   "TSR03q.frx":0576
         Left            =   1620
         List            =   "TSR03q.frx":0578
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1740
         Width           =   3945
      End
      Begin VB.ComboBox Cbo_A0911 
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
         ItemData        =   "TSR03q.frx":057A
         Left            =   1620
         List            =   "TSR03q.frx":057C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
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
         TabIndex        =   31
         Top             =   2550
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
            TabIndex        =   14
            Text            =   " "
            Top             =   660
            Width           =   5235
         End
         Begin Threed.SSOption Opt_File 
            Height          =   360
            Left            =   3060
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   13
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
      Begin VB.TextBox Txt_A0904e 
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
         TabIndex        =   8
         Top             =   2190
         Width           =   1515
      End
      Begin VB.TextBox Txt_A0904s 
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
         TabIndex        =   7
         Top             =   2190
         Width           =   1515
      End
      Begin VB.TextBox Txt_A0901e 
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
         MaxLength       =   10
         TabIndex        =   2
         Top             =   510
         Width           =   1515
      End
      Begin VB.TextBox Txt_A0902e 
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
         TabIndex        =   3
         Top             =   510
         Width           =   1515
      End
      Begin VB.TextBox Txt_A0902s 
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
         TabIndex        =   1
         Top             =   120
         Width           =   1515
      End
      Begin VB.TextBox Txt_A0901s 
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
         MaxLength       =   10
         TabIndex        =   0
         Top             =   120
         Width           =   1515
      End
      Begin Threed.SSCommand Cmd_Help 
         Height          =   405
         Left            =   5670
         TabIndex        =   15
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
         TabIndex        =   18
         Top             =   3270
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
         TabIndex        =   16
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
      Begin Threed.SSCommand Cmd_Set 
         Height          =   405
         Left            =   5670
         TabIndex        =   17
         Top             =   1020
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "���]�w F9"
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
         TabIndex        =   30
         Top             =   2250
         Width           =   300
      End
      Begin VB.Label Lbl_A0906 
         Caption         =   "�{���N�X"
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
         Left            =   30
         TabIndex        =   29
         Top             =   1410
         Width           =   1560
      End
      Begin VB.Label Lbl_A0904 
         Caption         =   "User ID"
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
         Left            =   30
         TabIndex        =   28
         Top             =   2220
         Width           =   1560
      End
      Begin VB.Label Lbl_A0905 
         Caption         =   "�s�եN�X"
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
         Left            =   30
         TabIndex        =   27
         Top             =   1830
         Width           =   1560
      End
      Begin VB.Label Lbl_A0901e 
         Caption         =   "�I����/�ɶ�"
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
         Left            =   30
         TabIndex        =   26
         Top             =   570
         Width           =   1560
      End
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "/"
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
         Index           =   1
         Left            =   3450
         TabIndex        =   25
         Top             =   570
         Width           =   300
      End
      Begin VB.Label Lbl_A0911 
         Caption         =   "�t�ΥN��"
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
         Left            =   30
         TabIndex        =   23
         Top             =   990
         Width           =   1560
      End
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "/"
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
         Left            =   3450
         TabIndex        =   22
         Top             =   180
         Width           =   300
      End
      Begin VB.Label Lbl_A0901s 
         Caption         =   "�_�l���/�ɶ�"
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
         Left            =   30
         TabIndex        =   21
         Top             =   180
         Width           =   1560
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   3780
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
Attribute VB_Name = "frm_PATTERNR2q"
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
Dim m_A0911Flag%
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

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
       Sts_MsgLine.Panels(1) = GetCaption("PgmMsg", "path_not_found", "���|���s�b !")
       m_FieldError% = Txt_FileName.TabIndex
       Txt_FileName.SetFocus
    End If
End Function

Private Function CheckRoutine_A0904s() As Boolean
    CheckRoutine_A0904s = False
    
'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Txt_A0904s.text = "" Then
       Txt_A0904s.Tag = ""
    Else
        If UCase$(Trim$(Txt_A0904s)) <> UCase$(Trim$(Txt_A0904s.Tag)) Then
           If Not CheckRoutine_A08(Trim(Txt_A0904s)) Then
              Sts_MsgLine.Panels(1) = Lbl_A0904 & G_FieldErr
              m_FieldError% = Txt_A0904s.TabIndex
              Txt_A0904s.Tag = ""
              Txt_A0904s.SetFocus
              Exit Function
           Else
              Txt_A0904s.Tag = Txt_A0904s
           End If
           If Not CheckDataRange(Sts_MsgLine, Txt_A0904s.text, Txt_A0904e.text) Then
              If ActiveControl.TabIndex = Txt_A0904e.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
                 m_FieldError% = Txt_A0904e.TabIndex
              Else
                 m_FieldError% = Txt_A0904s.TabIndex
                 Txt_A0904s.SetFocus
              End If
              Exit Function
           End If
        End If
    End If
    
    CheckRoutine_A0904s = True
End Function

Private Function CheckRoutine_A0904e() As Boolean
    CheckRoutine_A0904e = False
    
'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Txt_A0904e.text = "" Then
       Txt_A0904e.Tag = ""
       If Trim(Txt_A0904s) <> "" Then
          Txt_A0904e = Txt_A0904s
          Txt_A0904e.Tag = Txt_A0904s.Tag
       End If
    Else
        If UCase$(Trim$(Txt_A0904e)) <> UCase$(Trim$(Txt_A0904e.Tag)) Then
           If Not CheckRoutine_A08(Trim(Txt_A0904e)) Then
              Sts_MsgLine.Panels(1) = Lbl_A0904 & G_FieldErr
              m_FieldError% = Txt_A0904e.TabIndex
              Txt_A0904e.Tag = ""
              Txt_A0904e.SetFocus
              Exit Function
           Else
              Txt_A0904e.Tag = Txt_A0904e
           End If
           If Not CheckDataRange(Sts_MsgLine, Txt_A0904s.text, Txt_A0904e.text) Then
              If ActiveControl.TabIndex = Txt_A0904s.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
                 m_FieldError% = Txt_A0904s.TabIndex
              Else
                 m_FieldError% = Txt_A0904e.TabIndex
                 Txt_A0904e.SetFocus
              End If
              Exit Function
           End If
        End If
    End If
    
    CheckRoutine_A0904e = True
End Function


Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    If Not CheckRoutine_A0901s() Then Exit Function
    If Not CheckRoutine_A0901e() Then Exit Function
    If Not CheckRoutine_A0904s() Then Exit Function
    If Not CheckRoutine_A0904e() Then Exit Function
    If Not CheckRoutine_FileName() Then Exit Function
    If Trim$(Txt_A0902s) = "" Then Txt_A0902s = GetCurrentTime()
    If Trim$(Txt_A0902e) = "" Then Txt_A0902e = GetCurrentTime()
    DoEvents
    IsAllFieldsCheck = True
End Function

Private Sub KeepFieldsValue()
'??? ��ƪ��C�L�ӷ���RecordSet�Ө�
    G_ReportDataFrom = G_FromRecordSet
    
'Keep�C�L����
    G_A0901s$ = Trim(Txt_A0901s)
    G_A0901e$ = Trim(Txt_A0901e)
    G_A0902s$ = Trim(Txt_A0902s) & "00"
    G_A0902e$ = Trim(Txt_A0902e) & "00"
    G_A0904s$ = Trim$(Txt_A0904s)
    G_A0904e$ = Trim$(Txt_A0904e)
    StrCut Cbo_A0905.text, Space(1), G_A0905$, G_A0905o$
    StrCut Cbo_A0906.text, Space(1), G_A0906$, G_A0906o$
    StrCut Cbo_A0911.text, Space(1), G_A0911$, G_A0911o$
    G_OutFile = Trim$(Txt_FileName)
    If Opt_Printer.Value Then G_PrintSelect = G_Print2Printer
    If Opt_Scrn.Value Then G_PrintSelect = G_Print2Screen
    If Opt_File.Value Then G_PrintSelect = G_Print2File
    If Opt_Excel.Value Then G_PrintSelect = G_Print2Excel
End Sub

Private Sub OpenMainFile()
On Local Error GoTo MY_Error
Dim A_Sql$

    A_Sql$ = "SELECT * FROM A09"
    A_Sql$ = A_Sql$ & " WHERE NOT (A0901<'" & DateIn(G_A0901s$) & "'"
    A_Sql$ = A_Sql$ & " OR (A0901='" & DateIn(G_A0901s$) & "' AND A0902<'" & G_A0902s$ & "'))"
    A_Sql$ = A_Sql$ & " AND NOT (A0901>'" & DateIn(G_A0901e$) & "'"
    A_Sql$ = A_Sql$ & " OR (A0901='" & DateIn(G_A0901e$) & "' AND A0902>'" & G_A0902e$ & "'))"
    If G_A0904s$ <> "" Then A_Sql$ = A_Sql$ & " and A0904>='" & G_A0904s$ & "'"
    If G_A0904e$ <> "" Then A_Sql$ = A_Sql$ & " and A0904<='" & G_A0904e$ & "'"
    If G_A0905$ <> "" Then A_Sql$ = A_Sql$ & " and A0905='" & G_A0905$ & "'"
    If G_A0906$ <> "" Then A_Sql$ = A_Sql$ & " and A0906='" & G_A0906$ & "'"
    If G_A0911$ <> "" Then A_Sql$ = A_Sql$ & " and A0911='" & G_A0911$ & "'"
'??? �]�w�Ƨ����(�ĤG�ӰѼƬ��{���w�]���Ƨ����)
    A_Sql$ = A_Sql$ & GetOrderCols(tSpd_PATTERNR2, "A0909,A0911,A0901,A0902")
    CreateDynasetODBC DB_ARTHGUI, DY_A09, A_Sql$, "DY_A09", True
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub




Private Sub Cbo_A0905_DropDown()
    DoEvents

'�N�ثeCombo Box�W���N�XKeep�U��
    Dim A_A0905$
    StrCut Cbo_A0905.text, Space(1), A_A0905$, ""
    
'���s�ǳƦ�Combo Box�����e
    Cbo_A0905_Prepare
    
'�NCombo Box�W��ListIndex���VKeep�U�Ӫ����
    CboStrCut Cbo_A0905, A_A0905$, Space(1)
End Sub

Sub Cbo_A0911_Prepare()
On Local Error GoTo MY_Error
Dim A_Sql$

'���M��Combo Box���e
    Cbo_A0911.Clear

'�}�_�ɮ�
    A_Sql$ = "SELECT A1201,A1202 FROM A12 ORDER BY A1201"
    CreateDynasetODBC DB_ARTHGUI, DY_A12, A_Sql$, "DY_A12", True
    
'�N����\�JCombo Box��
    Do While Not DY_A12.EOF
       Cbo_A0911.AddItem Format$(Trim$(DY_A12.Fields("A1201") & ""), "!@@@@@@@@@@@") & Trim$(DY_A12.Fields("A1202") & "")
       DY_A12.MoveNext
    Loop

'��Default���
    Cbo_A0911.ListIndex = -1
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Sub Cbo_A0905_Prepare()
On Local Error GoTo MY_Error
Dim A_Sql$

'���M��Combo Box���e
    Cbo_A0905.Clear

'�}�_�ɮ�
    A_Sql$ = "SELECT A0601,A0602 FROM A06 ORDER BY A0601"
    CreateDynasetODBC DB_ARTHGUI, DY_A06, A_Sql$, "DY_A06", True
    
'�N����\�JCombo Box��
    Do While Not DY_A06.EOF
       Cbo_A0905.AddItem Format$(Trim$(DY_A06.Fields("A0601") & ""), "!@@@@") & Trim$(DY_A06.Fields("A0602") & "")
       DY_A06.MoveNext
    Loop

'��Default���
    Cbo_A0905.ListIndex = -1
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub


Function CheckRoutine_A08(ByVal A_A0826$) As Boolean
On Local Error GoTo MY_Error
Dim A_Sql$

    CheckRoutine_A08 = False
    A_Sql$ = "Select A0801 From A08"
    A_Sql$ = A_Sql$ & " where A0826='" & A_A0826$ & "'"
    CreateDynasetODBC DB_ARTHGUI, DY_A08, A_Sql$, "DY_A08", True
    If Not (DY_A08.BOF And DY_A08.EOF) Then CheckRoutine_A08 = True
    Exit Function
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Private Sub Cbo_A0905_GotFocus()
    TextGotFocus
End Sub


Private Sub Cbo_A0905_LostFocus()
    TextLostFocus
End Sub


Private Sub Cbo_A0911_Click()

    If m_A0911Flag% Then Exit Sub
    
'�Y����Ƥ��e���I��B�ܰʮ�, �Ҷ��B�z���ƶ�
    If Trim(Cbo_A0911.text) <> Trim(Cbo_A0911.Tag) Then
       Cbo_A0906_Prepare
       Cbo_A0911.Tag = Cbo_A0911.text
    End If
End Sub

Private Sub Cbo_A0911_DropDown()
    DoEvents
    
    m_A0911Flag% = True

'�N�ثeCombo Box�W���N�XKeep�U��
    Dim A_A0911$
    StrCut Cbo_A0911.text, Space(1), A_A0911$, ""
    
'���s�ǳƦ�Combo Box�����e
    Cbo_A0911_Prepare
    
'�NCombo Box�W��ListIndex���VKeep�U�Ӫ����
    CboStrCut Cbo_A0911, A_A0911$, Space(1)
    
    m_A0911Flag% = False
End Sub


Private Sub Cbo_A0911_GotFocus()
    TextGotFocus
End Sub


Private Sub Cbo_A0911_LostFocus()
    TextLostFocus
End Sub


Private Sub Cmd_Exit_Click()
'�зǼg�k,���i�ק�
    Unload Me
End Sub

Private Sub Cmd_Help_Click()
Dim a$

'�бNPATTERNRq�אּ��Form�W�r�Y�i, ��l���зǫ��O, ���o�ק�
    a$ = "notepad " + G_Help_Path + "PATTERNR2q.HLP"
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
    If DY_A09.BOF And DY_A09.EOF Then

'�L��Ƥ����C�L
       Sts_MsgLine.Panels(1) = G_NoQueryData
    Else

'����RepSet Form������,���|Ĳ�oForm_Activate
       If G_PrintSelect = G_Print2Printer Then
          G_FormFrom$ = "RptSet"
       End If
       
       If Not Opt_Scrn.Value Then

'??? �}�l�C�L����,�ĤT�ӰѼƶǤJV Screen��Spread
          PrePare_Data frm_PATTERNR2q, Prb_Percent, frm_PATTERNR2!Spd_PATTERNR2, m_ExitTrigger%

'��Esc��QĲ�o,�����C�L�ʧ@
          If m_ExitTrigger% Then Exit Sub
       Else
          DoEvents
          Me.Hide
          frm_PATTERNR2.Show
          Sts_MsgLine.Panels(1) = G_PrintOk
       End If
    End If
    Cmd_Print.Enabled = True
    Me.MousePointer = Default
End Sub

Private Sub Cbo_A0906_DropDown()
    DoEvents

'�N�ثeCombo Box�W���N�XKeep�U��
    Dim A_A0906$
    StrCut Cbo_A0906.text, Space(1), A_A0906$, ""
    
'���s�ǳƦ�Combo Box�����e
    Cbo_A0906_Prepare
    
'�NCombo Box�W��ListIndex���VKeep�U�Ӫ����
    CboStrCut Cbo_A0906, A_A0906$, Space(1)
End Sub


Private Sub Cbo_A0906_GotFocus()
    TextGotFocus
End Sub


Private Sub Cbo_A0906_LostFocus()
    TextLostFocus
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
    frm_PATTERNR2q.ZOrder 0
    If frm_PATTERNR2q.Visible Then Txt_A0901s.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case KEY_DELETE
                If TypeOf ActiveControl Is ComboBox Then
                   ActiveControl.ListIndex = -1
                End If
                
           Case KEY_F1
                If ActiveControl.TabIndex = Txt_A0904s.TabIndex Then Exit Sub
                If ActiveControl.TabIndex = Txt_A0904e.TabIndex Then Exit Sub
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
                
           Case KEY_F9
                KeyCode = 0
                If Cmd_Set.Visible And Cmd_Set.Enabled Then
                   Cmd_Set.SetFocus
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

Sub Cbo_A0906_Prepare()
On Local Error GoTo MY_Error
Dim A_Sql$, A_A1003$

'���M��Combo Box���e
    Cbo_A0906.Clear
    
'�}�_�ɮ�
    StrCut Cbo_A0911.text, Space(1), A_A1003$, ""
    A_Sql$ = "Select A1001,A1002 From A10"
    A_Sql$ = A_Sql$ & " where A1003='" & A_A1003$ & "'"
    A_Sql$ = A_Sql$ & " order by A1001"
    CreateDynasetODBC DB_ARTHGUI, DY_A10, A_Sql$, "DY_A10", True

'�N����\�JCombo Box��
    Do While Not DY_A10.EOF
       Cbo_A0906.AddItem Format$(Trim$(DY_A10.Fields("A1001") & ""), "!@@@@@@@@@@@") & Trim$(DY_A10.Fields("A1002") & "")
       DY_A10.MoveNext
    Loop

'��Default���
    Cbo_A0906.ListIndex = -1
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
'�D�ʱN��ƿ�J�Ѥp�g�ର�j�g
'  �Y���Y����줣�ݭn�ഫ��, �����H���L
   If ActiveControl.TabIndex <> Txt_A0904s.TabIndex And _
   ActiveControl.TabIndex <> Txt_A0904e.TabIndex Then _
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
Private Sub Set_Property()
'??? �]�w��Form�����D,�r�ΤΦ�t
    Form_Property frm_PATTERNR2q, GetCaption("FormTitle", "PATTERNRq", "�ϥΰO���C�L"), G_Font_Name
    
'??? �]�wForm���Ҧ�TextBox,ComboBox,ListBox���r�ΤΥi��H����,�i�P�ɳ]�w��ҹ�����Label������ݩ�
'    �ѼƤ@ : Control Name                              �ѼƤG : ���󪺳̤j����,�DTextBox�п�J0
'    �ѼƤT : ����Label��Control Name,�]�w������ݩ�    �Ѽƥ| : �]�wLabel��Caption,�Y�۸�Ʈw�줣��Caption�h�H���]�w��Label��Caption
'    �ѼƤ� : ��J��쪺�榡,�Ω����μƭȿ�J         �ѼƤ� : �ƭ���쪺�W��
'    �ѼƤC : �ƭ���쪺�U��                            �ѼƤK : Database Name,�󦹸�Ʈw�U��MLabel��Caption
'    �ѼƤE : Table Name,����U��MLabel��Caption     �ѼƤQ : Field Name,�H������MLabel��Caption
    Field_Property Txt_A0901s, 8, Lbl_A0901s, GetCaption("PATTERNR", "startdate", "�_�l���/�ɶ�")
    Field_Property Txt_A0902s, 6
    Field_Property Txt_A0901e, 8, Lbl_A0901e, GetCaption("PATTERNR", "enddate", "�I����/�ɶ�")
    Field_Property Txt_A0902e, 6
    Field_Property Txt_A0904s, 10, Lbl_A0904, GetCaption("PATTERNR", "userid", "User ID"), , , , "ARTHGUI", "A09", "A0904"
    Field_Property Txt_A0904e, Txt_A0904s.MaxLength
    Field_Property Txt_FileName, 60
    Field_Property Cbo_A0905, 0, Lbl_A0905, GetCaption("PanelDescpt", "groupid", "�s�եN�X"), , , , "ARTHGUI", "A09", "A0905"
    Field_Property Cbo_A0906, 0, Lbl_A0906, GetCaption("PanelDescpt", "groupid", "�{���N�X"), , , , "ARTHGUI", "A09", "A0906"
    Field_Property Cbo_A0911, 0, Lbl_A0911, GetCaption("PATTERNR", "systemid", "�t�ΥN��"), , , , "ARTHGUI", "A09", "A0911"
    Txt_FileName.Visible = False
    
'??? �]�wForm���Ҧ�Panel,Label,OptionButton,CheckBox,Frame�����D, �r�ΤΦ�t
'    �ѼƤ@ : Control Name              �ѼƤG : �]�wControl��Caption
'    �ѼƤT : �O�_���                  �Ѽƥ| : �]�w�I���C��
'    �ѼƤ� : �]�w�r���j�p              �ѼƤ� : �]�w�r���W��
    Control_Property Lbl_Sign(0), "/"
    Control_Property Lbl_Sign(1), "/"
    Control_Property Lbl_Sign(2), GetCaption("PanelDescpt", "dash", "~")
    Control_Property Opt_Printer, GetCaption("PanelDescpt", "printer", "�L���")
    Control_Property Opt_Scrn, GetCaption("PanelDescpt", "screen", "�ù����")
    Control_Property Opt_File, GetCaption("PanelDescpt", "file", "�ɮ�")
    Control_Property Opt_Excel, GetCaption("PanelDescpt", "excel", "Excel")
    Control_Property Fra_PrintType, GetCaption("PanelDescpt", "printtype", "�C�L�覡")
    Control_Property Fra_Help, "", False, COLOR_SKY
        
'�]Form���Ҧ�Command�����D�Φr��
    Command_Property Cmd_Help, G_CmdHelp, G_Font_Name
    Command_Property Cmd_Print, G_CmdPrint, G_Font_Name
    Command_Property Cmd_Set, G_CmdSet, G_Font_Name
    Command_Property Cmd_exit, G_CmdExit, G_Font_Name
    
'�H�U���зǫ��O, ���o�ק�
    ProgressBar_Property Prb_Percent
    VSElastic_Property Vse_Background
    StatusBar_ProPerty Sts_MsgLine
End Sub
Sub PrePare_ComboBox()
    Cbo_A0905_Prepare
    Cbo_A0906_Prepare
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'�YUser��������Windows, �ӥ��{�����b����, ���{���|���߰ݬO�_�n�������ۤv?
'�H�U���зǫ��O, ���o�ק�
Dim A_Msg$, A_Flag%

    A_Flag% = False
    If UnloadMode = vbAppWindows Then
       A_Msg$ = GetSIniStr("PgmMsg", "g_gui_run")   ' If exiting the application.
       'If user clicks the 'No' button, stop QueryUnload.
       If MsgBox(A_Msg$, 36, Me.Caption) = 7 Then
          Cancel = True
       Else
          A_Flag% = True
       End If
    Else
       A_Flag% = True
    End If
    
    If A_Flag% Then BeforeUnloadForm
    
End Sub

Sub BeforeUnloadForm()
'�������{���e,���B�z���ʧ@�b���[�J
    
'����Excel����
    CloseExcelFile
    
'�����ثe����,���X��L�B�z�{��
    m_ExitTrigger% = True
    
'??? ����V Screen
    DoEvents
    Unload frm_PATTERNR2
    
'??? �x�s�ثe����榡
    SaveSpreadDefault tSpd_Help, "frm_PATTERNR2q", "Spd_Help"
    SaveSpreadDefault tSpd_PATTERNR2, "frm_PATTERNR2", "Spd_PATTERNR2"
    
'���������ϱҪ�Recordset��Database
    CloseFileDB
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



Private Sub Spd_Help_Click(ByVal Col As Long, ByVal Row As Long)
'??? �Ѧ�����Spread�O�_���ѱƧǥ\��
    If Not tSpd_Help.SortEnable Then Exit Sub
    
'��Column Heading Click��, �̸����Ƨ�
    If Row = 0 And Col > 0 Then
    
'??? Update Spread Type�����Ƨ����
       SpdSortIndexReBuild tSpd_Help, Col
       
'??? �Q��Spread Type��Sort
       SpreadColsSort Spd_Help, tSpd_Help
       
    End If
End Sub

Private Sub Spd_Help_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim A_Code$

    Me.MousePointer = HOURGLASS
    
'KEEP�ۻ��U�����I�諸���
    With Spd_Help
    
'??? �H�ۭq�����W�٨��o����
         A_Code$ = GetSpdText(Spd_Help, tSpd_Help, "A0826", Row)
    
'�NKEEP����Ʊa�J�e��
         Select Case Val(.Tag)
           Case Txt_A0904s.TabIndex
                Txt_A0904s = A_Code$
           Case Txt_A0904e.TabIndex
                Txt_A0904e = A_Code$
         End Select
    End With
    
'���û��U����
    Fra_Help.Visible = False
    
    Me.MousePointer = Default
End Sub


Private Sub Spd_Help_DragDropBlock(ByVal Col As Long, ByVal Row As Long, ByVal Col2 As Long, ByVal Row2 As Long, ByVal newcol As Long, ByVal NewRow As Long, ByVal NewCol2 As Long, ByVal NewRow2 As Long, ByVal Overwrite As Boolean, Action As Integer, DataOnly As Boolean, Cancel As Boolean)
'??? �NSpread�W������첾�ʦܥت����
    SpreadColumnMove Spd_Help, tSpd_Help, Col, newcol, NewRow, Cancel
    
'�b�P�@���DragDrop���B�z�ܦ�
    If Col = newcol Then Exit Sub
    
'�M������쪺�C��
    SpreadLostFocus Col, Row
    
'�]�w�s��쪺�C��
    SpreadGotFocus newcol, NewRow
End Sub

Private Sub Spd_Help_GotFocus()
    SpreadGotFocus Spd_Help.ActiveCol, Spd_Help.ActiveRow
End Sub


Private Sub Spd_Help_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_RETURN Then
       Spd_Help_DblClick Spd_Help.ActiveCol, Spd_Help.ActiveRow
    End If
End Sub


Private Sub Spd_Help_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal newcol As Long, ByVal NewRow As Long, Cancel As Boolean)
'�зǫ��O,���o�ק�
    SpreadLostFocus Col, Row
    If newcol > 0 Then SpreadGotFocus newcol, NewRow
End Sub


Private Sub Spd_Help_LostFocus()
    Fra_Help.Visible = False
    Select Case Val(Spd_Help.Tag)
      Case Txt_A0904s.TabIndex
           Txt_A0904s.SetFocus
      Case Txt_A0904e.TabIndex
           Txt_A0904e.SetFocus
    End Select
End Sub


Private Sub Cmd_Set_Click()
'??? Load���]�w�����
'    �ѼƤ@ : ���]�w��Form Name
'    �ѼƤG : �п�J������User�]�w��vaSpread��Spread Type Name
'    �ѼƤT : �O�_�B�zSpread�Ƨ���첧�ʪ���s
    ShowRptDefForm frm_RptDef, tSpd_PATTERNR2
    
'??? �۪��]�w����^��,�B�zSpread�W����ƭ���
'    �ѼƤ@ : ��Ʊ����㪺Spread Name
'    �ѼƤG : �п�J�ѼƤ@��Spread Type Name
    RefreshSpreadData frm_PATTERNR2.Spd_PATTERNR2, tSpd_PATTERNR2
End Sub

Private Sub Txt_A0901e_GotFocus()
    TextGotFocus
End Sub


Private Sub Txt_A0901e_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If TypeOf ActiveControl Is SSCommand Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0901e.TabIndex Then Exit Sub

    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0901e()
End Sub


Private Sub Txt_A0902e_GotFocus()
    TextGotFocus
End Sub


Private Sub Txt_A0902e_LostFocus()
    TextLostFocus
    If Trim$(Txt_A0902e) = "" Then Txt_A0902e = GetCurrentTime()
End Sub


Private Sub Txt_A0902s_GotFocus()
    TextGotFocus
End Sub


Private Sub Txt_A0902s_LostFocus()
    TextLostFocus
    If Trim$(Txt_A0902s) = "" Then Txt_A0902s = GetCurrentTime()
End Sub


Private Sub Txt_A0901s_GotFocus()
    TextGotFocus
End Sub


Private Sub Txt_A0901s_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If TypeOf ActiveControl Is SSCommand Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0901s.TabIndex Then Exit Sub

    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0901s()
End Sub


Private Function CheckRoutine_A0901s() As Boolean
    CheckRoutine_A0901s = False
    
'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim$(Txt_A0901s) = "" Then
       Txt_A0901s = GetCurrentDay(0)
    Else
       If Not IsDateValidate(Trim$(Txt_A0901s)) Then
          Sts_MsgLine.Panels(1) = Lbl_A0901s & G_FieldErr
          m_FieldError% = Txt_A0901s.TabIndex
          Txt_A0901s.SetFocus
          Exit Function
       End If
    End If
    If Not CheckDateRange(Sts_MsgLine, Txt_A0901s.text, Txt_A0901e.text) Then
       If ActiveControl.TabIndex = Txt_A0901e.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A0901e.TabIndex
       Else
          m_FieldError% = Txt_A0901s.TabIndex
          Txt_A0901s.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A0901s = True
End Function
Private Function CheckRoutine_A0901e() As Boolean
    CheckRoutine_A0901e = False
    
'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim$(Txt_A0901e) = "" Then
       Txt_A0901e = GetCurrentDay(0)
    Else
       If Not IsDateValidate(Trim$(Txt_A0901e)) Then
          Sts_MsgLine.Panels(1) = Lbl_A0901e & G_FieldErr
          m_FieldError% = Txt_A0901e.TabIndex
          Txt_A0901e.SetFocus
          Exit Function
       End If
    End If
    If Not CheckDateRange(Sts_MsgLine, Txt_A0901s.text, Txt_A0901e.text) Then
       If ActiveControl.TabIndex = Txt_A0901s.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A0901s.TabIndex
       Else
          m_FieldError% = Txt_A0901e.TabIndex
          Txt_A0901e.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A0901e = True
End Function


Private Sub Txt_A0904e_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A0904e_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A0904e_GotFocus()
    TextHelpGotFocus
End Sub


Private Sub Txt_A0904e_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If KeyCode = KEY_F1 Then DataPrepare_A08 Txt_A0904e
End Sub


Private Sub Txt_A0904e_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If Fra_Help.Visible Then Exit Sub
    If TypeOf ActiveControl Is SSCommand Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0904e.TabIndex Then Exit Sub

    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0904e()
End Sub

Private Sub Txt_A0904s_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A0904s_KeyDown KEY_F1, 0
End Sub

Private Sub DataPrepare_A08(Txt As TextBox)
Dim A_Sql$

    Me.MousePointer = HOURGLASS

'�}�_�ɮ�
    A_Sql$ = "Select A0802,A0826 from A08"
    A_Sql$ = A_Sql$ & " where A0826<>''"
    A_Sql$ = A_Sql$ & " order by A0826"
    CreateDynasetODBC DB_ARTHGUI, DY_A08, A_Sql$, "DY_A08", True
    If DY_A08.BOF And DY_A08.EOF Then
       Me.MousePointer = Default
       Sts_MsgLine.Panels(1) = G_NoReference
       Exit Sub
    End If
    
    With Spd_Help

'�]�w���U����������ݩ�
         .UnitType = 2
         
'??? �]�w��Spread���U����D����ܼe��,�U���ݩʤ���ܦr��
'    �ѼƤ@ : Spread Name                                   �ѼƤG : �ѼƤ@���ݪ�Spead Type Name
'    �ѼƤT : �ۭq�����W��                                �Ѽƥ| : �]�w��e
'    �ѼƤ� : �w�]�������D                                �ѼƤ� : ��쪺��ƫ��A
'    �ѼƤC : �ƭ���쪺�U��                                �ѼƤK : �ƭ���쪺�U��
'    �ѼƤE : ��r��ƫ��A���̤j����                        �ѼƤQ : �����ܦbSpread�W������覡
'    �Ѽ�11 : �����ܦb����W������覡                    �Ѽ�12 : Database Name,�󦹸�Ʈw�U��MLabel��Caption
'    �Ѽ�13 : Field Name,�H������MLabel��Caption         �Ѽ�14 : Table Name,����U��MLabel��Caption
         Spread_Property Spd_Help, 0, UBound(tSpd_Help.Columns), WHITE, G_Font_Size, G_Font_Name
         SpdFldProperty Spd_Help, tSpd_Help, "A0826", TextWidth("X") * 9, GetCaption("PATTERNR", "userid", "User ID"), SS_CELL_TYPE_EDIT, "", "", 10, SS_CELL_H_ALIGN_CENTER
         SpdFldProperty Spd_Help, tSpd_Help, "A0802", TextWidth("X") * 10, GetCaption("PATTERNR", "username", "�ϥΪ�"), SS_CELL_TYPE_EDIT, "", "", 40

'�]�w��Spread���\Cell�����즲
         .AllowDragDrop = True
         
'���Spread���i�ק�
         .Row = -1: .Col = -1: .Lock = True
    
'�N����\�JSpread��
         Do While Not DY_A08.EOF
            .MaxRows = .MaxRows + 1
            
'??? �N��ƶ�JSpread��
'    �ѼƤ@ : Spread Name                               �ѼƤG : �ѼƤ@���ݪ�Spead Type Name
'    �ѼƤT : �ۭq�����W��                            �Ѽƥ| : ��ƦC
'    �ѼƤ� : ��J��
            SetSpdText Spd_Help, tSpd_Help, "A0826", .MaxRows, Trim(DY_A08.Fields("A0826") & "")
            SetSpdText Spd_Help, tSpd_Help, "A0802", .MaxRows, Trim(DY_A08.Fields("A0802") & "")
            DY_A08.MoveNext
         Loop
       
'??? �Q��Spread Type��Sort
       If tSpd_Help.SortEnable Then SpreadColsSort Spd_Help, tSpd_Help
    
'�]�w���U��������ܦ�m
         SetHelpWindowPos Fra_Help, Spd_Help, 1300, 120, 4265, 2085
         .Tag = Txt.TabIndex
         .SetFocus
    End With
    
    Me.MousePointer = Default
End Sub

Private Sub Txt_A0904s_GotFocus()
    TextHelpGotFocus
End Sub


Private Sub Txt_A0904s_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If KeyCode = KEY_F1 Then DataPrepare_A08 Txt_A0904s
End Sub

Private Sub Txt_A0904s_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If Fra_Help.Visible Then Exit Sub
    If TypeOf ActiveControl Is SSCommand Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0904s.TabIndex Then Exit Sub

    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0904s()
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


