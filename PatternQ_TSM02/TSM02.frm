VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_TSM02 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   Caption         =   "�|�p��إN�X��ƺ޲z"
   ClientHeight    =   4890
   ClientLeft      =   5520
   ClientTop       =   2880
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "TSM02.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   9135
   Begin VsOcxLib.VideoSoftElastic Vse_background 
      Height          =   4515
      Left            =   0
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   0
      Width           =   9135
      _Version        =   327680
      _ExtentX        =   16113
      _ExtentY        =   7964
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
      Picture         =   "TSM02.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "TSM02.frx":0326
      Begin VB.TextBox Txt_A0203 
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
         Left            =   5895
         MaxLength       =   12
         TabIndex        =   2
         Top             =   180
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0209 
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
         Left            =   7515
         MaxLength       =   8
         TabIndex        =   15
         Top             =   3915
         Width           =   1545
      End
      Begin VB.TextBox Txt_A0208 
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
         Left            =   4455
         MaxLength       =   8
         TabIndex        =   14
         Top             =   3915
         Width           =   1545
      End
      Begin VB.TextBox Txt_A0216 
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
         Left            =   1305
         MaxLength       =   8
         TabIndex        =   13
         Top             =   3915
         Width           =   1545
      End
      Begin VB.TextBox Txt_A0215 
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
         Left            =   5190
         MaxLength       =   20
         TabIndex        =   12
         Top             =   3465
         Width           =   2445
      End
      Begin VB.TextBox Txt_A0217 
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
         Left            =   1305
         MaxLength       =   20
         TabIndex        =   11
         Top             =   3465
         Width           =   2445
      End
      Begin VB.TextBox Txt_A0219 
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
         Left            =   5190
         MaxLength       =   15
         TabIndex        =   10
         Top             =   3015
         Width           =   2445
      End
      Begin VB.TextBox Txt_A0218 
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
         Left            =   1305
         MaxLength       =   15
         TabIndex        =   9
         Top             =   3015
         Width           =   2445
      End
      Begin VB.TextBox Txt_A0214 
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
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   8
         Text            =   " "
         Top             =   2565
         Width           =   6360
      End
      Begin VB.TextBox Txt_A0213 
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
         Left            =   1305
         MaxLength       =   50
         TabIndex        =   7
         Text            =   " "
         Top             =   2115
         Width           =   6360
      End
      Begin VB.TextBox Txt_A0207 
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
         Left            =   4545
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1620
         Width           =   3120
      End
      Begin VB.TextBox Txt_A0206 
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
         Left            =   1305
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1635
         Width           =   2085
      End
      Begin VB.TextBox Txt_A0201 
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
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   0
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox Txt_A0205 
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
         Left            =   1305
         MaxLength       =   40
         TabIndex        =   4
         Text            =   " "
         Top             =   1170
         Width           =   6360
      End
      Begin VB.TextBox Txt_A0204 
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
         Left            =   1305
         MaxLength       =   40
         TabIndex        =   3
         Top             =   675
         Width           =   6360
      End
      Begin VB.TextBox Txt_A0202 
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
         Left            =   3255
         MaxLength       =   12
         TabIndex        =   1
         Top             =   180
         Width           =   1770
      End
      Begin Threed.SSCommand cmd_ok 
         Height          =   405
         Left            =   7770
         TabIndex        =   19
         Top             =   1500
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�T�{F11"
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
      Begin Threed.SSCommand cmd_exit 
         Height          =   405
         Left            =   7770
         TabIndex        =   20
         Top             =   3375
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
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
      Begin Threed.SSCommand cmd_help 
         Height          =   405
         Left            =   7770
         TabIndex        =   16
         Top             =   150
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "����F1"
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
      Begin Threed.SSCommand Cmd_Next 
         Height          =   405
         Left            =   7770
         TabIndex        =   18
         Top             =   1050
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "����F8"
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
      Begin Threed.SSCommand Cmd_Previous 
         Height          =   405
         Left            =   7770
         TabIndex        =   17
         Top             =   600
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�e��F7"
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
      Begin VB.Label Lbl_A0203 
         Caption         =   "²��(�^)"
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
         Left            =   5085
         TabIndex        =   23
         Top             =   255
         Width           =   1470
      End
      Begin VB.Label Lbl_A0209 
         Caption         =   "�������"
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
         Left            =   6390
         TabIndex        =   36
         Top             =   4005
         Width           =   930
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
         Left            =   3330
         TabIndex        =   35
         Top             =   4005
         Width           =   930
      End
      Begin VB.Label Lbl_A0216 
         Caption         =   "�l���ϸ�"
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
         Left            =   60
         TabIndex        =   34
         Top             =   4005
         Width           =   1470
      End
      Begin VB.Label Lbl_A0215 
         Caption         =   "����"
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
         Left            =   4080
         TabIndex        =   33
         Top             =   3540
         Width           =   1470
      End
      Begin VB.Label Lbl_A0217 
         Caption         =   "��a"
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
         Left            =   60
         TabIndex        =   32
         Top             =   3555
         Width           =   1470
      End
      Begin VB.Label Lbl_A0219 
         Caption         =   "�s���ǯu"
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
         Left            =   4095
         TabIndex        =   31
         Top             =   3090
         Width           =   1470
      End
      Begin VB.Label Lbl_A0206 
         Caption         =   "�t�d�H(��)"
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
         Left            =   60
         TabIndex        =   26
         Top             =   1710
         Width           =   1470
      End
      Begin VB.Label Lbl_A0202 
         Caption         =   "²��(��)"
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
         Left            =   2400
         TabIndex        =   22
         Top             =   270
         Width           =   1470
      End
      Begin VB.Label Lbl_A0214 
         Caption         =   "�^��a�}"
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
         Left            =   60
         TabIndex        =   29
         Top             =   2640
         Width           =   1470
      End
      Begin VB.Label Lbl_A0218 
         Caption         =   "�s���q��"
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
         Left            =   60
         TabIndex        =   30
         Top             =   3105
         Width           =   1470
      End
      Begin VB.Label Lbl_A0207 
         Caption         =   "(�^)"
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
         Left            =   4005
         TabIndex        =   27
         Top             =   1710
         Width           =   615
      End
      Begin VB.Label Lbl_A0213 
         Caption         =   "����a�}"
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
         Left            =   60
         TabIndex        =   28
         Top             =   2190
         Width           =   1470
      End
      Begin VB.Label Lbl_A0205 
         Caption         =   "�^����W"
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
         Left            =   60
         TabIndex        =   25
         Top             =   1245
         Width           =   1470
      End
      Begin VB.Label Lbl_A0204 
         Caption         =   "������W"
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
         Left            =   60
         TabIndex        =   24
         Top             =   750
         Width           =   1470
      End
      Begin VB.Label Lbl_A0201 
         Caption         =   "�����N��"
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
         Left            =   60
         TabIndex        =   21
         Top             =   285
         Width           =   1470
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   38
      Top             =   4515
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_TSM02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================
' Module    : frm_TSM02
' Author    : Mike_chang
' Date      : 2015/8/28
' Purpose   :
'========================================================================
Option Explicit
Option Compare Text

'========================================================================
'   Coding Rules
'========================================================================
'�b���B�w�q���Ҧ��ܼ�, �@�ߥHM�}�Y, �pM_AAA$, M_BBB#, M_CCC&
'�B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
' $: ��r
' #: �Ҧ��Ʀr�B��(���B�μƶq)
' &: �{���j���ܼ�
' %: ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
' �ť�: �N��VARIENT, �ʺA�ܼ�
'========================================================================

'�۩w�ܼ�
'Dim m_A1501Flag%
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

'���n�ܼ�
Dim m_FieldError%    '���ܼƦb�P�_���O�_���~, �����^�����줧�ʧ@
Dim m_ExitTrigger%   '���ܼƦb�P�_������O�_�QĲ�o, �N����ثe���b�B�z���@�~
'Dim m_RecordChange% '���ܼƦb�P�_��ƬO�_������, �����N���ܬO�_�s�ɰT��
Dim m_TabGotFocus%   '����Tab_ClickAfter �u�B�z�@��
Dim m_TabMouseDown%  '�����Help Control DblClick��Ĳ�o��Tab_ClickAfter���B�z
'========================================================================

'====================================
'   User Defined Functions
'====================================

'========================================================================
' Module    : frm_TSM02
' Procedure : Set_Property
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Setup Properties
' Details   : 1. Form(caption, font, color)
'             2. Label(caption, font, color)
'             3. TextBox(font, maxlength)
'             4. command buttom(caption, font)
'========================================================================
Private Sub Set_Property()
    Me.FontBold = False
    
'�]�w��Form�����D,�r�ΤΦ�t
    Form_Property Me, G_Form_TSM02, G_Font_Name

'�]Form���Ҧ�Label�����D, �r�ΤΦ�t
    Label_Property Lbl_A0201, G_Pnl_A0201$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0202, G_Pnl_A0202$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0203, G_Pnl_A0203$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0204, G_Pnl_A0204$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0205, G_Pnl_A0205$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0206, G_Pnl_A0206$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0207, G_Pnl_A0207$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0208, G_Pnl_A0208$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0209, G_Pnl_A0209$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0213, G_Pnl_A0213$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0214, G_Pnl_A0214$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0215, G_Pnl_A0215$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0216, G_Pnl_A0216$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0217, G_Pnl_A0217$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0218, G_Pnl_A0218$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0219, G_Pnl_A0219$, G_Label_Color, G_Font_Size, G_Font_Name
    
'�]Form���Ҧ�TextBox���r�ΤΥi��J����
    Text_Property Txt_A0201, 6, G_Font_Name
    Text_Property Txt_A0202, 12, G_Font_Name
    Text_Property Txt_A0203, 12, G_Font_Name
    Text_Property Txt_A0204, 40, G_Font_Name
    Text_Property Txt_A0205, 40, G_Font_Name
    Text_Property Txt_A0206, 12, G_Font_Name
    Text_Property Txt_A0207, 40, G_Font_Name
    Text_Property Txt_A0213, 50, G_Font_Name
    Text_Property Txt_A0214, 50, G_Font_Name
    Text_Property Txt_A0218, 15, G_Font_Name
    Text_Property Txt_A0219, 15, G_Font_Name
    Text_Property Txt_A0217, 20, G_Font_Name
    Text_Property Txt_A0215, 20, G_Font_Name
    Text_Property Txt_A0216, 8, G_Font_Name
    Text_Property Txt_A0208, 8, G_Font_Name
    Text_Property Txt_A0209, 8, G_Font_Name

'�]Form���Ҧ�Command�����D�Φr��
    Command_Property cmd_help, G_CmdHelp, G_Font_Name
    Command_Property Cmd_Previous, G_CmdPrevious, G_Font_Name
    Command_Property Cmd_Next, G_CmdNext, G_Font_Name
    Command_Property cmd_ok, G_CmdOk, G_Font_Name
    Command_Property cmd_exit, G_CmdExit, G_Font_Name

    
    
'�H�U���зǫ��O, ���o�ק�
    VSElastic_Property Vse_background
    StatusBar_ProPerty Sts_MsgLine
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : SetCommand
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Setup command buttom enables
' Details   : Pkey(A0201) is only allow to be insert while Adding record,
'             Otherwise, at updating & delete, it's no meaning to change
'             Pkey since it doesn't stand for a specific literal meaning
'========================================================================
Sub SetCommand()
'�]�w�C�@�@�~���A�U, CONTROL�O�_�i�@��
    Select Case G_AP_STATE
        Case G_AP_STATE_ADD
            'while Adding, Pkey(A0201) is allowed to input
            cmd_help.Enabled = True
            Cmd_Previous.Enabled = False
            Cmd_Next.Enabled = False
            cmd_ok.Enabled = True
            cmd_exit.Enabled = True
            Txt_A0201.Enabled = True
        Case G_AP_STATE_UPDATE
            'while update, no meaning to change Pkey
            cmd_help.Enabled = True
            Cmd_Previous.Enabled = (G_ActiveRow# > 1)
            Cmd_Next.Enabled = (G_ActiveRow# < G_MaxRows#)
            cmd_ok.Enabled = True
            cmd_exit.Enabled = True
            Txt_A0201.Enabled = False
        Case G_AP_STATE_DELETE
            'while delete, no meaning to change Pkey
            cmd_help.Enabled = True
            Cmd_Previous.Enabled = (G_ActiveRow# > 1)
            Cmd_Next.Enabled = (G_ActiveRow# < G_MaxRows#)
            cmd_ok.Enabled = True
            cmd_exit.Enabled = True
            Txt_A0201.Enabled = False
     End Select
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : SetButtonEnable
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Set All Command Buttom to FALSE or restore previous
' Details   : If Set True, Store Current Enable state and set to false
'             Else restore from tag to previous state
'========================================================================
Sub SetButtonEnable(ByVal A_Enable%)
    If Not A_Enable% Then
       Vse_background.TabStop = True
       Cmd_Previous.Tag = Cmd_Previous.Enabled
       Cmd_Next.Tag = Cmd_Next.Enabled
       cmd_ok.Tag = cmd_ok.Enabled
       cmd_exit.Tag = cmd_exit.Enabled
       
       Cmd_Previous.Enabled = A_Enable%
       Cmd_Next.Enabled = A_Enable%
       cmd_ok.Enabled = A_Enable%
       cmd_exit.Enabled = A_Enable%
    Else
       Cmd_Previous.Enabled = CBool(Cmd_Previous.Tag)
       Cmd_Next.Enabled = CBool(Cmd_Next.Tag)
       cmd_ok.Enabled = CBool(cmd_ok.Tag)
       cmd_exit.Enabled = CBool(cmd_exit.Tag)
    End If
End Sub


'========================================================================
' Module    : frm_TSM02
' Procedure : IsKeyExist
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Check Primary Key Existed in DB
' Details   :
'========================================================================
Private Function IsKeyExist() As Boolean
On Local Error GoTo My_Error
Dim A_Sql$

    IsKeyExist = False
    
    A_Sql$ = "Select A0201 From A02"
    A_Sql$ = A_Sql$ & " where A0201='" & Trim(Txt_A0201) & "'"
    A_Sql$ = A_Sql$ & " Order by A0201"
    
    CreateDynasetODBC DB_ARTHGUI, DY_A021, A_Sql$, "DY_A021", True
    If Not (DY_A021.BOF And DY_A021.EOF) Then IsKeyExist = True
    Exit Function
    
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

'========================================================================
' Module    : frm_TSM02
' Procedure : IsRecordChange
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Detect if data changed by Global Variable "G_DataChange"
' Details   : set FALSE while Deleting
'========================================================================
Function IsRecordChange() As Boolean
'�Y�@�~���A���R���h����Check
    If G_AP_STATE = G_AP_STATE_DELETE Then
       IsRecordChange = False
       Exit Function
    End If
       
'�P�_Record��ƬO�_����
    IsRecordChange = G_DataChange%
End Function

'========================================================================
' Module    : frm_TSM02
' Procedure : IsAllFieldsCheck
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Boolean Function checking whether all txtBox pass value check
' Details   :
'========================================================================
Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    
'����s�ɫe���N�Ҧ��ˮ����A���@��
    If G_AP_STATE = G_AP_STATE_ADD Then
        If Not CheckRoutine_A0201() Then Exit Function
    End If
    If Not CheckRoutine_A0202() Then Exit Function
    If Not CheckRoutine_A0203() Then Exit Function
    If Not CheckRoutine_A0208() Then Exit Function
    If Not CheckRoutine_A0209() Then Exit Function
    
    IsAllFieldsCheck = True
End Function

'========================================================================
' Module    : frm_TSM02
' Procedure : OpenMainFile
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Open Dynaset "DY_A02"
' Details   :
'========================================================================
Private Sub OpenMainFile()
On Local Error GoTo My_Error
Dim A_Sql$

    'A02
    A_Sql$ = "SELECT * FROM A02"
    A_Sql$ = A_Sql$ & " where A0201='" & G_A0201$ & "'"
    A_Sql$ = A_Sql$ & " order by A0201"
    CreateDynasetODBC DB_ARTHGUI, DY_A02, A_Sql$, "DY_A02", True
    
'    'A15
'    A_Sql$ = "SELECT * FROM A15"
'    A_Sql$ = A_Sql$ & " where A1501='" & G_A0201$ & "'"
'    A_Sql$ = A_Sql$ & " and A1502='" & Mid$(G_A1502$, 1, 4) & "'"
'    A_Sql$ = A_Sql$ & " and A1503='" & Mid$(G_A1502$, 5) & "'"
'    A_Sql$ = A_Sql$ & " order by A1501,A1502,A1503"
'    CreateDynasetODBC DB_ARTHGUI, DY_A02, A_Sql$, "DY_A02", True
'    'A14
'    A_Sql$ = "SELECT * FROM A14"
'    A_Sql$ = A_Sql$ & " where A1406='" & G_A0201$ & "'"
'    A_Sql$ = A_Sql$ & " and A1402='" & Mid$(G_A1502$, 1, 4) & "'"
'    A_Sql$ = A_Sql$ & " and A1403='" & Mid$(G_A1502$, 5) & "'"
'    A_Sql$ = A_Sql$ & " order by A1401,A1406,A1402,A1403"
'    CreateDynasetODBC DB_ARTHGUI, DY_A14, A_Sql$, "DY_A14", True
'    'A20
'    A_Sql$ = "SELECT * FROM A20"
'    A_Sql$ = A_Sql$ & " where A2001='" & G_A0201$ & "'"
'    A_Sql$ = A_Sql$ & " and A2002='" & G_A1502$ & "'"
'    A_Sql$ = A_Sql$ & " order by A2001,A2002,A2003"
'    CreateDynasetODBC DB_ARTHGUI, DY_A20, A_Sql$, "DY_A20", True
    Exit Sub

My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Move2Menu
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   :
' Details   :
'========================================================================
Private Sub Move2Menu()
'�N���ʸ��UPDATE�^V�e����SPREAD�W
    With Frm_TSM02v!Spd_TSM02v
         If G_AP_STATE = G_AP_STATE_UPDATE Then
            .Row = G_ActiveRow#
         Else
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .Action = SS_ACTION_ACTIVE_CELL
         End If
         .Col = 1
         .text = Trim$(Txt_A0201 & "")
         .Col = 2
         .text = Trim$(Txt_A0202 & "")
         .Col = 3
         .text = Trim$(Txt_A0206 & "")
         .Col = 4
         .text = DateFormat2(Trim$(Txt_A0208 & ""))
         .Col = 5
         .text = Trim$(Txt_A0218 & "")
    End With
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : MoveDB2Field
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   :
' Details   :
'========================================================================
Private Sub MoveDB2Field()
On Local Error GoTo My_Error

'�N��ظ����ܦܵe���W
    Txt_A0201.text = Trim$(DY_A02.Fields("A0201") & "")
    Txt_A0202.text = Trim$(DY_A02.Fields("A0202") & "")
    Txt_A0203.text = Trim$(DY_A02.Fields("A0203") & "")
    Txt_A0204.text = Trim$(DY_A02.Fields("A0204") & "")
    Txt_A0205.text = Trim$(DY_A02.Fields("A0205") & "")
    Txt_A0206.text = Trim$(DY_A02.Fields("A0206") & "")
    Txt_A0207.text = Trim$(DY_A02.Fields("A0207") & "")
    Txt_A0208.text = Trim$(DateOut(Trim$(DY_A02.Fields("A0208") & "")))
    Txt_A0209.text = Trim$(DateOut(Trim$(DY_A02.Fields("A0209") & "")))
    Txt_A0213.text = Trim$(DY_A02.Fields("A0213") & "")
    Txt_A0214.text = Trim$(DY_A02.Fields("A0214") & "")
    Txt_A0215.text = Trim$(DY_A02.Fields("A0215") & "")
    Txt_A0216.text = Trim$(DY_A02.Fields("A0216") & "")
    Txt_A0217.text = Trim$(DY_A02.Fields("A0217") & "")
    Txt_A0218.text = Trim$(DY_A02.Fields("A0218") & "")
    Txt_A0219.text = Trim$(DY_A02.Fields("A0219") & "")

    Exit Sub

My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub



'========================================================================
' Module    : frm_TSM02
' Procedure : MoveField2DB
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   :
' Details   :
'========================================================================
Private Sub MoveField2DB()
On Local Error GoTo My_Error

    G_Str = ""
    If G_AP_STATE = G_AP_STATE_ADD Then
       InsertFields "A02001", GetCurrentDate(), G_Data_String   'G_Data_Numeric
       InsertFields "A02002", GetCurrentTime(), G_Data_String
       InsertFields "A02003", GetWorkStation(), G_Data_String
       InsertFields "A02004", GetUserId(), G_Data_String
       InsertFields "A02005", " ", G_Data_String
       InsertFields "A02006", " ", G_Data_String
       InsertFields "A02007", " ", G_Data_String
       InsertFields "A02008", " ", G_Data_String
       
       InsertFields "A0201", Trim(Txt_A0201.text & ""), G_Data_String
       InsertFields "A0202", Trim(Txt_A0202.text & ""), G_Data_String
       InsertFields "A0203", Trim(Txt_A0203.text), G_Data_String
       InsertFields "A0204", Trim(Txt_A0204.text), G_Data_String
       InsertFields "A0205", Trim(Txt_A0205.text), G_Data_String
       InsertFields "A0206", Trim(Txt_A0206.text), G_Data_String
       InsertFields "A0207", Trim(Txt_A0207.text), G_Data_String
       InsertFields "A0208", Trim(DateIn(Trim(Txt_A0208.text))), G_Data_String
       InsertFields "A0209", Trim(DateIn(Trim(Txt_A0209.text))), G_Data_String
       InsertFields "A0213", Trim(Txt_A0213.text), G_Data_String
       InsertFields "A0214", Trim(Txt_A0214.text), G_Data_String
       InsertFields "A0215", Trim(Txt_A0215.text), G_Data_String
       InsertFields "A0216", Trim(Txt_A0216.text), G_Data_String
       InsertFields "A0217", Trim(Txt_A0217.text), G_Data_String
       InsertFields "A0218", Trim(Txt_A0218.text), G_Data_String
       InsertFields "A0219", Trim(Txt_A0219.text), G_Data_String
       
       SQLInsert DB_ARTHGUI, "A02"
    Else
       UpdateString "A02005", GetCurrentDate(), G_Data_String
       UpdateString "A02006", GetCurrentTime(), G_Data_String
       UpdateString "A02007", GetWorkStation(), G_Data_String
       UpdateString "A02008", GetUserId(), G_Data_String
       
       UpdateString "A0201", Trim(Txt_A0201.text), G_Data_String
       UpdateString "A0202", Trim(Txt_A0202.text), G_Data_String
       UpdateString "A0203", Trim(Txt_A0203.text), G_Data_String
       UpdateString "A0204", Trim(Txt_A0204.text), G_Data_String
       UpdateString "A0205", Trim(Txt_A0205.text), G_Data_String
       UpdateString "A0206", Trim(Txt_A0206.text), G_Data_String
       UpdateString "A0207", Trim(Txt_A0207.text), G_Data_String
       UpdateString "A0208", Trim(DateIn(Trim(Txt_A0208.text))), G_Data_String
       UpdateString "A0209", Trim(DateIn(Trim(Txt_A0209.text))), G_Data_String
       UpdateString "A0213", Trim(Txt_A0213.text), G_Data_String
       UpdateString "A0214", Trim(Txt_A0214.text), G_Data_String
       UpdateString "A0215", Trim(Txt_A0215.text), G_Data_String
       UpdateString "A0216", Trim(Txt_A0216.text), G_Data_String
       UpdateString "A0217", Trim(Txt_A0217.text), G_Data_String
       UpdateString "A0218", Trim(Txt_A0218.text), G_Data_String
       UpdateString "A0219", Trim(Txt_A0219.text), G_Data_String
       
       G_Str = G_Str & " where A0201='" & G_A0201$ & "'"
       
       SQLUpdate DB_ARTHGUI, "A02"
    End If
    
    Exit Sub
    
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : SaveCheck
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   :
' Details   :
'========================================================================
Function SaveCheck(Optional A_PassQuestion% = False) As Boolean
    SaveCheck = False
    
    '�s�W�ק窱�A�U,���������e��,�Y�e����Ʀ����ʮ�,�߰ݬO�_�s��.
    'retcode = IDYES     , �s�ɫ��^V�e��
    'retcode = IDNO      , ���s�ɪ�^V�e��
    'retcode = IDCANCEL  , ���s�ɰ��d�b��e��
    If A_PassQuestion% Then
    '���T�{��n���s�ɰʧ@
        retcode = IDYES
    Else
    '��������ɥ�User�M�w�n�B�z���ʧ@
        retcode = MsgBox(G_Save_Check, vbYesNoCancel + vbQuestion, Me.Caption)
    End If
    
    If retcode = IDCANCEL Then
       Exit Function
    ElseIf retcode = IDYES Then
       If Not IsAllFieldsCheck() Then Exit Function
       Me.Refresh
       MoveField2DB
       Move2Menu
    End If
    
    SaveCheck = True
End Function

'========================================================================
' Module    : frm_TSM02
' Procedure : CheckRoutine_A0201
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Check Pkey valid
' Details   : Check: 1.Pkey not empty  2.Pkey not duplicate
'========================================================================
Private Function CheckRoutine_A0201() As Boolean
    CheckRoutine_A0201 = True
    m_FieldError% = -1
    
'�ˮָ����O�_��J
    If Txt_A0201.text = "" Then
        Sts_MsgLine.Panels(1) = G_Pnl_A0201$ & G_MustInput
        CheckRoutine_A0201 = False
        
        G_DataChange% = True
        
        m_FieldError% = Txt_A0201.TabIndex
        Txt_A0201.SetFocus
        Exit Function
    End If

'�ˮָ�ƬO�_�w�s�b
    If G_AP_STATE = G_AP_STATE_ADD Then
        If IsKeyExist() Then
             Sts_MsgLine.Panels(1) = G_Pnl_A0201$ & G_RecordExist
             CheckRoutine_A0201 = False
             G_DataChange% = True
             m_FieldError% = Txt_A0201.TabIndex
             Txt_A0201.SetFocus
        End If
    End If
End Function


'========================================================================
' Module    : frm_TSM02
' Procedure : CheckRoutine_A0202
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Check MustInput Constraint
' Details   :
'========================================================================
Private Function CheckRoutine_A0202() As Boolean
    CheckRoutine_A0202 = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Txt_A0202.text = "" Then
       Sts_MsgLine.Panels(1) = G_Pnl_A0202$ & G_MustInput
       m_FieldError% = Txt_A0202.TabIndex
       Txt_A0202.SetFocus
       Exit Function
    End If
       
    CheckRoutine_A0202 = True
End Function

'========================================================================
' Module    : frm_TSM02
' Procedure : CheckRoutine_A0203
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Check MustInput Constraint
' Details   :
'========================================================================
Private Function CheckRoutine_A0203() As Boolean
    CheckRoutine_A0203 = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Txt_A0203.text = "" Then
       Sts_MsgLine.Panels(1) = G_Pnl_A0203$ & G_MustInput
       m_FieldError% = Txt_A0203.TabIndex
       Txt_A0203.SetFocus
       Exit Function
    End If
       
    CheckRoutine_A0203 = True
End Function


'========================================================================
' Module    : frm_TSM02
' Procedure : CheckRoutine_A0208
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Function CheckRoutine_A0208() As Boolean
    CheckRoutine_A0208 = False
    
'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd

    'Check Date Formate Valid
    If Trim(Txt_A0208) <> "" Then
        If Not IsDateValidate(Txt_A0208) Then
            Sts_MsgLine.Panels(1) = G_Pnl_A0208$ & G_DateError
            m_FieldError% = Txt_A0208.TabIndex
            Txt_A0208.SetFocus
        End If
    End If
    
    'If the Apartment isn't dismiss, A0209 will be empty
    'As A0209 is empty, no need to check date range
    If Trim(Txt_A0209) = "" Then
        CheckRoutine_A0208 = True
        Exit Function
    End If
    
    'Check Data Range Not Exceed
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A0208), Trim$(Txt_A0209)) Then
        'Check whether Entering End Date
        If ActiveControl.TabIndex = Txt_A0209.TabIndex Then
            m_FieldError% = Txt_A0208.TabIndex
        Else
            m_FieldError% = Txt_A0208.TabIndex
            Txt_A0208.SetFocus
        End If
        'On Error & Exit
        Exit Function
    End If
    
    CheckRoutine_A0208 = True
End Function

'========================================================================
' Module    : frm_TSM02
' Procedure : CheckRoutine_A0209
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Function CheckRoutine_A0209() As Boolean
    CheckRoutine_A0209 = False
    
'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
    'Check Date Formate Valid
    If Trim(Txt_A0209) <> "" Then
        If Not IsDateValidate(Txt_A0208) Then
            Sts_MsgLine.Panels(1) = G_Pnl_A0209$ & G_DateError
            m_FieldError% = Txt_A0209.TabIndex
            Txt_A0209.SetFocus
        End If
    End If
    
    'If A0208 is empty, no need to check date range
    If Trim(Txt_A0208) = "" Then
        CheckRoutine_A0209 = True
        Exit Function
    End If
    
    'Check Data Range Not Exceed
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A0208), Trim$(Txt_A0209)) Then
        'Check whether Entering End Date
        If ActiveControl.TabIndex = Txt_A0209.TabIndex Then
            m_FieldError% = Txt_A0209.TabIndex
        Else
            m_FieldError% = Txt_A0209.TabIndex
            Txt_A0209.SetFocus
        End If
        'On Error & Exit
        Exit Function
    End If
    
    
    CheckRoutine_A0209 = True
End Function

'========================================================================
' Module    : frm_TSM02
' Procedure : ClearFieldsValue
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub ClearFieldsValue()
'�N���ȲM��
    Txt_A0201.text = ""
    Txt_A0201.Tag = ""
    Txt_A0202.text = ""
    Txt_A0203.text = ""
    Txt_A0204.text = ""
    Txt_A0205.text = ""
    Txt_A0206.text = ""
    Txt_A0207.text = ""
    Txt_A0213.text = ""
    Txt_A0214.text = ""
    Txt_A0218.text = ""
    Txt_A0219.text = ""
    Txt_A0217.text = ""
    Txt_A0215.text = ""
    Txt_A0216.text = ""
    Txt_A0208.text = ""
    Txt_A0209.text = ""

End Sub


'========================================================================
' Module    : frm_TSM02
' Procedure : Delete_From_Menu
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Delete_From_Menu()
'�NV�e���W���ӵ���ƦC�R��
    With Frm_TSM02v!Spd_TSM02v
        .Row = G_ActiveRow#
        .Action = SS_ACTION_DELETE_ROW
        .MaxRows = .MaxRows - 1
    End With
End Sub


'========================================================================
' Module    : frm_TSM02
' Procedure : Delete_Process_A02
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Do Deletion by Pkey: A0201
' Details   :
'========================================================================
Private Sub Delete_Process_A02()
On Local Error GoTo My_Error

    G_Str = "DELETE FROM A02"
    G_Str = G_Str & " WHERE A0201='" & G_A0201$ & "'"
    ExecuteProcess DB_ARTHGUI, G_Str
    Exit Sub
    
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub


'====================================
'   Command Buttom Events
'====================================

'========================================================================
' Module    : frm_TSM02
' Procedure : Cmd_Help_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : call HLP file
' Details   :
'========================================================================
Private Sub Cmd_Help_Click()
Dim a$

    a$ = "notepad " + G_Help_Path + "PATTERNQ.HLP"
    retcode = Shell(a$, 4)
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Cmd_Next_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Move to next record of the V-form
' Details   :
'========================================================================
Private Sub Cmd_Next_Click()
'�L�U�@����Ƥ����B�z
    If G_ActiveRow# >= G_MaxRows# Then
       Sts_MsgLine.Panels(1) = G_AP_NONEXT
       Exit Sub
    End If
    
    Me.MousePointer = HOURGLASS
    
'�]�w�|�v�T��Ʀs�ɪ��Ҧ�Button��Enabled Property = False
    SetButtonEnable False
    
'�Y�ثeRecord��Ʀ�����, ���ܬO�_�s��
    If IsRecordChange() Then
        If SaveCheck() = False Then
            'If Dialog's cancel buttom click:
            Me.MousePointer = Default
            SetButtonEnable True
            Txt_A0202.SetFocus
            Exit Sub
        End If
    End If

'���o�U�@����ƪ�P-KEY
    With Frm_TSM02v!Spd_TSM02v
         G_ActiveRow# = G_ActiveRow# + 1
         .Row = G_ActiveRow#
'         .Col = 1: StrCut Trim$(.text), Space(1), G_A0201$, ""
         .Col = 1: G_A0201$ = Trim$(.text)
        
'�NV�e������в���U�@��
         .Action = SS_ACTION_ACTIVE_CELL
    End With
    
'�a�X�U�@�����
    OpenMainFile
    ClearFieldsValue
    MoveDB2Field
    G_DataChange% = False
    
'�٭�Ҧ�Button��Enabled Property
    SetButtonEnable True
    
    SetCommand
    Txt_A0202.SetFocus
    Me.MousePointer = Default
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Cmd_Previous_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Move to Previous record of V-Form
' Details   :
'========================================================================
Private Sub Cmd_Previous_Click()
'�L�W�@����Ƥ����B�z
    If G_ActiveRow# <= 1 Then
       Sts_MsgLine.Panels(1) = G_AP_NOPRVS
       Exit Sub
    End If
    Me.MousePointer = HOURGLASS
    
'�]�w�|�v�T��Ʀs�ɪ��Ҧ�Button��Enabled Property = False
    SetButtonEnable False
    
'�Y�ثeRecord��Ʀ�����, ���ܬO�_�s��
    If IsRecordChange() Then
       If SaveCheck() = False Then
          SetButtonEnable True
          Me.MousePointer = Default
          Txt_A0202.SetFocus
          Exit Sub
       End If
    End If
    
'���o�W�@����ƪ�P-KEY
    With Frm_TSM02v!Spd_TSM02v
         G_ActiveRow# = G_ActiveRow# - 1
         .Row = G_ActiveRow#
'         .Col = 1: StrCut Trim$(.text), Space(1), G_A0201$, ""
         .Col = 1: G_A0201$ = Trim$(.text)
        
'�NV�e������в���U�@��
         .Action = SS_ACTION_ACTIVE_CELL
    End With
    
'�a�X�W�@�����
    OpenMainFile
    ClearFieldsValue
    MoveDB2Field
    G_DataChange% = False
    
'�٭�Ҧ�Button��Enabled Property
    SetButtonEnable True
    
    SetCommand
    Txt_A0202.SetFocus
    Me.MousePointer = Default
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Cmd_Ok_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Do Insert ,Update Or Delete
' Details   :
'========================================================================
Private Sub Cmd_Ok_Click()
    Me.MousePointer = HOURGLASS
    
'�]�w�|�v�T��Ʀs�ɪ��Ҧ�Button��Enabled Property = False
    SetButtonEnable False
    
'�̨C�ӧ@�~���A���U�O���B�z
    Select Case G_AP_STATE
      Case G_AP_STATE_ADD
           If SaveCheck(True) = False Then
              SetButtonEnable True
              Me.MousePointer = Default
              Exit Sub
           End If
           Txt_A0201.text = ""
           Sts_MsgLine.Panels(1) = G_Add_Ok
           If frm_TSM02.Visible Then Txt_A0201.SetFocus

      Case G_AP_STATE_UPDATE
           If IsRecordChange() Then
              If SaveCheck(True) = False Then
                 SetButtonEnable True
                 Me.MousePointer = Default
                 Exit Sub
              End If
              Sts_MsgLine.Panels(1) = G_Update_Ok
           End If

      Case G_AP_STATE_DELETE
            Delete_Process_A02
            Delete_From_Menu
            Sts_MsgLine.Panels(1) = G_Delete_Ok
    End Select
    G_DataChange% = False
    
'�٭�Ҧ�Button��Enabled Property
    SetButtonEnable True
    
    Me.MousePointer = Default

'�@�~���A�Y���ק�,�R��, �h��^V�e��
    If G_AP_STATE <> G_AP_STATE_ADD Then
       DoEvents
       Me.Hide
       Frm_TSM02v.Show
    End If
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Cmd_Exit_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Cmd_Exit_Click()
    Me.MousePointer = HOURGLASS

'�����ثe����,���X��L�B�z�{��
    m_ExitTrigger% = True
    
'�]�w�|�v�T��Ʀs�ɪ��Ҧ�Button��Enabled Property = False
    SetButtonEnable False
    
'�Y��Ʀ�����, ���ܬO�_�n�s��
    If IsRecordChange() Then
       If SaveCheck() = False Then
          SetButtonEnable True
          Me.MousePointer = Default
          Exit Sub
       End If
    End If

'�٭�Ҧ�Button��Enabled Property
    SetButtonEnable True

'���åثe�e��, ���V�e��
    DoEvents
    Me.Hide
    Frm_TSM02v.Show
    Me.MousePointer = Default
End Sub

'====================================
'   Form Events
'====================================

'========================================================================
' Module    : frm_TSM02
' Procedure : Form_Activate
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   :
' Details   :
'========================================================================
Private Sub Form_Activate()
    Me.MousePointer = HOURGLASS
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    Sts_MsgLine.Panels(1) = G_Process
    Me.Refresh
    
'Initial Form�������n�ܼ�
    m_FieldError% = -1
    m_ExitTrigger% = False
    G_DataChange% = False
    
'�P�_�O�_�Ѩ�L���U�e���^��, �ӫD��������
    If Trim(G_FormFrom$) <> "" Then
       Me.MousePointer = Default
       G_FormFrom$ = ""
       '.....                '�[�J�ҭn�]�w���ʧ@
       '.....
       Exit Sub
    Else
       '.....                '�Ĥ@������ɤ��ǳưʧ@
       '.....
       
        ClearFieldsValue
       
        Select Case G_AP_STATE
            Case G_AP_STATE_ADD
                'while from Q to D
                
            Case G_AP_STATE_UPDATE, G_AP_STATE_DELETE
                'while from V to D
                OpenMainFile
                MoveDB2Field
        End Select
        
        SetCommand          'set command buttom according to State
        
        If G_AP_STATE = G_AP_STATE_ADD Then
            If frm_TSM02.Visible Then Txt_A0201.SetFocus
        Else
            If frm_TSM02.Visible Then Txt_A0202.SetFocus
        End If
        Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    End If
    
    '�NForm��m��ù������h
    frm_TSM02.ZOrder 0
    Me.MousePointer = Default
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Form_KeyDown
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Handle Key Event
' Details   : Handling: F1����, F7�W��, F8�U��, F11�T�{, ESC���}
'========================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
           Case KEY_F1
                KeyCode = 0
                If cmd_help.Visible = True And cmd_help.Enabled = True Then
                   cmd_help.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F7
                KeyCode = 0
                If Cmd_Previous.Visible = True And Cmd_Previous.Enabled = True Then
                   Cmd_Previous.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F8
                KeyCode = 0
                If Cmd_Next.Visible = True And Cmd_Next.Enabled = True Then
                   Cmd_Next.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F11
                KeyCode = 0
                If cmd_ok.Visible = True And cmd_ok.Enabled = True Then
                   cmd_ok.SetFocus
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

'========================================================================
' Module    : frm_TSM02
' Procedure : Form_KeyPress
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Manage Uppercase input(A0201), and Record data changed
' Details   :
'========================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    
'�D�ʱN��ƿ�J�Ѥp�g�ର�j�g
'  �Y���Y����줣�ݭn�ഫ��, �����H���L
'    If ActiveControl.TabIndex = Txt_A0204.TabIndex Then GoTo Form_KeyPress_A
'    If ActiveControl.TabIndex = Txt_A0205.TabIndex Then GoTo Form_KeyPress_A
'    If ActiveControl.TabIndex = Txt_A1514.TabIndex Then GoTo Form_KeyPress_A
'    If ActiveControl.TabIndex = txt_xxx.TabIndex Then GoTo Form_KeyPress_A
    If ActiveControl.TabIndex <> Txt_A0201.TabIndex Then GoTo Form_KeyPress_A
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Form_KeyPress_A:
'��J���N�r��(ENTER���~), �N��Ʋ����ܼƳ]��TRUE
    If Not TypeOf ActiveControl Is SSCommand Then
       If KeyAscii <> KEY_RETURN Then G_DataChange% = True
    End If

    'If ActiveControl.TabIndex <> Spd_PATTERNM.TabIndex Then
       KeyPress KeyAscii           'Enter�ɦ۰ʸ���U�@���, spread���~
    'End If
    
'�R���@�~�U, �NKeyBoard���, ��������Ʋ���
    If G_AP_STATE = G_AP_STATE_DELETE Then KeyAscii = 0
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Form_Load
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Form_Load()
    FormCenter Me
    Set_Property
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Form_Unload
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    If cmd_exit.Enabled Then cmd_exit.SetFocus: Cmd_Exit_Click
End Sub

'====================================
'   TextBox Evnets
'====================================

'========================================================================
' Module    : frm_TSM02
' Procedure : Txt_A0201_GotFocus
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A0201_GotFocus()
    TextGotFocus
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Txt_A0201_LostFocus
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A0201_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = cmd_exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0202.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0201()
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Txt_A0202_GotFocus
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A0202_GotFocus()
    TextGotFocus
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Txt_A0202_LostFocus
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A0202_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = cmd_exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0202.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0202()
End Sub

'========================================================================
' Module    : frm_TSM02
' Procedure : Txt_A0203_GotFocus
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A0203_GotFocus()
    TextGotFocus
End Sub
'========================================================================
' Module    : frm_TSM02
' Procedure : Txt_A0203_LostFocus
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A0203_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = cmd_exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0203.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0203()
End Sub

Private Sub Txt_A0204_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0204_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0205_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0205_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0206_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0206_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0207_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0207_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0213_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0213_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0214_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0214_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0218_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0218_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0219_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0219_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0217_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0217_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0215_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0215_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0216_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0216_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0208_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0208_LostFocus()
    TextLostFocus
End Sub


Private Sub Txt_A0209_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0209_LostFocus()
    TextLostFocus
End Sub



Private Sub Vse_background_GotFocus()
    Vse_background.TabStop = False
End Sub

