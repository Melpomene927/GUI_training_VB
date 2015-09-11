VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_EXAM02 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   Caption         =   "員工基本資料管理"
   ClientHeight    =   4755
   ClientLeft      =   5520
   ClientTop       =   2880
   ClientWidth     =   10125
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
   Icon            =   "EXAM02.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4755
   ScaleWidth      =   10125
   Begin VB.Frame Fra_A0822 
      Height          =   585
      Left            =   5535
      TabIndex        =   34
      Top             =   2970
      Width           =   3120
      Begin Threed.SSOption Opt_A0822_M 
         Height          =   360
         Left            =   225
         TabIndex        =   20
         Top             =   180
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   635
         _StockProps     =   78
         Caption         =   "已婚"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption Opt_A0822_N 
         Height          =   360
         Left            =   1620
         TabIndex        =   21
         Top             =   180
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   635
         _StockProps     =   78
         Caption         =   "未婚"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
   Begin VsOcxLib.VideoSoftElastic Vse_background 
      Height          =   4380
      Left            =   0
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   0
      Width           =   10125
      _Version        =   327680
      _ExtentX        =   17859
      _ExtentY        =   7726
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
      Picture         =   "EXAM02.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "EXAM02.frx":0326
      Begin VB.Frame Fra_Help 
         BackColor       =   &H00FFFF80&
         Height          =   915
         Left            =   8955
         TabIndex        =   65
         Top             =   2205
         Visible         =   0   'False
         Width           =   855
         Begin FPSpread.vaSpread Spd_Help 
            Height          =   495
            Left            =   120
            OleObjectBlob   =   "EXAM02.frx":0342
            TabIndex        =   66
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.TextBox Txt_A0809 
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
         TabIndex        =   8
         Top             =   1143
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0823 
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
         Left            =   4110
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1143
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0808 
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
         Left            =   6885
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1143
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0820 
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
         Left            =   6885
         MaxLength       =   8
         TabIndex        =   27
         Top             =   3915
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0805 
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
         TabIndex        =   25
         Top             =   3915
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0806 
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
         Left            =   4110
         MaxLength       =   8
         TabIndex        =   26
         Top             =   3915
         Width           =   1770
      End
      Begin VB.Frame Fra_A0821 
         Height          =   585
         Left            =   1305
         TabIndex        =   33
         Top             =   2970
         Width           =   3165
         Begin Threed.SSOption Opt_A0821_M 
            Height          =   360
            Left            =   270
            TabIndex        =   18
            Top             =   180
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "Male"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption Opt_A0821_F 
            Height          =   360
            Left            =   1575
            TabIndex        =   19
            Top             =   180
            Width           =   1035
            _Version        =   65536
            _ExtentX        =   1826
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "Female"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.TextBox Txt_A0818 
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
         Left            =   6885
         MaxLength       =   15
         TabIndex        =   15
         Top             =   2241
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0819 
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
         Left            =   4770
         MaxLength       =   15
         TabIndex        =   17
         Top             =   2610
         Width           =   3885
      End
      Begin VB.TextBox Txt_A0817 
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
         Left            =   4110
         MaxLength       =   15
         TabIndex        =   14
         Top             =   2241
         Width           =   1770
      End
      Begin VB.ComboBox Cbo_A0824 
         Height          =   360
         IntegralHeight  =   0   'False
         ItemData        =   "EXAM02.frx":0572
         Left            =   6885
         List            =   "EXAM02.frx":0574
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   411
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0803 
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
         Left            =   6885
         MaxLength       =   12
         TabIndex        =   2
         Top             =   45
         Width           =   1770
      End
      Begin Threed.SSPanel Pnl_A0602 
         Height          =   360
         Left            =   2340
         TabIndex        =   36
         Top             =   780
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   635
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
      Begin VB.TextBox Txt_A0813 
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
         Left            =   6885
         MaxLength       =   8
         TabIndex        =   24
         Top             =   3555
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0812 
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
         Left            =   4110
         MaxLength       =   20
         TabIndex        =   23
         Top             =   3555
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0814 
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
         TabIndex        =   22
         Top             =   3555
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0816 
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
         TabIndex        =   16
         Top             =   2610
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0815 
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
         TabIndex        =   13
         Top             =   2241
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0811 
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
         TabIndex        =   12
         Text            =   " "
         Top             =   1875
         Width           =   7350
      End
      Begin VB.TextBox Txt_A0810 
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
         TabIndex        =   11
         Text            =   " "
         Top             =   1509
         Width           =   7350
      End
      Begin VB.TextBox Txt_A0804 
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
         Left            =   5535
         MaxLength       =   40
         TabIndex        =   7
         Top             =   780
         Width           =   1050
      End
      Begin VB.TextBox Txt_A0825 
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
         TabIndex        =   6
         Top             =   777
         Width           =   1050
      End
      Begin VB.TextBox Txt_A0801 
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
         Top             =   45
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0807 
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
         IMEMode         =   3  'DISABLE
         Left            =   4110
         MaxLength       =   40
         TabIndex        =   4
         Text            =   " "
         Top             =   411
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0826 
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
         Top             =   411
         Width           =   1770
      End
      Begin VB.TextBox Txt_A0802 
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
         Left            =   4110
         MaxLength       =   12
         TabIndex        =   1
         Top             =   45
         Width           =   1770
      End
      Begin Threed.SSCommand cmd_ok 
         Height          =   405
         Left            =   8760
         TabIndex        =   31
         Top             =   1320
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "確認F11"
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
         Left            =   8760
         TabIndex        =   32
         Top             =   3870
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "結束Esc"
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
         Left            =   8760
         TabIndex        =   28
         Top             =   45
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "說明F1"
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
         Left            =   8760
         TabIndex        =   30
         Top             =   895
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "次頁F8"
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
         Left            =   8760
         TabIndex        =   29
         Top             =   470
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "前筆F7"
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
      Begin Threed.SSPanel Pnl_A0202 
         Height          =   360
         Left            =   6570
         TabIndex        =   35
         Top             =   780
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   635
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
      Begin VB.Label Lbl_A0809 
         Caption         =   "身分證號碼"
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
         Left            =   90
         TabIndex        =   45
         Top             =   1218
         Width           =   1245
      End
      Begin VB.Label Lbl_A0823 
         Caption         =   "職稱"
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
         Left            =   3195
         TabIndex        =   46
         Top             =   1218
         Width           =   1245
      End
      Begin VB.Label Lbl_A0808 
         Caption         =   "生日"
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
         Left            =   5895
         TabIndex        =   47
         Top             =   1218
         Width           =   1245
      End
      Begin VB.Label Lbl_A0820 
         Caption         =   "有效日期"
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
         Left            =   5895
         TabIndex        =   62
         Top             =   3990
         Width           =   1245
      End
      Begin VB.Label Lbl_A0805 
         Caption         =   "到職日期"
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
         Left            =   90
         TabIndex        =   60
         Top             =   3990
         Width           =   1245
      End
      Begin VB.Label Lbl_A0806 
         Caption         =   "離職日期"
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
         Left            =   3195
         TabIndex        =   61
         Top             =   3990
         Width           =   1245
      End
      Begin VB.Label Lbl_A0822 
         Caption         =   "婚姻狀況"
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
         Left            =   4545
         TabIndex        =   56
         Top             =   3240
         Width           =   930
      End
      Begin VB.Label Lbl_A0821 
         Caption         =   "性別"
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
         Left            =   90
         TabIndex        =   55
         Top             =   3240
         Width           =   1245
      End
      Begin VB.Label Lbl_A0818 
         Caption         =   "行動電話"
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
         Left            =   5895
         TabIndex        =   52
         Top             =   2310
         Width           =   1245
      End
      Begin VB.Label Lbl_A0819 
         Caption         =   "Email Address"
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
         Left            =   3195
         TabIndex        =   54
         Top             =   2685
         Width           =   1425
      End
      Begin VB.Label Lbl_A0817 
         Caption         =   "BB Call"
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
         Left            =   3195
         TabIndex        =   51
         Top             =   2295
         Width           =   930
      End
      Begin VB.Label Lbl_A0824 
         Caption         =   "所屬公司"
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
         Left            =   5895
         TabIndex        =   42
         Top             =   486
         Width           =   1245
      End
      Begin VB.Label Lbl_A0803 
         Caption         =   "英文姓名"
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
         Left            =   5895
         TabIndex        =   39
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Lbl_A0813 
         Caption         =   "郵遞區號"
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
         Left            =   5895
         TabIndex        =   59
         Top             =   3630
         Width           =   1245
      End
      Begin VB.Label Lbl_A0812 
         Caption         =   "城市"
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
         Left            =   3195
         TabIndex        =   58
         Top             =   3630
         Width           =   1245
      End
      Begin VB.Label Lbl_A0814 
         Caption         =   "國家"
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
         Left            =   90
         TabIndex        =   57
         Top             =   3630
         Width           =   1245
      End
      Begin VB.Label Lbl_A0816 
         Caption         =   "連絡傳真"
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
         Left            =   90
         TabIndex        =   53
         Top             =   2685
         Width           =   1245
      End
      Begin VB.Label Lbl_A0825 
         Caption         =   "群組代號"
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
         Left            =   90
         TabIndex        =   43
         Top             =   852
         Width           =   1245
      End
      Begin VB.Label Lbl_A0802 
         Caption         =   "中文姓名"
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
         Left            =   3195
         TabIndex        =   38
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Lbl_A0811 
         Caption         =   "英文地址"
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
         Left            =   90
         TabIndex        =   49
         Top             =   1950
         Width           =   1245
      End
      Begin VB.Label Lbl_A0815 
         Caption         =   "連絡電話"
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
         Left            =   90
         TabIndex        =   50
         Top             =   2316
         Width           =   1245
      End
      Begin VB.Label Lbl_A0804 
         Caption         =   "部門代號"
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
         Left            =   4545
         TabIndex        =   44
         Top             =   855
         Width           =   1245
      End
      Begin VB.Label Lbl_A0810 
         Caption         =   "中文地址"
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
         Left            =   90
         TabIndex        =   48
         Top             =   1584
         Width           =   1245
      End
      Begin VB.Label Lbl_A0807 
         Caption         =   "密碼"
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
         Left            =   3195
         TabIndex        =   41
         Top             =   486
         Width           =   1245
      End
      Begin VB.Label Lbl_A0826 
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
         Left            =   90
         TabIndex        =   40
         Top             =   486
         Width           =   1245
      End
      Begin VB.Label Lbl_A0801 
         Caption         =   "員工編號"
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
         Left            =   90
         TabIndex        =   37
         Top             =   120
         Width           =   1245
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   64
      Top             =   4380
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_EXAM02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

'在此處定義之所有變數, 一律以M開頭, 如M_AAA$, M_BBB#, M_CCC&
'且變數之形態, 一律在最後一碼區別, 範例如下:
' $: 文字
' #: 所有數字運算(金額或數量)
' &: 程式迴圈變數
' %: 給一些使用於是或否用途之變數 (TRUE / FALSE )
' 空白: 代表VARIENT, 動態變數


'自定變數
'Dim m_A1501Flag%
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

'必要變數
Dim m_FieldError%    '此變數在判斷欄位是否有誤, 必須回到該欄位之動作
Dim m_ExitTrigger%   '此變數在判斷結束鍵是否被觸發, 將停止目前正在處理的作業
'Dim m_RecordChange% '此變數在判斷資料是否有異動, 結束將提示是否存檔訊息
Dim m_TabGotFocus%   '控制Tab_ClickAfter 只處理一次
Dim m_TabMouseDown%  '控制由Help Control DblClick所觸發的Tab_ClickAfter不處理
Dim m_A0821%         'Option of A0821
Dim m_A0822%         'Option of A0822
Const m_Male% = 1
Const m_Female% = 2
Const m_Married% = 1
Const m_NotMarried = 2

Private Sub CBO_A0824_Prepare()
On Local Error GoTo MyError
Dim A_Sql$
Dim DY_Tmp As Recordset

    '先清空Combo Box內容
    Cbo_A0824.Clear
    
    '加入空白選項
    Cbo_A0824.AddItem ""
    
    '開起檔案
    A_Sql$ = "Select A0101,A0102 From A01 ORDER BY A0101"
    CreateDynasetODBC DB_ARTHGUI, DY_Tmp, A_Sql$, "DY_TMP", True

    '將資料擺入Combo Box中
    Do While Not DY_Tmp.EOF
       Cbo_A0824.AddItem Format(Trim$(DY_Tmp.Fields("A0101") & ""), "!@@@") & Trim$(DY_Tmp.Fields("A0102") & "")
       DY_Tmp.MoveNext
    Loop
    DY_Tmp.Close

    '若Combo Box中有資料, 停在第一筆
    If Cbo_A0824.ListCount > 0 Then Cbo_A0824.ListIndex = 0
    Exit Sub
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Function CheckRoutine_A0801() As Boolean
    CheckRoutine_A0801 = False
    m_FieldError% = -1
            G_DataChange% = True

    '檢核該欄位是否輸入
    If Txt_A0801.text = "" Then
        Sts_MsgLine.Panels(1) = G_Pnl_A0801$ & G_MustInput
        m_FieldError% = Txt_A0801.TabIndex
        Txt_A0801.SetFocus
        Exit Function
    End If

    '檢核資料是否已存在
    If G_AP_STATE = G_AP_STATE_ADD Then
        If IsKeyExist(Txt_A0801) Then
             Sts_MsgLine.Panels(1) = G_Pnl_A0801$ & G_RecordExist
             m_FieldError% = Txt_A0801.TabIndex
             Txt_A0801.SetFocus
        End If
    End If
    
    CheckRoutine_A0801 = True
End Function

Private Function CheckRoutine_A0802() As Boolean
    CheckRoutine_A0802 = False

    '設定變數初始值
    m_FieldError% = -1
    
    '檢核該欄位是否輸入
    If Txt_A0802.text = "" Then
       Sts_MsgLine.Panels(1) = G_Pnl_A0802$ & G_MustInput
       m_FieldError% = Txt_A0802.TabIndex
       Txt_A0802.SetFocus
       Exit Function
    End If
       
    CheckRoutine_A0802 = True
End Function

Private Function CheckRoutine_A0804() As Boolean
    CheckRoutine_A0804 = False
    m_FieldError% = -1
    
    If Trim$(Txt_A0804.text & "") = "" Then GoTo Check_Pass

    '檢核資料是否不存在
    If Not IsKeyExist(Txt_A0804) Then
         Sts_MsgLine.Panels(1) = G_Pnl_A0804$ & G_RecordNotExist$
         m_FieldError% = Txt_A0804.TabIndex
         Pnl_A0202.Caption = ""
         Txt_A0804.SetFocus
    Else
         Pnl_A0202.Caption = Trim(DY_A02.Fields("A0202") & "")
    End If
    
Check_Pass:
    CheckRoutine_A0804 = True
End Function

Private Function CheckRoutine_A0805() As Boolean
    CheckRoutine_A0805 = False

    '設定變數初始值
    m_FieldError% = -1
    
    '檢核該日期格式是否正確
    If Trim(Txt_A0805) <> "" Then
       If Not IsDateValidate(Txt_A0805) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A0805$ & G_DateError
          m_FieldError% = Txt_A0805.TabIndex
          Txt_A0805.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A0805), Trim$(Txt_A0806)) Then
       If ActiveControl.TabIndex = Txt_A0806.TabIndex Then
          '若有錯誤, 將變數值設定為該Control之TabIndex
          m_FieldError% = Txt_A0805.TabIndex
       Else
          m_FieldError% = Txt_A0805.TabIndex
          Txt_A0805.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A0805 = True
End Function

Private Function CheckRoutine_A0806() As Boolean
    CheckRoutine_A0806 = False

    '設定變數初始值
    m_FieldError% = -1
    
    '檢核該日期格式是否正確
    If Trim(Txt_A0806) <> "" Then
       If Not IsDateValidate(Txt_A0806) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A0806$ & G_DateError
          m_FieldError% = Txt_A0806.TabIndex
          Txt_A0806.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A0805), Trim$(Txt_A0806)) Then
       If ActiveControl.TabIndex = Txt_A0805.TabIndex Then
          '若有錯誤, 將變數值設定為該Control之TabIndex
          m_FieldError% = Txt_A0806.TabIndex
       Else
          m_FieldError% = Txt_A0806.TabIndex
          Txt_A0806.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A0806 = True
End Function

Private Function CheckRoutine_A0808() As Boolean
    CheckRoutine_A0808 = False

    '設定變數初始值
    m_FieldError% = -1
    
    '檢核該日期格式是否正確
    If Trim(Txt_A0808) <> "" Then
       If Not IsDateValidate(Txt_A0808) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A0808$ & G_DateError
          m_FieldError% = Txt_A0808.TabIndex
          Txt_A0808.SetFocus
          Exit Function
       End If
    End If
    
    CheckRoutine_A0808 = True
End Function

Private Function CheckRoutine_A0820() As Boolean
    CheckRoutine_A0820 = False

    '設定變數初始值
    m_FieldError% = -1
    
    '檢核該日期格式是否正確
    If Trim(Txt_A0820) <> "" Then
       If Not IsDateValidate(Txt_A0820) Then
          Sts_MsgLine.Panels(1) = G_Pnl_A0820$ & G_DateError
          m_FieldError% = Txt_A0820.TabIndex
          Txt_A0820.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(Sts_MsgLine, Trim$(Txt_A0805), Trim$(Txt_A0820)) Then
       If ActiveControl.TabIndex = Txt_A0805.TabIndex Then
          '若有錯誤, 將變數值設定為該Control之TabIndex
          m_FieldError% = Txt_A0820.TabIndex
       Else
          m_FieldError% = Txt_A0820.TabIndex
          Txt_A0820.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A0820 = True
End Function


Private Function CheckRoutine_A0825() As Boolean
    CheckRoutine_A0825 = False
    m_FieldError% = -1
    
    If Trim(Txt_A0825.text & "") = "" Then GoTo Check_Pass
    
    '檢核資料是否不存在
    If Not IsKeyExist(Txt_A0825) Then
         Sts_MsgLine.Panels(1) = G_Pnl_A0825$ & G_RecordNotExist$
         m_FieldError% = Txt_A0825.TabIndex
         Txt_A0825.SetFocus
         Pnl_A0602.Caption = ""
    Else
         Pnl_A0602.Caption = Trim(DY_A06.Fields("A0602") & "")
    End If
    
Check_Pass:
    CheckRoutine_A0825 = True
End Function

Private Sub ClearFieldsValue()
'將欄位值清空
    Txt_A0801.text = ""
    Txt_A0802.text = ""
    Txt_A0803.text = ""
    Txt_A0804.text = ""
    Txt_A0805.text = ""
    Txt_A0806.text = ""
    Txt_A0807.text = ""
    Txt_A0808.text = ""
    Txt_A0809.text = ""
    Txt_A0810.text = ""
    Txt_A0811.text = ""
    Txt_A0812.text = ""
    Txt_A0813.text = ""
    Txt_A0814.text = ""
    Txt_A0815.text = ""
    Txt_A0816.text = ""
    Txt_A0817.text = ""
    Txt_A0818.text = ""
    Txt_A0819.text = ""
    Txt_A0820.text = ""
    Opt_A0821_M_Click 0
    Opt_A0822_N_Click 0
    Txt_A0823.text = ""
    CboStrCut Cbo_A0824, "", Space(1)
    Txt_A0825.text = ""
    Txt_A0825.text = ""
    Spd_Help.Tag = ""
    Pnl_A0602.Caption = ""
    Pnl_A0202.Caption = ""
End Sub

Private Sub DataPrepare_A02(Txt As TextBox)
'PrepareData for Txt_A0804
Dim A_Sql$                  'SQL Message
Dim DY_Tmp As Recordset     'Temporary Dynaset
    Me.MousePointer = HOURGLASS
    
    '開起檔案
    'concate SQL Message
    A_Sql$ = "Select A0201 ,A0202 From A02"
    
    'generate wildcard compare SQL Statement
    If Txt.text <> "" Then
        A_Sql$ = A_Sql$ & " Where A0201 Like '" & Txt.text & _
            GetLikeStr(DB_ARTHGUI, True) & "'"
    End If
    A_Sql$ = A_Sql$ & " Order by A0201"
    
    'open dynaset of A02
    CreateDynasetODBC DB_ARTHGUI, DY_Tmp, A_Sql$, "DY_TMP", True
    If DY_Tmp.BOF And DY_Tmp.EOF Then
       Me.MousePointer = Default
       Sts_MsgLine.Panels(1) = G_NoReference
       Exit Sub
    End If
    
    With Spd_Help
         '設定輔助視窗(Spd_Help)的欄位屬性
         .UnitType = 2
         Spread_Property Spd_Help, 0, 2, WHITE, G_Font_Size, G_Font_Name
         Spread_Col_Property Spd_Help, 1, TextWidth("X") * 10, G_Pnl_A0201$
         Spread_Col_Property Spd_Help, 2, TextWidth("X") * 12, G_Pnl_A0201$
         Spread_DataType_Property Spd_Help, 1, SS_CELL_TYPE_EDIT, "", "", 6
         Spread_DataType_Property Spd_Help, 2, SS_CELL_TYPE_EDIT, "", "", 12
         
         .Row = -1
         .Col = -1: .Lock = True
         .Col = 1: .TypeHAlign = 2
    
         '將資料擺入Spread中
         Do Until DY_Tmp.EOF
            .MaxRows = .MaxRows + 1
            .Row = Spd_Help.MaxRows
            .Col = 1
            .text = Trim(DY_Tmp.Fields("A0201") & "")
            .Col = 2
            .text = Trim(DY_Tmp.Fields("A0202") & "")
            DY_Tmp.MoveNext
         Loop
         DY_Tmp.Close
         
         '設定輔助視窗的顯示位置
         SetHelpWindowPos Fra_Help, Spd_Help, 330, 90, 8000, 2025
         .Tag = Txt.TabIndex    'set return control tab index
         .SetFocus
    End With
    Me.MousePointer = Default
End Sub

Private Sub DataPrepare_A06(Txt As TextBox)
'PrepareData for Txt_A0825
Dim A_Sql$                  'SQL Message
Dim DY_Tmp As Recordset     'Temporary Dynaset
    Me.MousePointer = HOURGLASS
    
    '開起檔案
    'concate SQL Message
    A_Sql$ = "Select A0601 ,A0602 From A06"
    
    'generate wildcard compare SQL Statement
    If Txt.text <> "" Then
       A_Sql$ = A_Sql$ & " Where A0601 Like '" & Txt.text & _
           GetLikeStr(DB_ARTHGUI, True) & "'"
    End If
    A_Sql$ = A_Sql$ & " Order by A0601"
    
    'open dynaset of A06
    CreateDynasetODBC DB_ARTHGUI, DY_Tmp, A_Sql$, "DY_TMP", True
    If DY_Tmp.BOF And DY_Tmp.EOF Then
       Me.MousePointer = Default
       Sts_MsgLine.Panels(1) = G_NoReference
       Exit Sub
    End If
    
    With Spd_Help
         '設定輔助視窗(Spd_Help)的欄位屬性
         .UnitType = 2
         Spread_Property Spd_Help, 0, 2, WHITE, G_Font_Size, G_Font_Name
         Spread_Col_Property Spd_Help, 1, TextWidth("X") * 10, G_Pnl_A0601$
         Spread_Col_Property Spd_Help, 2, TextWidth("X") * 20, G_Pnl_A0602$
         Spread_DataType_Property Spd_Help, 1, SS_CELL_TYPE_EDIT, "", "", 3
         Spread_DataType_Property Spd_Help, 2, SS_CELL_TYPE_EDIT, "", "", 40
         
         .Row = -1
         .Col = -1: .Lock = True
         .Col = 1: .TypeHAlign = 2
    
         '將資料擺入Spread中
         Do Until DY_Tmp.EOF
            .MaxRows = .MaxRows + 1
            .Row = Spd_Help.MaxRows
            .Col = 1
            .text = Trim(DY_Tmp.Fields("A0601") & "")
            .Col = 2
            .text = Trim(DY_Tmp.Fields("A0602") & "")
            DY_Tmp.MoveNext
         Loop
         DY_Tmp.Close
         
         '設定輔助視窗的顯示位置
         SetHelpWindowPos Fra_Help, Spd_Help, 330, 90, 8000, 2025
         .Tag = Txt.TabIndex    'set return control tab index
         .SetFocus
    End With
    Me.MousePointer = Default
End Sub
Private Sub Delete_From_Menu()
'將V畫面上的該筆資料列刪除
    With Frm_EXAM02v.Spd_EXAM02v
        .Row = G_ActiveRow#
        .Action = SS_ACTION_DELETE_ROW
        .MaxRows = .MaxRows - 1
    End With
End Sub

Private Sub Delete_Process_A08()
On Local Error GoTo My_Error

    G_Str = "DELETE FROM A08"
    G_Str = G_Str & " WHERE A0801='" & G_A0801$ & "'"
    ExecuteProcess DB_ARTHGUI, G_Str
    Exit Sub
    
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    
'執行存檔前須將所有檢核欄位再做一次
    If G_AP_STATE = G_AP_STATE_ADD Then
        If Not CheckRoutine_A0801() Then Exit Function
    End If
    If Not CheckRoutine_A0802() Then Exit Function
    If Not CheckRoutine_A0805() Then Exit Function
    If Not CheckRoutine_A0806() Then Exit Function
    If Not CheckRoutine_A0820() Then Exit Function
    
    IsAllFieldsCheck = True
End Function

Private Function IsKeyExist(Txt As TextBox) As Boolean
On Local Error GoTo My_Error
Dim A_Sql$
    IsKeyExist = False
    
    If Txt.Name = "Txt_A0801" Then
        A_Sql$ = "Select A0801 From A08"
        A_Sql$ = A_Sql$ & " where A0801='" & Trim(Txt.text) & "'"
        A_Sql$ = A_Sql$ & " Order by A0801"
        CreateDynasetODBC DB_ARTHGUI, DY_A081, A_Sql$, "DY_A081", True
        If Not (DY_A081.BOF And DY_A081.EOF) Then IsKeyExist = True
    End If
    
    If Txt.Name = "Txt_A0804" Then
        A_Sql$ = "Select A0201, A0202 From A02"
        A_Sql$ = A_Sql$ & " where A0201='" & Trim(Txt.text) & "'"
        A_Sql$ = A_Sql$ & " Order by A0201"
        CreateDynasetODBC DB_ARTHGUI, DY_A02, A_Sql$, "DY_A02", True
        If Not (DY_A02.BOF And DY_A02.EOF) Then IsKeyExist = True
    End If
    
    If Txt.Name = "Txt_A0825" Then
        A_Sql$ = "Select A0601, A0602 From A06"
        A_Sql$ = A_Sql$ & " where A0601='" & Trim(Txt.text) & "'"
        A_Sql$ = A_Sql$ & " Order by A0601"
        CreateDynasetODBC DB_ARTHGUI, DY_A06, A_Sql$, "DY_A06", True
        If Not (DY_A06.BOF And DY_A06.EOF) Then IsKeyExist = True
    End If
    
    Exit Function
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Function IsRecordChange() As Boolean
'若作業狀態為刪除則不做Check
    If G_AP_STATE = G_AP_STATE_DELETE Then
       IsRecordChange = False
       Exit Function
    End If
       
'判斷Record資料是否異動
    IsRecordChange = G_DataChange%
End Function

Private Sub Move2Menu()
'將異動資料UPDATE回V畫面的SPREAD上
    With Frm_EXAM02v.Spd_EXAM02v
         If G_AP_STATE = G_AP_STATE_UPDATE Then
            .Row = G_ActiveRow#
         Else
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .Action = SS_ACTION_ACTIVE_CELL
         End If
         'Write to Spread
         .Col = 1
         .text = Trim$(Txt_A0804 & "")
         .Col = 2
         .text = Trim$(Txt_A0801 & "")
         .Col = 3
         .text = Trim$(Txt_A0802 & "")
         .Col = 4
         .text = Trim$(Txt_A0826 & "")
         .Col = 5
         .text = DateFormat2(Trim$(Txt_A0805 & ""))
         .Col = 6
         .text = Trim$(Txt_A0815 & "")
         .Col = 7
         .text = Trim$(Txt_A0818 & "")
         .Col = 8
         .text = Trim$(Txt_A0810 & "")
    End With
End Sub

Private Sub MoveDB2Field()
On Local Error GoTo My_Error

'將科目資料顯示至畫面上
    Txt_A0801.text = Trim$(DY_A08.Fields("A0801") & "")
    Txt_A0802.text = Trim$(DY_A08.Fields("A0802") & "")
    Txt_A0803.text = Trim$(DY_A08.Fields("A0803") & "")
    Txt_A0804.text = Trim$(DY_A08.Fields("A0804") & "")
    Pnl_A0202.Caption = Trim(DY_A08.Fields("A0804") & "")
    Txt_A0805.text = Trim$(DateOut(DY_A08.Fields("A0805") & ""))
    Txt_A0806.text = Trim$(DateOut(DY_A08.Fields("A0806") & ""))
    Txt_A0807.text = Trim$(Word(DY_A08.Fields("A0807")) & "")
    Txt_A0808.text = Trim$(DateOut(DY_A08.Fields("A0808") & ""))
    Txt_A0809.text = Trim$(DY_A08.Fields("A0809") & "")
    Txt_A0810.text = Trim$(DY_A08.Fields("A0810") & "")
    Txt_A0811.text = Trim$(DY_A08.Fields("A0811") & "")
    Txt_A0812.text = Trim$(DY_A08.Fields("A0812") & "")
    Txt_A0813.text = Trim$(DY_A08.Fields("A0813") & "")
    Txt_A0814.text = Trim$(DY_A08.Fields("A0814") & "")
    Txt_A0815.text = Trim$(DY_A08.Fields("A0815") & "")
    Txt_A0816.text = Trim$(DY_A08.Fields("A0816") & "")
    Txt_A0817.text = Trim$(DY_A08.Fields("A0817") & "")
    Txt_A0818.text = Trim$(DY_A08.Fields("A0818") & "")
    Txt_A0819.text = Trim$(DY_A08.Fields("A0819") & "")
    Txt_A0820.text = Trim$(DateOut(DY_A08.Fields("A0820") & ""))
    If DY_A08.Fields("A0821") = 1 Then
        Opt_A0821_M_Click 0
    Else
        Opt_A0821_F_Click 0
    End If
    If DY_A08.Fields("A0822") = 1 Then
        Opt_A0822_M_Click 0
    Else
        Opt_A0822_N_Click 0
    End If
    Txt_A0823.text = Trim$(DY_A08.Fields("A0823") & "")
    CboStrCut Cbo_A0824, Trim$(DY_A08.Fields("A0824") & ""), Space(1)
    Txt_A0825.text = Trim$(DY_A08.Fields("A0825") & "")
    Pnl_A0602.Caption = Trim(DY_A08.Fields("A0825") & "")
    Txt_A0826.text = Trim$(DY_A08.Fields("A0826") & "")

    Exit Sub

My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Sub MoveField2DB()
On Local Error GoTo My_Error
Dim A_A0824$
    StrCut Cbo_A0824.text, Space(1), A_A0824$, ""
    
    G_Str = ""
    If G_AP_STATE = G_AP_STATE_ADD Then
       InsertFields "A08001", GetCurrentDate(), G_Data_String   'G_Data_Numeric
       InsertFields "A08002", GetCurrentTime(), G_Data_String
       InsertFields "A08003", GetWorkStation(), G_Data_String
       InsertFields "A08004", GetUserId(), G_Data_String
       InsertFields "A08005", " ", G_Data_String
       InsertFields "A08006", " ", G_Data_String
       InsertFields "A08007", " ", G_Data_String
       InsertFields "A08008", " ", G_Data_String
       
       InsertFields "A0801", Trim(Txt_A0801.text & ""), G_Data_String
       InsertFields "A0802", Trim(Txt_A0802.text & ""), G_Data_String
       InsertFields "A0803", Trim(Txt_A0803.text & ""), G_Data_String
       InsertFields "A0804", Trim(Txt_A0804.text & ""), G_Data_String
       InsertFields "A0805", Trim(DateIn(Txt_A0805.text & "")), G_Data_String
       InsertFields "A0806", Trim(DateIn(Txt_A0806.text & "")), G_Data_String
       InsertFields "A0807", Val(Num(Txt_A0807.text & "")), G_Data_Numeric
       InsertFields "A0808", Trim(DateIn(Txt_A0808.text & "")), G_Data_String
       InsertFields "A0809", Trim(Txt_A0809.text & ""), G_Data_String
       InsertFields "A0810", Trim(Txt_A0810.text & ""), G_Data_String
       InsertFields "A0811", Trim(Txt_A0811.text & ""), G_Data_String
       InsertFields "A0812", Trim(Txt_A0812.text & ""), G_Data_String
       InsertFields "A0813", Trim(Txt_A0813.text & ""), G_Data_String
       InsertFields "A0814", Trim(Txt_A0814.text & ""), G_Data_String
       InsertFields "A0815", Trim(Txt_A0815.text & ""), G_Data_String
       InsertFields "A0816", Trim(Txt_A0816.text & ""), G_Data_String
       InsertFields "A0817", Trim(Txt_A0817.text & ""), G_Data_String
       InsertFields "A0818", Trim(Txt_A0818.text & ""), G_Data_String
       InsertFields "A0819", Trim(Txt_A0819.text & ""), G_Data_String
       InsertFields "A0820", Trim(DateIn(Txt_A0820.text & "")), G_Data_String
       InsertFields "A0821", Trim(Str(m_A0821%) & ""), G_Data_String
       InsertFields "A0822", Trim(Str(m_A0822%) & ""), G_Data_String
       InsertFields "A0823", Trim(Txt_A0823.text & ""), G_Data_String
       InsertFields "A0824", Trim(A_A0824$ & ""), G_Data_String
       InsertFields "A0825", Trim(Txt_A0825.text & ""), G_Data_String
       InsertFields "A0826", Trim(Txt_A0826.text & ""), G_Data_String
       
       SQLInsert DB_ARTHGUI, "A08"
    Else
       UpdateString "A08005", GetCurrentDate(), G_Data_String
       UpdateString "A08006", GetCurrentTime(), G_Data_String
       UpdateString "A08007", GetWorkStation(), G_Data_String
       UpdateString "A08008", GetUserId(), G_Data_String
       
       UpdateString "A0801", Trim(Txt_A0801.text & ""), G_Data_String
       UpdateString "A0802", Trim(Txt_A0802.text & ""), G_Data_String
       UpdateString "A0803", Trim(Txt_A0803.text & ""), G_Data_String
       UpdateString "A0804", Trim(Txt_A0804.text & ""), G_Data_String
       UpdateString "A0805", Trim(DateIn(Txt_A0805.text & "")), G_Data_String
       UpdateString "A0806", Trim(DateIn(Txt_A0806.text & "")), G_Data_String
       UpdateString "A0807", Val(Num(Txt_A0807.text & "")), G_Data_Numeric
       UpdateString "A0808", Trim(DateIn(Txt_A0808.text & "")), G_Data_String
       UpdateString "A0809", Trim(Txt_A0809.text & ""), G_Data_String
       UpdateString "A0810", Trim(Txt_A0810.text & ""), G_Data_String
       UpdateString "A0811", Trim(Txt_A0811.text & ""), G_Data_String
       UpdateString "A0812", Trim(Txt_A0812.text & ""), G_Data_String
       UpdateString "A0813", Trim(Txt_A0813.text & ""), G_Data_String
       UpdateString "A0814", Trim(Txt_A0814.text & ""), G_Data_String
       UpdateString "A0815", Trim(Txt_A0815.text & ""), G_Data_String
       UpdateString "A0816", Trim(Txt_A0816.text & ""), G_Data_String
       UpdateString "A0817", Trim(Txt_A0817.text & ""), G_Data_String
       UpdateString "A0818", Trim(Txt_A0818.text & ""), G_Data_String
       UpdateString "A0819", Trim(Txt_A0819.text & ""), G_Data_String
       UpdateString "A0820", Trim(DateIn(Txt_A0820.text & "")), G_Data_String
       UpdateString "A0821", Trim(Str(m_A0821%) & ""), G_Data_String
       UpdateString "A0822", Trim(Str(m_A0822%) & ""), G_Data_String
       UpdateString "A0823", Trim(Txt_A0823.text & ""), G_Data_String
       UpdateString "A0824", Trim(A_A0824$ & ""), G_Data_String
       UpdateString "A0825", Trim(Txt_A0825.text & ""), G_Data_String
       UpdateString "A0826", Trim(Txt_A0826.text & ""), G_Data_String
       
       G_Str = G_Str & " where A0801='" & G_A0801$ & "'"
       
       SQLUpdate DB_ARTHGUI, "A08"
    End If
    
    Exit Sub
    
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Sub OpenMainFile()
On Local Error GoTo My_Error
Dim A_Sql$

    'A08
    A_Sql$ = "Select A08.*, ISNULL(A01.A0102,'') As A0102 From A08"
    A_Sql$ = A_Sql$ & " LEFT JOIN A01 On A08.A0824 = A01.A0101"
    A_Sql$ = A_Sql$ & " where A0801='" & G_A0801$ & "'"
    A_Sql$ = A_Sql$ & " order by A0801"
    CreateDynasetODBC DB_ARTHGUI, DY_A08, A_Sql$, "DY_A08", True

    Exit Sub
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Function SaveCheck(Optional A_PassQuestion% = False) As Boolean
    SaveCheck = False
    
    '新增修改狀態下,欲結束此畫面,若畫面資料有異動時,詢問是否存檔.
    'retcode = IDYES     , 存檔後返回V畫面
    'retcode = IDNO      , 不存檔返回V畫面
    'retcode = IDCANCEL  , 不存檔停留在原畫面
    If A_PassQuestion% Then
    '按確認鍵要做存檔動作
        retcode = IDYES
    Else
    '按結束鍵時由User決定要處理的動作
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

Sub SetButtonEnable(ByVal A_Enable%)
    If Not A_Enable% Then
       Vse_Background.TabStop = True
       cmd_previous.Tag = cmd_previous.Enabled
       cmd_next.Tag = cmd_next.Enabled
       Cmd_Ok.Tag = Cmd_Ok.Enabled
       Cmd_Exit.Tag = Cmd_Exit.Enabled
       
       cmd_previous.Enabled = A_Enable%
       cmd_next.Enabled = A_Enable%
       Cmd_Ok.Enabled = A_Enable%
       Cmd_Exit.Enabled = A_Enable%
    Else
       cmd_previous.Enabled = CBool(cmd_previous.Tag)
       cmd_next.Enabled = CBool(cmd_next.Tag)
       Cmd_Ok.Enabled = CBool(Cmd_Ok.Tag)
       Cmd_Exit.Enabled = CBool(Cmd_Exit.Tag)
    End If
End Sub

Sub SetCommand()
'設定每一作業狀態下, CONTROL是否可作用
    Select Case G_AP_STATE
        Case G_AP_STATE_ADD
            'while Adding, Pkey(A0801) is allowed to input
            Cmd_Help.Enabled = True
            cmd_previous.Enabled = False
            cmd_next.Enabled = False
            Cmd_Ok.Enabled = True
            Cmd_Exit.Enabled = True
            Txt_A0801.Enabled = True
        Case G_AP_STATE_UPDATE
            'while update, no meaning to change Pkey
            Cmd_Help.Enabled = True
            cmd_previous.Enabled = (G_ActiveRow# > 1)
            cmd_next.Enabled = (G_ActiveRow# < G_MaxRows#)
            Cmd_Ok.Enabled = True
            Cmd_Exit.Enabled = True
            Txt_A0801.Enabled = False
        Case G_AP_STATE_DELETE
            'while delete, no meaning to change Pkey
            Cmd_Help.Enabled = True
            cmd_previous.Enabled = (G_ActiveRow# > 1)
            cmd_next.Enabled = (G_ActiveRow# < G_MaxRows#)
            Cmd_Ok.Enabled = True
            Cmd_Exit.Enabled = True
            Txt_A0801.Enabled = False
     End Select
End Sub

Private Sub Set_Property()
    Me.FontBold = False
    
'設定本Form之標題,字形及色系
    Form_Property Me, G_Form_EXAM02, G_Font_Name

'設Form中所有Label之標題, 字形及色系
    Label_Property Lbl_A0801, G_Pnl_A0801$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0802, G_Pnl_A0802$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0803, G_Pnl_A0803$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0804, G_Pnl_A0804$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0805, G_Pnl_A0805$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0806, G_Pnl_A0806$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0807, G_Pnl_A0807$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0808, G_Pnl_A0808$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0809, G_Pnl_A0809$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0810, G_Pnl_A0810$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0811, G_Pnl_A0811$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0812, G_Pnl_A0812$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0813, G_Pnl_A0813$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0814, G_Pnl_A0814$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0815, G_Pnl_A0815$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0816, G_Pnl_A0816$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0817, G_Pnl_A0817$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0818, G_Pnl_A0818$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0819, G_Pnl_A0819$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0820, G_Pnl_A0820$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0821, G_Pnl_A0821$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0822, G_Pnl_A0822$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0823, G_Pnl_A0823$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0824, G_Pnl_A0824$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0825, G_Pnl_A0825$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0826, G_Pnl_A0826$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Pnl_A0602, "", G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Pnl_A0202, "", G_Label_Color, G_Font_Size, G_Font_Name
    
'設Form中所有TextBox之字形及可輸入長度
    Text_Property Txt_A0801, 10, G_Font_Name
    Text_Property Txt_A0802, 12, G_Font_Name
    Text_Property Txt_A0803, 40, G_Font_Name
    Text_Property Txt_A0804, 6, G_Font_Name
    Text_Property Txt_A0805, 8, G_Font_Name
    Text_Property Txt_A0806, 8, G_Font_Name
    Text_Property Txt_A0807, 8, G_Font_Name
    Text_Property Txt_A0808, 8, G_Font_Name
    Text_Property Txt_A0809, 20, G_Font_Name
    Text_Property Txt_A0810, 50, G_Font_Name
    Text_Property Txt_A0811, 50, G_Font_Name
    Text_Property Txt_A0812, 20, G_Font_Name
    Text_Property Txt_A0813, 8, G_Font_Name
    Text_Property Txt_A0814, 20, G_Font_Name
    Text_Property Txt_A0815, 15, G_Font_Name
    Text_Property Txt_A0816, 15, G_Font_Name
    Text_Property Txt_A0817, 15, G_Font_Name
    Text_Property Txt_A0818, 15, G_Font_Name
    Text_Property Txt_A0819, 50, G_Font_Name
    Text_Property Txt_A0820, 8, G_Font_Name
    Text_Property Txt_A0823, 20, G_Font_Name
    Text_Property Txt_A0825, 3, G_Font_Name
    Text_Property Txt_A0826, 10, G_Font_Name
    
    Txt_A0807.PasswordChar = "*"
    
'設Form中所有Combo Box 之字形
    ComboBox_Property Cbo_A0824, G_Font_Size, G_Font_Name
    
'設Form中Help Frame之標題, 字形及色系
    Label_Property Fra_Help, "", COLOR_SKY, G_Font_Size, G_Font_Name
    Fra_Help.Visible = False
    
'??? 設定Form中所有Panel,Label,OptionButton,CheckBox,Frame之標題, 字形及色系
'    參數一 : Control Name              參數二 : 設定Control的Caption
'    參數三 : 是否顯示                  參數四 : 設定背景顏色
'    參數五 : 設定字型大小              參數六 : 設定字型名稱
    Control_Property Opt_A0821_M, GetCaption("EXAM02", "Male", "Male")
    Control_Property Opt_A0821_F, GetCaption("EXAM02", "Female", "Female")
    Control_Property Opt_A0822_M, GetCaption("EXAM02", "Married", "已婚")
    Control_Property Opt_A0822_N, GetCaption("EXAM02", "NotMarried", "未婚")
    
'設Form中所有Command之標題及字形
    Command_Property Cmd_Help, G_CmdHelp, G_Font_Name
    Command_Property cmd_previous, G_CmdPrevious, G_Font_Name
    Command_Property cmd_next, G_CmdNext, G_Font_Name
    Command_Property Cmd_Ok, G_CmdOk, G_Font_Name
    Command_Property Cmd_Exit, G_CmdExit, G_Font_Name

'以下為標準指令, 不得修改
    VSElastic_Property Vse_Background
    StatusBar_ProPerty Sts_MsgLine
End Sub

Private Sub Cbo_A0824_DropDown()
Dim A_A0824$
    DoEvents
    
    '將目前Combo Box上的代碼Keep下來
    StrCut Cbo_A0824.text, Space(1), A_A0824$, ""
    
    '重新準備此Combo Box之內容
    CBO_A0824_Prepare
    
    '將Combo Box上的ListIndex指向Keep下來的資料
    CboStrCut Cbo_A0824, A_A0824$, Space(1)
End Sub

Private Sub Cbo_A0824_GotFocus()
    TextGotFocus
End Sub

Private Sub Cbo_A0824_LostFocus()
    TextLostFocus
End Sub

Private Sub Cmd_Help_Click()
Dim a$

    a$ = "notepad " + G_Help_Path + "EXAM02.HLP"
    retcode = Shell(a$, 4)
End Sub

Private Sub Cmd_Next_Click()
    '無下一筆資料不做處理
    If G_ActiveRow# >= G_MaxRows# Then
       Sts_MsgLine.Panels(1) = G_AP_NONEXT
       Exit Sub
    End If
    
    Me.MousePointer = HOURGLASS
    
    '設定會影響資料存檔的所有Button的Enabled Property = False
    SetButtonEnable False
    
    '若目前Record資料有異動, 提示是否存檔
    If IsRecordChange() Then
        If SaveCheck() = False Then
            'If Dialog's cancel buttom click:
            Me.MousePointer = Default
            SetButtonEnable True
            Txt_A0802.SetFocus
            Exit Sub
        End If
    End If

    '取得下一筆資料的P-KEY
    With Frm_EXAM02v.Spd_EXAM02v
         G_ActiveRow# = G_ActiveRow# + 1
         .Row = G_ActiveRow#
         .Col = 2: G_A0801$ = Trim$(.text)
        
         '將V畫面的游標移到下一筆
         .Action = SS_ACTION_ACTIVE_CELL
    End With
    
    '帶出下一筆資料
    OpenMainFile
    ClearFieldsValue
    MoveDB2Field
    G_DataChange% = False
    
    '還原所有Button的Enabled Property
    SetButtonEnable True
    
    SetCommand
    Txt_A0802.SetFocus
    Me.MousePointer = Default
End Sub

Private Sub Cmd_Previous_Click()
    '無上一筆資料不做處理
    If G_ActiveRow# <= 1 Then
       Sts_MsgLine.Panels(1) = G_AP_NOPRVS
       Exit Sub
    End If
    Me.MousePointer = HOURGLASS
    
    '設定會影響資料存檔的所有Button的Enabled Property = False
    SetButtonEnable False
    
    '若目前Record資料有異動, 提示是否存檔
    If IsRecordChange() Then
       If SaveCheck() = False Then
          SetButtonEnable True
          Me.MousePointer = Default
          Txt_A0802.SetFocus
          Exit Sub
       End If
    End If
    
    '取得上一筆資料的P-KEY
    With Frm_EXAM02v.Spd_EXAM02v
         G_ActiveRow# = G_ActiveRow# - 1
         .Row = G_ActiveRow#
         .Col = 2: G_A0801$ = Trim$(.text)
        
    '將V畫面的游標移到下一筆
         .Action = SS_ACTION_ACTIVE_CELL
    End With
    
    '帶出上一筆資料
    OpenMainFile
    ClearFieldsValue
    MoveDB2Field
    G_DataChange% = False
    
    '還原所有Button的Enabled Property
    SetButtonEnable True
    
    SetCommand
    Txt_A0802.SetFocus
    Me.MousePointer = Default
End Sub

Private Sub Cmd_Ok_Click()
    Me.MousePointer = HOURGLASS
    
    '設定會影響資料存檔的所有Button的Enabled Property = False
    SetButtonEnable False
    
    '依每個作業狀態做各別的處理
    Select Case G_AP_STATE
      Case G_AP_STATE_ADD
           If SaveCheck(True) = False Then
              SetButtonEnable True
              Me.MousePointer = Default
              Exit Sub
           End If
           Txt_A0801.text = ""
           Sts_MsgLine.Panels(1) = G_Add_Ok
           If frm_EXAM02.Visible Then Txt_A0801.SetFocus

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
            Delete_Process_A08
            Delete_From_Menu
            Sts_MsgLine.Panels(1) = G_Delete_Ok
    End Select
    G_DataChange% = False
    
    '還原所有Button的Enabled Property
    SetButtonEnable True
    
    Me.MousePointer = Default

    '作業狀態若為修改,刪除, 則返回V畫面
    If G_AP_STATE <> G_AP_STATE_ADD Then
       DoEvents
       Me.Hide
       Frm_EXAM02v.Show
    End If
End Sub

Private Sub Cmd_Exit_Click()
    Me.MousePointer = HOURGLASS

    '結束目前視窗,跳出其他處理程序
    m_ExitTrigger% = True
    
    '設定會影響資料存檔的所有Button的Enabled Property = False
    SetButtonEnable False
    
    '若資料有異動, 提示是否要存檔
    If IsRecordChange() Then
       If SaveCheck() = False Then
          SetButtonEnable True
          Me.MousePointer = Default
          Exit Sub
       End If
    End If

    '還原所有Button的Enabled Property
    SetButtonEnable True

    '隱藏目前畫面, 顯示V畫面
    DoEvents
    Me.Hide
    Frm_EXAM02v.Show
    Me.MousePointer = Default
End Sub

Private Sub Form_Activate()
    Me.MousePointer = HOURGLASS
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    Sts_MsgLine.Panels(1) = G_Process
    Me.Refresh
    
    'Initial Form中的必要變數
    m_FieldError% = -1
    m_ExitTrigger% = False
    G_DataChange% = False
    
    '判斷是否由其他輔助畫面回來, 而非首次執行
    If Trim(G_FormFrom$) <> "" Then
       Me.MousePointer = Default
       G_FormFrom$ = ""
       '.....                '加入所要設定之動作
       '.....
       Exit Sub
    Else
       '.....                '第一次執行時之準備動作
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
            If frm_EXAM02.Visible Then Txt_A0801.SetFocus
        Else
            If frm_EXAM02.Visible Then Txt_A0802.SetFocus
        End If
        Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    End If
    
    '將Form放置到螢幕的頂層
    frm_EXAM02.ZOrder 0
    Me.MousePointer = Default
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
           Case KEY_F1
                If ActiveControl.TabIndex = Txt_A0804.TabIndex Then Exit Sub
                If ActiveControl.TabIndex = Txt_A0825.TabIndex Then Exit Sub
                KeyCode = 0
                If Cmd_Help.Visible = True And Cmd_Help.Enabled = True Then
                   Cmd_Help.SetFocus
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
           Case KEY_F11
                KeyCode = 0
                If Cmd_Ok.Visible = True And Cmd_Ok.Enabled = True Then
                   Cmd_Ok.SetFocus
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
    
    '主動將資料輸入由小寫轉為大寫
    '  若有某些欄位不需要轉換時, 須予以跳過
    If ActiveControl.TabIndex <> Txt_A0801.TabIndex Then GoTo Form_KeyPress_A
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Form_KeyPress_A:
    '輸入任意字元(ENTER除外), 將資料異動變數設成TRUE
    If Not TypeOf ActiveControl Is SSCommand Then
       If KeyAscii <> KEY_RETURN Then G_DataChange% = True
    End If

    'If ActiveControl.TabIndex <> Spd_PATTERNM.TabIndex Then
       KeyPress KeyAscii           'Enter時自動跳到下一欄位, spread除外
    'End If
    
    '刪除作業下, 將KeyBoard鎖住, 不接受資料異動
    If G_AP_STATE = G_AP_STATE_DELETE Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    FormCenter Me
    Set_Property
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    If Cmd_Exit.Enabled Then Cmd_Exit.SetFocus: Cmd_Exit_Click
End Sub

Private Sub Opt_A0821_F_Click(Value As Integer)
    m_A0821% = m_Female%
End Sub

Private Sub Opt_A0821_M_Click(Value As Integer)
    m_A0821% = m_Male%
End Sub

Private Sub Opt_A0822_M_Click(Value As Integer)
    m_A0822% = m_Married%
End Sub

Private Sub Opt_A0822_N_Click(Value As Integer)
    m_A0822% = m_NotMarried%
End Sub

Private Sub Spd_Help_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim A_Code$, A_Name$

    Me.MousePointer = HOURGLASS
    
    'KEEP自輔助視窗點選的資料
    With Spd_Help
         .Row = .ActiveRow
         .Col = 1
         A_Code$ = Trim(.text)
         .Col = 2
         A_Name$ = Trim(.text)
    
         '將KEEP的資料帶入畫面
         Select Case Val(.Tag)
                Case Txt_A0804.TabIndex
                     Txt_A0804.text = A_Code$
                     Pnl_A0202.Caption = A_Name$
                Case Txt_A0825.TabIndex
                     Txt_A0825.text = A_Code$
                     Pnl_A0602.Caption = A_Name$
         End Select
         G_DataChange% = True
    End With
    
    '隱藏輔助視窗
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
'標準指令,不得修改
    SpreadLostFocus Col, Row
    If NewCol > 0 Then SpreadGotFocus NewCol, NewRow
End Sub

Private Sub Spd_Help_LostFocus()
    Fra_Help.Visible = False
    Select Case Val(Spd_Help.Tag)
      Case Txt_A0804.TabIndex
           Txt_A0804.SetFocus
      Case Txt_A0825.TabIndex
           Txt_A0825.SetFocus
    End Select
End Sub

Private Sub Txt_A0801_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0801_LostFocus()
    TextLostFocus
    
'判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = Cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0801.TabIndex Then Exit Sub
    If Spd_Help.Visible = True Then Exit Sub
    ' ....

'自我檢查
    retcode = CheckRoutine_A0801()
End Sub

Private Sub Txt_A0802_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0802_LostFocus()
    TextLostFocus
    
'判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = Cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0802.TabIndex Then Exit Sub
    If Spd_Help.Visible = True Then Exit Sub
    ' ....

'自我檢查
    retcode = CheckRoutine_A0802()
End Sub

Private Sub Txt_A0803_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0803_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0804_DblClick()
'若欄位有提供輔助資料,按下滑鼠, 所須處理之事項
    Txt_A0804_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A0804_KeyDown(KeyCode As Integer, Shift As Integer)
'若欄位有提供輔助資料,按下F1, 所須處理之事項
    If KeyCode = KEY_F1 Then DataPrepare_A02 Txt_A0804
End Sub

Private Sub Txt_A0804_LostFocus()
    TextLostFocus
    '判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = Cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0804.TabIndex Then Exit Sub
    If Spd_Help.Visible = True Then Exit Sub
    ' ....

    '自我檢查
    retcode = CheckRoutine_A0804()
End Sub

Private Sub Txt_A0804_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A0805_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0805_LostFocus()
    TextLostFocus
    
    '判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = Cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0805.TabIndex Then Exit Sub
    If Spd_Help.Visible = True Then Exit Sub
    ' ....

    '自我檢查
    retcode = CheckRoutine_A0805()
End Sub

Private Sub Txt_A0806_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0806_LostFocus()
    TextLostFocus
    '判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = Cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0806.TabIndex Then Exit Sub
    If Spd_Help.Visible = True Then Exit Sub
    ' ....

    '自我檢查
    retcode = CheckRoutine_A0806()
End Sub

Private Sub Txt_A0807_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0807_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0808_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0808_LostFocus()
    TextLostFocus
    
    '判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = Cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0808.TabIndex Then Exit Sub
    If Spd_Help.Visible = True Then Exit Sub
    ' ....

    '自我檢查
    retcode = CheckRoutine_A0808()
End Sub

Private Sub Txt_A0809_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0809_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0810_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0810_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0811_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0811_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0812_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0812_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0813_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0813_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0814_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0814_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0815_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0815_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0816_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0816_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0817_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0817_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0818_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0818_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0819_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0819_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0820_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0820_LostFocus()
    TextLostFocus
    
    '判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = Cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0820.TabIndex Then Exit Sub
    If Spd_Help.Visible = True Then Exit Sub
    ' ....

    '自我檢查
    retcode = CheckRoutine_A0820()
    
End Sub

Private Sub Txt_A0823_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0823_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0825_DblClick()
'若欄位有提供輔助資料,按下滑鼠, 所須處理之事項
    Txt_A0825_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A0825_KeyDown(KeyCode As Integer, Shift As Integer)
'若欄位有提供輔助資料,按下F1, 所須處理之事項
    If KeyCode = KEY_F1 Then DataPrepare_A06 Txt_A0825
End Sub

Private Sub Txt_A0825_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A0825_LostFocus()
    TextLostFocus
    '判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = Cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0825.TabIndex Then Exit Sub
    If Spd_Help.Visible = True Then Exit Sub
    ' ....

    '自我檢查
    retcode = CheckRoutine_A0825()
End Sub

Private Sub Vse_background_GotFocus()
    Vse_Background.TabStop = False
End Sub

