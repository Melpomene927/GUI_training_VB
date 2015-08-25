VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form1 
   Caption         =   "ComboBox-SSPanel-SSCommand Try Out"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel2 
      Height          =   465
      Left            =   1590
      TabIndex        =   2
      Top             =   1050
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   820
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   465
      Left            =   270
      TabIndex        =   1
      Top             =   1050
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   820
      _StockProps     =   15
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "combo_panel_comm.frx":0000
      Left            =   270
      List            =   "combo_panel_comm.frx":0002
      TabIndex        =   0
      Text            =   "please choose one"
      Top             =   240
      Width           =   1995
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   600
      Left            =   4200
      TabIndex        =   3
      Top             =   200
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1058
      _StockProps     =   78
      Caption         =   "SSCommand1"
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   600
      Left            =   4200
      TabIndex        =   4
      Top             =   950
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1058
      _StockProps     =   78
      Caption         =   "SSCommand2"
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   600
      Left            =   4200
      TabIndex        =   5
      Top             =   1700
      Width           =   1200
      _Version        =   65536
      _ExtentX        =   2117
      _ExtentY        =   1058
      _StockProps     =   78
      Caption         =   "SSCommand3"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim listItem() As Variant

Private Sub Combo1_Click()
    Dim strLine As Variant
    
    strLine = Split(Combo1.Text, " ")
    If UBound(strLine) < 1 Then
        Exit Sub
    End If
    SSPanel1.Caption = strLine(0)
    SSPanel2.Caption = strLine(1)
End Sub

Private Sub Form_Load()
    
    listItem = Array("1 萬泰", "2 金旭", "3 龍巖", "4 資生堂", "5 天漢", "6 統一期貨")
    
    Me.load_list Combo1
End Sub

Private Sub SSCommand1_Click()
    Combo1.Clear
    SSPanel1.Caption = ""
    SSPanel2.Caption = ""
End Sub

Private Sub SSCommand2_Click()
    If Combo1.ListCount = 0 Then
        Combo1.Text = "Please Choose One"
        Me.load_list Combo1
    End If
End Sub


Private Sub SSCommand3_Click()
    Unload Me
End Sub


Public Sub load_list(cmb As ComboBox)
    Dim i%
    
    For i% = 0 To UBound(listItem, 1)
        cmb.AddItem listItem(i%)
    Next
End Sub

