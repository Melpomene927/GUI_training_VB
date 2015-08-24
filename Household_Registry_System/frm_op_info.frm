VERSION 5.00
Begin VB.Form frm_op_info 
   Caption         =   "Personal Information Register"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fm_input 
      Caption         =   "Informations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   5175
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1380
         TabIndex        =   9
         Top             =   1975
         Width           =   1185
      End
      Begin VB.ComboBox cmb_gend 
         Height          =   315
         ItemData        =   "frm_op_info.frx":0000
         Left            =   1380
         List            =   "frm_op_info.frx":000A
         TabIndex        =   5
         Text            =   "Male"
         Top             =   1470
         Width           =   1185
      End
      Begin VB.TextBox txt_lname 
         Height          =   405
         Left            =   1380
         TabIndex        =   4
         Top             =   875
         Width           =   1185
      End
      Begin VB.TextBox txt_fname 
         Height          =   405
         Index           =   0
         Left            =   1380
         TabIndex        =   2
         Top             =   325
         Width           =   1185
      End
      Begin VB.Label lbl_gend 
         Alignment       =   1  'Right Justify
         Caption         =   "Gender¡G"
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label lbl_lname 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Name¡G"
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   950
         Width           =   1155
      End
      Begin VB.Label lbl_ssid 
         Alignment       =   1  'Right Justify
         Caption         =   "SS ID¡G"
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   2050
         Width           =   1155
      End
      Begin VB.Label lbl_fname 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name¡G"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   400
         Width           =   1155
      End
   End
   Begin VB.Frame fm_buttom 
      Height          =   2565
      Left            =   5940
      TabIndex        =   0
      Top             =   270
      Width           =   1755
   End
End
Attribute VB_Name = "frm_op_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label1_Click()

End Sub

Private Sub lbl_lname_Click()

End Sub

Private Sub lbl_ssid_Click()

End Sub
