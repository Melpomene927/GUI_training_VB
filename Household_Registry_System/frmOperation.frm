VERSION 5.00
Begin VB.Form frmOperation 
   Caption         =   "Operations"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fm_txt 
      Caption         =   "Welcom"
      Height          =   2745
      Left            =   390
      TabIndex        =   1
      Top             =   1530
      Width           =   4065
      Begin VB.Label lbl_intro 
         Caption         =   "------------------------------------------------------------------------"
         Height          =   2055
         Left            =   390
         TabIndex        =   8
         Top             =   390
         Width           =   3315
      End
   End
   Begin VB.Frame fm_Functions 
      Caption         =   "Functions"
      Height          =   4125
      Left            =   4830
      TabIndex        =   0
      Top             =   180
      Width           =   1785
      Begin VB.CommandButton cmd_marry 
         Caption         =   "   &Marriage    System"
         Height          =   600
         Left            =   180
         TabIndex        =   7
         Top             =   3315
         Width           =   1500
      End
      Begin VB.CommandButton cmd_res 
         Caption         =   "&Residence Information"
         Height          =   600
         Left            =   180
         TabIndex        =   6
         Top             =   2565
         Width           =   1500
      End
      Begin VB.CommandButton cmd_work 
         Caption         =   "     &Work     Information"
         Height          =   600
         Left            =   180
         TabIndex        =   5
         Top             =   1815
         Width           =   1500
      End
      Begin VB.CommandButton cmd_prop 
         Caption         =   " &Properties   Register"
         Height          =   600
         Left            =   180
         TabIndex        =   4
         Top             =   1065
         Width           =   1500
      End
      Begin VB.CommandButton cmd_info 
         Caption         =   "Personal &Information"
         Height          =   600
         Left            =   180
         TabIndex        =   3
         Top             =   315
         Width           =   1500
      End
   End
   Begin VB.Label lbl_title 
      Alignment       =   1  'Right Justify
      Caption         =   "Household                Registry System  "
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   390
      TabIndex        =   2
      Top             =   180
      Width           =   4065
   End
End
Attribute VB_Name = "frmOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_info_Click()
    Me.Hide
    frm_op_info.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    MainForm.Show
End Sub
