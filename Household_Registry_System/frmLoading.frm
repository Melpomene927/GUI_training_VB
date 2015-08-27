VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLoading 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loading"
   ClientHeight    =   2640
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_loading 
      Left            =   840
      Top             =   2100
   End
   Begin VB.Frame Fm_process 
      Caption         =   "Current Process"
      Height          =   855
      Left            =   450
      TabIndex        =   2
      Top             =   300
      Width           =   5205
      Begin VB.Label Lbl_process_msg 
         Alignment       =   1  'Right Justify
         Caption         =   "xxxxxx"
         Height          =   375
         Left            =   1410
         TabIndex        =   3
         Top             =   330
         Width           =   3555
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar_Load 
      Height          =   405
      Left            =   443
      TabIndex        =   1
      Top             =   1350
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2408
      TabIndex        =   0
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Label Lbl_symbol_persent 
      Caption         =   "%"
      Height          =   285
      Left            =   5430
      TabIndex        =   5
      Top             =   1890
      Width           =   255
   End
   Begin VB.Label Lbl_persent 
      Alignment       =   1  'Right Justify
      Caption         =   "xxx"
      Height          =   285
      Left            =   4890
      TabIndex        =   4
      Top             =   1890
      Width           =   495
   End
End
Attribute VB_Name = "FrmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================
' Module    : FrmLoading
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Loading Dialog
'========================================================================
Option Explicit
Option Compare Text

'========================================================================
Private fin As Boolean
Public checkPoint As Integer
Public accomplishPoint As Integer


'====================================
'   User Defined Function
'====================================

'========================================================================
' Procedure : LoadFin
' @ Author  : Mike_chang
' @ Date    : 2015/8/26
' Purpose   : Boolean function that determind whether procedure is finished
' Details   : N/A
'========================================================================
Public Function LoadFin() As Boolean
    LoadFin = fin
End Function


'========================================================================
' Procedure : ShowMsg
' @ Author  : Mike_chang
' @ Date    : 2015/8/26
' Purpose   : Update Label Caption above Process Bar
' Details   : N/A
'========================================================================
Public Sub ShowMsg(str As String)
    Me.Lbl_process_msg.Caption = str
End Sub


'========================================================================
' Procedure : SetCheckPoint
' @ Author  : Mike_chang
' @ Date    : 2015/8/26
' Purpose   : Change the accomplishPoint that advaced the Process Bar
' Details   : N/A
'========================================================================
Public Sub SetCheckPoint(pNum As Integer)
    Me.checkPoint = pNum
    Me.accomplishPoint = 0
    Me.Lbl_persent.Caption = Me.ProgressBar_Load.Value
End Sub

'========================================================================
' Procedure : AdvProccess
' @ Author  : Mike_chang
' @ Date    : 2015/8/26
' Purpose   : Advancing the Process Bar
' Details   :
'========================================================================
Public Sub AdvProccess()
Dim progress As Integer
    
    If Not fin Then
        'advance apPoint
        Me.accomplishPoint = Me.accomplishPoint + 1
        
        progress = Int(Me.accomplishPoint / Me.checkPoint * Me.ProgressBar_Load.Max)
        
        'update process bar & label
        Me.ProgressBar_Load.Value = progress
        Me.Lbl_persent.Caption = progress
        
        
        If Me.accomplishPoint >= Me.checkPoint Then
            fin = True
            Me.Lbl_process_msg.Caption = "Link Database Success"
        End If
    End If
End Sub

'====================================
'   Command Buttom Events
'====================================

'========================================================================
' Procedure : Cmd_ok_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/26
' Purpose   : Branch to MainForm & Unload itself
' Details   :
'========================================================================
Private Sub Cmd_ok_Click()
    Unload Me
    MainForm.Show
End Sub


'====================================
'   Form Events
'====================================

'========================================================================
' Procedure : Form_Load
' @ Author  : Mike_chang
' @ Date    : 2015/8/26
' Purpose   : Initialize
' Details   :
'========================================================================
Private Sub Form_Load()
    'Initialize
    Me.ProgressBar_Load.Value = 0
    SetCheckPoint (3)
    ShowMsg ("Connect To System Database")
    
    Timer_loading.Interval = 150
    Timer_loading.Enabled = True
    AdvProccess
End Sub

'========================================================================
' Procedure : Form_Unload
' @ Author  : Mike_chang
' @ Date    : 2015/8/26
' Purpose   :
' Details   :
'========================================================================
Private Sub Form_Unload(Cancel As Integer)
    Unload MainForm
End Sub

'====================================
'   Timer Events
'====================================


'========================================================================
' Procedure : Timer_loading_Timer
' @ Author  : Mike_chang
' @ Date    : 2015/8/26
' Purpose   :
' Details   :
'========================================================================
Private Sub Timer_loading_Timer()
    If LinkDB Then
        FrmLoading.AdvProccess
    End If
    
    If FrmLoading.LoadFin Then
        FrmLoading.Cmd_ok.Enabled = True
        Timer_loading.Enabled = False
    End If
End Sub


