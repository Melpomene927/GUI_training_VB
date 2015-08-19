VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoading 
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
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   2100
   End
   Begin VB.Frame Frame_Process 
      Caption         =   "Current Process"
      Height          =   855
      Left            =   450
      TabIndex        =   2
      Top             =   300
      Width           =   5205
      Begin VB.Label Label_proccess_msg 
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
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2408
      TabIndex        =   0
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Label Label_symble_persent 
      Caption         =   "%"
      Height          =   285
      Left            =   5430
      TabIndex        =   5
      Top             =   1890
      Width           =   255
   End
   Begin VB.Label Label_persentage 
      Alignment       =   1  'Right Justify
      Caption         =   "xxx"
      Height          =   285
      Left            =   4890
      TabIndex        =   4
      Top             =   1890
      Width           =   495
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public fin As Boolean
Public checkPoint As Integer
Public accomplishPoint As Integer


Private Sub Form_Load()
    Me.ProgressBar_Load.Value = 0
End Sub

Public Sub AdvProccess()
    Dim progress As Integer
    If Not fin Then
        Me.accomplishPoint = Me.accomplishPoint + 1
        
        progress = Int(Me.accomplishPoint / Me.checkPoint * Me.ProgressBar_Load.Max)
        
        Me.ProgressBar_Load.Value = progress
        Me.Label_persentage.Caption = progress
        
        If Me.accomplishPoint >= Me.checkPoint Then
            fin = True
            Me.Label_proccess_msg.Caption = "Link Database Success"
        End If
    End If
End Sub

Public Sub SetCheckPoint(pNum As Integer)
    Me.checkPoint = pNum
    Me.accomplishPoint = 0
    Me.Label_persentage.Caption = Me.ProgressBar_Load.Value
End Sub


Public Sub ShowMsg(str As String)
    Me.Label_proccess_msg.Caption = str
End Sub


Public Function LoadFin() As Boolean
    LoadFin = fin
End Function

Private Sub Form_Unload(Cancel As Integer)
    MainForm.Abort
End Sub

Private Sub OKButton_Click()
    Me.Hide
    MainForm.Show
End Sub

