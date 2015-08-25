VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Begin VB.Form MainForm 
   Caption         =   "System Login"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer ticktock 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4410
      Top             =   720
   End
   Begin VB.Frame Frame_main_msg 
      Height          =   1665
      Left            =   210
      TabIndex        =   2
      Top             =   840
      Width           =   2805
      Begin VB.Label Label_msg 
         Caption         =   "       Welcome, Please Login as User or Create A New Account"
         Height          =   945
         Left            =   270
         TabIndex        =   3
         Top             =   480
         Width           =   2235
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton Command_login 
      Caption         =   "Login"
      Height          =   525
      Left            =   3270
      TabIndex        =   1
      Top             =   1980
      Width           =   1545
   End
   Begin VB.CommandButton Command_create 
      Caption         =   "Create"
      Height          =   525
      Left            =   3270
      TabIndex        =   0
      Top             =   1200
      Width           =   1545
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   2700
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VsOcxLib.VideoSoftElastic VideoSoftElastic1 
      Height          =   2715
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5055
      _Version        =   327680
      _ExtentX        =   8916
      _ExtentY        =   4789
      _StockProps     =   70
      ConvInfo        =   1418783674
      Picture         =   "MainForm.frx":0000
      MouseIcon       =   "MainForm.frx":001C
      Begin VB.Label Label_title 
         Alignment       =   2  'Center
         Caption         =   "Household Regestry System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   263
         TabIndex        =   6
         Top             =   240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public HRDB As Database


Private Sub Command_create_Click()
    frmCreateAccount.Show
    Me.Hide
End Sub

Private Sub Command_login_Click()
    frmLogin.Show
    Me.Hide
End Sub

Private Sub Form_Load()
    
    Me.Hide
    frmLoading.Show
    frmLoading.SetCheckPoint (3)
    frmLoading.ShowMsg ("Connect To System Database")
    ticktock.Interval = 150
    ticktock.Enabled = True
    frmLoading.AdvProccess
    
End Sub

Public Sub Abort()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    
    If Not Me.HRDB Is Nothing Then
        Me.HRDB.Close
    End If
    
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
End Sub

Private Sub ticktock_Timer()
    If Me.LinkDB Then
        frmLoading.AdvProccess
    End If
    
    If frmLoading.LoadFin Then
        frmLoading.OkButton.Enabled = True
        ticktock.Enabled = False
    End If
End Sub



Public Function LinkDB() As Boolean
    frmLoading.AdvProccess
    Dim RS As Recordset
    Dim ret As Boolean
    
    ret = False
    
    On Error GoTo ErrHandler
    Set Me.HRDB = OpenDatabase("", False, False, _
        "ODBC;DSN=FamilyGroup;UID=SA;PWD=7669588")
    
    

    LinkDB = ret
    Exit Function
ErrHandler:
    Dim sMsg As String
    If Err.Number <> 0 Then
        
        MsgBox "Error Occur While Access Database", vbCritical, "Error"
    End If
End Function

