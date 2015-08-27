VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1965
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1160.987
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fm_keep 
      Height          =   465
      Left            =   1290
      TabIndex        =   6
      Top             =   870
      Width           =   1905
      Begin VB.OptionButton Opt_keep 
         Caption         =   "No"
         Height          =   255
         Index           =   1
         Left            =   1110
         TabIndex        =   9
         Top             =   150
         Width           =   765
      End
      Begin VB.OptionButton Opt_keep 
         Caption         =   "Yes"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   150
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.TextBox Txt_UID 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   465
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton Cmd_cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2070
      TabIndex        =   5
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox Txt_PW 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label Lbl_info 
      Caption         =   "&Keep:"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label Lbl_info 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label Lbl_info 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================
' Module    : FrmLogin
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   :
'========================================================================
Option Explicit
Option Compare Text

Public LoginSucceeded As Boolean
Public keepData As Boolean

Private Sub Cmd_cancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    MainForm.Show
End Sub

Private Sub Cmd_ok_Click()
Dim RS As Recordset
    'Retreive User Data
    Set RS = HRDB.OpenRecordset("Select * from User_login " & _
       "Where UID='" & Me.Txt_UID.Text & _
       "' And PW='" & Me.Txt_PW.Text & "'" _
       , dbOpenSnapshot)
    
    'Check login succeed
    If Not (RS.BOF And RS.EOF) Then
        LoginSucceeded = True
        MsgBox "Login Success", vbInformation, "Login"
        keep_data
        Me.Hide
        FrmOperation.Show
    End If

    'If not succeeded
    If Not LoginSucceeded Then
        MsgBox "Invalid Password Or UID, try again!", , "Login"
        Txt_PW.Text = ""
        Me.Txt_UID.SetFocus
        SendKeys "{Home}+{End}"
    End If
    
    If Not RS Is Nothing Then
        RS.Close
    End If
End Sub

Private Sub Form_Load()
    If Me.Opt_keep(0).Value = True Then
        Me.keepData = True
    End If
    
    If Me.keepData Then
        Me.Load_kept_data
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm.Show
End Sub

Private Sub opt_keep_Click(Index As Integer)
    If Index = 0 Then
        Me.keepData = True
    Else
        Me.keepData = False
    End If
End Sub

Public Sub Load_kept_data()
Dim file_name As String
Dim fnum As Integer
Dim whole_file As String
Dim lines As Variant
Dim one_line As Variant
Dim num_rows As Integer
Dim i As Integer

    'build file path
    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = _
        file_name & "\"
    file_name = file_name & "login.ini"
    
    ' Load the file.
    fnum = FreeFile
    Open file_name For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum
    
    ' Break the file into lines.
    lines = Split(whole_file, vbCrLf)
    
    ' Get Number of Rows
    num_rows = UBound(lines)
    
    For i = 0 To num_rows
        If Len(lines(i)) > 0 Then
            one_line = Split(lines(i), "=")
            If UBound(one_line) = 0 Then
                Exit For
            End If
            Select Case Trim(one_line(0))
                Case "UID"
                    Me.Txt_UID.Text = Trim(one_line(1))
                Case "PW"
                    Me.Txt_PW.Text = Trim(one_line(1))
                    
            End Select
        End If
    Next
    
End Sub

Public Sub keep_data()
Dim strEmpFileName As String
Dim strBackSlash As String
Dim intEmpFileNbr As Integer
Dim strEmpName$, strUID$, strPW$


    'Open File by function: FreeFile()
    strBackSlash = IIf(Right$(App.Path, 1) = "\", "", "\")
    strEmpFileName = App.Path & strBackSlash & "EMPLOYEE.DAT"
    intEmpFileNbr = FreeFile
    
    
    Open strEmpFileName For Output As #intEmpFileNbr
    
    If Me.keepData Then
        strUID = Me.Txt_UID.Text
        strPW = Me.Txt_PW.Text
    Else
        strUID = ""
        strPW = ""
    End If
    
    strEmpName = "UID = " & strUID & _
        vbCrLf & "PW = " & strPW
    
    Write #intEmpFileNbr, strEmpName
        
    Close #intEmpFileNbr
    
    
    
End Sub
