VERSION 5.00
Begin VB.Form frmCreateAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create New User"
   ClientHeight    =   1695
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRetypePw 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1455
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1080
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1455
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   690
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1455
      TabIndex        =   2
      Top             =   300
      Width           =   2325
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Retype Password:"
      Height          =   450
      Index           =   2
      Left            =   270
      TabIndex        =   6
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   705
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   315
      Width           =   1080
   End
End
Attribute VB_Name = "frmCreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RS As Recordset

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm.Show
End Sub

Private Sub OKButton_Click()
    If Me.checkDataComplete Then
        MainForm.HRDB.Execute "Insert into User_login values ('" _
            & Me.txtUserName.Text & "','" & Me.txtPassword.Text & "')"
        
        Set RS = MainForm.HRDB.OpenRecordset("Select * From User_login Where UID = '" _
            & Me.txtUserName.Text & "'", dbOpenSnapshot)
            
        If Not (RS.BOF Or RS.EOF) Then
            MsgBox "Account Create Success", , "Congratulations!"
            Me.Hide
            MainForm.Show
            Exit Sub
        Else
            MsgBox "User Account Reject, UID must be started with alphabets.", , "Error"
            Me.txtUserName.SetFocus
            Exit Sub
        End If
    End If
    
    MsgBox "User Information not Completed", , "Error"
End Sub



Private Sub txtRetypePw_LostFocus()
    If Me.txtPassword.Text <> "" And Me.txtRetypePw.Text <> Me.txtPassword.Text Then
        MsgBox "Password aren't the same, Please Type Again", , "Error"
        Me.txtRetypePw.Text = ""
        Me.txtPassword.Text = ""
        Me.txtPassword.SetFocus
    End If
End Sub

Private Sub txtUserName_LostFocus()
    Dim collision As Boolean
    Dim strCheck As String
    
    collision = False
    
    Set RS = MainForm.HRDB.OpenRecordset("Select * from User_login", dbOpenSnapshot)
    
    If Not (RS.BOF Or RS.EOF) Then
        RS.MoveFirst
        Do Until RS.EOF
            If RS.Fields("UID") = txtUserName.Text Then
                collision = True
                Exit Do
            End If
            RS.MoveNext
        Loop
        
    End If
    
    If collision Then
        Me.txtUserName.SetFocus
        SendKeys "{Home}+{End}"
        MsgBox "User Account Already Used, Try Again", , "Fail!"
'    Else
'        RS.Close
'        strCheck = "Select [FamilyGroup].[dbo].Check_user_id_valid('" _
'            & Me.txtUserName.Text & "') AS ans"
'
'        Set RS = MainForm.HRDB.OpenRecordset(strCheck, dbOpenSnapshot)
'
'        If Not (RS.BOF Or RS.EOF) Then
'            If RS.Fields("ans") > 0 Then
'                MsgBox "User Account Valid", , "Congratulations!"
'            End If
'        End If
    End If
    
    RS.Close
End Sub

Public Function checkDataComplete() As Boolean
    Dim ret As Boolean
    ret = False
    
    If Me.txtUserName.Text = "" Then
        'do nothing
    ElseIf Me.txtPassword.Text = "" Then
        'do nothing
    ElseIf Me.txtRetypePw.Text = "" Then
        'do nothing
    ElseIf Me.txtPassword.Text <> Me.txtRetypePw.Text Then
        'do nothing
    Else
        ret = True
    End If
    
    checkDataComplete = ret
End Function
