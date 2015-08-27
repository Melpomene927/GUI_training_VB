VERSION 5.00
Begin VB.Form FrmCreateAccount 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create New User"
   ClientHeight    =   2160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmb_type 
      Height          =   315
      ItemData        =   "CreateAccount.frx":0000
      Left            =   1455
      List            =   "CreateAccount.frx":0002
      TabIndex        =   9
      Top             =   1598
      Width           =   1455
   End
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
      Left            =   3960
      TabIndex        =   1
      Top             =   810
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Top             =   330
      Width           =   1215
   End
   Begin VB.Label lbl_type 
      Alignment       =   1  'Right Justify
      Caption         =   "Accout &Type:"
      Height          =   270
      Index           =   3
      Left            =   270
      TabIndex        =   8
      Top             =   1620
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Retype     Password:"
      Height          =   450
      Index           =   2
      Left            =   270
      TabIndex        =   6
      Top             =   1080
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   705
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   270
      TabIndex        =   4
      Top             =   315
      Width           =   1080
   End
End
Attribute VB_Name = "FrmCreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================
' Module    : FrmCreateAccount
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   :
'========================================================================
Option Explicit
Option Compare Text

Public acType As Variant
Public RS As Recordset


Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
    Me.acType = Array("Administrator", "User")
    For i = 0 To UBound(acType)
        Me.cmb_type.AddItem acType(i)
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MainForm.Show
End Sub

Private Sub OKButton_Click()
Dim auth$
    'Check information Fulfill
    If Me.checkDataComplete Then
        'check acount type
        Select Case Me.cmb_type.Text
            Case acType(0) 'Administrator
                auth = "rwx"
            Case acType(1) 'User
                auth = "r__"
            Case Else
                auth = "___"
        End Select
        
        HRDB.Execute "INSERT INTO User_login values ('" _
            & Me.txtUserName.Text & "','" & Me.txtPassword.Text & "','" _
            & auth & "')"
        
        Set RS = HRDB.OpenRecordset("Select * From User_login Where UID = '" _
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
    
    MsgBox "User Information not Completed", vbCritical + vbOKOnly, "Error"
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
    
    Set RS = HRDB.OpenRecordset("Select * from User_login", dbOpenSnapshot)
    
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
'        Set RS = HRDB.OpenRecordset(strCheck, dbOpenSnapshot)
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
    ElseIf Me.cmb_type.Text = "" Then
        'do nothing
    Else
        ret = True
    End If
    
    checkDataComplete = ret
End Function
