VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calulator"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ButtomBack 
      Caption         =   "¡ö|"
      Height          =   465
      Left            =   2130
      TabIndex        =   17
      Top             =   1590
      Width           =   915
   End
   Begin VB.CommandButton ButtomReset 
      Caption         =   "Reset"
      Height          =   465
      Left            =   2130
      TabIndex        =   18
      Top             =   2100
      Width           =   915
   End
   Begin VB.CommandButton ButtomEqual 
      Caption         =   "="
      Height          =   465
      Left            =   2130
      TabIndex        =   19
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton ButtomDiv 
      Caption         =   "/"
      Height          =   465
      Left            =   1650
      TabIndex        =   16
      Top             =   2640
      Width           =   435
   End
   Begin VB.CommandButton ButtomMulti 
      Caption         =   "*"
      Height          =   465
      Left            =   1650
      TabIndex        =   15
      Top             =   2100
      Width           =   435
   End
   Begin VB.CommandButton ButtomMinus 
      Caption         =   "-"
      Height          =   465
      Left            =   1650
      TabIndex        =   14
      Top             =   1590
      Width           =   435
   End
   Begin VB.CommandButton ButtomPlus 
      Caption         =   "+"
      Height          =   465
      Left            =   1650
      TabIndex        =   13
      Top             =   1110
      Width           =   435
   End
   Begin VB.CommandButton buttom9 
      Caption         =   "9"
      Height          =   465
      Left            =   1170
      TabIndex        =   11
      Top             =   1110
      Width           =   435
   End
   Begin VB.CommandButton buttom8 
      Caption         =   "8"
      Height          =   465
      Left            =   690
      TabIndex        =   10
      Top             =   1110
      Width           =   435
   End
   Begin VB.CommandButton buttom7 
      Caption         =   "7"
      Height          =   465
      Left            =   210
      TabIndex        =   9
      Top             =   1110
      Width           =   435
   End
   Begin VB.CommandButton ButtomDot 
      Caption         =   "."
      Height          =   465
      Left            =   1140
      TabIndex        =   12
      Top             =   2640
      Width           =   435
   End
   Begin VB.CommandButton buttom6 
      Caption         =   "6"
      Height          =   465
      Left            =   1170
      TabIndex        =   8
      Top             =   1590
      Width           =   435
   End
   Begin VB.CommandButton buttom5 
      Caption         =   "5"
      Height          =   465
      Left            =   690
      TabIndex        =   7
      Top             =   1590
      Width           =   435
   End
   Begin VB.CommandButton buttom4 
      Caption         =   "4"
      Height          =   465
      Left            =   210
      TabIndex        =   6
      Top             =   1590
      Width           =   435
   End
   Begin VB.CommandButton Buttom3 
      Caption         =   "3"
      Height          =   465
      Left            =   1170
      TabIndex        =   5
      Top             =   2100
      Width           =   435
   End
   Begin VB.CommandButton Buttom2 
      Caption         =   "2"
      Height          =   465
      Left            =   690
      TabIndex        =   4
      Top             =   2100
      Width           =   435
   End
   Begin VB.CommandButton Buttom0 
      Caption         =   "0"
      Height          =   465
      Left            =   210
      TabIndex        =   2
      Top             =   2640
      Width           =   915
   End
   Begin VB.CommandButton Buttom1 
      Caption         =   "1"
      Height          =   465
      Left            =   210
      TabIndex        =   3
      Top             =   2100
      Width           =   435
   End
   Begin VB.Frame Frame_result 
      Caption         =   "Results"
      Height          =   705
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3645
      Begin VB.TextBox Text_result 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   510
         TabIndex        =   1
         Text            =   "0"
         Top             =   210
         Width           =   3075
      End
      Begin VB.Label Label_operation 
         Height          =   315
         Left            =   60
         TabIndex        =   20
         Top             =   240
         Width           =   405
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------
'   Global Variables
'--------------------------------------------------------------------------------
Public vStack As myStack        'value stack
Public oStack As myStack        'operator stack
Public init As Boolean          'init flag
Public inCompute As Boolean     'Calculator in computation


'--------------------------------------------------------------------------------
'   Event Handle Sub
'--------------------------------------------------------------------------------
Private Sub Buttom0_Click()
    If Not StrComp(Text_result.Text, "0") = 0 Then
        NumInput (Buttom0.Caption)
    End If
End Sub

Private Sub Buttom1_Click()
    NumInput (Buttom1.Caption)
End Sub

Private Sub Buttom2_Click()
    NumInput (Buttom2.Caption)
End Sub

Private Sub Buttom3_Click()
    NumInput (Buttom3.Caption)
End Sub

Private Sub buttom4_Click()
    NumInput (buttom4.Caption)
End Sub

Private Sub buttom5_Click()
    NumInput (buttom5.Caption)
End Sub

Private Sub buttom6_Click()
    NumInput (buttom6.Caption)
End Sub

Private Sub buttom7_Click()
    NumInput (buttom7.Caption)
End Sub

Private Sub buttom8_Click()
    NumInput (buttom8.Caption)
End Sub

Private Sub buttom9_Click()
    NumInput (buttom9.Caption)
End Sub


Private Sub ButtomDot_Click()
    NumInput (ButtomDot.Caption)
End Sub

Private Sub ButtomEqual_Click()
    '-------------------------------
    '  Compute key
    '-------------------------------
    Dim result As String
    
    'debug msg
    Debug.Print "v-size=" & vStack.getSize & " o-size="; oStack.getSize
    
    'handle the condition
    If inCompute = True And vStack.getSize >= 2 Then
        oStack.pop
        inCompute = False
        Me.Label_operation.Caption = ""
        
    ElseIf vStack.getSize <> 0 And oStack.getSize = vStack.getSize Then
        vStack.push (Me.Text_result.Text)
        
    ElseIf vStack.getSize <> oStack.getSize + 1 Then
        Exit Sub
    End If
    
    'prevent sequence of '=' press
    If init Then
        Exit Sub
    End If
    
    'do the calculation in the stack
    Do While Not oStack.isEmpty
        result = Me.computeDue
    Loop
    
    
    Me.Text_result.Text = result
    init = True
    
End Sub

Private Sub ButtomReset_Click()
    '-------------------------------
    '  Reset key
    '-------------------------------
    Initial
    stackReset
End Sub

Private Sub ButtomBack_Click()
    '-------------------------------
    '  ¡ö| Backspace key
    '-------------------------------
    Dim length As Double
    length = Len(Text_result.Text)
    
    
    If length > 1 Then
        'do backspace
        Text_result.Text = Mid(Text_result.Text, 1, length - 1)
        
    ElseIf length = 1 Then
        'when last one char, replace with 0
        Text_result.Text = "0"
        init = True
    End If
    
End Sub


Private Sub ButtomPlus_Click()
    '-------------------------------
    '  + Plus key
    '-------------------------------
    Dim result As String
    
    Me.Label_operation = ButtomPlus.Caption
    
    If inCompute = True Then
        Exit Sub
    End If
    
    vStack.push (Me.Text_result.Text)
    If vStack.getSize >= 2 And oStack.getSize >= 1 Then
        vStack.push (Me.computeDue)
    End If

    oStack.push ("+")
    inCompute = True
    init = True
    
    
    
End Sub

Private Sub ButtomMinus_Click()
    '-------------------------------
    '  - Minus key
    '-------------------------------
    Dim result As String
    
    Me.Label_operation = ButtomMinus.Caption
    
    If inCompute = True Then
        Exit Sub
    End If
    
    vStack.push (Me.Text_result.Text)
    
    If vStack.getSize >= 2 And oStack.getSize >= 1 Then
        vStack.push (Me.computeDue)
    End If
    
    oStack.push ("-")
    
    inCompute = True
    init = True
    
    
    
End Sub

Private Sub ButtomMulti_Click()
    '-------------------------------
    '  * Multiply key
    '-------------------------------
    
    Me.Label_operation = ButtomMulti.Caption
    
    If inCompute = True Then
        Exit Sub
    End If
    
    vStack.push (Me.Text_result.Text)
    
    If vStack.getSize >= 2 And oStack.getSize >= 1 Then
        vStack.push (Me.computeDue)
    End If
    
    oStack.push ("*")
    
    inCompute = True
    init = True
    
End Sub

Private Sub ButtomDiv_Click()
    '-------------------------------
    '  / Divide key
    '-------------------------------
    
    Me.Label_operation = ButtomDiv.Caption
    
    If inCompute = True Then
        Exit Sub
    End If
    vStack.push (Me.Text_result.Text)
    
    If vStack.getSize >= 2 And oStack.getSize >= 1 Then
        vStack.push (Me.computeDue)
    End If
    
    oStack.push ("/")

    inCompute = True
    init = True
    
End Sub

Private Sub Form_Load()
    Dim RS As Recordset
    
    Text_result.MaxLength = 20
    
    Set vStack = New myStack
    Set oStack = New myStack
    
    Initial
    inCompute = False
    
End Sub

'--------------------------------------------------------------------------------
'   User Defined Sub/Function
'--------------------------------------------------------------------------------

Friend Sub NumInput(num As String)

    If init = True Then
        Text_result.Text = ""
        init = False
    End If
    
    Me.Label_operation.Caption = ""
    inCompute = False
    
    Text_result.Text = Text_result.Text & CStr(num)

End Sub


Public Sub Initial()
    init = True
    inCompute = False
    Me.Text_result.Text = "0"
    Me.Label_operation.Caption = ""
End Sub


Public Sub stackReset()
    '-------------------------------
    '  Reset value stack & op stack
    '-------------------------------
    If Not vStack.getSize = 0 Then
        vStack.reset
    End If
    If Not oStack.getSize = 0 Then
        oStack.reset
    End If
End Sub


Friend Function checkOStack() As Boolean
    Dim result As String
    Dim ret As Boolean
    
    ret = False
    
    result = oStack.check(oStack.getSize)
    
    If result = "*" Or result = "/" Then
        ret = True
    End If
    
    checkOStack = ret
End Function



Friend Function computeDue() As String
    '-------------------------------
    '  Pop 2 operand and operate
    '  Push the result back to stack
    '-------------------------------
    Dim num1 As Double, num2 As Double
    Dim op As String, ret As String
    
    Debug.Print "v-size=" & vStack.getSize & " o-size="; oStack.getSize
    
    num2 = CDbl(vStack.pop)
    num1 = CDbl(vStack.pop)
    op = oStack.pop
    
    Select Case op
        Case "+"
            ret = str(num1 + num2)
        Case "-"
            ret = str(num1 - num2)
        Case "*"
            ret = str(num1 * num2)
        Case "/"
            ret = str(num1 / num2)
        Case Else
            
    End Select
    
'    vStack.push (ret)
    computeDue = ret
    
End Function

