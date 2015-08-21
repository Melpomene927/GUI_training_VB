VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WTF"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "²Ó©úÅé"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Form3"
      Height          =   465
      Left            =   2190
      TabIndex        =   1
      Top             =   2700
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Form2"
      Height          =   465
      Left            =   3540
      TabIndex        =   0
      Top             =   2700
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Form2.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    Form3.Show
    Unload Me
End Sub

Private Sub Form_Activate()
Dim i%, j%, value$
Dim strS$
    Me.Refresh
    Debug.Print "====================="
    Debug.Print "@Test1: "
    test1
    Debug.Print "@Test2: "
    test2
    Debug.Print "@Test3: "
    
    test3
    
    For i = -1 To 9
        strS = ""
        
        If i = 0 Then
            Print String(11 * 3, "---")
            GoTo Con_I_loop
        End If
        
        For j = -1 To 9
            value = Str(Abs(i * j))
            If i = -1 And j = -1 Then
                value = "   "
            End If
            If j = 0 Then
                strS = strS + "  |"
                GoTo Con_J_loop
            End If
            
            strS = strS + Space(3 - Len(value)) + value
Con_J_loop:
        Next
        '================
        'fin 1 row:
        '================
        Print strS
Con_I_loop:
    Next
End Sub

Public Sub test1()
    Debug.Print -2 ^ 2 + (-2) ^ 2 <> -3 * 3 + (-3) ^ 2 And "a" = "A" Or Not 1 = 1
End Sub

Public Sub test2()
    Debug.Print Not 10 \ 3 < 10 / 3 Or 10 Mod 3 = 10 - 3 * 3 And Fix(-3.5) = Int(-3.5)
End Sub

Public Sub test3()
Dim i%, sum&, avg#
    sum = 0
    avg = 0
    For i = 1 To 100
        sum = sum + i
    Next
    avg = sum / 100
    
    Debug.Print "sum: " + Str(sum) + "; avg: " + Str(avg)
    
End Sub
