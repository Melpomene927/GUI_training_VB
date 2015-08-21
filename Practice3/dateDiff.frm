VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Count Diff"
      Height          =   555
      Left            =   1500
      TabIndex        =   2
      Top             =   1500
      Width           =   1485
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   2250
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "Date 2"
      Top             =   450
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   270
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "Date1"
      Top             =   450
      Width           =   1365
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim dateDiffer$
    If Text1.Text <> "Date 1" And Text2.Text <> "Date 2 " Then
        dateDiffer = dateDiff("m", Text1.Text, Text2.Text)
        MsgBox "Date Diff: " + dateDiffer, vbOKOnly, "ANSWER"
    End If
End Sub

Private Sub Text2_Click()
    Text2.Text = ""
End Sub
Private Sub Text1_click()
    Text1.Text = ""
End Sub

Private Sub Text1_LostFocus()
    If Not IsDate(Text1.Text) Then
        Text1.Text = "Date 1"
        MsgBox "Error Date Format", vbCritical + vbOKOnly, "Error"
    End If
    
ErrHandle1:

End Sub

Private Sub Text2_LostFocus()
    If Not IsDate(Text2.Text) Then
        Text2.Text = "Date 2"
        MsgBox "Error Date Format", vbCritical + vbOKOnly, "Error"
    End If
    
ErrHandle1:

End Sub
