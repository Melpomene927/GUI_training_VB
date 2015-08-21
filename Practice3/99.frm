VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "²Ó©úÅé"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4380
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4230
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim vArray&(1 To 11, 1 To 9)
Dim i%, j%, val&, strV$

    'initialize summary cells
    For j = 1 To 9
        vArray(11, j) = 0
    Next
    
    'compute values to array
    For i = 1 To 11
        strV = ""
        If i = 10 Then
            Print String(9 * 4, "----")
            GoTo Continue_I
        Else
            
        End If
        
        For j = 1 To 9
            val = i * j
            
            If i < 10 Then
                vArray(i, j) = val
                vArray(11, j) = vArray(11, j) + val
            Else
                val = vArray(i, j)
            End If
            strV = strV + Space(4 - Len(Str(val))) + Str(val)
            
Continue_J:
        Next
        Print strV
Continue_I:
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
