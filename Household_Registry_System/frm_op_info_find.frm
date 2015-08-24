VERSION 5.00
Begin VB.Form frm_op_info_find 
   Caption         =   "Find Person"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_lname 
      Height          =   375
      Left            =   1260
      MaxLength       =   12
      TabIndex        =   12
      Top             =   2070
      Width           =   1485
   End
   Begin VB.CommandButton cmd_cnl 
      Caption         =   "&Cansel"
      Height          =   465
      Left            =   2970
      TabIndex        =   10
      Top             =   1770
      Width           =   1395
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&Ok"
      Height          =   465
      Left            =   2970
      TabIndex        =   9
      Top             =   1110
      Width           =   1395
   End
   Begin VB.CommandButton cmd_search 
      Caption         =   "&Search"
      Height          =   465
      Left            =   2970
      TabIndex        =   8
      Top             =   450
      Width           =   1395
   End
   Begin VB.ListBox lst 
      Height          =   2400
      ItemData        =   "frm_op_info_find.frx":0000
      Left            =   210
      List            =   "frm_op_info_find.frx":0002
      TabIndex        =   7
      Top             =   2610
      Width           =   4155
   End
   Begin VB.TextBox txt_fname 
      Height          =   375
      Left            =   1260
      MaxLength       =   12
      TabIndex        =   6
      Top             =   1530
      Width           =   1485
   End
   Begin VB.TextBox txt_id 
      Height          =   375
      Left            =   1260
      MaxLength       =   10
      TabIndex        =   4
      Top             =   990
      Width           =   1485
   End
   Begin VB.Frame fm_search 
      Caption         =   "Search By:"
      Height          =   675
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   2625
      Begin VB.OptionButton opt_search 
         Caption         =   "SSID"
         Height          =   285
         Index           =   1
         Left            =   1500
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton opt_search 
         Caption         =   "PersonID"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Label lbl_lname 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name"
      Height          =   255
      Left            =   210
      TabIndex        =   11
      Top             =   2130
      Width           =   855
   End
   Begin VB.Label lbl_fname 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name"
      Height          =   255
      Left            =   210
      TabIndex        =   5
      Top             =   1590
      Width           =   855
   End
   Begin VB.Label lbl_ID 
      Alignment       =   1  'Right Justify
      Caption         =   "P&ID / SSID"
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   1050
      Width           =   855
   End
End
Attribute VB_Name = "frm_op_info_find"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public searchID As Integer

Private Sub cmd_cnl_Click()
    Unload Me
    
End Sub

Private Sub cmd_search_Click()
Dim strID$
Dim strFName$, strLName$
Dim whereClause$
Dim RS As Recordset

    'Capture Search Conditions: ID / NAME
    strID = Me.txt_id.Text
    strFName = Me.txt_fname.Text
    strLName = Me.txt_lname.Text
    
    'Check Empty Input
    If strID = "" And strFName = "" And strLName = "" Then
        Exit Sub
    Else
        'have some input
        strID = IIf(strID = "", _
            "", IIf(Me.searchID = 0, "PID = '", "ID_card_num = '") & strID & "' ")
        strFName = IIf(strFName = "", _
            "", "First_name = '" & strFName & "' ")
        strLName = IIf(strLName = "", _
            "", "Last_name = '" & strLName & "' ")
        
        'generate where clause
'        whereClause = "Where " & IIf(strID = "", _
'            strID, _
'            strID & "And ") & _
'            IIf(strFName = "" Or strLName = "", _
'                strFName, _
'                strFName & "And ") & strLName
        
        'get data
        Set RS = MainForm.HRDB.OpenRecordset( _
            "Select First_name, Last_name, ID_card_num " & _
            "From E_Personal_information " & whereClause _
            , dbOpenSnapshot)
            
            
        If Not RS Is Nothing Then
            RS.Close
        End If

    End If
    
    
End Sub

Private Sub Form_Load()

    If Me.opt_search(0).Value = True Then
        searchID = 0
    Else
        searchID = 1
    End If
    
End Sub

Private Sub lbl_name_Click()

End Sub
