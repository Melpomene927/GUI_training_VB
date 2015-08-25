VERSION 5.00
Begin VB.Form frm_op_info_find 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Person"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtb 
      Height          =   375
      Index           =   2
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   11
      Top             =   2070
      Width           =   1485
   End
   Begin VB.TextBox txtb 
      Height          =   375
      Index           =   1
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   10
      Top             =   1530
      Width           =   1485
   End
   Begin VB.CommandButton cmd_cnl 
      Caption         =   "&Cansel"
      Height          =   465
      Left            =   3690
      TabIndex        =   8
      Top             =   840
      Width           =   1395
   End
   Begin VB.CommandButton cmd_search 
      Caption         =   "&Search"
      Height          =   465
      Left            =   3690
      TabIndex        =   7
      Top             =   300
      Width           =   1395
   End
   Begin VB.ListBox lst 
      Height          =   2205
      ItemData        =   "frm_op_info_find.frx":0000
      Left            =   210
      List            =   "frm_op_info_find.frx":0002
      TabIndex        =   6
      Top             =   2610
      Width           =   4875
   End
   Begin VB.TextBox txtb 
      Height          =   375
      Index           =   0
      Left            =   1290
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
      TabIndex        =   9
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
Public chosed As Boolean
Public person As Variant
Public RS As Recordset

Const TXT_ID = 0
Const txt_fname = 1
Const txt_lname = 2

Private Sub cmd_cnl_Click()
    Unload Me
    
End Sub



Private Sub cmd_search_Click()
    Me.doSearch
End Sub

Private Sub Form_Load()
    Me.chosed = False
    If Me.opt_search(0).Value = True Then
        searchID = 0
    Else
        searchID = 1
    End If
    
End Sub


Private Sub lst_DblClick()
    
    Debug.Print lst.ListIndex
    person = Split(lst.List(lst.ListIndex), vbTab)
    If Not UBound(person) < 4 Then
        Unload Me
    End If
End Sub

Public Function showDialog() As Variant

    Me.Show vbModal
    showDialog = person
End Function


Public Sub doSearch()

Dim strID$
Dim strFName$, strLName$
Dim whereClause$
    

    'Capture Search Conditions: ID / NAME
    strID = Me.txtb(TXT_ID).Text
    strFName = Me.txtb(txt_fname).Text
    strLName = Me.txtb(txt_lname).Text
    
    
    'Check Empty Input
    If strID = "" And strFName = "" And strLName = "" Then
        Exit Sub
    Else
        'have some input
        strID = IIf(strID = "", _
            "", IIf(Me.searchID = 0, "PID = '", "ID_card_num = '") & strID & "' ")
        strFName = IIf(strFName = "", _
            "", "FiRSt_name = '" & strFName & "' ")
        strLName = IIf(strLName = "", _
            "", "Last_name = '" & strLName & "' ")
        
        'generate where clause
        whereClause = "Where " & IIf(strID = "" Or (strFName = "" And strLName = ""), _
            strID, _
            strID & "And ") & _
            IIf((strID = "" And strFName = "") Or strLName = "", _
                strFName, _
                strFName & "And ") & strLName

        'get data
        
        Set RS = MainForm.HRDB.OpenRecordset( _
            "Select * " & _
            "From E_Personal_information " & whereClause _
            , dbOpenSnapshot)
        
        'put to list
        If Not (RS.BOF And RS.EOF) Then
            RS.MoveFirst
            lst.Clear
            Do Until RS.EOF
                
                Me.lst.AddItem RS.Fields("PID") & vbTab & _
                    RS.Fields("First_name") & vbTab & _
                    RS.Fields("Last_name") & vbTab & _
                    RS.Fields("Gender") & vbTab & _
                    RS.Fields("ID_card_num")
                RS.MoveNext
            Loop
        End If
        
        RS.Close
        
    End If
    
    
End Sub


