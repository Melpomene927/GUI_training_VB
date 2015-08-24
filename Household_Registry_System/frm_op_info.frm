VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_op_info 
   Caption         =   "Personal Information Register"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fm_born 
      Caption         =   "Born Place"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   3810
      TabIndex        =   23
      Top             =   180
      Width           =   1755
      Begin VB.ComboBox cmb_city 
         Height          =   315
         Left            =   180
         TabIndex        =   27
         Text            =   "Choose"
         Top             =   1710
         Width           =   1365
      End
      Begin VB.ComboBox cmb_country 
         Height          =   315
         Left            =   180
         TabIndex        =   26
         Text            =   "Choose"
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label lbl_city 
         Caption         =   "City"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1380
         Width           =   765
      End
      Begin VB.Label lbl_country 
         Caption         =   "Country"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   420
         Width           =   765
      End
   End
   Begin VB.Frame fm_input 
      Caption         =   "Personal Informations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   210
      TabIndex        =   1
      Top             =   180
      Width           =   3405
      Begin VB.CommandButton cmd_find_mom 
         Caption         =   "Find Person"
         Height          =   315
         Left            =   1380
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Frame fm_mom 
         Height          =   1275
         Left            =   420
         TabIndex        =   16
         Top             =   3570
         Width           =   2595
         Begin Threed.SSPanel pnl_mom_ssid 
            Height          =   315
            Left            =   960
            TabIndex        =   20
            Top             =   750
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin Threed.SSPanel pnl_mon_name 
            Height          =   315
            Left            =   960
            TabIndex        =   22
            Top             =   210
            Width           =   1425
            _Version        =   65536
            _ExtentX        =   2514
            _ExtentY        =   556
            _StockProps     =   15
            BackColor       =   15790320
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin VB.Label lbl_mom_name 
            Alignment       =   1  'Right Justify
            Caption         =   "Name¡G"
            Height          =   315
            Left            =   270
            TabIndex        =   21
            Top             =   270
            Width           =   675
         End
         Begin VB.Label lbl_mom_ssid 
            Alignment       =   1  'Right Justify
            Caption         =   "SSID¡G"
            Height          =   315
            Left            =   270
            TabIndex        =   19
            Top             =   825
            Width           =   675
         End
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Top             =   2570
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   108986369
         CurrentDate     =   32874
         MaxDate         =   401768
      End
      Begin VB.TextBox txt_ssid 
         Height          =   315
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1975
         Width           =   1455
      End
      Begin VB.ComboBox cmb_gend 
         Height          =   315
         ItemData        =   "frm_op_info.frx":0000
         Left            =   1380
         List            =   "frm_op_info.frx":000A
         TabIndex        =   5
         Text            =   "Male"
         Top             =   1470
         Width           =   915
      End
      Begin VB.TextBox txt_lname 
         Height          =   315
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   4
         Top             =   875
         Width           =   1455
      End
      Begin VB.TextBox txt_fname 
         Height          =   315
         Index           =   0
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   2
         Top             =   325
         Width           =   1455
      End
      Begin VB.Label lbl_mom 
         Alignment       =   1  'Right Justify
         Caption         =   "Mother¡G"
         Height          =   255
         Left            =   210
         TabIndex        =   17
         Top             =   3150
         Width           =   1155
      End
      Begin VB.Label lbl_birth 
         Alignment       =   1  'Right Justify
         Caption         =   "Birthday¡G"
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   2600
         Width           =   1155
      End
      Begin VB.Label lbl_gend 
         Alignment       =   1  'Right Justify
         Caption         =   "Gender¡G"
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label lbl_lname 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Name¡G"
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   950
         Width           =   1155
      End
      Begin VB.Label lbl_ssid 
         Alignment       =   1  'Right Justify
         Caption         =   "SS ID¡G"
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   2050
         Width           =   1155
      End
      Begin VB.Label lbl_fname 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name¡G"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   400
         Width           =   1155
      End
   End
   Begin VB.Frame fm_buttom 
      Height          =   2445
      Left            =   3810
      TabIndex        =   0
      Top             =   2790
      Width           =   1755
      Begin VB.CommandButton cmd_ext 
         Enabled         =   0   'False
         Height          =   405
         Left            =   210
         TabIndex        =   15
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmd_del 
         Caption         =   "Delete"
         Height          =   405
         Left            =   210
         TabIndex        =   14
         Top             =   1300
         Width           =   1335
      End
      Begin VB.CommandButton cmd_create_update 
         Caption         =   "Create/&Update"
         Height          =   405
         Left            =   210
         TabIndex        =   13
         Top             =   800
         Width           =   1335
      End
      Begin VB.CommandButton cmd_load 
         Caption         =   "&Load Info"
         Height          =   405
         Left            =   210
         TabIndex        =   12
         Top             =   300
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_op_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RS As Recordset
Public countryChosen As Boolean
Public place_id As Long


Private Sub cmb_city_Click()

    'Check Place select complete
    If (Me.cmb_city.Text = "Choose" Or Me.cmb_country.Text = "Choose") Then
        Exit Sub
    End If

    'Retrieve Place_id from DB
    Set RS = MainForm.HRDB.OpenRecordset( _
        "Select Place_id From E_Place " & _
        "Where Country = '" & Me.cmb_country.Text & "' And " & _
        "City = '" & Me.cmb_city.Text & "'" _
        , dbOpenSnapshot)
    
    If Not (RS.EOF Or RS.BOF) Then
        place_id = RS.Fields("Place_id")
    Else
        MsgBox "Database Corrupt @ Place_id", vbCritical + vbOKOnly, "Error"
    End If
    
    'Close Recordset
    If Not RS Is Nothing Then
        RS.Close
    End If
    
End Sub

Private Sub cmb_country_Click()

Dim i%
    'Check if default value is chosen
    If Me.cmb_country.Text = "Choose" Then
        Exit Sub
    End If

    'Load City Data From database
    Set RS = MainForm.HRDB.OpenRecordset( _
        "Select City From E_Place " & _
        "Where Country = '" & Me.cmb_country.Text & "'" _
        , dbOpenSnapshot)
    
    
    If Not (RS.BOF Or RS.EOF) Then
        RS.MoveFirst
        Do Until RS.EOF     'Load data until end of file
            Me.cmb_city.AddItem RS.Fields("City")
            RS.MoveNext
        Loop
    Else
        MsgBox "Database Corrupt @ City", vbCritical + vbOKOnly, "Error"
    End If
    
    'Close Recordset
    If Not RS Is Nothing Then
        RS.Close
    End If
End Sub

Private Sub cmd_find_mom_Click()
    frm_op_info_find.Show
End Sub

Private Sub Form_Load()

Dim i%
    'Load Country Data From databasse
    Set RS = MainForm.HRDB.OpenRecordset( _
        "Select Distinct Country From E_Place" _
        , dbOpenSnapshot)
    
    If Not (RS.BOF Or RS.EOF) Then
        RS.MoveFirst
        Do Until RS.EOF     'Load data until end of file
            Me.cmb_country.AddItem RS.Fields("Country")
            RS.MoveNext
        Loop
    Else
        MsgBox "Database Corrupt @ Country", vbCritical + vbOKOnly, "Error"
    End If
    
    If Not RS Is Nothing Then
        RS.Close
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    frmOperation.Show
End Sub


Private Sub txt_ssid_LostFocus()
    'Check if SSID already exist
    Set RS = MainForm.HRDB.OpenRecordset( _
        "Select * From [E_Personal_information] " & _
        "Where [ID_card_num] = '" & Me.txt_ssid.Text & "'" _
        , dbOpenSnapshot)
    
    If Not (RS.BOF Or RS.EOF) Then
        MsgBox "Person's SSID already Exists", vbCritical + vbOKOnly, "Error"
        Me.txt_ssid.Text = ""
        Me.txt_ssid.SetFocus
    End If
        
    If Not RS Is Nothing Then
        RS.Close
    End If
End Sub
