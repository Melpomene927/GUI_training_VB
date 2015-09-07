VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_GD 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0FFFF&
   Caption         =   "客戶廠商資料查尋"
   ClientHeight    =   6435
   ClientLeft      =   -4095
   ClientTop       =   870
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   136
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Gd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6435
   ScaleWidth      =   9480
   Begin VB.ListBox prtval 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10000
      TabIndex        =   18
      Top             =   5700
      Visible         =   0   'False
      Width           =   375
   End
   Begin ComctlLib.StatusBar sts_msgline 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   6060
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VsOcxLib.VideoSoftElastic vse_background 
      Height          =   6060
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9480
      _Version        =   327680
      _ExtentX        =   16722
      _ExtentY        =   10689
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Align           =   5
      Picture         =   "Gd.frx":030A
      MouseIcon       =   "Gd.frx":0326
      Begin FPSpread.vaSpread spd_gd 
         Height          =   4620
         Left            =   60
         OleObjectBlob   =   "Gd.frx":0342
         TabIndex        =   10
         Top             =   1365
         Width           =   7830
      End
      Begin Threed.SSPanel Pnl_Source 
         Height          =   375
         Left            =   3825
         TabIndex        =   29
         Top             =   930
         Width           =   4065
         _Version        =   65536
         _ExtentX        =   7170
         _ExtentY        =   661
         _StockProps     =   15
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   11.99
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         Begin VB.CheckBox Chk_Source 
            BackColor       =   &H00C0C0C0&
            Caption         =   "銀行"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   2730
            TabIndex        =   9
            Top             =   40
            Width           =   1305
         End
         Begin VB.CheckBox Chk_Source 
            BackColor       =   &H00C0C0C0&
            Caption         =   "廠商"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1425
            TabIndex        =   8
            Top             =   40
            Width           =   1305
         End
         Begin VB.CheckBox Chk_Source 
            BackColor       =   &H00C0C0C0&
            Caption         =   "客戶"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   40
            Width           =   1305
         End
      End
      Begin VB.TextBox txt_Condition 
         Height          =   380
         Index           =   6
         Left            =   6375
         TabIndex        =   2
         Top             =   80
         Width           =   1515
      End
      Begin VB.TextBox txt_Condition 
         Height          =   380
         Index           =   5
         Left            =   6375
         TabIndex        =   5
         Top             =   500
         Width           =   1515
      End
      Begin VB.TextBox txt_Condition 
         Height          =   380
         Index           =   4
         Left            =   3825
         TabIndex        =   4
         Top             =   500
         Width           =   1515
      End
      Begin VB.TextBox txt_Condition 
         Height          =   380
         Index           =   3
         Left            =   1290
         TabIndex        =   3
         Top             =   500
         Width           =   1515
      End
      Begin VB.TextBox txt_Condition 
         Height          =   380
         Index           =   2
         Left            =   1290
         TabIndex        =   6
         Top             =   930
         Width           =   1515
      End
      Begin VB.TextBox txt_Condition 
         Height          =   380
         Index           =   1
         Left            =   3825
         TabIndex        =   1
         Top             =   80
         Width           =   1515
      End
      Begin VB.TextBox txt_Condition 
         Height          =   380
         Index           =   0
         Left            =   1290
         TabIndex        =   0
         Top             =   80
         Width           =   1515
      End
      Begin VB.CheckBox chk_sort 
         BackColor       =   &H00C0C0C0&
         Caption         =   "由大到小"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7965
         TabIndex        =   16
         Top             =   2448
         Width           =   1270
      End
      Begin Threed.SSCommand cmd_Help 
         Height          =   390
         Left            =   7920
         TabIndex        =   11
         Top             =   75
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   688
         _StockProps     =   78
         Caption         =   "輔助 F1"
         ForeColor       =   -2147483630
      End
      Begin Threed.SSCommand cmd_Sort 
         Height          =   390
         Left            =   7920
         TabIndex        =   13
         Top             =   915
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   688
         _StockProps     =   78
         Caption         =   "排序 F5"
         ForeColor       =   -2147483630
      End
      Begin Threed.SSCommand cmd_Query 
         Height          =   390
         Left            =   7920
         TabIndex        =   12
         Top             =   495
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   688
         _StockProps     =   78
         Caption         =   "查詢 F2"
         ForeColor       =   -2147483630
      End
      Begin Threed.SSCommand cmd_Exit 
         Height          =   390
         Left            =   7920
         TabIndex        =   17
         Top             =   5595
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   688
         _StockProps     =   78
         Caption         =   "結束 Esc"
         ForeColor       =   0
      End
      Begin Threed.SSCommand cmd_previous 
         Height          =   390
         Left            =   7920
         TabIndex        =   14
         Top             =   1335
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   688
         _StockProps     =   78
         Caption         =   "前頁 F7"
         ForeColor       =   -2147483630
      End
      Begin Threed.SSCommand cmd_next 
         Height          =   390
         Left            =   7920
         TabIndex        =   15
         Top             =   1755
         Width           =   1515
         _Version        =   65536
         _ExtentX        =   2672
         _ExtentY        =   688
         _StockProps     =   78
         Caption         =   "次頁 F8"
         ForeColor       =   -2147483630
      End
      Begin VB.Label Lbl_Source 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "資料類別"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2835
         TabIndex        =   28
         Top             =   990
         Width           =   990
      End
      Begin VB.Label pnl_Condition 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "中文全名"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   6
         Left            =   5385
         TabIndex        =   27
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label pnl_Condition 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "起始編號"
         ForeColor       =   &H00000000&
         Height          =   276
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   135
         Width           =   1200
      End
      Begin VB.Label pnl_Condition 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "統一編號"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   5
         Left            =   5385
         TabIndex        =   26
         Top             =   555
         Width           =   1200
      End
      Begin VB.Label pnl_Condition 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "地址"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   2835
         TabIndex        =   25
         Top             =   555
         Width           =   990
      End
      Begin VB.Label pnl_Condition 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "電話"
         ForeColor       =   &H00000000&
         Height          =   276
         Index           =   3
         Left            =   60
         TabIndex        =   24
         Top             =   550
         Width           =   1200
      End
      Begin VB.Label pnl_Condition 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "連絡人"
         ForeColor       =   &H00000000&
         Height          =   276
         Index           =   2
         Left            =   60
         TabIndex        =   23
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label pnl_Condition 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "中文簡稱"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   2835
         TabIndex        =   22
         Top             =   135
         Width           =   990
      End
   End
End
Attribute VB_Name = "frm_GD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''DESC: 客戶廠商輔助查詢(design by Jerry)
'Input: frm_GD.Tag  --  "1" 是客戶, "2" 是廠商, "3" 是銀行, 其他為全部
'Return:frm_GD.Tag  --  如有選擇, 傳回"(客戶編號)"＋"[Tab]" ＋"(客戶名稱)"
'                       否則傳回"!".
Option Explicit
Dim DY_A161 As Recordset
Dim DB_LGUI_GD As Database
Dim DY_INI_GD As Recordset

Dim m_Flag$  '1:客戶, 2:廠商, 3:銀行, other is all.

''以下是參考範例 (歡迎複製使用)
''(1) 傳入時動作
''Sub txt_exf02_KeyDown (KeyCode As Integer, Shift As Integer)
''    If KeyCode = KEY_F1 And G_AP_STATE <> G_AP_STATE_DELETE Then
''       frm_GD.Tag = "2"    'Supplier
''       frm_GD.Show 1
''    End If
''End Sub
''''''''
''(2) 回傳時動作
''Sub Form_Activate ()
''    Dim Tmp1$, Tmp2$
''    If frm_GD.Tag = "!" Then Exit Sub  '無挑選客戶
''    If frm_GD.Tag > "!" Then
''        StrCut frm_GD.Tag, Chr$(KEY_TAB), Tmp1$, Tmp2$
''        txt_EXF02.Text = Tmp1$     'Customer ID
''        pnl_GD02.Caption = Tmp2$   'Customer Name
''        m_RecordChange = True
''        txt_EXF02.SetFocus
''        Exit Sub
''    End If
''    ....
''End Sub

Private Sub Cmd_Exit_Click()
    Me.Tag = "!"
    Me.Hide
End Sub

Private Sub Cmd_Help_Click()
''DESC:When cmd_help click
    Dim a$, retcode
    a$ = "notepad " + G_Help_Path + "GD.HLP"
    retcode = Shell(a$, 4)
End Sub

Private Sub Cmd_Next_Click()
    spd_gd.SetFocus
    SendKeys "{PGDN}"
End Sub

Private Sub cmd_query_Click()
Dim a_cnt
    Me.MousePointer = HOURGLASS
    cmd_exit.Enabled = False
    sts_msgline.Panels(1) = G_Process
    OpenMainFile
    If DY_A161.EOF And DY_A161.BOF Then
        sts_msgline.Panels(1) = G_NoQueryData
        spd_gd.MaxRows = 0
        cmd_exit.Enabled = True
        txt_Condition(0).SetFocus
        Me.MousePointer = 0
        Exit Sub
    End If
    '
    DY_A161.MoveLast
    a_cnt = DY_A161.RecordCount
    DY_A161.MoveFirst
    '
    If a_cnt > 200 Then
       retcode = MsgBox(GetSINISTR_GD("PGMMSG", "TOO_MUCH"), MB_YESNO, Me.Caption)
       If retcode = IDNO Then
          cmd_exit.Enabled = True
          Me.MousePointer = 0
          sts_msgline.Panels(1) = SetMessage(G_AP_STATE_QUERY)
          txt_Condition(0).SetFocus
          Exit Sub
       End If
    End If
    MoveDB2Field
    sts_msgline.Panels(1) = G_Query_Ok
    cmd_exit.Enabled = True
    Me.MousePointer = 0
End Sub

Private Sub Cmd_Previous_Click()
    spd_gd.SetFocus
    SendKeys "{PGUP}"
End Sub

Private Sub cmd_Sort_Click()
    spd_gd.Row = -1
    spd_gd.Col = -1
    spd_gd.SortBy = 0   'by row
    spd_gd.SortKey(1) = spd_gd.ActiveCol
    If chk_Sort.Value Then
        spd_gd.SortKeyOrder(1) = 2  'desc
    Else
        spd_gd.SortKeyOrder(1) = 1  'asc
    End If
    spd_gd.Action = SS_ACTION_SORT
    spd_gd.SetFocus
End Sub

Private Sub Form_Activate()
Dim SQL$
Dim a_recordcnt&
    
    Chk_Source(0).Value = False
    Chk_Source(1).Value = False
    Chk_Source(2).Value = False
    
    sts_msgline.Panels(2) = GetCurrentDay(1)
    Me.MousePointer = 11
    frm_GD.Refresh
    If frm_GD.Tag = "2" Then     '廠商
        Label_Property pnl_Condition(0), GetSINISTR_GD("PanelDescpt", "makerid"), G_Label_Color, G_Font_Size, G_Font_Name
'        Label_Property pnl_Condition(1), GetSIniStr("GD", "GDS1"), G_Label_Color, G_Font_Size, G_Font_Name
'        Label_Property pnl_Condition(2), GetSIniStr("GD", "GDS2"), G_Label_Color, G_Font_Size, G_Font_Name
        Spread_Col_Property spd_gd, 1, TextWidth("A") * 10, GetSINISTR_GD("PanelDescpt", "makerid")
        Chk_Source(1).Value = "1"
    ElseIf frm_GD.Tag = "1" Then   '客戶
        Label_Property pnl_Condition(0), GetSINISTR_GD("PanelDescpt", "buyerid"), G_Label_Color, G_Font_Size, G_Font_Name
'        Label_Property pnl_Condition(1), GetSIniStr("PanelDescpt", "buyername"), G_Label_Color, G_Font_Size, G_Font_Name
'        Label_Property pnl_Condition(2), GetSIniStr("PanelDescpt", "buyertrade"), G_Label_Color, G_Font_Size, G_Font_Name
        Spread_Col_Property spd_gd, 1, TextWidth("A") * 10, GetSINISTR_GD("PanelDescpt", "buyerid")
        Chk_Source(0).Value = "1"
    ElseIf frm_GD.Tag = "3" Then  '銀行
        Label_Property pnl_Condition(0), GetSINISTR_GD("MCFGDB", "order1"), G_Label_Color, G_Font_Size, G_Font_Name
        Spread_Col_Property spd_gd, 1, TextWidth("A") * 10, GetSINISTR_GD("MCFGDB", "order1")
        Chk_Source(2).Value = "1"
    ElseIf frm_GD.Tag = "4" Then  '客戶+廠商
        Label_Property pnl_Condition(0), GetSINISTR_GD("PanelDescpt", "makerid"), G_Label_Color, G_Font_Size, G_Font_Name
        Spread_Col_Property spd_gd, 1, TextWidth("A") * 10, GetSINISTR_GD("PanelDescpt", "makerid")
        Chk_Source(0).Value = "1"
        Chk_Source(1).Value = "1"
    Else
        Spread_Col_Property spd_gd, 1, TextWidth("A") * 10, GetSINISTR_GD("PanelDescpt", "buyerid")
    End If
    Spread_Col_Property spd_gd, 2, TextWidth("A") * 12, GetSINISTR_GD("PanelDescpt", "abb_name_c")
    Spread_Col_Property spd_gd, 3, TextWidth("A") * 20, GetSINISTR_GD("PanelDescpt", "full_name_c")
    DoEvents
    m_Flag$ = frm_GD.Tag
'    If spd_gd.DataRowCnt <= 0 Then
'        sts_msgline.Panels(1) = GetSIniStr("PgmMsg", "g_data_search")
'    End If
    sts_msgline.Panels(1) = SetMessage(G_AP_STATE_QUERY)
    If txt_Condition(0).Enabled = True And txt_Condition(0).Visible = True Then
       txt_Condition(0).SetFocus
    End If
    Me.MousePointer = 0
End Sub

Sub OpenLGuiDB()
    Dim A_Path As String
    Dim A_ConnectMethod As String
    
    On Local Error Resume Next
    Screen.MousePointer = HOURGLASS
   'Pick Local INI DataPath String (GL.mdb)
    A_Path = GetIniStr("DBPath", "Path3", "GUI.INI")
    A_ConnectMethod = GetIniStr("DBPath", "Connect3", "GUI.INI")
    Set DB_LGUI_GD = GetEngine.OpenDatabase(A_Path, False, False, A_ConnectMethod)
    If Err Then
       If Trim$(A_ConnectMethod) = "" Then   'Access DataBase
          If Err = 3043 Then
             Err = 0
             DB_LGUI_GD.Close
             Set DB_LGUI_GD = GetEngine.OpenDatabase(A_Path, False, False, A_ConnectMethod)
          ElseIf Err = 3049 Then
             Err = 0
             RepairDatabase A_Path
             Set DB_LGUI_GD = GetEngine.OpenDatabase(A_Path, False, False, A_ConnectMethod)
          End If
       End If
    End If
    If Err Then
       MsgBox Error(Err), MB_ICONEXCLAMATION, App.Title
       End
    End If
    If Trim$(A_ConnectMethod) <> "" Then DB_LGUI_GD.QueryTimeout = 0
    'Open Table
    If Trim(DB_LGUI_GD.Connect) = "" Then
        Set DY_INI_GD = DB_LGUI_GD.OpenRecordset("INI", dbOpenTable)
        DY_INI_GD.index = "INI"
    Else
    
    End If
    Screen.MousePointer = Default

End Sub

Function GetSINISTR_GD(Section$, Topic$) As String
    GetSINISTR_GD = " "
    If Trim(DB_LGUI_GD.Connect) <> "" Then
        Dim A_Sql$
        A_Sql$ = "SELECT TOPICVALUE FROM INI"
        A_Sql$ = A_Sql$ & " WHERE SECTION='" & Section$ & "'"
        A_Sql$ = A_Sql$ & " AND TOPIC='" & Topic$ & "'"
        Set DY_INI_GD = DB_LGUI_GD.OpenRecordset(A_Sql$, dbOpenSnapshot, dbSQLPassThrough)
        If Not (DY_INI_GD.BOF And DY_INI_GD.EOF) Then
            GetSINISTR_GD = Trim(DY_INI_GD.Fields("TOPICVALUE") & "")
        End If
        DY_INI_GD.Close
    Else
        DY_INI_GD.Seek "=", Section$, Topic$
        If Not DY_INI_GD.NoMatch Then
           GetSINISTR_GD = DY_INI_GD.Fields("TOPICVALUE") & ""
        End If
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case KEY_F1
            KeyCode = 0
            If cmd_help.Enabled = True Then
                cmd_help.SetFocus
                DoEvents
                SendKeys "{Enter}"
            End If
        Case KEY_F5
            KeyCode = 0
            If cmd_Sort.Enabled = True Then
                cmd_Sort.SetFocus
                DoEvents
                SendKeys "{Enter}"
            End If
        Case KEY_F7
            KeyCode = 0
            If cmd_previous.Enabled Then
                cmd_previous.SetFocus
                DoEvents
                SendKeys "{Enter}"
            End If
        Case KEY_F8
            KeyCode = 0
            If cmd_next.Enabled Then
                cmd_next.SetFocus
                DoEvents
                SendKeys "{Enter}"
            End If
        Case KEY_F2
            KeyCode = 0
            If cmd_query.Enabled Then
                cmd_query.SetFocus
                DoEvents
                SendKeys "{Enter}"
            End If
        Case KEY_PAUSE, KEY_ESCAPE
            KeyCode = 0
            If cmd_exit.Enabled Then
                cmd_exit.SetFocus
                DoEvents
                SendKeys "{Enter}"
            End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    sts_msgline.Panels(1) = SetMessage(G_AP_STATE)
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
    KeyPress KeyAscii           'Enter時自動跳到下一欄位

End Sub

Private Sub Form_Load()
    FormCenter Me
    Me.MousePointer = 11
    OpenLGuiDB
    Set_Property
End Sub

Private Sub MoveDB2Field()
Dim a_count&, A_Index%
Dim A_A1603$, A_Section$

    spd_gd.MaxRows = 0
    a_count& = 0: A_Index% = 0
    Do Until DY_A161.EOF
        a_count& = a_count& + 1
        spd_gd.MaxRows = a_count&
        spd_gd.Row = a_count&
        spd_gd.Col = 1
        spd_gd.text = Trim(DY_A161.Fields("A1601") & "")
        spd_gd.Col = 2
        spd_gd.text = Trim(DY_A161.Fields("A1602") & "")
        spd_gd.Col = 3
        ''
        A_Section$ = " A1603Title_" & ReplaceSingleSign(Trim(DY_A161.Fields("A1601") & ""))
        A_A1603$ = GetFldStrFromSINI(DB_ARTHGUI, A_Section$, Trim$(DY_A161.Fields("A1603") & ""))
        ''
        spd_gd.text = A_A1603$
        spd_gd.Col = 4
        spd_gd.text = Trim(DY_A161.Fields("A1614") & "")
        spd_gd.Col = 5
        spd_gd.text = Trim(DY_A161.Fields("A1605") & "")
        spd_gd.Col = 6
        spd_gd.text = Trim(DY_A161.Fields("A16041") & "")
        spd_gd.Col = 7
        spd_gd.text = Trim(DY_A161.Fields("A1609") & "")
        'S000801046增加傳真號碼
        spd_gd.Col = 8
        spd_gd.text = Trim(DY_A161.Fields("A1606") & "")
        'A_Index% = A_Index% + 1
        spd_gd.TopRow = SetSpreadTopRow(spd_gd)
'        If a_index% > 15 Then
'           If spd_gd.Visible = True Then
'              spd_gd.SetFocus
'           End If
'           SendKeys "{PgDn}"
'           retcode = DoEvents()
'           a_index% = 0
'        End If
        DY_A161.MoveNext
    Loop
    spd_gd.SetFocus
    spd_gd.TopRow = 1
'    spd_gd.Row = 1: spd_gd.Col = 1
'    spd_gd.Action = SS_ACTION_ACTIVE_CELL
'    sts_msgline.Panels(1) = GetSIniStr("PgmMsg", "g_ap_query")
End Sub

Private Sub OpenMainFile()
On Local Error GoTo MY_Error
Dim A_Sql$, A_A1613$

    A_Sql$ = ""
    If Chk_Source(0).Value = "1" Then A_A1613$ = "'" & 1 & "'"
    If Chk_Source(1).Value = "1" Then
        If Trim(A_A1613$) <> "" Then
            A_A1613$ = Trim(A_A1613$) & "," & "'" & 2 & "'"
        Else
            A_A1613$ = "'" & 2 & "'"
        End If
    End If
    If Chk_Source(2).Value = "1" Then
        If Trim(A_A1613$) <> "" Then
            A_A1613$ = Trim(A_A1613$) & "," & "'" & 3 & "'"
        Else
            A_A1613$ = "'" & 3 & "'"
        End If
    End If
    
    A_Sql$ = ""
    If Trim$(txt_Condition(0).text) <> "" Then
       A_Sql$ = A_Sql$ & " WHERE A1601 LIKE '" & txt_Condition(0).text & GetLikeStr(DB_ARTHGUI, True) & "'"
    End If
    If Trim$(txt_Condition(1).text) <> "" Then
       If Trim(A_Sql$) <> "" Then
          A_Sql$ = A_Sql$ & " AND A1602 LIKE '" & txt_Condition(1).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       Else
          A_Sql$ = A_Sql$ & " WHERE A1602 LIKE '" & txt_Condition(1).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       End If
    End If
    If Trim$(txt_Condition(6).text) <> "" Then
       If Trim(A_Sql$) <> "" Then
          A_Sql$ = A_Sql$ & " AND A1603 LIKE '" & txt_Condition(6).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       Else
          A_Sql$ = A_Sql$ & " WHERE A1603 LIKE '" & txt_Condition(6).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       End If
    End If
    If Trim$(txt_Condition(2).text) <> "" Then
       If Trim(A_Sql$) <> "" Then
          A_Sql$ = A_Sql$ & " AND A1614 LIKE '" & txt_Condition(2).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       Else
          A_Sql$ = A_Sql$ & " WHERE A1614 LIKE '" & txt_Condition(2).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       End If
    End If
    If Trim$(txt_Condition(3).text) <> "" Then
       If Trim(A_Sql$) <> "" Then
          A_Sql$ = A_Sql$ & " AND A1605 LIKE '" & txt_Condition(3).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       Else
          A_Sql$ = A_Sql$ & " WHERE A1605 LIKE '" & txt_Condition(3).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       End If
    End If
    If Trim$(txt_Condition(4).text) <> "" Then
       If Trim(A_Sql$) <> "" Then
          A_Sql$ = A_Sql$ & " AND A16041 LIKE '" & txt_Condition(4).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       Else
          A_Sql$ = A_Sql$ & " WHERE A16041 LIKE '" & txt_Condition(4).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       End If
    End If
    If Trim$(txt_Condition(5).text) <> "" Then
       If Trim(A_Sql$) <> "" Then
          A_Sql$ = A_Sql$ & " AND A1609 LIKE '" & txt_Condition(5).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       Else
          A_Sql$ = A_Sql$ & " WHERE A1609 LIKE '" & txt_Condition(5).text & GetLikeStr(DB_ARTHGUI, True) & "'"
       End If
    End If
    'If Trim$(txt_Condition(0).Text) <> "" Or A_Sql$ = "" Then
        If Trim(A_A1613$) <> "" Then
            If Trim(A_Sql$) <> "" Then
                A_Sql$ = A_Sql$ & " AND A1613 In (" & Trim(A_A1613$) & ")"
            Else
                A_Sql$ = A_Sql$ & " Where A1613 In (" & Trim(A_A1613$) & ")"
            End If
        End If
    'End If
    If chk_Sort.Value Then
        A_Sql$ = A_Sql$ & "ORDER BY A1601 DESC"
    Else
        A_Sql$ = A_Sql$ & "ORDER BY A1601"
    End If
    A_Sql$ = "SELECT * FROM  A16 " & A_Sql$
    CreateDynasetODBC DB_ARTHGUI, DY_A161, A_Sql$, "DY_A161", True
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Sub Set_Property()
    frm_GD.FontBold = False
    
    Form_Property frm_GD, GetSINISTR_GD("formtitle", "GD"), G_Font_Name
    
    Label_Property pnl_Condition(0), GetSINISTR_GD("PanelDescpt", "buyerid"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property pnl_Condition(1), GetSINISTR_GD("PanelDescpt", "abb_name_c"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property pnl_Condition(2), GetSINISTR_GD("PanelDescpt", "liaison"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property pnl_Condition(3), GetSINISTR_GD("TSM02", "TEL_NO"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property pnl_Condition(4), GetSINISTR_GD("MCFGD", "address0"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property pnl_Condition(5), GetSINISTR_GD("PanelDescpt", "unifyno"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property pnl_Condition(6), GetSINISTR_GD("PanelDescpt", "full_name_c"), G_Label_Color, G_Font_Size, G_Font_Name
    
    Label_Property Lbl_Source, GetSINISTR_GD("MCFGD", "GD13P"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Pnl_Source, "", G_Label_Color, G_Font_Size, G_Font_Name
    
    Text_Property txt_Condition(0), 10, G_Font_Name
    Text_Property txt_Condition(1), 12, G_Font_Name
    Text_Property txt_Condition(2), 10, G_Font_Name
    Text_Property txt_Condition(3), 40, G_Font_Name
    Text_Property txt_Condition(4), 15, G_Font_Name
    Text_Property txt_Condition(5), 10, G_Font_Name
    Text_Property txt_Condition(6), 40, G_Font_Name

    'fra_Condition.Caption = GetSIniStr("A16", "sortcondition")
    Checkbox_Property chk_Sort, GetSINISTR_GD("GD", "desc"), G_Font_Size, G_Font_Name
    
    Checkbox_Property Chk_Source(0), GetSINISTR_GD("PanelDescpt", "customer"), G_Font_Size, G_Font_Name
    Checkbox_Property Chk_Source(1), GetSINISTR_GD("PanelDescpt", "firm"), G_Font_Size, G_Font_Name
    Checkbox_Property Chk_Source(2), GetSINISTR_GD("GDP01", "bank"), G_Font_Size, G_Font_Name
    
    
    Command_Property cmd_help, G_CmdHelp, G_Font_Name
    Command_Property cmd_Sort, G_CmdSort, G_Font_Name
    Command_Property cmd_exit, G_CmdExit, G_Font_Name
    Command_Property cmd_previous, G_CmdPrvPage, G_Font_Name
    Command_Property cmd_query, G_CmdQuery, G_Font_Name
    Command_Property cmd_next, G_CmdNxtPage, G_Font_Name
    VSElastic_Property vse_background
    StatusBar_ProPerty sts_msgline
    Set_Spread_Property
End Sub

Private Sub Set_Spread_Property()
    'S000801046增加傳真號碼
    spd_gd.UnitType = 2
    Spread_Property spd_gd, 500, 8, WHITE, G_Font_Size, G_Font_Name

    Spread_Col_Property spd_gd, 4, TextWidth("A") * 14, GetSINISTR_GD("PanelDescpt", "liaison")
    Spread_Col_Property spd_gd, 5, TextWidth("A") * 15, GetSINISTR_GD("TSM02", "TEL_NO")
    Spread_Col_Property spd_gd, 6, TextWidth("A") * 20, GetSINISTR_GD("MCFGD", "address0")
    Spread_Col_Property spd_gd, 7, TextWidth("A") * 10, GetSINISTR_GD("PanelDescpt", "unifyno")
    Spread_Col_Property spd_gd, 8, TextWidth("A") * 10, GetSINISTR_GD("PanelDescpt", "faxno")
    'retcode = DoEvents()
    Spread_DataType_Property spd_gd, 1, SS_CELL_TYPE_STATIC_TEXT, "", "", 10
    Spread_DataType_Property spd_gd, 2, SS_CELL_TYPE_STATIC_TEXT, "", "", 12
    Spread_DataType_Property spd_gd, 3, SS_CELL_TYPE_STATIC_TEXT, "", "", 100 '欄位加大
    Spread_DataType_Property spd_gd, 4, SS_CELL_TYPE_STATIC_TEXT, "", "", 10
    Spread_DataType_Property spd_gd, 5, SS_CELL_TYPE_STATIC_TEXT, "", "", 15
    Spread_DataType_Property spd_gd, 6, SS_CELL_TYPE_STATIC_TEXT, "", "", 40
    Spread_DataType_Property spd_gd, 7, SS_CELL_TYPE_STATIC_TEXT, "", "", 10
    Spread_DataType_Property spd_gd, 8, SS_CELL_TYPE_STATIC_TEXT, "", "", 15
    'retcode = DoEvents()
    spd_gd.ColsFrozen = 1
End Sub

Private Sub Form_Resize()
'    vse_background.AutoSizeChildren = azProportional
'    vse_background.Align = asFill

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Cmd_Exit_Click
End Sub

Private Sub spd_gd_Click(ByVal Col As Long, ByVal Row As Long)
'於Column Heading Click時, 依該欄位排序
    If Row = 0 Then SpreadSort spd_gd, Col
End Sub

Private Sub spd_GD_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim A_Row As Long, A_GD01$, A_GD02$, A_GD03$
    
    If Row <= 0 Then Exit Sub
    '
    A_Row = Row
    With spd_gd
        .Row = A_Row
        .Col = 1: A_GD01$ = Trim(.text)
        .Col = 2: A_GD02$ = Trim(.text)
        .Col = 3: A_GD03$ = Trim(.text)
    End With
    Me.MousePointer = HOURGLASS
    frm_GD.Tag = A_GD01$ & Chr$(KEY_TAB) & A_GD02$ & Chr$(KEY_TAB) & A_GD03$
    Me.Hide
    Me.MousePointer = Default
End Sub

Private Sub spd_GD_GotFocus()
    SpreadGotFocus spd_gd.ActiveCol, spd_gd.ActiveRow
End Sub

Private Sub spd_gd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_RETURN Then
       KeyCode = 0
       spd_GD_DblClick spd_gd.ActiveCol, spd_gd.ActiveRow
    End If
End Sub

Private Sub spd_GD_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    spd_gd.Row = Row
    spd_gd.Col = -1
    spd_gd.BackColor = G_TextLostBack_Color
    spd_gd.ForeColor = G_TextLostFore_Color
    If NewRow > 0 And NewCol > 0 Then
        spd_gd.Row = NewRow
        spd_gd.Col = -1
        spd_gd.BackColor = G_TextGotBack_Color
        spd_gd.ForeColor = G_TextGotFore_Color
    End If
End Sub

Private Sub spd_gd_LostFocus()
    SpreadLostFocus spd_gd.ActiveCol, spd_gd.ActiveRow
End Sub


Private Sub txt_Condition_GotFocus(index As Integer)
    txt_Condition(index).BackColor = G_TextGotBack_Color
    txt_Condition(index).ForeColor = G_TextGotFore_Color
    'txt_Condition(Index).Text = Trim$(txt_Condition(Index).Text)
    txt_Condition(index).SelStart = 0
    txt_Condition(index).SelLength = Len(txt_Condition(index).text)
End Sub

Private Sub txt_Condition_LostFocus(index As Integer)
    txt_Condition(index).BackColor = G_TextLostBack_Color
    txt_Condition(index).ForeColor = G_TextLostFore_Color
    txt_Condition(index).text = Trim$(txt_Condition(index).text)
End Sub

Private Sub Vse_background_Click()
    vse_background.AutoSizeChildren = azNone

End Sub

