VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Begin VB.Form frm_RptDef 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "表格設定"
   ClientHeight    =   5040
   ClientLeft      =   300
   ClientTop       =   435
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "RptDef.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5040
   ScaleWidth      =   8310
   Tag             =   "rptset"
   Begin VsOcxLib.VideoSoftElastic Vse_background 
      Height          =   4665
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8310
      _Version        =   327680
      _ExtentX        =   14658
      _ExtentY        =   8229
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Align           =   5
      BevelOuter      =   6
      Picture         =   "RptDef.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "RptDef.frx":0326
      Begin VB.CommandButton Cmd_Cancel 
         Appearance      =   0  'Flat
         Caption         =   "取消(Esc)"
         Height          =   360
         Left            =   6990
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4200
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         Height          =   4500
         Left            =   60
         TabIndex        =   11
         Top             =   90
         Width           =   6855
         Begin VsOcxLib.VideoSoftElastic vse_background2 
            Height          =   4500
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   6855
            _Version        =   327680
            _ExtentX        =   12091
            _ExtentY        =   7937
            _StockProps     =   70
            ConvInfo        =   1418783674
            Align           =   5
            BevelOuter      =   6
            Picture         =   "RptDef.frx":0342
            MouseIcon       =   "RptDef.frx":035E
            Begin FPSpread.vaSpread Spd_ColSort 
               Height          =   1635
               Left            =   3900
               OleObjectBlob   =   "RptDef.frx":037A
               TabIndex        =   2
               Top             =   2760
               Width           =   2845
            End
            Begin FPSpread.vaSpread Spd_ColSelect 
               Height          =   1995
               Left            =   3900
               OleObjectBlob   =   "RptDef.frx":058B
               TabIndex        =   1
               Top             =   420
               Width           =   2845
            End
            Begin FPSpread.vaSpread Spd_Cols 
               Height          =   3975
               Left            =   90
               OleObjectBlob   =   "RptDef.frx":079C
               TabIndex        =   0
               Top             =   420
               Width           =   2445
            End
            Begin VB.CommandButton Cmd_RemoveS 
               Appearance      =   0  'Flat
               Caption         =   "移除(&R)"
               Height          =   360
               Left            =   2610
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   3840
               Width           =   1200
            End
            Begin VB.CommandButton Cmd_AddC 
               Appearance      =   0  'Flat
               Caption         =   "新增(&C)"
               Height          =   360
               Left            =   2610
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   1230
               Width           =   1200
            End
            Begin VB.CommandButton Cmd_AddS 
               Appearance      =   0  'Flat
               Caption         =   "新增(&S)"
               Height          =   360
               Left            =   2610
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   3420
               Width           =   1200
            End
            Begin VB.CommandButton Cmd_RemoveC 
               Appearance      =   0  'Flat
               Caption         =   "移除(&D)"
               Height          =   360
               Left            =   2610
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   1650
               Width           =   1200
            End
            Begin VB.Label Lbl_Cols 
               Appearance      =   0  'Flat
               Caption         =   "報表所有欄位清單"
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   90
               TabIndex        =   15
               Top             =   90
               Width           =   2130
            End
            Begin VB.Label Lbl_ColSelect 
               Appearance      =   0  'Flat
               Caption         =   "已選列印欄位清單"
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   3900
               TabIndex        =   14
               Top             =   90
               Width           =   2400
            End
            Begin VB.Label Lbl_ColSort 
               Appearance      =   0  'Flat
               Caption         =   "排序欄位清單(最多三欄)"
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   3900
               TabIndex        =   13
               Top             =   2460
               Width           =   2800
            End
         End
      End
      Begin VB.CommandButton cmd_ok 
         Appearance      =   0  'Flat
         Caption         =   "確認(F11)"
         Height          =   360
         Left            =   6990
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   1275
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4665
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_RptDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim DB_LGUI As Database
Dim TB_CommIni As Recordset

Dim tSpd_Cols As Spread
Dim tSpd_ColSelect As Spread
Dim tSpd_ColSort As Spread

Dim M_OrgCols As String
Dim M_OrgSorts As String

Private Sub AddSpreadRows(Spd As vaSpread, tSPD As Spread)
'新增報表欄位顯示或排序欄位
Dim I&, j&, a_count&, A_Row&, A_Row2&, A_Col&, A_STR$

    '選取Spread上的所有區塊
    Spd_Cols.Action = SS_ACTION_GET_MULTI_SELECTION
    
    '如果無選取區塊,則以目前所在列處理
    If Spd_Cols.MultiSelCount = 0 Then
       
       '設定區塊範圍列為目前所在列
       A_Row& = Spd_Cols.ActiveRow
       A_Row2& = A_Row&
        
       '比對ID欄位
       GoSub CompareID
       Exit Sub
        
    End If
    
    '取得區塊數
    a_count& = Spd_Cols.MultiSelCount
    
    '取得ID欄位名稱的欄數
    A_Col& = GetSpdColIndex(tSpd_RptDef, "ID")
    
    '加入欄位至顯示欄位或排序欄位,若已存在則不加入
    For I& = 0 To a_count& - 1
    
        '設定目前欲讀取的區塊索引值(由0開始)
        Spd_Cols.MultiSelIndex = I&
        
        '利用Row及Row2屬性取得目前區塊的起始及截止的列數
        A_Row& = Spd_Cols.Row
        A_Row2& = Spd_Cols.Row2
        
        '若選取整個Spread,則設定起始列為1,截止列為總列數
        If A_Row& = -1 And A_Row2& = -1 Then
           A_Row& = 1
           A_Row2& = Spd_Cols.MaxRows
        End If
        
        '比對ID欄位
        GoSub CompareID
        
    Next I&
    
ProcessA:

    '取消Spread上的所有區塊
    Spd_Cols.Action = SS_ACTION_DESELECT_BLOCK
    Exit Sub
    
CompareID:

    '比對ID欄位,不存在才加入
    For j& = A_Row& To A_Row2&
    
        '排序欄位最多只能挑選三個欄位(Spread的限制)
        If (Spd Is Spd_ColSort) And Spd.MaxRows >= 3 Then GoTo ProcessA
        
        '取得待選欄位清單中的ID欄位值
        A_STR$ = GetSpdText(Spd_Cols, tSpd_Cols, "ID", j&)
        
        '若已選欄位清單中,尚無此欄位則加入
        If Not IsExist(Spd, tSPD, A_STR$) Then
        
           '將總列數加一
           Spd.MaxRows = Spd.MaxRows + 1
           
           '寫入已選欄位清單的ID欄位值
           SetSpdText Spd, tSPD, "ID", Spd.MaxRows, A_STR$
           
           '取得待選欄位清單中的Name欄位值
           A_STR$ = GetSpdText(Spd_Cols, tSpd_Cols, "Name", j&)
           
           '寫入已選欄位清單的Name欄位值
           SetSpdText Spd, tSPD, "Name", Spd.MaxRows, A_STR$
           
        End If
    Next j&
    Return
End Sub


Sub CheckDefCols()
'自螢幕顯示畫面進入此表單時,於此表單結束前,判斷顯示欄位與排序欄位資料是否有異動
Dim A_Cols$, A_Sorts$
    
    '將目前設定資料Keep至變數中
    GetRptDefSet2Str A_Cols$, A_Sorts$
    
    '若設定無異動,則Screen無須重新Prepare
    tSpd_RptDef.RefreshCol = StrComp(A_Cols$, M_OrgCols, vbTextCompare) <> 0
    tSpd_RptDef.RefreshSort = (StrComp(A_Sorts$, M_OrgSorts, vbTextCompare) <> 0) And tSpd_RptDef.Refresh

    '還原初始值
    tSpd_RptDef.Refresh = False
End Sub

Sub GetRptDefSet2Str(Cols$, Sorts$)
'將Spread Type中的Columns及Sorts Type資料轉換成字串型態,以利比較於此表單中是否有異動其內容
Dim I%
    
    '將報表所有的顯示欄位組成一個字串,欄與欄間以分號區隔
    Cols$ = ""
    For I% = 1 To UBound(tSpd_RptDef.Columns)
        If tSpd_RptDef.Columns(I%).Hidden = 0 Then
           Cols$ = Cols$ & tSpd_RptDef.Columns(I%).Name & ";"
        End If
    Next I%
    
    '將報表所有的排序欄位組成一個字串,遞減欄位於欄位名稱前加一減號,欄與欄間以分號區隔
    Sorts$ = ""
    For I% = 1 To UBound(tSpd_RptDef.Sorts)
        If Trim(tSpd_RptDef.Sorts(I%).SortKey) = "" Then Exit For
        If tSpd_RptDef.Sorts(I%).SortOrder = SS_SORT_ORDER_DESCENDING Then
           Sorts$ = Sorts$ & "-"
        End If
        Sorts$ = Sorts$ & tSpd_RptDef.Sorts(I%).SortKey & ";"
    Next I%
End Sub


Function IsAllFieldCheck() As Boolean
    IsAllFieldCheck = False
    
    If Spd_ColSelect.MaxRows = 0 Then
       Sts_MsgLine.Panels(1) = Lbl_ColSelect & G_MustInput
       Spd_Cols.SetFocus
       Exit Function
    End If
    
    If Spd_ColSort.MaxRows = 0 Then
       Sts_MsgLine.Panels(1) = Lbl_ColSort & G_MustInput
       Spd_Cols.SetFocus
       Exit Function
    End If
    
    IsAllFieldCheck = True
End Function

Private Function IsExist(Spd As vaSpread, tSPD As Spread, ByVal Id$) As Boolean
'判斷欲加入的欄位是否已存在已選清單中
Dim I&, a_count&, A_STR$

    IsExist = False

    a_count& = Spd.MaxRows
    If a_count& = 0 Then Exit Function

    For I& = 1 To a_count&
        A_STR$ = GetSpdText(Spd, tSPD, "ID", I&)
        If StrComp(Id$, A_STR$, vbTextCompare) = 0 Then
           IsExist = True
           Exit For
        End If
    Next I&
End Function
Sub PrepareComboBox()
'Prepare排序欄位的遞增及遞減資料

    Spd_ColSort.Row = -1
    Spd_ColSort.Col = 2
    Spd_ColSort.TypeComboBoxList = GetSIniStr("RptDef", "ascending") & _
                                   Chr(KEY_TAB) & GetSIniStr("RptDef", "descending")
    Spd_ColSort.TypeComboBoxEditable = False
    Spd_ColSort.TypeComboBoxCurSel = 0
End Sub

Private Sub RemoveSpreadRows(Spd As vaSpread, tSPD As Spread)
'移除報表欄位顯示或排序欄位
Dim I&, j&, a_count&, A_Row&, A_Row2&
Dim A_Flag%

    '選取Spread上的所有區塊
    Spd.Action = SS_ACTION_GET_MULTI_SELECTION
    
    '如果無選取區塊,則以目前所在列處理
    If Spd.MultiSelCount = 0 Then
    
       'Spread中已無列數,則不處理刪除列的動作
       If Spd.ActiveRow < 1 Then Exit Sub
       
       '設定區塊範圍列為目前所在列
       A_Row& = Spd.ActiveRow
       A_Row2& = A_Row&
        
       '將區塊範圍列的列首,填入"D"字元,視為將刪除的列
       GoSub CompareID
       GoTo ProcessA
        
    End If
    
    '取得區塊數
    a_count& = Spd.MultiSelCount
    
    '取得每個區塊中的列範圍,並在列首填入"D"字元,以判斷是否刪除該列
    For I& = 0 To a_count& - 1
    
        '設定目前欲讀取的區塊索引值(由0開始)
        Spd.MultiSelIndex = I&
        
        '利用Row及Row2屬性取得目前區塊的起始及截止的列數
        A_Row& = Spd.Row
        A_Row2& = Spd.Row2
        
        '若選取整個Spread,則設定起始列為1,截止列為總列數
        If A_Row& = -1 And A_Row2& = -1 Then
           A_Row& = 1
           A_Row2& = Spd.MaxRows
        End If
        
        '將區塊範圍列的列首,填入"D"字元,視為將刪除的列
        GoSub CompareID
        
    Next I&
    
ProcessA:

    '取消Spread上的所有區塊
    Spd.Action = SS_ACTION_DESELECT_BLOCK
    
    '刪除列首字元為"D"的列
    For I& = 1 To Spd.MaxRows
        If I& > Spd.MaxRows Then Exit For
        Spd.Row = I&: Spd.Col = 0
        If StrComp(Spd.text, "D", vbTextCompare) = 0 Then
           Spd.Action = SS_ACTION_DELETE_ROW
           Spd.MaxRows = Spd.MaxRows - 1
           I& = I& - 1
        End If
    Next I&
    Exit Sub
    
CompareID:

    '將區塊範圍列的列首,填入"D"字元,視為將刪除的列
    For j& = A_Row& To A_Row2&
        A_Flag% = False
        If Spd Is Spd_ColSelect Then
           If Val(GetSpdText(Spd, tSPD, "BreakCol", j&)) > 0 Then
              A_Flag% = True
           End If
        End If
        If Not A_Flag% Then
           Spd.Row = j&
           Spd.Col = 0
           Spd.text = "D"
        End If
    Next j&
    Return
End Sub
Private Sub SaveDefaultValue()
'按下確認時,將目前設定Update至tSpd_RptDef Type中
Dim tSpd_Temp As Spread
Dim I%, A_Cols%

    A_Cols% = UBound(tSpd_RptDef.Columns)
    For I% = 1 To A_Cols%
        tSpd_RptDef.Columns(I%).ReportIndex = 0
        tSpd_RptDef.Columns(I%).ScreenIndex = 0
        If tSpd_RptDef.Columns(I%).Hidden <> 2 Then
           tSpd_RptDef.Columns(I%).Hidden = 1
        End If
    Next I%
    
    ReDim tCols(1 To Spd_ColSelect.MaxRows) As SpreadCol
    ReDim tSorts(1 To Spd_ColSort.MaxRows) As SpreadSort
    tSpd_Temp.Columns = tCols
    tSpd_Temp.Sorts = tSorts
    
    For I% = 1 To Spd_ColSelect.MaxRows
        tSpd_Temp.Columns(I%).Name = GetSpdText(Spd_ColSelect, tSpd_ColSelect, "ID", I%)
        tSpd_Temp.Columns(I%).ReportIndex = I%
    Next I%
    
    For I% = 1 To Spd_ColSort.MaxRows
        tSpd_Temp.Sorts(I%).SortKey = GetSpdText(Spd_ColSort, tSpd_ColSort, "ID", I%)
        tSpd_Temp.Sorts(I%).SortOrder = Val(GetSpdText(Spd_ColSort, tSpd_ColSort, "Order", I%, , , , 1)) + 1
    Next I%

    SetColPosition tSpd_RptDef, tSpd_Temp
End Sub


Sub SetDefaultValue()
'將報表的設定值,顯示到Spread上
Dim I%, A_Cols%
    
    '設定所有可供挑選的欄位清單
    A_Cols% = UBound(tSpd_RptDef.Columns)
    Spd_Cols.MaxRows = A_Cols%
    If A_Cols% > 0 Then
       For I% = 1 To A_Cols%
           If tSpd_RptDef.Columns(I%).Hidden <> 2 Then
              SetSpdText Spd_Cols, tSpd_Cols, "Name", tSpd_RptDef.Columns(I%).SelectIndex, tSpd_RptDef.Columns(I%).Caption
              SetSpdText Spd_Cols, tSpd_Cols, "ID", tSpd_RptDef.Columns(I%).SelectIndex, tSpd_RptDef.Columns(I%).Name
              SetSpdText Spd_Cols, tSpd_Cols, "BreakCol", tSpd_RptDef.Columns(I%).SelectIndex, tSpd_RptDef.Columns(I%).BreakIndex
              tSpd_RptDef.Columns(I%).TempIndex = tSpd_RptDef.Columns(I%).ScreenIndex
           End If
       Next I%
    End If
    Spd_Cols.MaxRows = Spd_Cols.DataRowCnt
    
    '設定已挑選的顯示欄位清單
    Spd_ColSelect.MaxRows = A_Cols%
    If A_Cols% > 0 Then
       For I% = 1 To A_Cols%
           If tSpd_RptDef.Columns(I%).Hidden = 0 Then
              SetSpdText Spd_ColSelect, tSpd_ColSelect, "Name", tSpd_RptDef.Columns(I%).ScreenIndex, tSpd_RptDef.Columns(I%).Caption
              SetSpdText Spd_ColSelect, tSpd_ColSelect, "ID", tSpd_RptDef.Columns(I%).ScreenIndex, tSpd_RptDef.Columns(I%).Name
              SetSpdText Spd_ColSelect, tSpd_ColSelect, "BreakCol", tSpd_RptDef.Columns(I%).ScreenIndex, tSpd_RptDef.Columns(I%).BreakIndex
              If tSpd_RptDef.Columns(I%).BreakIndex > 0 Then
                 Spd_ColSelect.Col = -1
                 Spd_ColSelect.FontBold = True
              End If
           End If
       Next I%
    End If
    Spd_ColSelect.MaxRows = Spd_ColSelect.DataRowCnt
    
    '設定已挑選的排序欄位清單
    A_Cols% = UBound(tSpd_RptDef.Sorts)
    Spd_ColSort.MaxRows = A_Cols%
    If A_Cols% > 0 Then
       For I% = 1 To A_Cols%
           If Trim(tSpd_RptDef.Sorts(I%).SortKey) = "" Then Exit For
           SetSpdText Spd_ColSort, tSpd_ColSort, "Name", I%, tSpd_RptDef.Columns(GetSpdColIndex(tSpd_RptDef, tSpd_RptDef.Sorts(I%).SortKey)).Caption
           SetSpdText Spd_ColSort, tSpd_ColSort, "Order", I%, tSpd_RptDef.Sorts(I%).SortOrder - 1, , , 1
           SetSpdText Spd_ColSort, tSpd_ColSort, "ID", I%, tSpd_RptDef.Sorts(I%).SortKey
       Next I%
    End If
    Spd_ColSort.MaxRows = Spd_ColSort.DataRowCnt
End Sub

Private Sub Set_Property()
    frm_RptDef.FontBold = False
    
'設定本Form之標題,字形及色系
    Form_Property frm_RptDef, GetRptSet("RptDef", "formtitle"), G_Font_Name
    
'設定Form中所有Panel,Label,OptionButton,CheckBox,Frame之標題, 字形及色系
    Control_Property Lbl_Cols, GetRptSet("RptDef", "rptallcols")
    Control_Property Lbl_ColSelect, GetRptSet("RptDef", "selectedcols")
    Control_Property Lbl_ColSort, GetRptSet("RptDef", "sortcols")

'設Form中所有Command之標題及字形
    Command_Property Cmd_AddC, GetRptSet("RptDef", "addcols"), G_Font_Name
    Command_Property Cmd_RemoveC, GetRptSet("RptDef", "removecols"), G_Font_Name
    Command_Property Cmd_AddS, GetRptSet("RptDef", "addsort"), G_Font_Name
    Command_Property Cmd_RemoveS, GetRptSet("RptDef", "removesort"), G_Font_Name
    Command_Property cmd_ok, GetRptSet("CmdDescpt", "cmd_ok"), G_Font_Name
    Command_Property Cmd_Cancel, GetRptSet("CmdDescpt", "cmd_cancel"), G_Font_Name
    
    Set_Spread_Property
    StatusBar_ProPerty Sts_MsgLine
    VSElastic_Property Vse_Background
    VSElastic_Property2 vse_background2
End Sub
Private Sub Set_Spread_Property()
    With Spd_Cols
         .UnitType = 2
         '設定本Spread之筆數及欄位數
         Spread_Property Spd_Cols, 0, 2, WHITE, G_Font_Size, G_Font_Name
         '設定本Spread之各欄標題及顯示寬度,各欄屬性及顯示字數
         SpdFldProperty Spd_Cols, tSpd_Cols, "Name", TextWidth("X") * 10, GetSIniStr("RptDef", "colname"), SS_CELL_TYPE_EDIT, "", "", 100
         SpdFldProperty Spd_Cols, tSpd_Cols, "ID", TextWidth("X") * 10, "ID", SS_CELL_TYPE_EDIT, "", "", 20
         .AllowMultiBlocks = True
         .AllowDragDrop = True
         .OperationMode = OperationModeNormal
         '設定Block選取方式為Row
         .SelectBlockOptions = 15
         '鎖住Spread不可修改
         .Row = -1: .Col = -1: .Lock = True
    End With
    
    With Spd_ColSelect
         .UnitType = 2
         '設定本Spread之筆數及欄位數
         Spread_Property Spd_ColSelect, 0, 3, WHITE, G_Font_Size, G_Font_Name
         '設定本Spread之各欄標題及顯示寬度,各欄屬性及顯示字數
         SpdFldProperty Spd_ColSelect, tSpd_ColSelect, "Name", TextWidth("X") * 10, GetSIniStr("RptDef", "colname"), SS_CELL_TYPE_EDIT, "", "", 100
         SpdFldProperty Spd_ColSelect, tSpd_ColSelect, "ID", TextWidth("X") * 10, "ID", SS_CELL_TYPE_EDIT, "", "", 20, , , True
         SpdFldProperty Spd_ColSelect, tSpd_ColSelect, "BreakCol", TextWidth("X") * 10, "BreakCol", SS_CELL_TYPE_FLOAT, "0", "999"
         .AllowMultiBlocks = True
         .AllowDragDrop = True
         .OperationMode = OperationModeNormal
         .SelectBlockOptions = 15
         '鎖住Spread不可修改
         .Row = -1: .Col = -1: .Lock = True
    End With
    
    With Spd_ColSort
         .UnitType = 2
         '設定本Spread之筆數及欄位數
         Spread_Property Spd_ColSort, 0, 3, WHITE, G_Font_Size, G_Font_Name
         '設定本Spread之各欄標題及顯示寬度,各欄屬性及顯示字數
         SpdFldProperty Spd_ColSort, tSpd_ColSort, "Name", TextWidth("X") * 10, GetSIniStr("RptDef", "colname"), SS_CELL_TYPE_EDIT, "", "", 100
         SpdFldProperty Spd_ColSort, tSpd_ColSort, "Order", TextWidth("X") * 6, GetSIniStr("RptDef", "sortorder"), SS_CELL_TYPE_COMBOBOX
         SpdFldProperty Spd_ColSort, tSpd_ColSort, "ID", TextWidth("X") * 10, "ID", SS_CELL_TYPE_EDIT, "", "", 20
         PrepareComboBox
         .AllowMultiBlocks = True
         .AllowDragDrop = (tSpd_RptDef.SortEnable)
         .DisplayRowHeaders = False
         .OperationMode = OperationModeNormal
         .SelectBlockOptions = 15
         '鎖住Spread不可修改
         .Row = -1: .Col = -1: .Lock = True
         If Not tSpd_RptDef.SortEnable Then
            Lbl_ColSort.Enabled = False
            Cmd_AddS.Enabled = False
            Cmd_RemoveS.Enabled = False
         Else
            .Col = 2: .Lock = False
         End If
    End With
End Sub


Sub SetReportCols()
    '宣告Spread型態的Columns及Sorts的陣列個數
    InitialCols tSpd_Cols, 2, False
    InitialCols tSpd_ColSelect, 3, False
    InitialCols tSpd_ColSort, 3, False
    
    '設定報表顯示的欄位及排序欄位至Spread Type中
    AddReportCol tSpd_Cols, "Name"
    AddReportCol tSpd_Cols, "ID", 2
    AddReportCol tSpd_ColSelect, "Name"
    AddReportCol tSpd_ColSelect, "ID", 2
    AddReportCol tSpd_ColSelect, "BreakCol", 2
    AddReportCol tSpd_ColSort, "Name"
    AddReportCol tSpd_ColSort, "Order"
    AddReportCol tSpd_ColSort, "ID", 2
    
    '抓取User自訂報表之欄位顯示順序及排序欄位
    GetSpreadDefault tSpd_Cols, "frm_RptDef", "Spd_Cols"
    GetSpreadDefault tSpd_ColSelect, "frm_RptDef", "Spd_ColSelect"
    GetSpreadDefault tSpd_ColSort, "frm_RptDef", "Spd_ColSort"
End Sub

Private Sub Cmd_AddC_Click()
'於報表加入新的欄位顯示
    AddSpreadRows Spd_ColSelect, tSpd_ColSelect
End Sub

Private Sub Cmd_AddS_Click()
'於報表加入新的排序欄位
    AddSpreadRows Spd_ColSort, tSpd_ColSort
End Sub

Private Sub Cmd_Cancel_Click()
    Unload Me
End Sub


Private Sub cmd_OK_Click()
    If Not IsAllFieldCheck() Then Exit Sub
    SaveDefaultValue
    Unload Me
End Sub

Private Sub Cmd_RemoveC_Click()
'自報表移除欄位不顯示
    RemoveSpreadRows Spd_ColSelect, tSpd_ColSelect
End Sub

Private Sub Cmd_RemoveS_Click()
'自報表移除排序欄位
    RemoveSpreadRows Spd_ColSort, tSpd_ColSort
End Sub


Private Sub Form_Activate()
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    Sts_MsgLine.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
         Case KEY_F11
              KeyCode = 0
              If cmd_ok.Visible And cmd_ok.Enabled Then
                 cmd_ok.SetFocus
                 DoEvents
                 SendKeys "{Enter}"
              End If
              
         Case KEY_ESCAPE
              KeyCode = 0
              If Cmd_Cancel.Visible And Cmd_Cancel.Enabled Then
                 Cmd_Cancel.SetFocus
                 DoEvents
                 SendKeys "{Enter}"
              End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Sts_MsgLine.Panels(1) = Me.Caption
    KeyPress KeyAscii
End Sub

Private Sub Form_Load()
    Me.MousePointer = HOURGLASS
    
    Sts_MsgLine.Panels(1) = Me.Caption
    Me.Move (Screen.Width - Me.Width) \ 6, (Screen.Height - Me.Height) \ 6

'開啟共用資料庫
    OpenCommLDB
    
'設定報表的所有欄位至Spread Type中
    SetReportCols
    
'設定表單所有物件的屬性
    Set_Property
    
'將原設定資料Keep至變數中
    GetRptDefSet2Str M_OrgCols, M_OrgSorts
    
'將初始值填入Spread中
    SetDefaultValue
    
    Me.MousePointer = Default
End Sub

Function GetRptSet(ByVal Section$, ByVal Topic$) As String
    GetRptSet = " "
    If Trim(DB_LGUI.Connect) <> "" Then
        Dim A_Sql$
        A_Sql$ = "SELECT TOPICVALUE FROM INI"
        A_Sql$ = A_Sql$ & " WHERE SECTION='" & Section$ & "'"
        A_Sql$ = A_Sql$ & " AND TOPIC='" & Topic$ & "'"
        Set DY_INICommon = DB_LGUI.OpenRecordset(A_Sql$, dbOpenSnapshot, dbSQLPassThrough)
        If Not (DY_INICommon.BOF And DY_INICommon.EOF) Then
            GetRptSet = Trim(DY_INICommon.Fields("TOPICVALUE") & "")
        End If
        DY_INICommon.Close
    Else
        TB_CommIni.Seek "=", Section$, Topic$
        If Not TB_CommIni.NoMatch Then
           GetRptSet = TB_CommIni.Fields("TOPICVALUE") & ""
        End If
    End If
End Function


Sub OpenCommLDB()
    Dim A_Path As String
    Dim A_ConnectMethod As String
    
    On Local Error Resume Next
    Screen.MousePointer = HOURGLASS
   'Pick Local INI DataPath String (GL.mdb)
    A_Path = GetIniStr("DBPath", "Path3", "GUI.INI")
    A_ConnectMethod = GetIniStr("DBPath", "Connect3", "GUI.INI")
    Set DB_LGUI = GetEngine.OpenDatabase(A_Path, False, False, A_ConnectMethod)
    If Err Then
       If Trim$(A_ConnectMethod) = "" Then   'Access DataBase
          If Err = 3043 Then
             Err = 0
             DB_LGUI.Close
             Set DB_LGUI = GetEngine.OpenDatabase(A_Path, False, False, A_ConnectMethod)
          ElseIf Err = 3049 Then
             Err = 0
             GetEngine.RepairDatabase A_Path
             Set DB_LGUI = GetEngine.OpenDatabase(A_Path, False, False, A_ConnectMethod)
          End If
       End If
    End If
    If Err Then
       MsgBox Error(Err), MB_ICONEXCLAMATION, App.Title
       End
    End If
    If Trim$(A_ConnectMethod) <> "" Then DB_LGUI.QueryTimeout = 0
    'Open Table
    If Trim(DB_LGUI.Connect) = "" Then
        Set TB_CommIni = DB_LGUI.OpenRecordset("INI", dbOpenTable)
        TB_CommIni.index = "INI"
        Screen.MousePointer = Default
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckDefCols
    G_FormFrom$ = "RptDef"
    If Not TB_CommIni Is Nothing Then TB_CommIni.Close
    DB_LGUI.Close
    Set TB_CommIni = Nothing
    Set DB_LGUI = Nothing
End Sub

Private Sub Spd_Cols_Click(ByVal Col As Long, ByVal Row As Long)
    Sts_MsgLine.Panels(1) = Me.Caption
End Sub

Private Sub Spd_Cols_GotFocus()
    SpreadGotFocus Spd_Cols.ActiveCol, Spd_Cols.ActiveRow
End Sub


Private Sub Spd_Cols_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal newcol As Long, ByVal NewRow As Long, Cancel As Boolean)
    SpreadLostFocus Col, Row
    If NewRow > 0 Then SpreadGotFocus newcol, NewRow
End Sub


Private Sub Spd_ColSelect_Click(ByVal Col As Long, ByVal Row As Long)
    Sts_MsgLine.Panels(1) = Me.Caption
End Sub

Private Sub Spd_ColSelect_DragDropBlock(ByVal Col As Long, ByVal Row As Long, ByVal Col2 As Long, ByVal Row2 As Long, ByVal newcol As Long, ByVal NewRow As Long, ByVal NewCol2 As Long, ByVal NewRow2 As Long, ByVal Overwrite As Boolean, Action As Integer, DataOnly As Boolean, Cancel As Boolean)
Dim A_STR$, A_Bold%

    Cancel = True
    
    Spd_ColSelect.Row = Row
    Spd_ColSelect.Col = 1
    Spd_ColSelect.Row2 = Row
    Spd_ColSelect.Col2 = Spd_ColSelect.MaxCols
    A_STR$ = Spd_ColSelect.Clip
    A_Bold% = Spd_ColSelect.FontBold
    Spd_ColSelect.Action = SS_ACTION_DELETE_ROW
    '
    Spd_ColSelect.Row = NewRow
    Spd_ColSelect.Action = SS_ACTION_INSERT_ROW
    '
    Spd_ColSelect.Row = NewRow
    Spd_ColSelect.Col = 1
    Spd_ColSelect.Row2 = NewRow
    Spd_ColSelect.Col2 = Spd_ColSelect.MaxCols
    Spd_ColSelect.Clip = A_STR$
    Spd_ColSelect.FontBold = A_Bold%
    '
    Spd_ColSelect.Row = NewRow
    Spd_ColSelect.Col = newcol
    Spd_ColSelect.Action = SS_ACTION_ACTIVE_CELL
    Spd_ColSelect_GotFocus
End Sub

Private Sub Spd_ColSelect_GotFocus()
    SpreadGotFocus Spd_ColSelect.ActiveCol, Spd_ColSelect.ActiveRow
End Sub


Private Sub Spd_ColSelect_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal newcol As Long, ByVal NewRow As Long, Cancel As Boolean)
    SpreadLostFocus Col, Row
    If NewRow > 0 Then SpreadGotFocus newcol, NewRow
End Sub


Private Sub Spd_ColSort_Click(ByVal Col As Long, ByVal Row As Long)
    Sts_MsgLine.Panels(1) = Me.Caption
End Sub

Private Sub Spd_ColSort_DragDropBlock(ByVal Col As Long, ByVal Row As Long, ByVal Col2 As Long, ByVal Row2 As Long, ByVal newcol As Long, ByVal NewRow As Long, ByVal NewCol2 As Long, ByVal NewRow2 As Long, ByVal Overwrite As Boolean, Action As Integer, DataOnly As Boolean, Cancel As Boolean)
Dim A_STR$

    Cancel = True
    
    Spd_ColSort.Row = Row
    Spd_ColSort.Col = 1
    Spd_ColSort.Row2 = Row
    Spd_ColSort.Col2 = Spd_ColSort.MaxCols
    A_STR$ = Spd_ColSort.Clip
    Spd_ColSort.Action = SS_ACTION_DELETE_ROW
    '
    Spd_ColSort.Row = NewRow
    Spd_ColSort.Action = SS_ACTION_INSERT_ROW
    '
    Spd_ColSort.Row = NewRow
    Spd_ColSort.Col = 1
    Spd_ColSort.Row2 = NewRow
    Spd_ColSort.Col2 = Spd_ColSort.MaxCols
    Spd_ColSort.Clip = A_STR$
    '
    Spd_ColSort.Row = NewRow
    Spd_ColSort.Col = newcol
    Spd_ColSort.Action = SS_ACTION_ACTIVE_CELL
    Spd_ColSort_GotFocus
End Sub


Private Sub Spd_ColSort_GotFocus()
    SpreadGotFocus Spd_ColSort.ActiveCol, Spd_ColSort.ActiveRow
End Sub


Private Sub Spd_ColSort_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal newcol As Long, ByVal NewRow As Long, Cancel As Boolean)
    SpreadLostFocus Col, Row
    If NewRow > 0 Then SpreadGotFocus newcol, NewRow
End Sub


