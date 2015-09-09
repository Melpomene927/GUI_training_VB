VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_TSR03 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "使用記錄列印"
   ClientHeight    =   6420
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "TSR03.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6420
   ScaleWidth      =   9480
   Begin VsOcxLib.VideoSoftElastic Vse_Background 
      Height          =   6045
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   9480
      _Version        =   327680
      _ExtentX        =   16722
      _ExtentY        =   10663
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
      Picture         =   "TSR03.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "TSR03.frx":0326
      Begin FPSpread.vaSpread Spd_TSR03 
         Height          =   4665
         Left            =   60
         OleObjectBlob   =   "TSR03.frx":0342
         TabIndex        =   0
         Top             =   540
         Width           =   7860
      End
      Begin ComctlLib.ProgressBar Prb_Percent 
         Height          =   210
         Left            =   1290
         TabIndex        =   13
         Top             =   5250
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Frame Fra_PrintType 
         Caption         =   "列印方式"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   14
         Top             =   5220
         Width           =   7875
         Begin VB.TextBox Txt_FileName 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3330
            TabIndex        =   4
            Text            =   " "
            Top             =   270
            Width           =   4440
         End
         Begin Threed.SSOption Opt_File 
            Height          =   360
            Left            =   1260
            TabIndex        =   2
            Top             =   270
            Width           =   1068
            _Version        =   65536
            _ExtentX        =   1884
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "檔案"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption Opt_Printer 
            Height          =   360
            Left            =   156
            TabIndex        =   1
            Top             =   276
            Width           =   1068
            _Version        =   65536
            _ExtentX        =   1884
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "印表機"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption Opt_Excel 
            Height          =   360
            Left            =   2340
            TabIndex        =   3
            Top             =   270
            Width           =   945
            _Version        =   65536
            _ExtentX        =   1667
            _ExtentY        =   635
            _StockProps     =   78
            Caption         =   "Excel "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSCommand Cmd_Help 
         Height          =   405
         Left            =   8010
         TabIndex        =   5
         Top             =   90
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "輔助 F1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Cmd_Print 
         Height          =   405
         Left            =   8010
         TabIndex        =   6
         Top             =   540
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "列印F6"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Cmd_Next 
         Height          =   405
         Left            =   8010
         TabIndex        =   8
         Top             =   1440
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "次頁 F8"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Cmd_Previous 
         Height          =   405
         Left            =   8010
         TabIndex        =   7
         Top             =   990
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "前頁 F7"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Cmd_exit 
         Height          =   405
         Left            =   8010
         TabIndex        =   10
         Top             =   5550
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "結束Esc"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand Cmd_Set 
         Height          =   405
         Left            =   8010
         TabIndex        =   9
         Top             =   1890
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "表格設定 F9"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSPanel Pnl_A1501 
         Height          =   390
         Left            =   1035
         TabIndex        =   15
         Top             =   90
         Width           =   465
         _Version        =   65536
         _ExtentX        =   820
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   15790320
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin Threed.SSPanel Pnl_A1501n 
         Height          =   390
         Left            =   1485
         TabIndex        =   16
         Top             =   90
         Width           =   1860
         _Version        =   65536
         _ExtentX        =   3281
         _ExtentY        =   688
         _StockProps     =   15
         BackColor       =   15790320
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
      End
      Begin VB.Label Lbl_A1501 
         Caption         =   "公司別"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   17
         Top             =   135
         Width           =   1635
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   6045
      Width           =   9480
      _ExtentX        =   16722
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
Attribute VB_Name = "frm_TSR03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'========================================================================
'   Coding Rule
'========================================================================
'在此處定義之所有變數, 一律以M開頭, 如M_AAA$, M_BBB#, M_CCC&
'且變數之形態, 一律在最後一碼區別, 範例如下:
' $: 文字
' #: 所有數字運算(金額或數量)
' &: 程式迴圈變數
' %: 給一些使用於是或否用途之變數 (TRUE / FALSE )
' 空白: 代表VARIENT, 動態變數
'========================================================================
'自定變數
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

'必要變數
Dim m_FieldError%    '此變數在判斷欄位是否有誤, 必須回到該欄位之動作
Dim m_ExitTrigger%   '此變數在判斷結束鍵是否被觸發, 將停止目前正在處理的作業



'========================================================================
' Procedure : BeforeUnloadForm (frm_TSR03)
' @ Author  : Mike_chang
' @ Date    : 2015/9/3
' Purpose   : 關閉本表單前,須處理的動作在此加入
' Details   :
'========================================================================
Sub BeforeUnloadForm()


'??? 取消Spread上的所有標識區塊
    Spd_TSR03.Action = SS_ACTION_DESELECT_BLOCK

'結束目前視窗,跳出其他處理程序
    m_ExitTrigger% = True

'??? Keep目前結束的表單名稱至變數中
    G_FormFrom$ = "TSR03"
    
'??? 隱藏V畫面,回到Q畫面
    DoEvents
    Me.Hide
    frm_TSR03q.Show
End Sub

Private Function CheckRoutine_FileName() As Boolean
    CheckRoutine_FileName = True
    
    If Opt_Printer.Value = True Then Exit Function
    
    '設定變數初始值
    m_FieldError% = -1
    
    '若選擇檔案列印,欄位若空白,則帶出 Default Value
    If Opt_File.Value Then
        SetDefaultFileName Txt_FileName, G_Print2File
    ElseIf Opt_Excel.Value Then
        SetDefaultFileName Txt_FileName, G_Print2Excel
    End If
    DoEvents
    
    '檢核路徑是否存在
    Dim a$
    a$ = Trim(Txt_FileName)
    If Not CheckDirectoryExist(a$) Then
       CheckRoutine_FileName = False
       Sts_MsgLine.Panels(1) = G_PathNotFound$
       m_FieldError% = Txt_FileName.TabIndex
       Txt_FileName.SetFocus
    End If
End Function

'========================================================================
' Procedure : IsAllFieldsCheck (frm_TSR03)
' @ Author  : Mike_chang
' @ Date    : 2015/9/3
' Purpose   :
' Details   :
'========================================================================
Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    If Not CheckRoutine_FileName() Then Exit Function
    IsAllFieldsCheck = True
End Function

'========================================================================
' Procedure : KeepFieldsValue (frm_TSR03)
' @ Author  : Mike_chang
' @ Date    : 2015/9/3
' Purpose   :
' Details   :
'========================================================================
Sub KeepFieldsValue()
    G_ReportDataFrom = G_FromScreen
    G_OutFile = Trim$(Txt_FileName)
    If Opt_Printer.Value Then G_PrintSelect = G_Print2Printer
    If Opt_File.Value Then G_PrintSelect = G_Print2File
    If Opt_Excel.Value Then G_PrintSelect = G_Print2Excel
End Sub


'========================================================================
' Procedure : Set_Property (frm_TSR03)
' @ Author  : Mike_chang
' @ Date    : 2015/9/3
' Purpose   :
' Details   :
'========================================================================
Private Sub Set_Property()

    '??? 設定本Form之標題,字形及色系
    Form_Property frm_TSR03, G_Form_TSR01$, G_Font_Name
    
    '========================================================================
    '???設定Form中所有TextBox,ComboBox,ListBox之字形及可輸人長度,
    '   可同時設定其所對應的Label控制項的屬性
    '
    '   參數一 : Control Name
    '   參數二 : 物件的最大長度,非TextBox請輸入0
    '   參數三 : 對應Label的Control Name,設定其相關屬性
    '   參數四 : 設定Label的Caption,若自資料庫抓不到Caption則以此設定為Label的Caption
    '   參數五 : 輸入欄位的格式,用於日期或數值輸入
    '   參數六 : 數值欄位的上限
    '   參數七 : 數值欄位的下限
    '   參數八 : Database Name,於此資料庫下找尋Label的Caption
    '   參數九 : Table Name,於表格下找尋Label的Caption
    '   參數十 : Field Name,以此欄位找尋Label的Caption
    '========================================================================
    Field_Property Txt_FileName, 60
    Txt_FileName.Visible = False
        
    '========================================================================
    '??? 設定Form中所有Panel,Label,OptionButton,CheckBox,Frame之標題, 字形及色系
    '    參數一 : Control Name              參數二 : 設定Control的Caption
    '    參數三 : 是否顯示                  參數四 : 設定背景顏色
    '    參數五 : 設定字型大小              參數六 : 設定字型名稱
    '========================================================================
    Control_Property Fra_PrintType, G_Pnl_PrtType$
    Control_Property Opt_Printer, G_Pnl_Printer$
    Control_Property Opt_File, G_Pnl_File$
    Control_Property Opt_Excel, G_Pnl_Excel$
    
    '========================================================================
    '   設Form中所有Command之標題及字形
    '========================================================================
    Command_Property Cmd_Help, G_CmdHelp, G_Font_Name
    Command_Property Cmd_Print, G_CmdPrint, G_Font_Name
    Command_Property cmd_exit, G_CmdExit, G_Font_Name
    Command_Property Cmd_Previous, G_CmdPrvPage, G_Font_Name
    Command_Property Cmd_Next, G_CmdNxtPage, G_Font_Name
    Command_Property Cmd_Set, G_CmdSet, G_Font_Name
    
    '========================================================================
    '   設Form中Spread之屬性
    '========================================================================
    Set_Spread_Property

    '========================================================================
    '   以下為標準指令, 不得修改
    '========================================================================
    ProgressBar_Property Prb_Percent
    VSElastic_Property Vse_background
    StatusBar_ProPerty Sts_MsgLine
End Sub

'========================================================================
' Procedure : Set_Spread_Property (frm_TSR03)
' @ Author  : Mike_chang
' @ Date    : 2015/9/3
' Purpose   :
' Details   :
'========================================================================
Private Sub Set_Spread_Property()
    With Spd_TSR03
         .UnitType = 2

        '??? 設定本Spread之筆數及欄位數(取Columns Type的上限值)
         Spread_Property Spd_TSR03, 0, UBound(tSpd_TSR03.Columns), WHITE, _
             G_Font_Size, G_Font_Name
         
        '========================================================================
        '??? 設定本Spread之各欄標題及顯示寬度,各欄屬性及顯示字數
        '    參數一 : Spread Name
        '    參數二 : 參數一所屬的Spead Type Name
        '    參數三 : 自訂的欄位名稱
        '    參數四 : 設定欄寬
        '    參數五 : 預設的欄位標題
        '    參數六 : 欄位的資料型態
        '    參數七 : 數值欄位的下限
        '    參數八 : 數值欄位的上限
        '    參數九 : 文字資料型態的最大長度
        '    參數十 : 欄位顯示在Spread上的對齊方式
        '    參數11 : 設定報表欄位標題及資料列印的Format
        '    參數12 : 報表輸出至Excel時,是否將日期欄位格式化成日期格式
        '    參數13 : Database Name,於此資料庫下找尋Label的Caption
        '    參數14 : Field Name,以此欄位找尋Label的Caption
        '    參數15 : Table Name,於表格下找尋Label的Caption
        '========================================================================
         SpdFldProperty Spd_TSR03, tSpd_TSR03, "A1507", TextWidth("X") * 10, _
             G_Pnl_A1507, SS_CELL_TYPE_EDIT, "", "", 20, SS_CELL_H_ALIGN_LEFT, _
             SS_CELL_H_ALIGN_LEFT
         SpdFldProperty Spd_TSR03, tSpd_TSR03, "A1502", TextWidth("X") * 6, _
             G_Pnl_A1502, SS_CELL_TYPE_EDIT, "", "", 6, SS_CELL_H_ALIGN_CENTER
         SpdFldProperty Spd_TSR03, tSpd_TSR03, "A1505", TextWidth("X") * 15, _
             G_Pnl_A1505, SS_CELL_TYPE_EDIT, "", "", 40, SS_CELL_H_ALIGN_LEFT
         SpdFldProperty Spd_TSR03, tSpd_TSR03, "A1504", TextWidth("X") * 8, _
             G_Pnl_A1504, SS_CELL_TYPE_EDIT, "", "", 8
         SpdFldProperty Spd_TSR03, tSpd_TSR03, "A1510", TextWidth("X") * 8, _
             G_Pnl_A1510, SS_CELL_TYPE_EDIT, "", "", 8
         SpdFldProperty Spd_TSR03, tSpd_TSR03, "A1512", TextWidth("X") * 8, _
             G_Pnl_A1512, SS_CELL_TYPE_EDIT, "", "", 8
         SpdFldProperty Spd_TSR03, tSpd_TSR03, "A1508", TextWidth("X") * 15, _
             G_Pnl_A1508, SS_CELL_TYPE_FLOAT, "-999999999.99", "999999999.99", 15, _
             SS_CELL_H_ALIGN_RIGHT, SS_CELL_H_ALIGN_RIGHT
         SpdFldProperty Spd_TSR03, tSpd_TSR03, "Flag", TextWidth("X") * 20, _
             "Flag", SS_CELL_TYPE_EDIT, "", "", 20

        '設定本Spread允許Cell間的拖曳
         .AllowDragDrop = False

        '設定本Spread允許資料跨欄顯示
         .AllowCellOverflow = True
         
         .EditEnterAction = SS_CELL_EDITMODE_EXIT_NONE

        '固定向右捲動時, 所凍住之欄位
         .ColsFrozen = 2

        '鎖住Spread不可修改
         .Row = -1: .Col = -1: .Lock = True
    End With
End Sub

Private Sub Cmd_Exit_Click()
'離開V Screen前的處理動作,標準寫法,不可修改
    BeforeUnloadForm
End Sub

Private Sub Cmd_Help_Click()
Dim a$

'請將PATTERNR改為此Form名字即可, 其餘為標準指令, 不得修改
    a$ = "notepad " + G_Help_Path + "TSR03.HLP"
    retcode = Shell(a$, 4)
End Sub

Private Sub Cmd_Next_Click()
    Cmd_Next.Enabled = False
    Spd_TSR03.SetFocus
    SendKeys "{PgDn}"
    DoEvents
    Cmd_Next.Enabled = True
End Sub

Private Sub Cmd_Print_Click()
    Me.MousePointer = HOURGLASS
    Cmd_Print.Enabled = False

'檢核欄位正確性
    If IsAllFieldsCheck() = False Then
       Me.MousePointer = Default
       Cmd_Print.Enabled = True
       Exit Sub
    End If

'Keep共用變數供印表用
    KeepFieldsValue
    
'處理列印動作
    Sts_MsgLine.Panels(1) = G_Process

'控制RepSet Form結束後,不會觸發Form_Activate
    If G_PrintSelect = G_Print2Printer Then
       G_FormFrom$ = "RptSet"
    End If
       
'??? 開始列印報表,第三個參數傳入V Screen的Spread
    PrePare_Data frm_TSR03, Prb_Percent, Spd_TSR03, m_ExitTrigger%
    
    Cmd_Print.Enabled = True
    Me.MousePointer = Default
End Sub

Private Sub Cmd_Previous_Click()
    Cmd_Previous.Enabled = False
    Spd_TSR03.SetFocus
    SendKeys "{PgUp}"
    DoEvents
    Cmd_Previous.Enabled = True
End Sub

Private Sub Cmd_Set_Click()
'??? Load表格設定的表單
'    參數一 : 表格設定的Form Name
'    參數二 : 請輸入欲提供User設定的Spread的Spread Type Name
'    參數三 : 是否處理Spread排序欄位異動的更新
    ShowRptDefForm frm_RptDef, tSpd_TSR03
    
'??? 自表格設定表單返回時,處理Spread上的資料重整
'    參數一 : 資料欲重整的Spread Name
'    參數二 : 請輸入參數一的Spread Type Name
    RefreshSpreadData frm_TSR03.Spd_TSR03, tSpd_TSR03
    
'??? 結束表格設定視窗,將Focus設定在Spread上
    Spd_TSR03.SetFocus
End Sub

Private Sub Form_Activate()
    Me.MousePointer = HOURGLASS
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    
'Initial Form中的必要變數
    m_FieldError% = -1
    m_ExitTrigger% = False
         
'判斷是否由其他輔助畫面回來, 而非首次執行
    If Trim(G_FormFrom$) <> "" Then
       Me.MousePointer = Default
       G_FormFrom$ = ""
       '.....                '加入所要設定之動作
       '.....
       Exit Sub
    Else
       '.....                '第一次執行時之準備動作
       '.....
'設定Spread屬性
       Sts_MsgLine.Panels(1) = G_Process
       Set_Spread_Property
       Cmd_Print.Enabled = False
       PrePare_Data frm_TSR03, Prb_Percent, Spd_TSR03, m_ExitTrigger%
       If m_ExitTrigger% Then Exit Sub
       Cmd_Print.Enabled = True
    End If
    
    '將Form放置到螢幕的頂層
    frm_TSR03.ZOrder 0
    If frm_TSR03.Visible Then Spd_TSR03.SetFocus
    Me.MousePointer = Default
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case KEY_F1
           KeyCode = 0
           If Cmd_Help.Visible And Cmd_Help.Enabled Then
              Cmd_Help.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
        
      Case KEY_F6
           KeyCode = 0
           If Cmd_Print.Visible And Cmd_Print.Enabled Then
              Cmd_Print.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
           
      Case KEY_F7
           KeyCode = 0
           If Cmd_Previous.Visible And Cmd_Previous.Enabled Then
              Cmd_Previous.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
           
      Case KEY_F8
           KeyCode = 0
           If Cmd_Next.Visible And Cmd_Next.Enabled Then
              Cmd_Next.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
                
      Case KEY_F9
           KeyCode = 0
           If Cmd_Set.Visible And Cmd_Set.Enabled Then
              Cmd_Set.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
        
      Case KEY_ESCAPE
           KeyCode = 0
           If cmd_exit.Visible And cmd_exit.Enabled Then
              cmd_exit.SetFocus
              DoEvents
              SendKeys "{Enter}"
           End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    KeyPress KeyAscii
End Sub

Private Sub Form_Load()
    FormCenter Me
    Set_Property
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'若非由Q畫面結束V畫面,則不結束此畫面.標準寫法不可修改.
    If UnloadMode <> vbFormCode Then
       Cancel = True
       BeforeUnloadForm
    End If
End Sub

Private Sub Opt_Excel_Click(Value As Integer)
    SetDefaultFileName Txt_FileName, G_Print2Excel
End Sub

Private Sub Opt_File_Click(Value As Integer)
    SetDefaultFileName Txt_FileName, G_Print2File
End Sub

Private Sub Opt_Printer_Click(Value As Integer)
    Txt_FileName.Visible = False
End Sub

Private Sub Spd_TSR03_Click(ByVal Col As Long, ByVal Row As Long)
'??? 由此控制報表是否提供排序功能
    If Not tSpd_TSR03.SortEnable Then Exit Sub
    
'於Column Heading Click時, 依該欄位排序
    If Row = 0 And Col > 0 Then
    
'??? Update Spread Type中的排序欄位
       SpdSortIndexReBuild tSpd_TSR03, Col
       
'??? 利用Spread Type做Sort
       SpreadColsSort Spd_TSR03, tSpd_TSR03
       
    End If
End Sub

Private Sub Spd_TSR03_DragDropBlock(ByVal Col As Long, ByVal Row As Long, ByVal Col2 As Long, ByVal Row2 As Long, ByVal newcol As Long, ByVal NewRow As Long, ByVal NewCol2 As Long, ByVal NewRow2 As Long, ByVal Overwrite As Boolean, Action As Integer, DataOnly As Boolean, Cancel As Boolean)
'??? 將Spread上的原欄位移動至目的欄位
    SpreadColumnMove Spd_TSR03, tSpd_TSR03, Col, newcol, NewRow, Cancel
    
'在同一欄位DragDrop不處理變色
    If Col = newcol Then Exit Sub
    
'清除原欄位的顏色
    SpreadLostFocus2 Spd_TSR03, -1, Row, , , ConnectSemiColon(CStr(COLOR_YELLOW))
    
'設定新欄位的顏色
    SpreadGotFocus -1, NewRow, , , ConnectSemiColon(CStr(COLOR_YELLOW))
End Sub

Private Sub Spd_TSR03_GotFocus()
    SpreadGotFocus -1, CLng(Spd_TSR03.ActiveRow), , , ConnectSemiColon(CStr(COLOR_YELLOW))
End Sub

Private Sub Spd_TSR03_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal newcol As Long, ByVal NewRow As Long, Cancel As Boolean)
'恢復前一欄位的顏色
    SpreadLostFocus2 Spd_TSR03, -1, Row, , , ConnectSemiColon(CStr(COLOR_YELLOW))

'改變新欄位的顏色
    If NewRow > 0 Then SpreadGotFocus -1, NewRow, , , ConnectSemiColon(CStr(COLOR_YELLOW))
End Sub

Private Sub Txt_FileName_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_FileName_LostFocus()
    TextLostFocus
    
'判斷以下狀況發生時, 不須做任何處理
    If TypeOf ActiveControl Is SSCommand Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_FileName.TabIndex Then Exit Sub

    ' ....

'自我檢查
    retcode = CheckRoutine_FileName()
End Sub

