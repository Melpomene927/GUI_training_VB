VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_EXAM01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFFF&
   Caption         =   "會計科目代碼資料管理"
   ClientHeight    =   4260
   ClientLeft      =   5520
   ClientTop       =   2880
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "EXAM01.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4260
   ScaleWidth      =   10350
   Begin VsOcxLib.VideoSoftElastic Vse_background 
      Height          =   3885
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   10350
      _Version        =   327680
      _ExtentX        =   18256
      _ExtentY        =   6853
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Align           =   5
      BevelOuter      =   6
      Picture         =   "EXAM01.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "EXAM01.frx":0326
      Begin FPSpread.vaSpread Spd_EXAM01 
         Height          =   2115
         Left            =   1080
         OleObjectBlob   =   "EXAM01.frx":0342
         TabIndex        =   22
         Top             =   1710
         Width           =   7620
      End
      Begin VB.TextBox Txt_A1628 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7200
         MaxLength       =   40
         TabIndex        =   20
         Top             =   675
         Width           =   1455
      End
      Begin VB.TextBox Txt_A1606 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4095
         MaxLength       =   40
         TabIndex        =   18
         Top             =   675
         Width           =   1680
      End
      Begin VB.TextBox Txt_A1609 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6705
         MaxLength       =   12
         TabIndex        =   2
         Top             =   180
         Width           =   1950
      End
      Begin VB.TextBox Txt_A1601 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1305
         MaxLength       =   6
         TabIndex        =   0
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox Txt_A1612 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1305
         MaxLength       =   40
         TabIndex        =   4
         Text            =   " "
         Top             =   1170
         Width           =   7350
      End
      Begin VB.TextBox Txt_A1605 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1305
         MaxLength       =   40
         TabIndex        =   3
         Top             =   675
         Width           =   1680
      End
      Begin VB.TextBox Txt_A1602 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3390
         MaxLength       =   12
         TabIndex        =   1
         Top             =   180
         Width           =   1680
      End
      Begin Threed.SSCommand cmd_ok 
         Height          =   405
         Left            =   8940
         TabIndex        =   8
         Top             =   1050
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "確認F11"
         ForeColor       =   0
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
      Begin Threed.SSCommand cmd_exit 
         Height          =   405
         Left            =   8940
         TabIndex        =   9
         Top             =   3375
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "結束Esc"
         ForeColor       =   0
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
      Begin Threed.SSCommand cmd_help 
         Height          =   405
         Left            =   8940
         TabIndex        =   5
         Top             =   135
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "說明F1"
         ForeColor       =   0
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
      Begin Threed.SSCommand Cmd_Next 
         Height          =   405
         Left            =   8940
         TabIndex        =   7
         Top             =   1950
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "次筆F8"
         ForeColor       =   0
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
      Begin Threed.SSCommand Cmd_Previous 
         Height          =   405
         Left            =   8940
         TabIndex        =   6
         Top             =   1500
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "前筆F7"
         ForeColor       =   0
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
      Begin Threed.SSCommand Cmd_Delete 
         Height          =   405
         Left            =   8940
         TabIndex        =   23
         Top             =   585
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "刪除列F3"
         ForeColor       =   0
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
      Begin VB.Label Lbl_A1628 
         Caption         =   "生日/成立日"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5940
         TabIndex        =   21
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Lbl_A1606 
         Caption         =   "傳真電話"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3060
         TabIndex        =   19
         Top             =   750
         Width           =   1470
      End
      Begin VB.Label Lbl_A1609 
         Caption         =   "身分證/統編"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         TabIndex        =   12
         Top             =   255
         Width           =   1470
      End
      Begin VB.Label Lbl_A0206 
         Caption         =   "關係人"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   15
         Top             =   1710
         Width           =   1470
      End
      Begin VB.Label Lbl_A1602 
         Caption         =   "客戶簡稱"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2445
         TabIndex        =   11
         Top             =   270
         Width           =   1470
      End
      Begin VB.Label Lbl_A1612 
         Caption         =   "聯絡地址"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   14
         Top             =   1245
         Width           =   1470
      End
      Begin VB.Label Lbl_A1605 
         Caption         =   "連絡電話"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   13
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Lbl_A1601 
         Caption         =   "客戶編號"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   10
         Top             =   285
         Width           =   1470
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   3885
      Width           =   10350
      _ExtentX        =   18256
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
Attribute VB_Name = "frm_EXAM01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'========================================================================
'   Coding Rules
'========================================================================
'在此處定義之所有變數, 一律以M開頭, 如M_AAA$, M_BBB#, M_CCC&
'且變數之形態, 一律在最後一碼區別, 範例如下:
'   $: 文字
'   #: 所有數字運算(金額或數量)
'   &: 程式迴圈變數
'   %: 給一些使用於是或否用途之變數 (TRUE / FALSE )
'   空白: 代表VARIENT, 動態變數
'========================================================================

'自定變數
'Dim m_A1501Flag%
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

'必要變數
Dim m_FieldError%    '此變數在判斷欄位是否有誤, 必須回到該欄位之動作
Dim m_ExitTrigger%   '此變數在判斷結束鍵是否被觸發, 將停止目前正在處理的作業
'Dim m_RecordChange% '此變數在判斷資料是否有異動, 結束將提示是否存檔訊息
Dim m_TabGotFocus%   '控制Tab_ClickAfter 只處理一次
Dim m_TabMouseDown%  '控制由Help Control DblClick所觸發的Tab_ClickAfter不處理
'========================================================================
'====================================
'   User Defined Functions
'====================================

'========================================================================
' Procedure : CheckRoutine_A1601 (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Check Pkey valid
' Details   : Check: 1.Pkey not empty  2.Pkey not duplicate
'========================================================================
Private Function CheckRoutine_A1601() As Boolean
    CheckRoutine_A1601 = True
    m_FieldError% = -1
    
'檢核該欄位是否輸入
    If Txt_A1601.text = "" Then
        sts_msgline.Panels(1) = G_Pnl_A1601$ & G_MustInput
        CheckRoutine_A1601 = False
        
        G_DataChange% = True
        
        m_FieldError% = Txt_A1601.TabIndex
        If Txt_A1601.Enabled Then Txt_A1601.SetFocus
        Exit Function
    End If

'檢核資料是否已存在
    If G_AP_STATE = G_AP_STATE_ADD Then
        If IsKeyExist(Txt_A1601) Then
             sts_msgline.Panels(1) = G_Pnl_A1601$ & G_RecordExist
             CheckRoutine_A1601 = False
             G_DataChange% = True
             m_FieldError% = Txt_A1601.TabIndex
             Txt_A1601.SetFocus
        End If
    End If
End Function

'========================================================================
' Procedure : CheckRoutine_A1602 (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Check MustInput Constraint
' Details   :
'========================================================================
Private Function CheckRoutine_A1602() As Boolean
    CheckRoutine_A1602 = False

'設定變數初始值
    m_FieldError% = -1
    
'增加想要做的檢查
    If Txt_A1602.text = "" Then
       sts_msgline.Panels(1) = G_Pnl_A1602$ & G_MustInput
       m_FieldError% = Txt_A1602.TabIndex
       Exit Function
    End If
       
    CheckRoutine_A1602 = True
End Function

'========================================================================
' Procedure : CheckRoutine_A1628 (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Check Date Formate Valid
' Details   :
'========================================================================
Private Function CheckRoutine_A1628() As Boolean
    CheckRoutine_A1628 = False
    
'設定變數初始值
    m_FieldError% = -1
    
'增加想要做的檢查

    'Check Date Formate Valid
    If Trim(Txt_A1628) <> "" Then
        If Not IsDateValidate(Txt_A1628) Then
            sts_msgline.Panels(1) = G_Pnl_A1628$ & G_DateError
            m_FieldError% = Txt_A1628.TabIndex
            Txt_A1628.SetFocus
        End If
    End If

    CheckRoutine_A1628 = True
End Function

'========================================================================
' Procedure : ClearFieldsValue (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Clear Fields
' Details   :
'========================================================================
Private Sub ClearFieldsValue()
    'Clear Txtbox
    Txt_A1601.text = ""
    Txt_A1601.Tag = ""
    Txt_A1602.text = ""
    Txt_A1609.text = ""
    Txt_A1605.text = ""
    Txt_A1606.text = ""
    Txt_A1628.text = ""
    Txt_A1612.text = ""
    
    'Clear Spread
    Spd_EXAM01.MaxRows = 0
    Spd_EXAM01.MaxRows = 1
End Sub

'========================================================================
' Procedure : Delete_From_Menu (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Delete_From_Menu()
'將V畫面上的該筆資料列刪除
    With frm_EXAM01v.Spd_EXAM01v
        .Row = G_ActiveRow#
        .Action = SS_ACTION_DELETE_ROW
        .MaxRows = .MaxRows - 1
    End With
End Sub

'========================================================================
' Procedure : Delete_Process_A16 (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Do Deletion by Pkey: A1601
' Details   :
'========================================================================
Private Sub Delete_Process_A16()
On Local Error GoTo MY_Error

    G_Str = "DELETE FROM A16"
    G_Str = G_Str & " WHERE A1601='" & G_A1601$ & "'"
    ExecuteProcess DB_ARTHGUI, G_Str
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Procedure : Delete_Process_A19 (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub Delete_Process_A19()
On Local Error GoTo MY_Error

    G_Str = "DELETE FROM A19"
    G_Str = G_Str & " WHERE A1901='" & G_A1601$ & "'"
    If Trim(G_A1902$) <> "" Then
        G_Str = G_Str & " And A1902 = '" & G_A1902$ & "'"
    End If
    ExecuteProcess DB_ARTHGUI, G_Str
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : IsAllFieldsCheck
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Boolean Function checking whether all txtBox pass value check
' Details   :
'========================================================================
Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    
'執行存檔前須將所有檢核欄位再做一次
    If G_AP_STATE = G_AP_STATE_ADD Then
        If Not CheckRoutine_A1601() Then
            Txt_A1601.SetFocus
            Exit Function
        End If
    End If
    If Not CheckRoutine_A1602() Then
        Txt_A1602.SetFocus
        Exit Function
    End If
    If Not CheckRoutine_A1628() Then
        Txt_A1628.SetFocus
        Exit Function
    End If
    
    IsAllFieldsCheck = True
End Function

'========================================================================
' Procedure : IsKeyChanged (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Function IsKeyChanged(ByVal A_A1902$, ByVal A_A1902o$) As Boolean

   IsKeyChanged = False
   If UCase$(A_A1902$) <> UCase$(A_A1902o$) Then
      IsKeyChanged = True
   End If
   
End Function

'========================================================================
' Procedure : IsKeyExist ()
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Check Primary Key Existed in DB
' Details   : N/A
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' @ Modified: change to A16 DB
'========================================================================
Private Function IsKeyExist(ByVal A_A1601$) As Boolean
On Local Error GoTo MY_Error
Dim DY_A16_TMP As Recordset
Dim A_Sql$

    IsKeyExist = False
    
    A_Sql$ = "Select A1601 From A16"
    A_Sql$ = A_Sql$ & " where A1601='" & Trim(A_A1601$) & "'"
    A_Sql$ = A_Sql$ & " Order by A1601"
    
    CreateDynasetODBC DB_ARTHGUI, DY_A16_TMP, A_Sql$, "DY_A16_TMP", True
    If Not (DY_A16_TMP.BOF And DY_A16_TMP.EOF) Then IsKeyExist = True
    
    DY_A16_TMP.Close
    Exit Function
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

'========================================================================
' Module    : frm_EXAM01
' Procedure : IsRecordChange
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Detect if data changed by Global Variable "G_DataChange"
' Details   : set FALSE while Deleting
'========================================================================
Function IsRecordChange() As Boolean
'若作業狀態為刪除則不做Check
    If G_AP_STATE = G_AP_STATE_DELETE Then
       IsRecordChange = False
       Exit Function
    End If
       
'判斷Record資料是否異動
    IsRecordChange = G_DataChange%
End Function

'========================================================================
' Procedure : Move2Menu (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Move fields to V-form after Insert, Update
' Details   : 將異動資料UPDATE回V畫面的SPREAD上
'========================================================================
Private Sub Move2Menu()
    With frm_EXAM01v.Spd_EXAM01v
         If G_AP_STATE = G_AP_STATE_UPDATE Then
            .Row = G_ActiveRow#         'Focus on the row to update
         Else
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .Action = SS_ACTION_ACTIVE_CELL
         End If
         .Col = 1
         .text = Trim$(Txt_A1601 & "")
         .Col = 2
         .text = Trim$(Txt_A1602 & "")
         .Col = 3
         .text = Trim$(Txt_A1609 & "")
         .Col = 4
         .text = DateFormat2(Trim$(Txt_A1628 & ""))
         .Col = 5
         .text = Trim$(Txt_A1605 & "")
         .Col = 6
         .text = Trim$(Txt_A1612 & "")
    End With
End Sub

'========================================================================
' Procedure : MoveDB2Field (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Fetch Data from the DY_A16, DY_A19
' Details   :
'========================================================================
Private Sub MoveDB2Field()
On Local Error GoTo MY_Error

    'textbox
    Txt_A1601.text = Trim$(DY_A16.Fields("A1601") & "")
    Txt_A1602.text = Trim$(DY_A16.Fields("A1602") & "")
    Txt_A1609.text = Trim$(DY_A16.Fields("A1609") & "")
    Txt_A1605.text = Trim$(DY_A16.Fields("A1605") & "")
    Txt_A1612.text = Trim$(DY_A16.Fields("A16121") & "")
    Txt_A1612.text = Txt_A1612.text & Trim$(DY_A16.Fields("A16122") & "")
    Txt_A1612.text = Txt_A1612.text & Trim$(DY_A16.Fields("A16123") & "")
    Txt_A1606.text = Trim$(DY_A16.Fields("A1606") & "")
    Txt_A1628.text = Trim$(DY_A16.Fields("A1628") & "")
   
   
    Exit Sub
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Procedure : MoveDB2Spread (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Function MoveDB2Spread()
On Local Error GoTo MY_Error
    
    'spread
    With Spd_EXAM01
        .MaxRows = 0
        Do Until DY_A19.EOF
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .Col = 1
            .text = Trim(DY_A19.Fields("A1902") & "")
            .Col = 2
            .text = Trim(DY_A19.Fields("A1602") & "")
            .Col = 3
            .text = Trim(DY_A19.Fields("A1903") & "")
            .Col = 4
            .text = Trim(DY_A19.Fields("A1902") & "")
        Loop
    End With
    
    Exit Function
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

'========================================================================
' Procedure : MoveField2DB (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub MoveField2DB()
On Local Error GoTo MY_Error
Dim A_A16121$, A_A16122$, A_A16123$
Dim I&

    'split client address to 3 fields
    If Len(Txt_A1612) > 40 Then
        A_A16123$ = Mid(Txt_A1612, 41, Len(Txt_A1612))
        A_A16122$ = Mid(Txt_A1612, 21, 40)
        A_A16121$ = Mid(Txt_A1612, 1, 20)
    ElseIf Len(Txt_A1612) > 20 And Len(Txt_A1612) <= 40 Then
        A_A16123$ = ""
        A_A16122$ = Mid(Txt_A1612, 21, Len(Txt_A1612))
        A_A16121$ = Mid(Txt_A1612, 1, 20)
    Else
        A_A16123$ = ""
        A_A16122$ = ""
        A_A16121$ = Mid(Txt_A1612, 1, Len(Txt_A1612))
    End If
    
    
    G_Str = ""
    If G_AP_STATE = G_AP_STATE_ADD Then
        'Add A16
        InsertFields "A16001", GetCurrentDate(), G_Data_String   'G_Data_Numeric
        InsertFields "A16002", GetCurrentTime(), G_Data_String
        InsertFields "A16003", GetWorkStation(), G_Data_String
        InsertFields "A16004", GetUserId(), G_Data_String
        InsertFields "A16005", " ", G_Data_String
        InsertFields "A16006", " ", G_Data_String
        InsertFields "A16007", " ", G_Data_String
        InsertFields "A16008", " ", G_Data_String
        
        InsertFields "A1601", Trim(Txt_A1601.text & ""), G_Data_String
        InsertFields "A1602", Trim(Txt_A1602.text & ""), G_Data_String
        InsertFields "A1605", Trim(Txt_A1605.text & ""), G_Data_String
        InsertFields "A1606", Trim(Txt_A1606.text & ""), G_Data_String
        InsertFields "A1628", Trim(DateIn(Txt_A1628.text & "")), G_Data_String
        InsertFields "A1609", Trim(Txt_A1609.text & ""), G_Data_String
        InsertFields "A1613", "1", G_Data_String
        
        InsertFields "A16121", A_A16121$, G_Data_String
        InsertFields "A16122", A_A16122$, G_Data_String
        InsertFields "A16123", A_A16123$, G_Data_String
       
        SQLInsert DB_ARTHGUI, "A16"
    Else
        UpdateString "A16005", GetCurrentDate(), G_Data_String
        UpdateString "A16006", GetCurrentTime(), G_Data_String
        UpdateString "A16007", GetWorkStation(), G_Data_String
        UpdateString "A16008", GetUserId(), G_Data_String
        
        UpdateString "A1601", Trim(Txt_A1601.text & ""), G_Data_String
        UpdateString "A1602", Trim(Txt_A1602.text & ""), G_Data_String
        UpdateString "A1605", Trim(Txt_A1605.text & ""), G_Data_String
        UpdateString "A1606", Trim(Txt_A1606.text & ""), G_Data_String
        UpdateString "A1628", Trim(DateIn(Txt_A1628.text & "")), G_Data_String
        UpdateString "A1609", Trim(Txt_A1609.text & ""), G_Data_String
        
        UpdateString "A16121", A_A16121$, G_Data_String
        UpdateString "A16122", A_A16122$, G_Data_String
        UpdateString "A16123", A_A16123$, G_Data_String
        
        G_Str = G_Str & " where A1601='" & G_A1601$ & "'"
        
        SQLUpdate DB_ARTHGUI, "A16"
    End If
    
    'Add A19
    If Spd_EXAM01.MaxRows > 1 Then
        For I& = 1 To Spd_EXAM01.MaxRows
            MoveSpread2DB (I&)
        Next
    End If
    
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Procedure : MoveSpread2DB (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub MoveSpread2DB(ByVal Row As Long)
On Local Error GoTo MY_Error
Dim A_A1901$, A_A1902$, A_A1903$
Dim A_A1902o$, A_Action$
    
    Me.MousePointer = HOURGLASS
    
    A_A1901$ = G_A1601
    
    With Spd_EXAM01
        .Row = Row
        .Col = 1: A_A1902$ = Trim(.text)
        .Col = 3: A_A1903$ = Trim(.text)
        .Col = 4: A_A1902o$ = Trim(.text)
        .Col = 5: A_Action$ = Trim(.text)
        '
        G_Str = ""
        If UCase$(A_Action$) = UCase$("U") Then
           UpdateString "A19005", GetCurrentDate(), G_Data_String
           UpdateString "A19006", GetCurrentTime(), G_Data_String
           UpdateString "A19007", GetWorkStation(), G_Data_String
           UpdateString "A19008", GetUserId(), G_Data_String
           UpdateString "A1901", A_A1901$, G_Data_String
           UpdateString "A1902", A_A1902$, G_Data_String
           UpdateString "A1903", A_A1903$, G_Data_String
           G_Str = G_Str & " where A1901='" & Trim(A_A1901$) & "'"
           G_Str = G_Str & " and A1902='" & A_A1902o$ & "'"
           SQLUpdate DB_ARTHGUI, "A19"
        Else
           InsertFields "A19001", GetCurrentDate(), G_Data_String
           InsertFields "A19002", GetCurrentTime(), G_Data_String
           InsertFields "A19003", GetWorkStation(), G_Data_String
           InsertFields "A19004", GetUserId(), G_Data_String
           InsertFields "A19005", " ", G_Data_String
           InsertFields "A19006", " ", G_Data_String
           InsertFields "A19007", " ", G_Data_String
           InsertFields "A19008", " ", G_Data_String
           InsertFields "A1901", A_A1901$, G_Data_String
           InsertFields "A1902", A_A1902$, G_Data_String
           InsertFields "A1903", A_A1903$, G_Data_String
           SQLInsert DB_ARTHGUI, "A19"
        End If
        '
        .Col = 4: .text = A_A1902$
        .Col = 5: .text = ""
    End With
    
    Me.MousePointer = Default
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : OpenMainFile
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Open Dynaset "DY_A19"
' Details   : A19 is a relation of A16-A16, A1901 is ClientA relats A1902
'             which is ClientB
'========================================================================
Private Sub OpenMainFile()
On Local Error GoTo MY_Error
Dim A_Sql$

    'A16
    A_Sql$ = "Select * From A16"
    A_Sql$ = A_Sql$ & " Where A1601='" & G_A1601$ & "'"
    CreateDynasetODBC DB_ARTHGUI, DY_A16, A_Sql$, "DY_A16", True

    'A19
    A_Sql$ = "Select A1902,A1903,A1602 From A19"
    A_Sql$ = A_Sql$ & " INNER JOIN A16"
    A_Sql$ = A_Sql$ & " ON A19.A1902 = A16.A1601"
    A_Sql$ = A_Sql$ & " Where A19.A1901 = '" & G_A1601$ & "'"
    CreateDynasetODBC DB_ARTHGUI, DY_A19, A_Sql$, "DY_A19", True

'    'A15
'    A_Sql$ = "SELECT * FROM A15"
'    A_Sql$ = A_Sql$ & " where A1501='" & G_A1601$ & "'"
'    A_Sql$ = A_Sql$ & " and A1502='" & Mid$(G_A1502$, 1, 4) & "'"
'    A_Sql$ = A_Sql$ & " and A1503='" & Mid$(G_A1502$, 5) & "'"
'    A_Sql$ = A_Sql$ & " order by A1501,A1502,A1503"
'    CreateDynasetODBC DB_ARTHGUI, DY_A16, A_Sql$, "DY_A16", True
'    'A14
'    A_Sql$ = "SELECT * FROM A14"
'    A_Sql$ = A_Sql$ & " where A1406='" & G_A1601$ & "'"
'    A_Sql$ = A_Sql$ & " and A1402='" & Mid$(G_A1502$, 1, 4) & "'"
'    A_Sql$ = A_Sql$ & " and A1403='" & Mid$(G_A1502$, 5) & "'"
'    A_Sql$ = A_Sql$ & " order by A1401,A1406,A1402,A1403"
'    CreateDynasetODBC DB_ARTHGUI, DY_A14, A_Sql$, "DY_A14", True
'    'A20
'    A_Sql$ = "SELECT * FROM A20"
'    A_Sql$ = A_Sql$ & " where A2001='" & G_A1601$ & "'"
'    A_Sql$ = A_Sql$ & " and A2002='" & G_A1502$ & "'"
'    A_Sql$ = A_Sql$ & " order by A2001,A2002,A2003"
'    CreateDynasetODBC DB_ARTHGUI, DY_A20, A_Sql$, "DY_A20", True
    Exit Sub

MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Procedure : Reference_A16 (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Function Reference_A16(ByVal A_A1601$) As String
On Local Error GoTo MyError
Dim DY_Tmp As Recordset
Dim A_Sql$

    Reference_A16 = ""
    A_Sql$ = "Select A1602 From A16"
    A_Sql$ = A_Sql$ & " where A1601='" & A_A1601$ & "'"
    A_Sql$ = A_Sql$ & " order by A1601"
    CreateDynasetODBC DB_ARTHGUI, DY_Tmp, A_Sql$, "DY_TMP", True
    
    If Not (DY_Tmp.BOF And DY_Tmp.EOF) Then
       Reference_A16 = Trim$(DY_Tmp.Fields("A1602") & "")
    End If
    DY_Tmp.Close
    
    Exit Function
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

'========================================================================
' Procedure : SaveCheck (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Open Dialog(or not) and ask whether save to DB
' Details   : Using "MoveField2DB", "Move2Menu" to Save
'========================================================================
Function SaveCheck(Optional A_PassQuestion% = False) As Boolean
    SaveCheck = False
    
    '新增修改狀態下,欲結束此畫面,若畫面資料有異動時,詢問是否存檔.
    'retcode = IDYES     , 存檔後返回V畫面
    'retcode = IDNO      , 不存檔返回V畫面
    'retcode = IDCANCEL  , 不存檔停留在原畫面
    If A_PassQuestion% Then
    '按確認鍵要做存檔動作
        retcode = IDYES
    Else
    '按結束鍵時由User決定要處理的動作
        retcode = MsgBox(G_Save_Check, vbYesNoCancel + vbQuestion, Me.Caption)
    End If
    
    If retcode = IDCANCEL Then
       Exit Function
    ElseIf retcode = IDYES Then
       If Not IsAllFieldsCheck() Then Exit Function
       Me.Refresh
       MoveField2DB
       Move2Menu
    End If
    
    SaveCheck = True
End Function

'========================================================================
' Module    : frm_EXAM01
' Procedure : SetButtonEnable
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Set All Command Buttom to FALSE or restore previous
' Details   : If Set True, Store Current Enable state and set to false
'             Else restore from tag to previous state
'========================================================================
Sub SetButtonEnable(ByVal A_Enable%)
    If Not A_Enable% Then
       vse_background.TabStop = True
       cmd_previous.Tag = cmd_previous.Enabled
       cmd_next.Tag = cmd_next.Enabled
       cmd_ok.Tag = cmd_ok.Enabled
       cmd_Exit.Tag = cmd_Exit.Enabled
       
       cmd_previous.Enabled = A_Enable%
       cmd_next.Enabled = A_Enable%
       cmd_ok.Enabled = A_Enable%
       cmd_Exit.Enabled = A_Enable%
    Else
       cmd_previous.Enabled = CBool(cmd_previous.Tag)
       cmd_next.Enabled = CBool(cmd_next.Tag)
       cmd_ok.Enabled = CBool(cmd_ok.Tag)
       cmd_Exit.Enabled = CBool(cmd_Exit.Tag)
    End If
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : SetCommand
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   : Setup command buttom enables
' Details   : Pkey(A1601) is only allow to be insert while Adding record,
'             Otherwise, at updating & delete, it's no meaning to change
'             Pkey since it doesn't stand for a specific literal meaning
'------------------------------------------------------------------------
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' @ Modified: Mark out all Pre&Next page controls
'========================================================================
Sub SetCommand()
'設定每一作業狀態下, CONTROL是否可作用
    Select Case G_AP_STATE
        Case G_AP_STATE_ADD
            'while Adding, Pkey(A1601) is allowed to input
            cmd_Help.Enabled = True
            cmd_delete.Enabled = True
            cmd_ok.Enabled = True
            cmd_Exit.Enabled = True
            Txt_A1601.Enabled = True
'            Cmd_Previous.Enabled = False
'            Cmd_Next.Enabled = False

        Case G_AP_STATE_UPDATE
            'while update, no meaning to change Pkey
            cmd_Help.Enabled = True
            cmd_delete.Enabled = True
            cmd_ok.Enabled = True
            cmd_Exit.Enabled = True
            Txt_A1601.Enabled = False
'            Cmd_Previous.Enabled = (G_ActiveRow# > 1)
'            Cmd_Next.Enabled = (G_ActiveRow# < G_MaxRows#)

        Case G_AP_STATE_DELETE
            'while delete, no meaning to change Pkey
            cmd_Help.Enabled = True
            cmd_delete.Enabled = True
            cmd_ok.Enabled = True
            cmd_Exit.Enabled = True
            Txt_A1601.Enabled = False
'            Cmd_Previous.Enabled = (G_ActiveRow# > 1)
'            Cmd_Next.Enabled = (G_ActiveRow# < G_MaxRows#)
     End Select
End Sub

'========================================================================
' Procedure : Set_Property (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Setup Properties
' Details   : 1. Form(caption, font, color)
'             2. Label(caption, font, color)
'             3. TextBox(font, maxlength)
'             4. command buttom(caption, font)
'------------------------------------------------------------------------
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' @ Modified: Add Spread setting
'========================================================================
Private Sub Set_Property()
    Me.FontBold = False
    
    '設定本Form之標題,字形及色系
    Form_Property Me, G_Form_EXAM01, G_Font_Name

    '設Form中所有Label之標題, 字形及色系
    Label_Property Lbl_A1601, G_Pnl_A1601$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1602, G_Pnl_A1602$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1609, G_Pnl_A1609$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1605, G_Pnl_A1605$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1606, G_Pnl_A1606$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1612, G_Pnl_A1612$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1628, G_Pnl_A1628$, G_Label_Color, G_Font_Size, G_Font_Name
    
    
    '設Form中所有TextBox之字形及可輸入長度
    Text_Property Txt_A1601, 10, G_Font_Name
    Text_Property Txt_A1602, 12, G_Font_Name
    Text_Property Txt_A1609, 15, G_Font_Name
    Text_Property Txt_A1605, 15, G_Font_Name
    Text_Property Txt_A1612, 120, G_Font_Name
    Text_Property Txt_A1606, 15, G_Font_Name
    Text_Property Txt_A1628, 8, G_Font_Name
    

    '設Form中所有Command之標題及字形
    Command_Property cmd_Help, G_CmdHelp, G_Font_Name
    Command_Property cmd_previous, G_CmdPrevious, G_Font_Name
    Command_Property cmd_next, G_CmdNext, G_Font_Name
    Command_Property cmd_ok, G_CmdOk, G_Font_Name
    Command_Property cmd_Exit, G_CmdExit, G_Font_Name

    cmd_next.Enabled = False
    cmd_next.Visible = False
    cmd_previous.Enabled = False
    cmd_previous.Visible = False
    
    '設Form中Spread之屬性
    Set_Spread_Property
    
    '以下為標準指令, 不得修改
    VSElastic_Property vse_background
    StatusBar_ProPerty sts_msgline
End Sub

'========================================================================
' Procedure : Set_Spread_Property (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub Set_Spread_Property()
    Spd_EXAM01.UnitType = 2

    '設定本Spread之筆數及欄位數
    Spread_Property Spd_EXAM01, 0, 5, WHITE, G_Font_Size, G_Font_Name

    '設定本Spread之各欄標題及顯示寬度, 0代表該欄位不顯示
    Spread_Col_Property Spd_EXAM01, 1, TextWidth("A") * 10, G_Pnl_A1902$
    Spread_Col_Property Spd_EXAM01, 2, TextWidth("A") * 12, G_Pnl_A1902n$
    Spread_Col_Property Spd_EXAM01, 3, TextWidth("A") * 20, G_Pnl_A1903$
    Spread_Col_Property Spd_EXAM01, 4, TextWidth("A") * 0, "A1901o" 'p-key
    Spread_Col_Property Spd_EXAM01, 5, TextWidth("A") * 0, "Change/Add/No Change"

    '====================================
    '設定本Spread之各欄屬性及顯示字數
    '   SS_CELL_TYPE_EDIT        = 文字可輸入
    '   SS_CELL_TYPE_FLOAT       = 數字可輸入
    '   SS_CELL_TYPE_STATIC_TEXT = 純顯示
    '   SS_CELL_TYPE_CHECKBOX    = 點選項目
    '====================================
    Spread_DataType_Property Spd_EXAM01, 1, SS_CELL_TYPE_EDIT, "", "", 10
    Spread_DataType_Property Spd_EXAM01, 2, SS_CELL_TYPE_EDIT, "", "", 12
    Spread_DataType_Property Spd_EXAM01, 3, SS_CELL_TYPE_EDIT, "", "", 20
    Spread_DataType_Property Spd_EXAM01, 4, SS_CELL_TYPE_EDIT, "", "", 10
    Spread_DataType_Property Spd_EXAM01, 5, SS_CELL_TYPE_EDIT, "", "", 1
    Spd_EXAM01.EditEnterAction = SS_CELL_EDITMODE_EXIT_DOWN

    '固定向右捲動時, 所凍住之欄位
    Spd_EXAM01.ColsFrozen = 2

    '定義某些欄置中位置之設定 0:左靠  1:右靠  2:置中
    Spd_EXAM01.Row = -1
    Spd_EXAM01.Col = 1: Spd_EXAM01.TypeHAlign = 2

    '定義某些欄置被保護無法顯示
    Spd_EXAM01.Col = 4:  Spd_EXAM01.ColHidden = True
    Spd_EXAM01.Col = 5:  Spd_EXAM01.ColHidden = True
End Sub

'========================================================================
' Procedure : SpreadLineCheck (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Check Spread Row data while the Row is modified
' Details   :
'========================================================================
Function SpreadLineCheck(ByVal Row As Long, Col As Long) As Boolean
    With Spd_EXAM01
        .Row = Row
        SpreadLineCheck = False
        
        '必要欄位檢查, 並傳回Col
        If SpreadCheck_1(Row) = False Then
            Col = 1
            Exit Function
        End If
        If SpreadCheck_3(Row) = False Then
            Col = 3
            Exit Function
        End If
        
        '全部過關
        SpreadLineCheck = True
    End With
End Function

'========================================================================
' Procedure : SpreadCheck_1 (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Check Column 1 must input & not relat to itself
' Details   :
'========================================================================
Function SpreadCheck_1(ByVal Row As Long) As Boolean
Dim A_A1902$, A_A1902o$, A_Action$
    
    SpreadCheck_1 = False
    With Spd_EXAM01
        .Row = Row
        .Col = 1
        A_A1902$ = Trim(.text)          '取得Pkey
        .Col = 4
        A_A1902o$ = Trim(.text)         '取得修改前Pkey
        .Col = 5
        A_Action$ = Trim(.text)         '取得Action Code(新增或修改)

        'Check Must Input Constraint
        .Col = 1
        If Trim(.text) = "" Then
            sts_msgline.Panels(1) = G_Pnl_A1902$ & G_MustInput
            Exit Function
        End If
        
        '====================================
        '@Adding: According to A19 is
        '   a self-relation of A16, A1901
        '   must not the same with A1902
        '   which is G_A1601 instead here
        '====================================
        If A_A1902$ = G_A1601$ Then
            sts_msgline.Panels(1) = G_Pnl_A1902$ & G_FieldErr
            Exit Function
        End If

        'Check Primary Key must exist
        If A_Action$ = "A" Then
            If Not IsKeyExist(A_A1902$) = True Then
                sts_msgline.Panels(1) = G_Pnl_A1902$ & "資料不存在"
                Exit Function
            End If
        ElseIf A_Action$ = "U" Then
            'Check whether Updating the Pkey of the record
            If IsKeyChanged(A_A1902$, A_A1902o$) = True Then
                If Not IsKeyExist(A_A1902$) = True Then
                    sts_msgline.Panels(1) = G_Pnl_A1902$ & "資料不存在"
                    Exit Function
                End If
            End If
        End If

        SpreadCheck_1 = True
    End With
End Function

'========================================================================
' Procedure : SpreadCheck_3 (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Check Column 3 Must Input
' Details   :
'========================================================================
Function SpreadCheck_3(ByVal Row As Long) As Boolean

    SpreadCheck_3 = False
    With Spd_EXAM01
        'Check Must Input Constraint
        .Col = 3
        If Not Trim(.text) = "" Then
            sts_msgline.Panels(1) = G_Pnl_A1903$ & G_MustInput
            Exit Function
        End If
    End With
End Function

'====================================
'   Command Buttom Events
'====================================

'========================================================================
' Procedure : Cmd_Delete_Click (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub cmd_delete_Click()
    Delete_Process_A19
    sts_msgline.Panels(1) = G_Delete_Ok
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : Cmd_Help_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : call HLP file
' Details   :
'========================================================================
Private Sub Cmd_Help_Click()
Dim a$

    a$ = "notepad " + G_Help_Path + "EXAM01.HLP"
    retcode = Shell(a$, 4)
End Sub

'========================================================================
' Procedure : Cmd_Next_Click (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Move to next record of the V-form
' Details   :
'========================================================================
Private Sub Cmd_Next_Click()
''無下一筆資料不做處理
'    If G_ActiveRow# >= G_MaxRows# Then
'       Sts_MsgLine.Panels(1) = G_AP_NONEXT
'       Exit Sub
'    End If
'
'    Me.MousePointer = HOURGLASS
'
''設定會影響資料存檔的所有Button的Enabled Property = False
'    SetButtonEnable False
'
''若目前Record資料有異動, 提示是否存檔
'    If IsRecordChange() Then
'        If SaveCheck() = False Then
'            'If Dialog's cancel buttom click:
'            Me.MousePointer = Default
'            SetButtonEnable True
'            Txt_A1602.SetFocus
'            Exit Sub
'        End If
'    End If
'
''取得下一筆資料的P-KEY
'    With frm_EXAM01v!Spd_EXAM01v
'         G_ActiveRow# = G_ActiveRow# + 1
'         .Row = G_ActiveRow#
''         .Col = 1: StrCut Trim$(.text), Space(1), G_A1601$, ""
'         .Col = 1: G_A1601$ = Trim$(.text)
'
''將V畫面的游標移到下一筆
'         .Action = SS_ACTION_ACTIVE_CELL
'    End With
'
''帶出下一筆資料
'    OpenMainFile
'    ClearFieldsValue
'    MoveDB2Field
'    G_DataChange% = False
'
''還原所有Button的Enabled Property
'    SetButtonEnable True
'
'    SetCommand
'    Txt_A1602.SetFocus
'    Me.MousePointer = Default
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : Cmd_Previous_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Move to Previous record of V-Form
' Details   :
'========================================================================
Private Sub Cmd_Previous_Click()
''無上一筆資料不做處理
'    If G_ActiveRow# <= 1 Then
'       Sts_MsgLine.Panels(1) = G_AP_NOPRVS
'       Exit Sub
'    End If
'    Me.MousePointer = HOURGLASS
'
''設定會影響資料存檔的所有Button的Enabled Property = False
'    SetButtonEnable False
'
''若目前Record資料有異動, 提示是否存檔
'    If IsRecordChange() Then
'       If SaveCheck() = False Then
'          SetButtonEnable True
'          Me.MousePointer = Default
'          Txt_A1602.SetFocus
'          Exit Sub
'       End If
'    End If
'
''取得上一筆資料的P-KEY
'    With frm_EXAM01v!Spd_EXAM01v
'         G_ActiveRow# = G_ActiveRow# - 1
'         .Row = G_ActiveRow#
''         .Col = 1: StrCut Trim$(.text), Space(1), G_A1601$, ""
'         .Col = 1: G_A1601$ = Trim$(.text)
'
''將V畫面的游標移到下一筆
'         .Action = SS_ACTION_ACTIVE_CELL
'    End With
'
''帶出上一筆資料
'    OpenMainFile
'    ClearFieldsValue
'    MoveDB2Field
'    G_DataChange% = False
'
''還原所有Button的Enabled Property
'    SetButtonEnable True
'
'    SetCommand
'    Txt_A1602.SetFocus
'    Me.MousePointer = Default
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : Cmd_Ok_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Do Insert ,Update Or Delete
' Details   :
'========================================================================
Private Sub Cmd_Ok_Click()
    Me.MousePointer = HOURGLASS
    
    '設定會影響資料存檔的所有Button的Enabled Property = False
    SetButtonEnable False
    
    '依每個作業狀態做各別的處理
    Select Case G_AP_STATE
        Case G_AP_STATE_ADD
            'SaveCheck without Dialog showed
            If SaveCheck(True) = False Then
                SetButtonEnable True
                Me.MousePointer = Default
                Exit Sub
            End If
            Txt_A1601.text = ""
            sts_msgline.Panels(1) = G_Add_Ok
            If frm_EXAM01.Visible Then Txt_A1601.SetFocus

        Case G_AP_STATE_UPDATE
            If IsRecordChange() Then
                'SaveCheck without Dialog showed
                If SaveCheck(True) = False Then
                    SetButtonEnable True
                    Me.MousePointer = Default
                    Exit Sub
                End If
                sts_msgline.Panels(1) = G_Update_Ok
            End If

        Case G_AP_STATE_DELETE
            Delete_Process_A16
            Delete_Process_A19
            Delete_From_Menu
            sts_msgline.Panels(1) = G_Delete_Ok
    End Select
    G_DataChange% = False
    
    '還原所有Button的Enabled Property
    SetButtonEnable True
    
    Me.MousePointer = Default

    '作業狀態若為修改,刪除, 則返回V畫面
    If G_AP_STATE <> G_AP_STATE_ADD Then
       DoEvents
       Me.Hide
       frm_EXAM01v.Show
    End If
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : Cmd_Exit_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Cmd_Exit_Click()
    Me.MousePointer = HOURGLASS

'結束目前視窗,跳出其他處理程序
    m_ExitTrigger% = True
    
'設定會影響資料存檔的所有Button的Enabled Property = False
    SetButtonEnable False
    
'若資料有異動, 提示是否要存檔
    If IsRecordChange() Then
       If SaveCheck() = False Then
          SetButtonEnable True
          Me.MousePointer = Default
          Exit Sub
       End If
    End If

'還原所有Button的Enabled Property
    SetButtonEnable True

'隱藏目前畫面, 顯示V畫面
    DoEvents
    Me.Hide
    frm_EXAM01v.Show
    Me.MousePointer = Default
End Sub

'====================================
'   Form Events
'====================================

'========================================================================
' Procedure : Form_Activate (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/8/31
' Purpose   :
' Details   :
'========================================================================
Private Sub Form_Activate()
Dim A_A1601$
    Me.MousePointer = HOURGLASS
    sts_msgline.Panels(2) = GetCurrentDay(1)
    sts_msgline.Panels(1) = G_Process
    Me.Refresh
    
'Initial Form中的必要變數
    m_FieldError% = -1
    m_ExitTrigger% = False
    G_DataChange% = False
    
'判斷是否由其他輔助畫面回來, 而非首次執行
    If Trim(G_FormFrom$) <> "" Then
        Me.MousePointer = Default
        G_FormFrom$ = ""
        '.....                '加入所要設定之動作
        '.....
        StrCut frm_GD.Tag, Chr$(KEY_TAB), A_A1601$, ""
        With Spd_EXAM01
            .text = A_A1601$
            .Action = SS_ACTION_ACTIVE_CELL
            .Col = .Col + 1
            .text = Reference_A16(A_A1601$)
        End With
        Exit Sub
    Else
       '.....                '第一次執行時之準備動作
       '.....
       
        ClearFieldsValue
       
        Select Case G_AP_STATE
            Case G_AP_STATE_ADD
                'while from Q to D
                
            Case G_AP_STATE_UPDATE, G_AP_STATE_DELETE
                'while from V to D
                OpenMainFile
                MoveDB2Field
                If Not (DY_A19.BOF And DY_A19.EOF) Then
                    MoveDB2Spread
                End If
        End Select
        
        SetCommand          'set command buttom according to State
        
        If G_AP_STATE = G_AP_STATE_ADD Then
            If frm_EXAM01.Visible Then Txt_A1601.SetFocus
        Else
            If frm_EXAM01.Visible Then Txt_A1602.SetFocus
        End If
        sts_msgline.Panels(1) = SetMessage(G_AP_STATE)
    End If
    
    '將Form放置到螢幕的頂層
    frm_EXAM01.ZOrder 0
    Me.MousePointer = Default
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : Form_KeyDown
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Handle Key Event
' Details   : Handling: F1說明, F7上筆, F8下筆, F11確認, ESC離開
'========================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
           Case KEY_F1
                KeyCode = 0
                If ActiveControl.TabIndex = Spd_EXAM01.TabIndex Then Exit Sub
                If cmd_Help.Visible = True And cmd_Help.Enabled = True Then
                   cmd_Help.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F7
                KeyCode = 0
                If cmd_previous.Visible = True And cmd_previous.Enabled = True Then
                   cmd_previous.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F8
                KeyCode = 0
                If cmd_next.Visible = True And cmd_next.Enabled = True Then
                   cmd_next.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F11
                KeyCode = 0
                If cmd_ok.Visible = True And cmd_ok.Enabled = True Then
                   cmd_ok.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_ESCAPE
                KeyCode = 0
                If cmd_Exit.Visible = True And cmd_Exit.Enabled = True Then
                   cmd_Exit.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    End Select
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : Form_KeyPress
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Manage Uppercase input(A1601), and Record data changed
' Details   :
'========================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
    sts_msgline.Panels(1) = SetMessage(G_AP_STATE)
    
'主動將資料輸入由小寫轉為大寫
'  若有某些欄位不需要轉換時, 須予以跳過
'    If ActiveControl.TabIndex = Txt_A1605.TabIndex Then GoTo Form_KeyPress_A
'    If ActiveControl.TabIndex = Txt_A1612.TabIndex Then GoTo Form_KeyPress_A
'    If ActiveControl.TabIndex = Txt_A1514.TabIndex Then GoTo Form_KeyPress_A
'    If ActiveControl.TabIndex = txt_xxx.TabIndex Then GoTo Form_KeyPress_A
    If ActiveControl.TabIndex <> Txt_A1601.TabIndex Then GoTo Form_KeyPress_A
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Form_KeyPress_A:
'輸入任意字元(ENTER除外), 將資料異動變數設成TRUE
    If Not TypeOf ActiveControl Is SSCommand Then
       If KeyAscii <> KEY_RETURN Then G_DataChange% = True
    End If

    'If ActiveControl.TabIndex <> Spd_EXAM01.TabIndex Then
       KeyPress KeyAscii           'Enter時自動跳到下一欄位, spread除外
    'End If
    
'刪除作業下, 將KeyBoard鎖住, 不接受資料異動
    If G_AP_STATE = G_AP_STATE_DELETE Then KeyAscii = 0
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : Form_Load
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Form_Load()
    FormCenter Me
    Set_Property
End Sub

'========================================================================
' Module    : frm_EXAM01
' Procedure : Form_Unload
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    If cmd_Exit.Enabled Then cmd_Exit.SetFocus: Cmd_Exit_Click
End Sub

'====================================
'   Spread Evnets
'====================================

'========================================================================
' Procedure : Spd_EXAM01_Change (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   : 如任何一欄位有所變更時, 在P-key是空白情況下, 視同新增,
'             否則為修改狀態
'========================================================================
Private Sub Spd_EXAM01_Change(ByVal Col As Long, ByVal Row As Long)
Dim A_A1902$, A_A1903$      'Column Value of Spd_EXAM01
Dim A_A1902o$               '

    With Spd_EXAM01
        .Row = Row
        .Col = 1: A_A1902$ = Trim(.text)
        .Col = 3: A_A1903$ = Trim(.text)
        .Col = 4: A_A1902o$ = Trim(.text)
        .Col = 5
        If A_A1902o$ <> "" Then
            .text = "U"
        Else
            If A_A1902$ + A_A1903$ + A_A1902o$ <> "" Then
                .text = "A"
            Else
                .text = ""
            End If
        End If
    End With
End Sub

'========================================================================
' Procedure : Spd_EXAM01_DblClick (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : When DblClick on the first column, call frm_GD to help input
' Details   :
'========================================================================
Private Sub Spd_EXAM01_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Col <> 1 Then Exit Sub
    If Row < 1 Then Exit Sub
    G_FormFrom$ = frm_GD.Name
    frm_GD.Tag = "1"
    frm_GD.Show vbModal
    G_Hlp_Return = Spd_EXAM01.TabIndex
End Sub

'========================================================================
' Procedure : Spd_EXAM01_GotFocus (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : change color
' Details   :
'========================================================================
Private Sub Spd_EXAM01_GotFocus()
    SpreadGotFocus Spd_EXAM01.ActiveCol, Spd_EXAM01.ActiveRow
End Sub

'========================================================================
' Procedure : Spd_EXAM01_KeyUp (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Hotfix
' Details   :
'========================================================================
Private Sub Spd_EXAM01_KeyUp(KeyCode As Integer, Shift As Integer)
'標準指令, 避免中文字第一個字上不去, 不得修改
    SpreadKeyPress Spd_EXAM01, KeyCode
End Sub

'========================================================================
' Procedure : Spd_EXAM01_LeaveCell (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub Spd_EXAM01_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
On Local Error GoTo MY_Error
    'change color
    SpreadLostFocus Col, Row

    '判斷以下狀況發生時, 不須做任何處理
    If ActiveControl.TabIndex = cmd_Exit.TabIndex Then Exit Sub
    If ActiveControl.TabIndex = cmd_delete.TabIndex Then Exit Sub
    
    With Spd_EXAM01
        .Row = Row: .Col = Col

        '若跳離該筆時, 先檢查所有欄位是否正確, 再存檔
        '先判斷該筆是否有異動

        .Row = Row
        .Col = 5
        If Row <> NewRow And Trim(.text) <> "" Then     '
            '標準指令, 避免修改
            Dim A_Col&
            If SpreadLineCheck(Row, A_Col&) = False Then
                Cancel = True
                .Row = Row: .Col = A_Col&
                .Action = SS_ACTION_ACTIVE_CELL
                .SetFocus
                SpreadGotFocus A_Col&, Row
                Exit Sub
            End If
'   @Modified: Not Insert into DB, do it while Cmd_OK is clicked
'            MoveField2DB Row '!!!Old code
            GoTo New_Cell
        End If
        
        '判斷在最後一筆有輸入時, 自動增加一行
        '標準指令, 避免修改
        If Trim(.text) <> "" And Row = .MaxRows Then
            .MaxRows = .MaxRows + 1
        End If

        '每欄位是否要檢查
        .Row = Row
        .Col = 5
        If Trim(.text) <> "" Then
            Select Case Col
                Case 1
                    retcode = SpreadCheck_1(Row)
                Case 3
                    retcode = SpreadCheck_3(Row)
            End Select
         End If
    End With
    
New_Cell:
'新欄位顏色處理, 標準指令, 不得修改
    If NewCol > 0 Then SpreadGotFocus NewCol, NewRow
    Exit Sub
    
MY_Error:
    sts_msgline.Panels(1) = Error(Err)
End Sub

'========================================================================
' Procedure : Spd_EXAM01_MouseDown (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Update status bar
' Details   :
'========================================================================
Private Sub Spd_EXAM01_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    sts_msgline.Panels(1) = SetMessage(G_AP_STATE)
End Sub

'====================================
'   TextBox Evnets
'====================================

Private Sub Txt_A1601_GotFocus()
    TextGotFocus
End Sub

'========================================================================
' Procedure : Txt_A1601_LostFocus(frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   : Do check pkey not empty & duplicate
' Details   :
'========================================================================
Private Sub Txt_A1601_LostFocus()
    TextLostFocus
    
'判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A1601.TabIndex Then Exit Sub
    ' ....

'自我檢查
    retcode = CheckRoutine_A1601()
End Sub

Private Sub Txt_A1628_GotFocus()
    TextGotFocus
End Sub

'========================================================================
' Procedure : Txt_A1628_LostFocus (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Do check date format
' Details   :
'========================================================================
Private Sub Txt_A1628_LostFocus()
    TextLostFocus
    '判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A1628.TabIndex Then Exit Sub
    ' ....

    '自我檢查
    retcode = CheckRoutine_A1628()
End Sub

Private Sub Txt_A1602_GotFocus()
    TextGotFocus
End Sub

'========================================================================
' Procedure : Txt_A1602_LostFocus (frm_EXAM01)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Do check must input
' Details   :
'========================================================================
Private Sub Txt_A1602_LostFocus()
    TextLostFocus
    
'判斷以下狀況發生時, 不須做任何處理
    If G_AP_STATE = G_AP_STATE_DELETE Then Exit Sub
    If ActiveControl.TabIndex = cmd_Exit.TabIndex Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A1602.TabIndex Then Exit Sub
    ' ....

'自我檢查
    retcode = CheckRoutine_A1602()
End Sub

Private Sub Txt_A1609_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A1609_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A1605_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A1605_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A1606_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A1606_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A1612_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A1612_LostFocus()
    TextLostFocus
End Sub

Private Sub Vse_background_GotFocus()
    vse_background.TabStop = False
End Sub

