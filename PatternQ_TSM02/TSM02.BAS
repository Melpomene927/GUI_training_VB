Attribute VB_Name = "mod_TSM02"
'========================================================================
' Module    : mod_TSM02
' Author    : Mike_chang
' Date      : 2015/8/27
' Purpose   :
'========================================================================
Option Explicit
Option Compare Text

'========================================================================
'   Coding Rule
'========================================================================
'在此處定義之所有變數, 一律以G開頭, 如G_AAA$, G_BBB#, G_CCC&
'且變數之形態, 一律在最後一碼區別, 範例如下:
' $: 文字
' #: 所有數字運算(金額或數量)
' &: 程式迴圈變數
' %: 給一些使用於是或否用途之變數 (TRUE / FALSE )
' 空白: 代表VARIENT, 動態變數
'========================================================================

'必要變數
Global G_FormFrom$    '空白代表首次執行

'定義各 Form 標題文字變數
Global G_Form_TSM02 As String
Global G_Form_TSM02v As String
Global G_Form_TSM02q As String

'定義各欄位標題文字
Global G_Pnl_Dash$
Global G_Pnl_A02001$
Global G_Pnl_A02005$
Global G_Pnl_A0201$
Global G_Pnl_A0202$
Global G_Pnl_A0203$
Global G_Pnl_A0204$
Global G_Pnl_A0205$
Global G_Pnl_A0206$
Global G_Pnl_A0207$
Global G_Pnl_A0208$
Global G_Pnl_A0209$
Global G_Pnl_A0213$
Global G_Pnl_A0214$
Global G_Pnl_A0215$
Global G_Pnl_A0216$
Global G_Pnl_A0217$
Global G_Pnl_A0218$
Global G_Pnl_A0219$
Global G_Pnl_Menu$
Global G_Pnl_Selected$

'Global G_Pnl_A15001$
'Global G_Pnl_A15005$
'Global G_Pnl_A1501$
'Global G_Pnl_A1502$
'Global G_Pnl_A1504$
'Global G_Pnl_A1505$
'Global G_Pnl_A1506$
'Global G_Pnl_A1507$
'Global G_Pnl_A1508$
'Global G_Pnl_A1509$
'Global G_Pnl_A1510$
'Global G_Pnl_A1511$
'Global G_Pnl_A1512$
'Global G_Pnl_A1513$
'Global G_Pnl_A1514$
'Global G_Pnl_A1519$
'Global G_Pnl_A1521$
'Global G_Pnl_A1523$
'Global G_Pnl_A1601$
'Global G_Pnl_A1602$


'Def 程式共用變數
Global G_A0201$                      'Keep部門別
Global G_A0201o$                     'Keep Pre 部門別

Global G_ActiveRow#                  'Keep資料目前所在列
Global G_MaxRows#                    'Keep資料總筆數
'SAMPLE
Global G_AA$
Global G_BB#
Global G_CC!

'========================================================================
' Module    : mod_TSM02
' Procedure : GetPanelCaption
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Get Panel Caption from "Sini" table
' Details   : Using GUI_common_module function:
'             "GetSIniStr" to get topicvalue from sini
'========================================================================
Sub GetPanelCaption()
'取FORM標題文字
    G_Form_TSM02 = GetSIniStr("FormTitle", "TSM02")
    G_Form_TSM02v = GetSIniStr("FormTitle", "TSM02V")
    G_Form_TSM02q = GetSIniStr("FormTitle", "TSM02Q")
'取欄位標題文字
    G_Pnl_A02001$ = GetSIniStr("PanelDescpt", "Date_Add")
    G_Pnl_A02005$ = GetSIniStr("PATTERNQ", "date_update")
    G_Pnl_A0201$ = GetSIniStr("TSM02", "A0201")
    G_Pnl_A0202$ = GetSIniStr("TSM02", "A0202")
    G_Pnl_A0203$ = GetSIniStr("TSM02", "A0203")
    G_Pnl_A0204$ = GetSIniStr("TSM02", "A0204")
    G_Pnl_A0205$ = GetSIniStr("TSM02", "A0205")
    G_Pnl_A0206$ = GetSIniStr("TSM02", "A0206")
    G_Pnl_A0207$ = GetSIniStr("TSM02", "A0207")
    G_Pnl_A0208$ = GetSIniStr("TSM02", "A0208")
    G_Pnl_A0209$ = GetSIniStr("TSM02", "A0209")
    
    G_Pnl_A0213$ = GetSIniStr("TSM02", "A0213")
    G_Pnl_A0214$ = GetSIniStr("TSM02", "A0214")
    G_Pnl_A0215$ = GetSIniStr("TSM02", "A0215")
    G_Pnl_A0216$ = GetSIniStr("TSM02", "A0216")
    G_Pnl_A0217$ = GetSIniStr("TSM02", "A0217")
    G_Pnl_A0218$ = GetSIniStr("TSM02", "A0218")
    G_Pnl_A0219$ = GetSIniStr("TSM02", "A0219")
    
'取其他變數內含值
'    G_Pnl_A1602$ = GetSIniStr("PATTERNQ", "bankname")
    G_Pnl_Dash$ = GetSIniStr("PanelDescpt", "dash")
    G_Pnl_Menu$ = GetSIniStr("PATTERNQ", "listmenu")
    G_Pnl_Selected$ = GetSIniStr("PATTERNQ", "listselected")
End Sub

'========================================================================
' Module    : mod_TSM02
' Procedure : main
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   :
' Details   :
'========================================================================
Sub main()
' 本模組中, 必須按照下列順序執行, 如果有特殊情況須將某些模組關閉時,
' 請在該模組前上 ' 即可, 不得刪除.

    Screen.MousePointer = HOURGLASS
    IsAppropriateCheck        ' 檢查本程式是否由MENU中呼叫執行
    DoubleRunCheck            ' 檢查本程式不得重覆執行
    GetSystemINIString        ' 抓取本系統使用之資料庫之路逕變數,
                              ' CHECK (C:\WINDOWS) LOCAL INI.
    OpenDB                    ' 開起本系統所有程式會使用之資料庫
    GetSystemDefault          ' 抓取本系統共同應有之環境參數設定,
                              ' CHECK LXXX.MDB中之INI TABLE.
    GetSvrDefault             ' 抓取本系統使用上, 特定限制條件, 如關帳日,
                              ' 底圖檔名, 日期格式, ...
    GetPanelCaption           ' 抓取本程式已設定共用變數之內含值
    Load frm_TSM02            ' 先將Detail畫面Load至Memory
    frm_TSM02q.Show           ' 首頁畫面顯示
    Screen.MousePointer = Default
End Sub
