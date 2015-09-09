Attribute VB_Name = "mod_EXAM01"
Option Explicit
Option Compare Text


'在此處定義之所有變數, 一律以G開頭, 如G_AAA$, G_BBB#, G_CCC&
'且變數之形態, 一律在最後一碼區別, 範例如下:
' $: 文字
' #: 所有數字運算(金額或數量)
' &: 程式迴圈變數
' %: 給一些使用於是或否用途之變數 (TRUE / FALSE )
' 空白: 代表VARIENT, 動態變數

'必要變數
Global G_FormFrom$    '空白代表首次執行

'定義各 Form 標題文字變數
Global G_Form_EXAM01 As String
Global G_Form_EXAM01v As String
Global G_Form_EXAM01q As String

'定義各欄位標題文字
Global G_Pnl_Dash$
Global G_Pnl_A1601$
Global G_Pnl_A1602$
Global G_Pnl_A1605$
Global G_Pnl_A1606$
Global G_Pnl_A1609$
Global G_Pnl_A1628$
Global G_Pnl_A1612$

Global G_Pnl_A1901$
Global G_Pnl_A1902$
Global G_Pnl_A1902n$
Global G_Pnl_A1903$


'SAMPLE
'Global G_AA$
'Global G_BB#
'Global G_CC!

'Def 程式共用變數
Global G_A1601$                      'Keep部門別
Global G_A1601o$                     'Keep Pre 部門別
Global G_Hlp_Return                  'Keep return value from form_GD

Global G_ActiveRow#                  'Keep資料目前所在列
Global G_MaxRows#                    'Keep資料總筆數

Sub GetPanelCaption()
    '取FORM標題文字
    G_Form_EXAM01 = GetCaption("FormTitle", "EXAM01", "客戶基本資料維護")
    G_Form_EXAM01v = GetCaption("FormTitle", "EXAM01q", "客戶基本資料瀏覽")
    G_Form_EXAM01q = GetCaption("FormTitle", "EXAM01V", "客戶基本資料查詢")
    '取欄位標題文字

    G_Pnl_A1601$ = GetCaption("EXAM01", "order1", "客戶編號")
    G_Pnl_A1602$ = GetCaption("paneldescpt", "custmer", "客戶簡稱")
    G_Pnl_A1605$ = GetCaption("PanelDescpt", "10", "聯絡電話")
    G_Pnl_A1606$ = GetCaption("KBM01", "fax", "傳真電話")
    G_Pnl_A1612$ = GetCaption("EXAM01", "address0", "聯絡地址")
    G_Pnl_A1609$ = GetCaption("EXAM01", "idcard", "身分證/統編")
    G_Pnl_A1628$ = GetCaption("EXAM01", "birthday", "生日/成立日")
    
    G_Pnl_A1901$ = GetCaption("EXAM01", "relPerson", "關係人")
    G_Pnl_A1902$ = GetCaption("EXAM01", "relPrnNum", "關係人編號")
    G_Pnl_A1902n$ = GetCaption("EXAM01", "relPrnBrf", "關係人簡稱")
    G_Pnl_A1903$ = GetCaption("PanelDescpt", "relationship", "關係")

    
    '取其他變數內含值
'    G_Pnl_A1602$ = GetCaption("EXAM01", "bankname")
    G_Pnl_Dash$ = GetCaption("PanelDescpt", "dash", "～")
End Sub

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
    Load frm_EXAM01q          ' 先將Detail畫面Load至Memory
    frm_EXAM01q.Show          ' 首頁畫面顯示
    Screen.MousePointer = Default
End Sub
