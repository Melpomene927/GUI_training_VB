Attribute VB_Name = "mod_TSM01"
Option Explicit
Option Compare Text

'========================================================================
' Coding Rules
'========================================================================
'在此處定義之所有變數, 一律以G開頭
'       Ex: G_AAA$, G_BBB#, G_CCC&
'
'且變數之形態, 一律在最後一碼區別, 範例如下:
' $: 文字
' #: 所有數字運算(金額或數量)
' &: 程式迴圈變數
' %: 給一些使用於是或否用途之變數 (TRUE / FALSE )
' 空白: 代表VARIENT, 動態變數
'========================================================================

'-- Fixed Variables (必要變數) :
Global G_FormFrom$    '空白代表首次執行


'定義各 Form 標題文字變數
Global G_Form_PATTERN$

'定義各欄位標題文字

Global G_Pnl_A0101$
Global G_Pnl_A0102$
Global G_Pnl_A0104$
Global G_Pnl_A0111$

'Def 程式共用變數
''SAMPLE
Global G_BB#
Global G_CC!


Sub GetPanelCaption()
'取FORM標題文字
    G_Form_PATTERN$ = GetSIniStr("FormTitle", "TSM01")
'取欄位標題文字
     G_Pnl_A0101$ = GetCaption("TSM01", "ik01", "公司代號")
     G_Pnl_A0102$ = GetCaption("TSM01", "ik02", "公司簡稱")
     G_Pnl_A0104$ = GetCaption("TSM01", "ik03", "公司全名")
     G_Pnl_A0111$ = GetCaption("TSM01", "ik04", "統一編號")
'取其他變數內含值
End Sub

Sub Main()
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
    
    Frm_TSM01.Show            ' 首頁畫面顯示
    Screen.MousePointer = Default
End Sub

