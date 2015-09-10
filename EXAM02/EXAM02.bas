Attribute VB_Name = "mod_EXAM02"
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
Global G_Form_EXAM02 As String
Global G_Form_EXAM02v As String
Global G_Form_EXAM02q As String

'定義各欄位標題文字
Global G_Pnl_Dash$
Global G_Pnl_A0801$
Global G_Pnl_A0802$
Global G_Pnl_A0803$
Global G_Pnl_A0804$
Global G_Pnl_A0805$
Global G_Pnl_A0806$
Global G_Pnl_A0807$
Global G_Pnl_A0808$
Global G_Pnl_A0809$
Global G_Pnl_A0810$
Global G_Pnl_A0811$
Global G_Pnl_A0812$
Global G_Pnl_A0813$
Global G_Pnl_A0814$
Global G_Pnl_A0815$
Global G_Pnl_A0816$
Global G_Pnl_A0817$
Global G_Pnl_A0818$
Global G_Pnl_A0819$
Global G_Pnl_A0820$
Global G_Pnl_A0821$
Global G_Pnl_A0822$
Global G_Pnl_A0823$
Global G_Pnl_A0824$
Global G_Pnl_A0825$
Global G_Pnl_A0826$
Global G_Pnl_A0201$
Global G_Pnl_A0202$
Global G_Pnl_A0601$
Global G_Pnl_A0602$
Global G_RecordNotExist$


'Def 程式共用變數
Global G_A0801$                      'Keep部門別
Global G_A0801o$                     'Keep Pre 部門別
Global G_A0824n$

Global G_ActiveRow#                  'Keep資料目前所在列
Global G_MaxRows#                    'Keep資料總筆數

Sub GetPanelCaption()
'取FORM標題文字
    G_Form_EXAM02 = GetCaption("FormTitle", "EXAM02", "員工基本資料管理")
    G_Form_EXAM02v = GetCaption("FormTitle", "EXAM02V", "員工基本資料目錄")
    G_Form_EXAM02q = GetCaption("FormTitle", "EXAM02Q", "員工基本資料查詢")
'取欄位標題文字
    G_Pnl_A0801$ = GetCaption("EXAM02", "A0801", "員工編號")
    G_Pnl_A0802$ = GetCaption("EXAM02", "A0802", "中文姓名")
    G_Pnl_A0803$ = GetCaption("EXAM02", "A0803", "英文姓名")
    G_Pnl_A0804$ = GetCaption("EXAM02", "A0804", "部門代號")
    G_Pnl_A0805$ = GetCaption("EXAM02", "A0805", "到職日期")
    G_Pnl_A0806$ = GetCaption("EXAM02", "A0806", "離職日期")
    G_Pnl_A0807$ = GetCaption("EXAM02", "A0807", "密碼")
    G_Pnl_A0808$ = GetCaption("EXAM02", "A0808", "出生日期")
    G_Pnl_A0809$ = GetCaption("EXAM02", "A0809", "身份證號碼")
    G_Pnl_A0810$ = GetCaption("EXAM02", "A0810", "中文地址")
    G_Pnl_A0811$ = GetCaption("EXAM02", "A0811", "英文地址")
    G_Pnl_A0812$ = GetCaption("EXAM02", "A0812", "城市")
    G_Pnl_A0813$ = GetCaption("EXAM02", "A0813", "郵遞區號")
    G_Pnl_A0814$ = GetCaption("EXAM02", "A0814", "國家")
    G_Pnl_A0815$ = GetCaption("EXAM02", "A0815", "連絡電話")
    G_Pnl_A0816$ = GetCaption("EXAM02", "A0816", "連絡傳真")
    G_Pnl_A0817$ = GetCaption("EXAM02", "A0817", "BB Call")
    G_Pnl_A0818$ = GetCaption("EXAM02", "A0818", "行動電話")
    G_Pnl_A0819$ = GetCaption("EXAM02", "A0819", "E-Mail Address")
    G_Pnl_A0820$ = GetCaption("EXAM02", "A0820", "有效日期")
    G_Pnl_A0821$ = GetCaption("EXAM02", "A0821", "性別")
    G_Pnl_A0822$ = GetCaption("EXAM02", "A0822", "婚姻狀況")
    G_Pnl_A0823$ = GetCaption("EXAM02", "A0823", "職稱")
    G_Pnl_A0824$ = GetCaption("EXAM02", "A0824", "公司別代碼")
    G_Pnl_A0825$ = GetCaption("EXAM02", "A0825", "群組代號")
    G_Pnl_A0826$ = GetCaption("EXAM02", "A0826", "User ID")
    
    G_Pnl_A0201$ = GetCaption("EXAM02", "A0201", "員工編號")
    G_Pnl_A0202$ = GetCaption("EXAM02", "A0202", "員工中文名稱")
    
    G_Pnl_A0601$ = GetCaption("EXAM02", "A0601", "群組代號")
    G_Pnl_A0602$ = GetCaption("EXAM02", "A0602", "群組說明")
    
    G_RecordNotExist$ = GetCaption("PgmMsg", "g_record_no_exist", "資料不存在! 請再查明!")
    
'取其他變數內含值
'    G_Pnl_A1602$ = GetCaption("EXAM02", "bankname")
    G_Pnl_Dash$ = GetCaption("PanelDescpt", "dash", "～")
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
    Load frm_EXAM02            ' 先將Detail畫面Load至Memory
    Frm_EXAM02q.Show           ' 首頁畫面顯示
    Screen.MousePointer = Default
End Sub
