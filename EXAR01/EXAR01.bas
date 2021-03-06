Attribute VB_Name = "mod_EXAR01"
Option Explicit             'Do not allow ambiguous declaration
Option Compare Text         'Set all compare method as text-compare

'在此處定義之所有變數, 一律以G開頭, 如G_AAA$, G_BBB#, G_CCC&
'且變數之形態, 一律在最後一碼區別, 範例如下:
' $: 文字
' #: 所有數字運算(金額或數量)
' &: 程式迴圈變數
' %: 給一些使用於是或否用途之變數 (TRUE / FALSE )
' 空白: 代表VARIENT, 動態變數

Global G_FormFrom$    '空白代表首次執行

'========================================================================
'   定義各 Form 標題文字變數
'========================================================================
Global G_Form_EXAR01$
Global G_Form_EXAR01q$

'========================================================================
'   定義各欄位標題文字
'========================================================================
Global G_Pnl_A1601$
Global G_Pnl_A1602$
Global G_Pnl_A1605$
Global G_Pnl_A1606$
Global G_Pnl_A1609$
Global G_Pnl_A1614$
Global G_Pnl_A1617$
Global G_Pnl_A1620$
Global G_Pnl_A1621$
Global G_Pnl_A1643$
Global G_Pnl_Sum$
Global G_Pnl_Total$
Global G_Pnl_Credit$

Global G_Pnl_A0801$
Global G_Pnl_A0802$

Global G_Pnl_Dash$
Global G_Pnl_PrtType$
Global G_Pnl_Printer$
Global G_Pnl_Screen$
Global G_Pnl_File$
Global G_Pnl_Excel$

'========================================================================
'   Def 程式共用變數
'========================================================================
Global G_PathNotFound$
Global G_Report_Heading$

Global G_A1620_Total#
Global G_A1621_Total#
Global G_A1643_Total#
Global G_Credit_Total#

Global G_A1601s$
Global G_A1601e$
Global G_A1617s$
Global G_A1617e$
Global G_A1609s$
Global G_A1609e$


'========================================================================
'??? 在此宣告此程式中所有的Spread自訂型態變數,每個提供User自訂欄位的vaSpread,
'    必須宣告一個Spread自訂型態變數,命名如下:
'    vaSread Name : Spd_EXAR01   Spread Type Name: tSpd_EXAR01
'========================================================================
Global tSpd_Help As Spread
Global tSpd_EXAR01 As Spread

'========================================================================
'   Def 報表格式
'========================================================================
'Global Const H0$ = "....5...10....5...20....5...30....5...40....5...50....5...60....5...70....5...80....5...90....5..100....5..110....5..120....5..130....5..140....5..150....5..160....5..170....5..180....5..190....5..."
'Global Const H1$ = " "
'Global Const H2$ = "  <SCR01>                                                     ***  使用日誌列印  ***"
'Global Const H3$ = "  起始日期/時間 : 89/02/15   / 10:01:01"
'Global Const H4$ = "  截止日期/時間 : 89/02/15   / 11:44:47"
'Global Const H5$ = "  系統代號:"
'Global Const H6$ = "  程式代碼      :                                                                                                    頁次：1"
'Global Const H7$ = "  群組代號      :                                                                                                    日期：89/02/15"
'Global Const H8$ = "  User ID       :            -                                                                                       時間：11:44:47"
'Global Const H9$ = "  ================================================================================================================================="
'Global Const FC$ = "  系統代號  日期       時間     登錄   程式名稱                                  備註                                              "
'Global Const B1$ = "  使用者 : "
'Global Const B2$ = "  ---------------------------------------------------------------------------------------------------------------------------------"
'Global Const B3$ = "  科目合計   : 2   Start : 1  Exit : 1"
'Global Const B3$ = "  科目小計 : 2   Start : 1  Exit : 1"
'Global Const B3$ = "  使用者合計 : 2   Start : 1  Exit : 1"
'Global Const N1$ = "                                                                 ... 續 下 頁 ...                          印表人 :                "
'Global Const N2$ = "                                                                                                           印表人 :                "

'Global Const H0$ = "....5...10....5...20....5...30....5...40....5...50....5...60....5...70....5...80....5...90....5..100....5..110....5..120....5..130....5..140....5..150....5..160....5..170....5..180....5..190....5..."
'Global Const H1$ = " "
'Global Const H2$ = "  <SCR01> ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
'Global Const H3$ = "  ##############: ########## / ########"
'Global Const H4$ = "  ##############: ########## / ########"
'Global Const H5$ = "  ##############: ########## ########################################"
'Global Const H6$ = "  ##############: ########## ########################################                                        ######## : #####"
'Global Const H7$ = "  ##############: ### #########################################                                              ######## : ##########"
'Global Const H8$ = "  ##############: ########## - ##########                                                                    ######## : ##########"
'Global Const H9$ = "  ================================================================================================================================"
'Global Const FC$ = "  ########## ######## ##### ######################################## ########## ##################################################"
'Global Const FD$ = "  ########## ######## ##### ######################################## ########## ##################################################"
'Global Const N1$ = "                         ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^######## : ############"
'Global Const N2$ = "                                                                                                           ######## : ############"
'Global Const B1$ = "  ############ : "
'Global Const B2$ = "  --------------------------------------------------------------------------------------------------------------------------------"
'Global Const B3$ = "  科目合計     : #######   Start : #######  Exit :#######"
'Global Const B3$ = "  科目小計   : #######   Start : #######  Exit :#######"
'Global Const B3$ = "  使用者合計   : #######   Start : #######  Exit :#######"

'========================================================================
'??? 宣告報表格式變數
'========================================================================
Global Const H0$ = "....5...10....5...20....5...30....5...40....5...50....5...60....5...70....5...80....5...90....5..100....5..110....5..120....5..130....5..140....5..150....5..160....5..170....5..180....5..190....5..."
Global Const H1$ = " "
Global H2$
Global H3$
Global H4$
Global H5$
Global H6$
Global H7$
Global H8$
Global H3l$
Global H4l$
Global H5l$
Global H6l$
Global H7l$
Global H8l$
Global HDate$
Global HPerson$
Global H9$              'Line: =========================
Global B1$              'Break Header Line
Global B11$             'Break value Format of 1 column
Global B2$              'Line: -------------------------
Global B3$              'Line: #########################
Global B31$             'Break value Format of 3 columns
Global FC$              'Column Header Description
Global fd$              'Column Header Data
Global N1$
Global N2$

Sub GetPanelCaption()
'取FORM標題文字
    G_Form_EXAR01$ = GetCaption("FormTitle", "", "客戶資料列印")
    G_Form_EXAR01q$ = GetCaption("FormTitle", "", "客戶資料列印")
    
'取欄位標題文字
    G_Pnl_A1609$ = GetCaption("PanelDescpt", "unifyno", "統一編號")
    G_Pnl_A1617$ = GetCaption("", "", "負責業務")
    G_Pnl_A1601$ = GetCaption("PanelDescpt", "buyerid", "客戶編號")
    G_Pnl_A1602$ = GetCaption("PanelDescpt", "custmer", "客戶簡稱")
    G_Pnl_A1614$ = GetCaption("PanelDescpt", "liaison", "聯絡人")
    G_Pnl_A1605$ = GetCaption("PanelDescpt", "telno", "電話號碼")
    G_Pnl_A1606$ = GetCaption("PanelDescpt", "faxno", "傳真號碼")
    G_Pnl_A1620$ = GetCaption("PanelDescpt", "credit_limit", "授信額度")
    G_Pnl_A1621$ = GetCaption("PanelDescpt", "current_credit", "應收帳款")
    G_Pnl_A1643$ = GetCaption("PanelDescpt", "nr", "應收票據")
    G_Pnl_Credit$ = GetCaption("PanelDescpt", "credit use", "可用額度")
    
    G_Pnl_Sum$ = GetCaption("EXAR01", "sum", "小計")
    G_Pnl_Total$ = GetCaption("EXAR01", "total", "合計")
    
    G_Pnl_A0801$ = GetCaption("PanelDescpt", "", "業務員編號")
    G_Pnl_A0802$ = GetCaption("PanelDescpt", "p_name_c", "可用額度")
    
    G_Pnl_Dash$ = GetCaption("PanelDescpt", "dash", "∼")
    G_Pnl_PrtType$ = GetCaption("PanelDescpt", "printtype", "列印方式")
    G_Pnl_Printer$ = GetCaption("PanelDescpt", "printer", "印表機")
    G_Pnl_Screen$ = GetCaption("PanelDescpt", "screen", "螢幕顯示")
    G_Pnl_File$ = GetCaption("PanelDescpt", "file", "檔案")
    G_Pnl_Excel$ = GetCaption("PanelDescpt", "excel", "Excel")

'取列印替代文字
'    G_SlipAttrib_1$ = Reference_SINI("SlipAttrib", "1")
'    G_SlipAttrib_2$ = Reference_SINI("SlipAttrib", "2")
'    G_AccountUse_1$ = Reference_SINI("AccountUse", "1")
'    G_AccountUse_2$ = Reference_SINI("AccountUse", "2")
'    G_AccountUse_3$ = Reference_SINI("AccountUse", "3")
'    G_SlipType_1$ = Reference_SINI("SlipType", "1")
'    G_SlipType_2$ = Reference_SINI("SlipType", "2")
'    G_SlipType_3$ = Reference_SINI("SlipType", "3")
'    G_SlipType_4$ = Reference_SINI("SlipType", "4")
'    G_SlipType_5$ = Reference_SINI("SlipType", "5")
'    G_SlipType_6$ = Reference_SINI("SlipType", "6")
'    G_SlipType_7$ = Reference_SINI("SlipType", "7")
'    G_SlipType_8$ = Reference_SINI("SlipType", "8")
    
'取其他變數內含值
    G_PathNotFound$ = GetCaption("PgmMsg", "path_not_found", "檔案路徑錯誤!")
    G_Report_Heading$ = GetCaption("ReportHeading", "EXAR01", "科目列印")
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
    SetReportCols             ' 設定報表的所有欄位至Spread Type中
    
'??? 將所有明細畫面先Load進Memory,請修改Form Name
    Load frm_EXAR01           ' 為在Q畫面的設定鍵觸發時,能抓取V畫面Spread上
                              ' 的Caption,故於程式執行時先Load V畫面
                              
'??? 請修改成第一個畫面的Form Name
    frm_EXAR01q.Show       ' 首頁畫面顯示
    Screen.MousePointer = Default
End Sub

Sub PageCheck(Spd As vaSpread, Optional Break As Boolean = False)
'   under 2 circumstances do jump as below
'   1. reach maximum line
'   2. change to next break column
'   !!! "Excel" & "Screen Spread" does not need to jump to next page

    If G_PrintSelect = G_Print2Excel And Not Break Then Exit Sub
    If G_PrintSelect = G_Print2Screen Then Exit Sub
    '跳頁處理
    If G_LineNo > G_OverFlow Or Break Then          '@Alter R2: Adding "Break" mechanism
        If G_PageNo > 0 Then
           If G_PrintSelect <> G_Print2Excel Then   '@Alter R2:
              PrintOut3 Spd, H1$, ""
              PrintOut3 Spd, H1$, ""
              PrintOut3 Spd, N1$, ""
           End If
           If G_PrintSelect = G_Print2Printer Then
              Printer.NewPage
           ElseIf G_PrintSelect = G_Print2Excel Then
              SetExcelNewPage
           Else
              Print #1, G_G1
           End If
        End If
        If G_PrintSelect <> G_Print2Excel Then ReportHeader Spd
    End If
End Sub

Sub PrePare_Data(Frm As Form, Prb As ProgressBar, Spd As vaSpread, A_Exit%)
On Local Error GoTo MY_Error
    
    '??? 設定ProgressBar最大值
    If G_ReportDataFrom = G_FromRecordSet Then
       Spd.MaxRows = 0
       DY_A16.MoveLast
       Prb.MAX = DY_A16.RecordCount
       DY_A16.MoveFirst
    Else
       Prb.MAX = Spd.MaxRows
    End If
    
    '開啟文字檔
    If G_PrintSelect = G_Print2File Then
        Open G_OutFile For Output As #1
    ElseIf G_PrintSelect = G_Print2Excel Then
        If Not OpenExcelFile(G_OutFile) Then
           Frm!Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
           Exit Sub
        End If
        '??? Excel設定初始化
        Set_Excel_Property Spd, tSpd_EXAR01
    End If

    '初始化tSpd中的欄位值
    InitialtSpdTextValue tSpd_EXAR01

    '設定動態的報表格式
    SetPrintFormatStr
    
    '設定報表字體,字型及印表機設定
    If Not ReportSet() Then Frm!Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE): Exit Sub

    '??? 若有Break欄位時,須重新調整報表的欄寬
    AdjustColWidth Spd, tSpd_EXAR01, "A1617", B31$
    
    
    '資料列印處理
    If G_ReportDataFrom = G_FromRecordSet Then
       Print2Spread Prb, Spd, A_Exit%
    Else
       PrintSub Prb, Spd, True, A_Exit%
    End If
    
    '當Esc鍵被觸發,結束列印動作
    If A_Exit% Then Exit Sub
    
    Frm!Sts_MsgLine.Panels(1) = G_PrintOk
    Exit Sub

MY_Error:
    Beep
    Select Case Err
      Case 55
           Resume Next
      Case Else
           MsgBox Error$(Err)
    End Select
    Err = 0
End Sub

Sub Print2Spread(Prb As ProgressBar, Spd As vaSpread, A_Exit%)
Dim A_FmtStr$                                   'Format String
Dim A_A1601$, A_A1602$, A_A1614$                'Output Column Values
Dim A_A1605$, A_A1606$, A_A1620$
Dim A_A1621$, A_A1643$, A_credit$
Dim A_A1617$, A_A1617_Brk$                      'Break Column(A1607科目大類) & Previous Value of it
Dim A_Row#, A_Index#                            'Statical Counter
Dim A_Break_Value#                              '授信額度小計 of A1620
Dim A_Break_Value2#                             '應收帳款小計 of A1621
Dim A_Break_Value3#                             '應收票據小計 of A1643
Dim A_Break_Value4#                             '可用額度小計 of credit

    'Initialize
    Prb.Visible = True
    Prb.Value = 0
    A_Row# = 0
    A_Index# = 0
    Spd.MaxRows = 0
    
    'Initialize Summary Counter
    A_Break_Value# = 0
    A_Break_Value2# = 0
    A_Break_Value3# = 0
    A_Break_Value4# = 0
    G_A1620_Total# = 0
    G_A1621_Total# = 0
    G_A1643_Total# = 0
    G_Credit_Total# = 0
    
    
    '列印表頭
    ReportHeader Spd
    
    'Keep Break Value
    A_A1617_Brk$ = Trim$(DY_A16.Fields("A0802") & "")
    A_A1617$ = A_A1617_Brk$
    
    'Setup Output format
    A_FmtStr$ = "FD$"   'Format: [Break Header] + [-------] + [Data]
'    A_FmtStr$ = "B1$;B2$;FD$"   'Format: [Break Header] + [-------] + [Data]

    'Loop to Dump Report Values
    Do While Not DY_A16.EOF And Not A_Exit%
       
        '累加目前處理的資料筆數
        A_Index# = A_Index# + 1
    
        'If change to another break
        If StrComp(A_A1617_Brk$, Trim$(DY_A16.Fields("A0802") & ""), _
            vbTextCompare) <> 0 Then
                 
            '列印小計
            PrintBreak Spd, A_Row#, G_Pnl_Sum$, A_Break_Value#, A_Break_Value2#, _
                A_Break_Value3#, A_Break_Value4#, "B2$;B3$;B2$"
          
            '變數歸零,以便重新累計
            A_Break_Value# = 0
            A_Break_Value2# = 0
            A_Break_Value3# = 0
            A_Break_Value4# = 0
          
            'Keep Break Value
            A_A1617_Brk$ = Trim$(DY_A16.Fields("A0802") & "")
            A_A1617$ = A_A1617_Brk$
            
            'Setup Output format
            A_FmtStr$ = "NP;FD$"    'Format: [NewPage] + [Data]
'            A_FmtStr$ = "NP;B1$;B2$;FD$"    'Format: [NewPage] + [Break Header] + [-------] + [Data]
        End If
              
        'Keep列印資料至變數
        'col2 客戶編號 10
        A_A1601$ = Trim$(DY_A16.Fields("A1601") & "")
        'col3 客戶簡稱 12
        A_A1602$ = Trim$(DY_A16.Fields("A1602") & "")
        'col4 聯絡人 20
        A_A1614$ = Trim$(DY_A16.Fields("A1614") & "")
        'col5 電話號碼 15
        A_A1605$ = Trim$(DY_A16.Fields("A1605") & "")
        'col6 傳真號碼 15
        A_A1606$ = Trim$(DY_A16.Fields("A1606") & "")
        'col7 授信額度 8
        A_A1620$ = Trim$(DY_A16.Fields("A1620") & "")
        'col8 應收帳款 8
        A_A1621$ = Trim$(DY_A16.Fields("A1621") & "")
        'col9 應收票據 8
        A_A1643$ = Trim$(DY_A16.Fields("A1643") & "")
        'col10 可用額度 8
        A_credit$ = Str(Val(A_A1620$) - Val(A_A1621$) - Val(A_A1643$))
        
        'sum up col7~10 to break value
        G_A1620_Total# = G_A1620_Total# + CDbl(A_A1620$)
        G_A1621_Total# = G_A1621_Total# + CDbl(A_A1621$)
        G_A1643_Total# = G_A1643_Total# + CDbl(A_A1643$)
        G_Credit_Total# = G_Credit_Total# + CDbl(A_credit$)

        A_Break_Value# = A_Break_Value# + CDbl(A_A1620$)
        A_Break_Value2# = A_Break_Value2# + CDbl(A_A1621$)
        A_Break_Value3# = A_Break_Value3# + CDbl(A_A1643$)
        A_Break_Value4# = A_Break_Value4# + CDbl(A_credit$)
        
       
        '將Spread上的MaxRows加一
        AddSpreadMaxRows Spd, A_Row#
        
        '========================================================================
        ' [Mechanism Desciption]:
        '??? 以欄位名稱設定欄位值至vaSpread
        '    參數一 : Spread Name           參數二 : 參數一所屬的Spead Type Name
        '    參數三 : 自訂的欄位名稱        參數四 : 資料列
        '    參數五 : 填入值
        '========================================================================
        SetSpdText Spd, tSpd_EXAR01, "A1617", A_Row#, A_A1617$   '負責業務
        SetSpdText Spd, tSpd_EXAR01, "A1601", A_Row#, A_A1601$   '客戶編號
        SetSpdText Spd, tSpd_EXAR01, "A1602", A_Row#, A_A1602$   '客戶簡稱
        SetSpdText Spd, tSpd_EXAR01, "A1614", A_Row#, A_A1614$   '聯絡人
        SetSpdText Spd, tSpd_EXAR01, "A1605", A_Row#, A_A1605$   '電話號碼
        SetSpdText Spd, tSpd_EXAR01, "A1606", A_Row#, A_A1606$   '傳真號碼
        SetSpdText Spd, tSpd_EXAR01, "A1620", A_Row#, A_A1620$   '授信額度
        SetSpdText Spd, tSpd_EXAR01, "A1621", A_Row#, A_A1621$   '應收帳款
        SetSpdText Spd, tSpd_EXAR01, "A1643", A_Row#, A_A1643$   '應收票據
        SetSpdText Spd, tSpd_EXAR01, "credit", A_Row#, A_credit$ '可用額度
        SetSpdText Spd, tSpd_EXAR01, "Flag", A_Row#, A_FmtStr$
'        SetSpdText Spd, tSpd_EXAR01, "TEST", A_Row#, "TEST"
        
        
       
        '設定Spread第一列的列數
        If G_PrintSelect = G_Print2Screen Then Spd.TopRow = SetSpreadTopRow(Spd)
       
        '========================================================================
        ' [Mechanism Desciption]:
        '   若於Q畫面選擇 "非螢幕顯示" 的列印方式
        '   亦先將資料Prepare至V Screen的Spread上.
        '   若Spread的MaxRows大於等於100筆,則先跳至PrintSub將Spread上的資料印出,
        '   並將MaxRows歸零,再繼續Prepare資料至V Screen.
        '========================================================================
        If (G_ReportDataFrom = G_FromRecordSet And G_PrintSelect <> _
            G_Print2Screen) And A_Row# >= 100 Then
            GoSub Print2SpreadA
        End If
       
        '清空,Break以後,不列印此欄位的資料
        A_A1617$ = ""
       
        '設定資料列的套用格式
        A_FmtStr$ = "FD$"   'Format: [ReportData]
       
        '顯示目前處理進度
        Prb.Value = A_Index#
       
        DoEvents
       
        '當Esc鍵被觸發,結束列印動作
        If A_Exit% Then Exit Do
       
        DY_A16.MoveNext
       
    Loop
    
    '已觸發結束鍵時, 跳出此程序
    If A_Exit% Then Exit Sub

    '列印表尾
    '列印科目合計的Break
    PrintBreak Spd, A_Row#, G_Pnl_Sum$, A_Break_Value#, A_Break_Value2#, _
        A_Break_Value3#, A_Break_Value4#, "B2$;B3$;H9$"
          
    '列印科目小計的Break
    PrintBreak Spd, A_Row#, G_Pnl_Total$, G_A1620_Total#, G_A1621_Total#, G_A1643_Total#, G_Credit_Total#, "B3$;H9$"
          
    '若於Q畫面選擇非螢幕顯示的列印方式,於資料處理結束,須再將Spread上的資料印出.
    If (G_ReportDataFrom = G_FromRecordSet And G_PrintSelect <> G_Print2Screen) _
        And Spd.MaxRows > 0 Then
       GoSub Print2SpreadA
    End If
    
    '處理資料列印完成後的結束動作
    PrintBottom Prb, Spd
    Exit Sub
    
Print2SpreadA:
    '將資料由Spread讀取列印至文字檔,印表機或Excel
    PrintSub Prb, Spd, False, A_Exit%
    ClearSpreadText Spd
    Spd.MaxRows = 0
    Return
End Sub

Sub PrintBottom(Prb As ProgressBar, Spd As vaSpread)
    
    '??? 列印印表人
    PrintOut3 Spd, H1$, "", -1
    PrintOut3 Spd, H1$, "", -1
    PrintOut3 Spd, N2$, "", -1


    '??? 將起始欄位中的資料,以G_G1字元將資料切割成多個欄位值
    SetExcelTextToColumns G_XlsStartCol%, 1, G_XlsHRows% + G_ExcelIndex#, _
        SetXlsFldDataType(tSpd_EXAR01)
    
    '設定Excel的報表格式
    SetExcelFormat

    '??? 處理各種列印方式之結束動作
    PrintEnd4 Spd, tSpd_EXAR01
    Prb.Visible = False
End Sub

Sub PrintBreak(Spd As vaSpread, A_Row#, ByVal A_Desc$, ByVal A_Break_Value#, _
    ByVal A_Break_Value2#, ByVal A_Break_Value3#, ByVal A_Break_Value4#, ByVal _
    A_FmtStr$)
Dim A_STR$, A_PrtStr$, A_Col&, A_Len&, A_Len2&

    'Keep列印資料至變數
    A_STR$ = A_Desc$ & G_G1 & Format(A_Break_Value#, "#,##0.00")
    A_STR$ = A_STR$ & G_G1 & Format(A_Break_Value2#, "#,##0.00")
    A_STR$ = A_STR$ & G_G1 & Format(A_Break_Value3#, "#,##0.00")
    A_STR$ = A_STR$ & G_G1 & Format(A_Break_Value4#, "#,##0.00")
    A_STR$ = PrintUse(B31$, A_STR$)
    
    '將Spread上的MaxRows加一
    AddSpreadMaxRows Spd, A_Row#
    
    '以欄位名稱設定欄位值至Spread
    SetSpdText Spd, tSpd_EXAR01, "A1617", A_Row#, A_STR$
    
    '設定該列列印時所套用的格式字串
    SetSpdText Spd, tSpd_EXAR01, "Flag", A_Row#, A_FmtStr$
    
    '設定Break資料列的顏色
    SetSpreadColor Spd, A_Row#, -1, CStr(COLOR_YELLOW), G_TextGotFore_Color
    
    '設定Spread第一列的列數
    If G_PrintSelect = G_Print2Screen Then Spd.TopRow = SetSpreadTopRow(Spd)
End Sub

Sub PrintSub(Prb As ProgressBar, Spd As vaSpread, ByVal ShowProgress%, A_Exit%)
Dim A_PrtStr$, A_A1617$, A_FmtStr$()
Dim A_Row#, I#

    '由V Screen執行的列印動作,才須處理的作業
    If ShowProgress% Then
        Prb.Visible = True
        Prb.Value = 0
        ReportHeader Spd
    End If
    
    '設定初始值
    A_Row# = 0
    
    '處理每列資料的列印
    Do While A_Row# < Spd.MaxRows And Not A_Exit%
        
        '累加目前處理的資料筆數
        A_Row# = A_Row# + 1
        '========================================================================
        '??? 以欄位名稱取得欄位值,置入Columns Type的Text屬性中
        '    參數一 : Spread Name           參數二 : 參數一所屬的Spead Type Name
        '    參數三 : 自訂的欄位名稱        參數四 : 資料列
        '========================================================================
        A_A1617$ = GetSpdText(Spd, tSpd_EXAR01, "A1617", A_Row#)
        GetSpdText Spd, tSpd_EXAR01, "A1601", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1602", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1614", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1605", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1606", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1620", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1621", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "A1643", A_Row#
        GetSpdText Spd, tSpd_EXAR01, "credit", A_Row#
        A_FmtStr$ = Split(GetSpdText(Spd, tSpd_EXAR01, "Flag", A_Row#), ";")
        
        For I# = 0 To UBound(A_FmtStr$)
       
            '累計目前列印行數,若超過一頁則跳頁
            G_LineNo = G_LineNo + 1
            PageCheck Spd
           
            '??? 將字串傳給PrintOut3處理列印動作
            Select Case UCase$(A_FmtStr$(I#))
                Case "H1$"      'Single white space
                    PrintOut3 Spd, H1$, "", -1
                    
                Case "H9$"      'Line: =========================
                    PrintOut3 Spd, H9$, "", -1
                    
                Case "B2$"      'Line: -------------------------
                    PrintOut3 Spd, B2$, "", -1
                    
                Case "B1$"      'Break Header
                    G_ExcelIndex# = G_ExcelIndex# + 1
                    If G_PrintSelect = G_Print2Excel Then
                       A_PrtStr$ = PrintUse(B1$, G_Pnl_A1617$ & G_G1 & A_A1617$)
                    Else
                       A_PrtStr$ = G_Pnl_A1617$ & G_G1 & A_A1617$
                    End If
                    PrintOut3 Spd, B1$, A_PrtStr$, G_ExcelIndex#
                    '若列印至Excel時,合併Break欄位的儲存格
                    SetCellAlignment GetMergeCols(1, G_ExcelIndex# + _
                        G_XlsHRows%, G_ExcelMaxCols%, G_ExcelMaxCols%, 0), xlLeft, _
                        xlCenter, True
                        
                Case "B3$"      'Break Value
                    G_ExcelIndex# = G_ExcelIndex# + 1
                    PrintOut3 Spd, B3$, A_A1617$, G_ExcelIndex#
                    
                    '設定Excel Cells Range的背景顏色
                    SetExcelRangeColor G_XlsHRows% + G_ExcelIndex#, G_XlsHRows% _
                        + G_ExcelIndex#, G_XlsStartCol%, G_ExcelMaxCols%, _
                        COLOR_YELLOW
                        
                    '若列印至Excel時,合併Break欄位的儲存格
                    SetCellAlignment GetMergeCols(1, G_ExcelIndex# + _
                        G_XlsHRows%, G_ExcelMaxCols%, G_ExcelMaxCols%, 0), xlRight, _
                        xlCenter, True
                        
                Case "FD$"      'Contents
                    G_ExcelIndex# = G_ExcelIndex# + 1
                    PrintOut3 Spd, fd$, PrintStrConnect(tSpd_EXAR01, 2), _
                        G_ExcelIndex#
                        
                Case "NP"       'New Page
                    PageCheck Spd, True
                    
            End Select
        Next I#
       
        '當Esc鍵被觸發,結束列印動作
        If A_Exit% Then Exit Do
       
        '由V Screen執行的列印動作,須顯示目前處理進度
        If ShowProgress% Then Prb.Value = A_Row#
        DoEvents
    Loop
    
    '已觸發結束鍵時, 跳出此程序
    If A_Exit% Then Exit Sub
    
    '由V Screen執行的列印動作完成後,須執行的結束動作
    If ShowProgress% Then PrintBottom Prb, Spd
End Sub

Sub ReDefineHeaderAlign()
'針對與預設值不同的欄位,重新設定報表抬頭欄位的對齊方式

    ChangeReportHeaderAlign tSpd_EXAR01, "A1617", SS_CELL_H_ALIGN_LEFT
'    ChangeReportHeaderAlign tSpd_EXAR01, "A0902", SS_CELL_H_ALIGN_CENTER
'    ChangeReportHeaderAlign tSpd_EXAR01, "A0906", SS_CELL_H_ALIGN_CENTER
'    ChangeReportHeaderAlign tSpd_EXAR01, "A0907", SS_CELL_H_ALIGN_RIGHT
'    ChangeReportHeaderAlign tSpd_EXAR01, "A0909", SS_CELL_H_ALIGN_RIGHT
'                                   :
'                                   :
End Sub

Sub ReDefineReportHeader()
'針對與預設值不同的欄位,重新設定報表抬頭欄位的Caption

'    ChangeReportHeader tSpd_EXAR01, "A0901", "Test"
'    ChangeReportHeader tSpd_EXAR01, "A0902", "Test"
'    ChangeReportHeader tSpd_EXAR01, "A0906", "Test"
'    ChangeReportHeader tSpd_EXAR01, "A0907", "Test"
'    ChangeReportHeader tSpd_EXAR01, "A0909", "Test"
'                                   :
'                                   :
End Sub

Private Function Reference_SINI(ByVal A_Section$, ByVal A_Topic$) As String
On Local Error GoTo MyError
Dim A_Sql$

    Reference_SINI = ""
    A_Sql$ = "Select TOPICVALUE From SINI"
    A_Sql$ = A_Sql$ & " where SECTION='" & A_Section$ & "'"
    A_Sql$ = A_Sql$ & " and TOPIC='type" & A_Topic$ & "'"
    A_Sql$ = A_Sql$ & " order by SECTION,TOPIC"
    CreateDynasetODBC DB_ARTHGUI, DY_INI, A_Sql$, "DY_INI", True
    If Not (DY_INI.BOF And DY_INI.EOF) Then
       Reference_SINI = Trim$(DY_INI.Fields("TOPICVALUE") & "")
    End If
    Exit Function
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Sub ReportHeader(Spd As vaSpread)
Dim A_H2$, A_H3$, A_H4$, A_H5$, A_FC$
Dim A_FirstColName$, A_LastColName$

    '選擇螢幕列印時,不列印表頭
    If G_PrintSelect = G_Print2Screen Then Exit Sub

    '??? 初始值頁目前列數=表頭總列數
    G_LineNo = 8
    
    '每列印一次表頭,頁數累加一
    G_PageNo = G_PageNo + 1

    '串接表頭資料至變數
    If G_PrintSelect = G_Print2Excel Then
        G_ExcelWkb.Visible = True
        '===========================================
        '???取得Excel首欄及最後一欄的自訂欄位名稱
        '===========================================
        A_FirstColName$ = GetRptColName(tSpd_EXAR01, 1)
        A_LastColName$ = GetRptColName(tSpd_EXAR01, GetReportCols(tSpd_EXAR01))
       
        '===========================================
        '???將Excel Header的資料,暫存至Spread Type
        '   再利用PrintStrConnect串接列印列資料至變數.
        '   於PrintStrConnect Function將參數二設為2
        '   以取得Header資料字串
        '===========================================
        SetTSpdText tSpd_EXAR01, A_FirstColName$, G_Report_Heading$
        A_H2$ = PrintStrConnect(tSpd_EXAR01, 2)
        SetTSpdText tSpd_EXAR01, A_FirstColName$, H3l$
        A_H3$ = PrintStrConnect(tSpd_EXAR01, 2)
        SetTSpdText tSpd_EXAR01, A_FirstColName$, H4l$
        SetTSpdText tSpd_EXAR01, A_LastColName$, PrintUse(HDate$, G_Print_Date & G_G1 & GetCurrentDay(1))
        A_H4$ = PrintStrConnect(tSpd_EXAR01, 2)
        SetTSpdText tSpd_EXAR01, A_FirstColName$, H5l$
        SetTSpdText tSpd_EXAR01, A_LastColName$, PrintUse(HDate$, G_Print_Time & G_G1 & Format(Now, "HH:MM:SS"))
        A_H5$ = PrintStrConnect(tSpd_EXAR01, 2)
        
'        SetTSpdText tSpd_EXAR01, A_FirstColName$, H6l$
'        A_H6$ = PrintStrConnect(tSpd_EXAR01, 2)
'        SetTSpdText tSpd_EXAR01, A_FirstColName$, H7l$
'        SetTSpdText tSpd_EXAR01, A_LastColName$, PrintUse(HDate$, G_Print_Date & G_G1 & GetCurrentDay(1))
'        A_H7$ = PrintStrConnect(tSpd_EXAR01, 2)
'        SetTSpdText tSpd_EXAR01, A_FirstColName$, H8l$
'        SetTSpdText tSpd_EXAR01, A_LastColName$, PrintUse(HDate$, G_Print_Time & G_G1 & Format(Now, "HH:MM:SS"))
'        A_H8$ = PrintStrConnect(tSpd_EXAR01, 2)
    Else
        '??? 設定列印至印表機或文字檔Header的字串變數,資料間以G_G1做區隔
        A_H2$ = G_Report_Heading$
        A_H3$ = G_Print_Page & G_G1 & Format(G_PageNo, "###0")
        A_H4$ = G_Print_Date & G_G1 & GetCurrentDay(1)
        A_H5$ = G_Print_Time & G_G1 & Format(Now, "HH:MM:SS")
    End If
    
    '??? 列印報表表頭
    PrintOut3 Spd, H1$, "", 1         '最後一個參數: 列印的行數
    PrintOut3 Spd, H2$, A_H2$, 2
    PrintOut3 Spd, H3$, A_H3$, 3
    PrintOut3 Spd, H4$, A_H4$, 4
    PrintOut3 Spd, H5$, A_H5$, 5
'    PrintOut3 Spd, H6$, A_H6$, 6
'    PrintOut3 Spd, H7$, A_H7$, 7
'    PrintOut3 Spd, H8$, A_H8$, 8
    PrintOut3 Spd, H9$, "", -1
    PrintOut3 Spd, FC$, FC$, 6
    PrintOut3 Spd, H9$, "", -1

    '??? 設定列印至Excel時,表頭所用的列數
    If G_PrintSelect = G_Print2Excel Then G_XlsHRows% = 6
End Sub

Function ReportSet() As Boolean
    ReportSet = True
    
    'Initial報表張數及每頁行數變數值
    G_PageNo = 0
    G_LineNo = 0
    
    'Initial報表是否有資料已列印的變數值
    G_HaveDataPrint% = False
    
    '選擇Excel or Screen列印,不須設定印表設定
    If G_PrintSelect = G_Print2Screen Then Exit Function
    If G_PrintSelect = G_Print2Excel Then Exit Function
    
    '設定報表屬性初始值
    G_FontName = GetIniStr("Font", "Name", "GUI.INI")
    G_FontSize = GetGLRptFont("Font3")
    G_PageSize = GetGLRptPageLine("Font3")
    G_OverFlow = GetGLRptPageSize("Font3")
    
    '若選擇印表機列印,則顯示印表設定視窗
    If G_PrintSelect = G_Print2Printer Then
        Printer.FontName = G_FontName
        Printer.FontSize = G_FontSize
        rptset.Show MODAL
        G_OverFlow = G_PageSize - 6
        If G_PageSize <= 0 Then G_PageSize = 60
        If G_OverFlow <= 0 Then G_OverFlow = 60
        ReportSet = G_RptSet
    End If
End Function

Sub Set_Excel_Property(Spd As vaSpread, tSPD As Spread)
'設定起始欄位為1,並將Headers的總列數歸零
    G_XlsStartCol% = 1: G_XlsHRows% = 0

'將目前列數歸零
    G_ExcelIndex# = 0
    
'取得Excel報表的總欄數
    G_ExcelMaxCols% = GetReportCols(tSPD)
    
 '設定Excel每個欄位的資料型態及對齊位置
    SetExcelDataType Spd, tSPD
End Sub

Sub SetExcelFormat()
Dim A_MaxCol$, A_Row#

    If G_PrintSelect <> G_Print2Excel Then Exit Sub

    '========================================================================
    ' Excel style Setting
    '========================================================================
    '??? 目前EXCEL印到第幾列
    A_Row# = G_ExcelIndex# + G_XlsHRows%
    
    '??? 該最大欄位數之名稱
    A_MaxCol$ = Chr(Asc("A") + G_ExcelMaxCols% - 1)

    '??? 在報表四周加上線條
    xlsDrawLine G_ExcelWkb, "A" & Trim(Str(G_XlsHRows%)) & ":" & A_MaxCol$ & _
        Trim(Str(A_Row#))

    '??? 標題跨欄置中(若有需要,請修改Rang中的範圍即可使用) FALSE-不合併儲存格
    SetCellAlignment "A2:" & A_MaxCol$ & "2", xlCenter, xlCenter, True

    '========================================================================
    ' Header Left Part Setting
    '========================================================================
    '??? N/A(儲存格合併)
    SetCellAlignment GetMergeCols(1, 3, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
        1), xlLeft, xlCenter, True

    '??? 公司別(儲存格合併)
    SetCellAlignment GetMergeCols(1, 4, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
        1), xlLeft, xlCenter, True

    '??? 科目範圍(儲存格合併)
    SetCellAlignment GetMergeCols(1, 5, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
        1), xlLeft, xlCenter, True

'    '??? 程式代碼(儲存格合併)
'    SetCellAlignment GetMergeCols(1, 6, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
'        1), xlLeft, xlCenter, True
'
'    '??? 群組代號(儲存格合併)
'    SetCellAlignment GetMergeCols(1, 7, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
'        1), xlLeft, xlCenter, True
'
'    '??? User ID(儲存格合併)
'    SetCellAlignment GetMergeCols(1, 8, G_ExcelMaxCols%, G_ExcelMaxCols% - 1, _
'        1), xlLeft, xlCenter, True
'
    
    '========================================================================
    ' Header Right Part Setting
    '========================================================================
    '??? 列印日期置右
    SetCellAlignment GetExcelColName(G_ExcelMaxCols% + G_XlsStartCol% - 1) & _
        "4", xlRight, xlCenter, True

    '??? 列印時間置右
    SetCellAlignment GetExcelColName(G_ExcelMaxCols% + G_XlsStartCol% - 1) & _
        "5", xlRight, xlCenter, True
    
    '??? 標題列置中
    SetCellAlignment Trim(Str(G_XlsHRows%)) + ":" + Trim(Str(G_XlsHRows%)), _
        xlCenter, xlCenter, False
    
    
    
    '========================================================================
    ' Other Setting
    '========================================================================
    '??? 設定Excel欄及列的預設大小,並調整報表欄寬至最適寬度
    SetExcelSize "A:" & A_MaxCol$
    
    '版面設定 (Orientation%參數值 - xlPortrait:直印  xlLandscape:橫印)
    SetExcelAllocate "$1:$" & Trim(Str(G_XlsHRows%))
    
    '設定作用儲存格
    SelectExcelCells "A1"
End Sub

Sub SetPrintFormatStr()

    '??? 對報表表頭格式進行變數初始值的動作
    H3l$ = "########## : ########## - ##########"
    H4l$ = "########## : ############ - ############"
    H5l$ = "########## : ############### - ###############"
    HDate$ = "######## : ##########"
    HPerson$ = "######## : ############"
    B31$ = "######## : ~~~~~~~~~~~~~~~ ~~~~~~~~~~~~~~~ ~~~~~~~~~~~~~~~ ~~~~~~~~~~~~~~~  "
    B11$ = "######## : ###############"

    '螢幕顯示不須設定報表格式
    If G_PrintSelect = G_Print2Screen Then Exit Sub

    '??? 設定報表左右側空間及欄間隔,若使用預設值可不輸入
    SetRptAllocate
    
    '??? 取得報表的最小寬度
    GetRptMinWidth H3l$ & Space(1) & HDate$
    
' 一列以上Header的格式設定 =====================================================================
'??? 取得標題或資料的字串格式(參數二表傳回的格式型態 -- 1:標題格式 2:資料顯示的格式)
'??? Multi Line 時使用
'    ' 針對與預設值不同的欄位,重新設定報表抬頭欄位的對齊方式
'    ReDefineHeaderAlign
'    ' 設定第一列Header的Caption
'    ReDefineReportHeader
'    '取得第一列Header的Format
    'FC$ = GetRptFormatStr(tSpd_EXAR01, 3)
'    ' 設定第二列Header的Caption
'    ReDefineReportHeader
'    '取得第二列Header的Format
'    FC$ = GetRptFormatStr(tSpd_EXAR01, 3)
'    fd$ = GetRptFormatStr(tSpd_EXAR01, 2)
' ==============================================================================================
   
    '??? 取得標題或資料的字串格式(參數二表傳回的格式型態 -- 1:標題格式 2:資料顯示的格式)
    '   針對與預設值不同的欄位,重新設定報表抬頭欄位的對齊方式
    ReDefineHeaderAlign

    '??? 表頭為Single Line 時使用
    FC$ = GetRptFormatStr(tSpd_EXAR01, 1)
    fd$ = GetRptFormatStr(tSpd_EXAR01, 2)

    '??? 取得報表抬頭的格式
    H2$ = GetRptTitleFormat()
    
    '??? 取得報表表頭資料的格式
    H3l$ = PrintUse(H3l$, G_Pnl_A1601$ & G_G1 & G_A1601s$ & G_G1 & G_A1601e$)
    H4l$ = PrintUse(H4l$, G_Pnl_A1617$ & G_G1 & G_A1617s$ & G_G1 & G_A1617e$)
    H5l$ = PrintUse(H5l$, G_Pnl_A1609$ & G_G1 & G_A1609s$ & G_G1 & G_A1609e$)
    H3$ = GetRptHeaderFormat(H3l$, HDate$)
    H4$ = GetRptHeaderFormat(H4l$, HDate$)
    H5$ = GetRptHeaderFormat(H5l$, HDate$)
    B31$ = GetRptHeaderFormat("", B31$)

    '??? 取得報表Break欄位的格式
    B1$ = GetRptHeaderFormat(B11$)
   
    '??? 取得續下頁及印表人的格式
    N1$ = GetRptFootFormat(HPerson$)
    N2$ = PrintUse(GetRptLineFormat("~"), HPerson$)
    
    '??? 取得區隔列的格式
    B2$ = GetRptLineFormat("-")
    B3$ = GetRptLineFormat("~")
    H9$ = GetRptLineFormat("=")
End Sub

Sub SetReportCols()
    '========================================================================
    '*** 設定Q Screen中的Spd_Help vaSpread **********************************
    '??? 宣告Spread型態的Columns及Sorts的陣列個數,
    '    參數一 : Spread Type Name
    '    參數二 : vaSpread上的欄位總數
    '    參數三 : 是否允許User自訂排序欄位及其順序
    '========================================================================
    InitialCols tSpd_Help, 2, False
    
    '========================================================================
    '??? 設定vaSpread上的所有欄位及排序欄位至Spread Type中
    '    參數一 : Spread Type Name
    '    參數二 : 設定用來存取vaSpread上欄位的欄位名稱
    '    參數三 : Optional - 設定隱藏欄位(0:顯示  1:暫時隱藏,預設值  2:永久隱藏)
    '    參數四 : Optional - 設定程式預設排序欄位的順序
    '    參數五 : Optional - 設定程式預設排序欄位的方向(1:遞增,預設值  2:遞減)
    '    參數六 : Optional - 設定Break欄位的順序
    '    參數七 : Optional - 設定Break欄位是否與其他欄位顯示於同一列上(True,預設值 / False)
    '========================================================================
    AddReportCol tSpd_Help, "A0801", , 1
    AddReportCol tSpd_Help, "A0802", , 2
    
    '========================================================================
    '??? 抓取User自訂報表之欄位顯示順序及排序欄位
    '    參數一 : Spread Type Name
    '    參數二 : vaSpread所在的Form Name
    '    參數三 : vaSpread Name
    '========================================================================
    GetSpreadDefault tSpd_Help, "frm_EXAR01q", "Spd_Help"

    '========================================================================
    '*** 設定V Screen中的Spd_EXAR01 vaSpread *********************************
    '??? 宣告Spread型態的Columns及Sorts的陣列個數,
    '    參數一 : Spread Type Name
    '    參數二 : vaSpread上的欄位總數
    '    參數三 : 是否允許User自訂排序欄位及其順序
    '========================================================================
    InitialCols tSpd_EXAR01, 11, False
    
    '========================================================================
    '??? 設定vaSpread上的所有欄位及排序欄位至Spread Type中
    '    參數一 : Spread Type Name
    '    參數二 : 設定用來存取vaSpread上欄位的欄位名稱
    '    參數三 : Optional - 設定隱藏欄位(0:顯示  1:暫時隱藏,預設值  2:永久隱藏)
    '    參數四 : Optional - 設定程式預設排序欄位的順序
    '    參數五 : Optional - 設定程式預設排序欄位的方向(1:遞增,預設值  2:遞減)
    '    參數六 : Optional - 設定Break欄位的順序
    '    參數七 : Optional - 設定Break欄位是否與其他欄位顯示於同一列上(True,預設值 / False)
    '========================================================================
    AddReportCol tSpd_EXAR01, "A1617", , 1, , 1
    AddReportCol tSpd_EXAR01, "A1601", , 2
    AddReportCol tSpd_EXAR01, "A1602"
    AddReportCol tSpd_EXAR01, "A1614"
    AddReportCol tSpd_EXAR01, "A1605"
    AddReportCol tSpd_EXAR01, "A1606"
    AddReportCol tSpd_EXAR01, "A1620"
    AddReportCol tSpd_EXAR01, "A1621"
    AddReportCol tSpd_EXAR01, "A1643"
    AddReportCol tSpd_EXAR01, "credit"
    AddReportCol tSpd_EXAR01, "Flag", 2
    
    '========================================================================
    '??? 抓取User自訂報表之欄位顯示順序及排序欄位
    '    參數一 : Spread Type Name
    '    參數二 : vaSpread所在的Form Name
    '    參數三 : vaSpread Name
    '========================================================================
    GetSpreadDefault tSpd_EXAR01, "frm_EXAR01", "Spd_EXAR01"
End Sub

