Attribute VB_Name = "mod_TSR01"
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
Global G_Form_TSR01$
Global G_Form_TSR01q$

'定義各欄位標題文字

Global G_Pnl_A1501$
Global G_Pnl_A1502$
Global G_Pnl_A15023$
Global G_Pnl_A1504$
Global G_Pnl_A1505$
Global G_Pnl_A1507$
Global G_Pnl_A1508$
Global G_Pnl_A1508_Sum$
Global G_Pnl_A1510$
Global G_Pnl_A1512$

Global G_Pnl_Dash$
Global G_Pnl_PrtType$
Global G_Pnl_Printer$
Global G_Pnl_Screen$
Global G_Pnl_File$
Global G_Pnl_Excel$

'Def 程式共用變數
Global G_SlipAttrib_1$
Global G_SlipAttrib_2$
Global G_AccountUse_1$
Global G_AccountUse_2$
Global G_AccountUse_3$
Global G_SlipType_1$
Global G_SlipType_2$
Global G_SlipType_3$
Global G_SlipType_4$
Global G_SlipType_5$
Global G_SlipType_6$
Global G_SlipType_7$
Global G_SlipType_8$

Global G_PathNotFound$
Global G_Report_Heading$
Global G_A1502s$
Global G_A1502e$
Global G_A1501$
Global G_A1501n$
Global G_A1508_Sum$

''SAMPLE
'Global G_BB#
'Global G_CC!

'Def 報表格式
'Global Const H0$ = "....5...10....5...20....5...30....5...40....5...50....5...60....5...70....5...80....5...90....5..100....5..110....5..120....5..130....5..140....5..150....5..160....5..170....5..180....5..190....5..."
'Global Const H1$ = " "
'Global Const H2$ = "  <TSR01>                             科目列印
'Global Const H3$ = "  "
'Global Const H4$ = "  "
'Global Const H5$ = "  系統代號:"
'Global Const H6$ = "  程式代碼      :                                                                                                                頁次：1"
'Global Const H7$ = "  群組代號      :                                                                                                                日期：89/02/15"
'Global Const H8$ = "  User ID       :            -                                                                                                   時間：11:44:47"
'Global Const H9$ = "  ============================================================================================================================================="
'Global Const HA$ = "  日期       時間     登錄  使用者       程式名稱                                 系統代號   備註                                              "
'Global Const N1$ = "                                                                 ... 續 下 頁 ...                                                              "

Global Const H0$ = "....5...10....5...20....5...30....5...40....5...50....5...60....5...70....5...80....5...90....5..100....5..110....5..120....5..130....5..140....5..150....5..160....5..170....5..180....5..190....5..."
Global Const H1$ = " "
Global Const H2$ = "  <TSR01> ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
Global Const H3$ = "                                                                                                 ######## : #####"
Global Const H4$ = "  ###########: ## ############                                                              ######## : ##########"
Global Const H5$ = "  ###########: ~~~~~~ - ######                                                              ######## : ##########"
Global Const H6$ = "  ==============================================================================================================="
Global Const HA$ = "  ######## ######################################## #### #################### ######## ########   ~~~~~~~~~~~~~~~"
Global Const HB$ = "  ######## ######################################## #### #################### ######## ########   ~~~~~~~~~~~~~~~"
Global Const N1$ = "  ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
Global Const SU$ = "                                                                                 ~~~~~~~~:   ~~~~~~~~~~~~~~~~~~~~"

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
    frm_TSR01q.Show        ' 首頁畫面顯示
    Screen.MousePointer = Default
End Sub

Sub GetPanelCaption()
'取FORM標題文字
    G_Form_TSR01$ = GetSIniStr("FormTitle", "TSR01")
    G_Form_TSR01q$ = GetSIniStr("FormTitle", "TSR01q")
    
'取欄位標題文字
    G_Pnl_A1501$ = GetSIniStr("TSR01", "A1501")
    G_Pnl_A1502$ = GetSIniStr("TSR01", "A1502")
    G_Pnl_A15023$ = GetSIniStr("TSR01", "acctcode")
    G_Pnl_A1504$ = GetSIniStr("TSR01", "A1504")
    G_Pnl_A1505$ = GetSIniStr("TSR01", "A1505")
    G_Pnl_A1507$ = GetSIniStr("TSR01", "A1507")
    G_Pnl_A1508$ = GetSIniStr("TSR01", "A1508")
    G_Pnl_A1508_Sum$ = GetSIniStr("TSR01", "A1508_Sum")
    G_Pnl_A1510$ = GetSIniStr("TSR01", "A1510")
    G_Pnl_A1512$ = GetSIniStr("TSR01", "A1512")
    
    G_Pnl_Dash$ = GetSIniStr("PanelDescpt", "dash")
    G_Pnl_PrtType$ = GetSIniStr("PanelDescpt", "printtype")
    G_Pnl_Printer$ = GetSIniStr("PanelDescpt", "printer")
    G_Pnl_Screen$ = GetSIniStr("PanelDescpt", "screen")
    G_Pnl_File$ = GetSIniStr("PanelDescpt", "file")
    G_Pnl_Excel$ = GetSIniStr("PanelDescpt", "excel")

'取列印替代文字
    G_SlipAttrib_1$ = Reference_INI("SlipAttrib", "1")
    G_SlipAttrib_2$ = Reference_INI("SlipAttrib", "2")
    G_AccountUse_1$ = Reference_INI("AccountUse", "1")
    G_AccountUse_2$ = Reference_INI("AccountUse", "2")
    G_AccountUse_3$ = Reference_INI("AccountUse", "3")
    G_SlipType_1$ = Reference_INI("SlipType", "1")
    G_SlipType_2$ = Reference_INI("SlipType", "2")
    G_SlipType_3$ = Reference_INI("SlipType", "3")
    G_SlipType_4$ = Reference_INI("SlipType", "4")
    G_SlipType_5$ = Reference_INI("SlipType", "5")
    G_SlipType_6$ = Reference_INI("SlipType", "6")
    G_SlipType_7$ = Reference_INI("SlipType", "7")
    G_SlipType_8$ = Reference_INI("SlipType", "8")
    
'取其他變數內含值
    G_PathNotFound$ = GetSIniStr("PgmMsg", "path_not_found")
    G_Report_Heading$ = GetSIniStr("ReportHeading", "TSR01")
End Sub

Sub PageCheck(Tmp As Object)
    If G_PrintSelect = G_Print2Excel Then Exit Sub
    If G_PrintSelect = G_Print2Screen Then Exit Sub
'跳頁處理
    If G_LineNo > G_OverFlow Then
        If G_PageNo > 0 Then
           PrintOut3 Tmp, H1$, ""
           PrintOut3 Tmp, H1$, ""
           PrintOut3 Tmp, N1$, G_Print_NextPage
           If G_PrintSelect = G_Print2Printer Then
              Printer.NewPage
           Else
              Print #1, G_G1
           End If
        End If
        ReportHeader Tmp
    End If
End Sub

Sub PrePare_Data(Frm As Form, Prb As ProgressBar, Tmp As Object, A_Exit%)
On Local Error GoTo MY_Error

'設定ProgressBar最大值
    DY_A15.MoveLast
    Prb.MAX = DY_A15.RecordCount
    DY_A15.MoveFirst
    
'開啟文字檔
    If G_PrintSelect = G_Print2File Then
       Open G_OutFile For Output As #1
    ElseIf G_PrintSelect = G_Print2Excel Then
        If Not OpenExcelFile(G_OutFile) Then
            Frm!Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
            Exit Sub
        End If
        Set_Excel_Property
    End If
    
'設定報表字體,字型及印表機設定
    If Not ReportSet() Then Frm!Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE): Exit Sub

'資料列印處理
    PrintSub Prb, Tmp, A_Exit%
        
'已觸發結束鍵時, 跳出此程序
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
    G_FontSize = GetGLRptFont("FontF")
    G_PageSize = GetGLRptPageLine("FontF")
    G_OverFlow = GetGLRptPageSize("FontF")
    
'若選擇印表機列印,則顯示印表設定視窗
    If G_PrintSelect = G_Print2Printer Then
       G_RptNeedWidth = Len(H6$)
       Printer.FontName = G_FontName
       Printer.FontSize = G_FontSize
       rptset.Show MODAL
       G_OverFlow = G_PageSize - 6
       If G_PageSize <= 0 Then G_PageSize = 60
       If G_OverFlow <= 0 Then G_OverFlow = 60
       ReportSet = G_RptSet
    End If
End Function

Sub ReportHeader(Tmp As Object)
Dim A_H2$
Dim A_H3$
Dim A_H4$
Dim A_H5$
Dim A_HA$

    If G_PrintSelect = G_Print2Screen Then Exit Sub

'列印報表表頭
    G_LineNo = 8                'Headers: H1~H5 + 2xH6 + HA
    G_PageNo = G_PageNo + 1     'Page# counting
    '
    If G_PrintSelect = G_Print2Excel Then   'Print to Excel
       A_H2$ = String(3, G_G1) & G_Report_Heading$
       A_H3$ = ""
       A_H4$ = G_Pnl_A1501$ & " : " & G_A1501 & " " & G_A1501n$
       A_H4$ = A_H4$ & String(5, G_G1) & G_Print_Date & " : " & GetCurrentDay(1)
       A_H5$ = G_Pnl_A15023$ & " : " & G_A1502s$ & " - " & G_A1502e$
       A_H5$ = A_H5$ & String(5, G_G1) & G_Print_Time & " : " & Format(Now, "HH:MM:SS")
       
    Else
       A_H2$ = G_Report_Heading$
       A_H3$ = G_Print_Page & G_G1 & Format(G_PageNo, "###0")
       A_H4$ = G_Pnl_A1501$ & G_G1 & G_A1501 & G_G1 & G_A1501n$ & G_G1 & G_Print_Date & G_G1 & GetCurrentDay(1)
       A_H5$ = G_Pnl_A15023$ & G_G1 & G_A1502s$ & G_G1 & G_A1502e$ & G_G1 & G_Print_Time & G_G1 & Format(Now, "HH:MM:SS")
       
    End If
    
    A_HA$ = G_Pnl_A1502$ & G_G1 & G_Pnl_A1505$ & G_G1 & G_Pnl_A1504$
    A_HA$ = A_HA$ & G_G1 & G_Pnl_A1507 & G_G1 & G_Pnl_A1510$
    A_HA$ = A_HA$ & G_G1 & G_Pnl_A1512$ & G_G1 & G_Pnl_A1508$
    '
    PrintOut3 Tmp, H1$, "", 1         '最後一個參數: 列印的行數
    PrintOut3 Tmp, H2$, A_H2$, 2
    PrintOut3 Tmp, H3$, A_H3$, 3
    PrintOut3 Tmp, H4$, A_H4$, 4
    PrintOut3 Tmp, H5$, A_H5$, 5
    PrintOut3 Tmp, H6$, "", 6
    PrintOut3 Tmp, HA$, A_HA$, 7
    PrintOut3 Tmp, H6$, "", -1
    '
    If G_PrintSelect = G_Print2Excel Then G_XlsHRows% = 10
End Sub

Sub PrintSub(Prb As ProgressBar, Tmp As Object, A_Exit%)
'Printing Procedure (Header, Body, Buttom)
Dim A_PrtStr$                                   'Output Str
Dim A_A1502$, A_A1505$, A_A1504$, A_A1507$      'Column Value
Dim A_A1510$, A_A1512$, A_A1508$
Dim A_A1508_Sum#                                'Sum of A1508
Dim A_Row#                                      'Print Line #

'Initialize
    Prb.Visible = True
    Prb.Value = 0
    G_A1508_Sum$ = ""
    A_A1508_Sum# = 0
    
'Print Header
    ReportHeader Tmp
    
'Print Body
    A_Row# = 0
    Do While Not DY_A15.EOF And Not A_Exit%
       A_Row# = A_Row# + 1
        'col1
        A_A1502$ = Trim$(DY_A15.Fields("A1502") & "")
        A_A1502$ = A_A1502$ & Trim$(DY_A15.Fields("A1503") & "")
        'col2
        A_A1505$ = Trim$(DY_A15.Fields("A1505") & "")
        'col3
        Select Case Trim$(DY_A15.Fields("A1504") & "")
            Case "1"
                A_A1504$ = G_SlipAttrib_1$
            Case "2"
                A_A1504$ = G_SlipAttrib_2$
        End Select
        'col4
        A_A1507 = Trim$(DY_A15.Fields("A1302") & "")
        'col5
        Select Case Trim$(DY_A15.Fields("A1510") & "")
            Case "1"
                A_A1510$ = G_AccountUse_1$
            Case "2"
                A_A1510$ = G_AccountUse_2$
            Case "3"
                A_A1510$ = G_AccountUse_3$
        End Select
        'col6
        Select Case Trim$(DY_A15.Fields("A1512") & "")
            Case "1"
                A_A1512$ = G_SlipType_1$
            Case "2"
                A_A1512$ = G_SlipType_2$
            Case "3"
                A_A1512$ = G_SlipType_3$
            Case "4"
                A_A1512$ = G_SlipType_4$
            Case "5"
                A_A1512$ = G_SlipType_5$
            Case "6"
                A_A1512$ = G_SlipType_6$
            Case "7"
                A_A1512$ = G_SlipType_7$
            Case "8"
                A_A1512$ = G_SlipType_8$
        End Select
        'col7
        A_A1508$ = Trim$(DY_A15.Fields("A1508") & "")
        'sum up col7
        A_A1508_Sum# = A_A1508_Sum# + CDbl(A_A1508$)

'串接列印列資料至變數
        A_PrtStr$ = A_A1502$ & G_G1 & A_A1505$ & G_G1 & A_A1504$
        A_PrtStr$ = A_PrtStr$ & G_G1 & A_A1507$ & G_G1 & A_A1510$
        A_PrtStr$ = A_PrtStr$ & G_G1 & A_A1512$ & G_G1 & A_A1508$
        

'累計目前列印行數,若超過一頁則跳頁
       G_LineNo = G_LineNo + 1
       PageCheck Tmp
       
'將字串傳給PrintOut3處理列印動作
       PrintOut3 Tmp, HB$, A_PrtStr$, A_Row#
       '
       If A_Exit% Then Exit Do
       Prb.Value = A_Row#
       DoEvents
       DY_A15.MoveNext
    Loop
    
'已觸發結束鍵時, 跳出此程序
    If A_Exit% Then Exit Sub
    
'處理資料列印完成後的結束動作
    G_A1508_Sum$ = Format(A_A1508_Sum#, "#,##0.00")
    PrintBottom Prb, Tmp, A_Row#
End Sub

Sub PrintBottom(Prb As ProgressBar, Tmp As Control, ByVal A_EndRow#)
'列印報表表尾
Dim A_SU$
'Prepare Information
    A_EndRow# = A_EndRow# + 1
    If G_PrintSelect = G_Print2Excel Then   'Print to Excel
        A_SU$ = String(5, G_G1) & G_Pnl_A1508_Sum$ & G_A1508_Sum$
    Else
        A_SU$ = G_Pnl_A1508_Sum$ & G_G1 & G_A1508_Sum$
    End If
    
    
'列印表尾
    PrintOut3 Tmp, H6$, "", -1
    PrintOut3 Tmp, SU$, A_SU$, A_EndRow#
    
'處理Excel列印的文字剖析動作
    ProcessExcelText2Column A_EndRow# + G_XlsHRows%
    
'處理各種列印方式之結束動作
    PrintEnd3 Tmp
    Prb.Visible = False
End Sub

Sub Set_Excel_Property()
'設定Excel 起始欄位值
    G_XlsStartCol% = 1
    
'將Excel Title Rows歸零
    G_XlsHRows% = 0
    
'設定欄位的資料型態
    SetColumnFormat "A", SS_CELL_TYPE_EDIT
'    SetColumnFormat "H", SS_CELL_TYPE_FLOAT, "#,##0.00"
End Sub

Sub ProcessExcelText2Column(ByVal A_EndRow#)
'處理Excel列印的文字剖析動作
Dim A_FldType(6, 1)

    If G_PrintSelect <> G_Print2Excel Then Exit Sub
    
'??? 產生Excel文字剖析須要的欄位型態陣列
'參數二 : 欄位行號
'參數三 : 欄位資料型態 (G_Data_Date:日期型態 G_Data_String:文字型態
'                       G_Data_Numeric:數值型態)
    AddXlsFldDataType A_FldType, 1, G_Data_String
    AddXlsFldDataType A_FldType, 2, G_Data_String
    AddXlsFldDataType A_FldType, 3, G_Data_String
    AddXlsFldDataType A_FldType, 4, G_Data_String
    AddXlsFldDataType A_FldType, 5, G_Data_String
    AddXlsFldDataType A_FldType, 6, G_Data_String
    AddXlsFldDataType A_FldType, 7, G_Data_String
    
'??? 將起始欄位中的資料,以G_G1字元將資料切割成多個欄位值
    SetExcelTextToColumns G_XlsStartCol%, 1, A_EndRow#, A_FldType
End Sub


Private Function Reference_INI(ByVal A_Section$, ByVal A_Topic$) As String
On Local Error GoTo MyError
Dim A_Sql$

    Reference_INI = ""
    A_Sql$ = "Select TOPICVALUE From SINI"
    A_Sql$ = A_Sql$ & " where SECTION='" & A_Section$ & "'"
    A_Sql$ = A_Sql$ & " and TOPIC='type" & A_Topic$ & "'"
    A_Sql$ = A_Sql$ & " order by SECTION,TOPIC"
    CreateDynasetODBC DB_ARTHGUI, DY_INI, A_Sql$, "DY_INI", True
    If Not (DY_INI.BOF And DY_INI.EOF) Then
       Reference_INI = Trim$(DY_INI.Fields("TOPICVALUE") & "")
    End If
    Exit Function
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function
