Attribute VB_Name = "mod_ArthGUI"
Option Explicit

'Customer Request Value
Global G_PrintFontSize As String

'Keep User Logon Date and Time
Global Const G_SystemID = "ARTHGUI"             '系統名稱
Global Const G_ProductName = "Security"
Global Const G_Pnt2Printer = 1
Global Const G_Pnt2Screen = 2
Global Const G_Pnt2File = 3

Global g_Sec_Date As String                     'Run PG Start Date
Global g_Sec_Time As String                     'Run PG Start Time

'GL/2000 Dynaset Name
Global DB_ARTHGUI As Database
Global DB_ARTHGL As Database
Global DB_LOCAL As Database
Global DY_SINI As Recordset
Global DY_SINI1 As Recordset
Global DY_A01 As Recordset
Global DY_A011 As Recordset
Global DY_A02 As Recordset
Global DY_A021 As Recordset
Global DY_A03 As Recordset
Global DY_A031 As Recordset
Global DY_A04 As Recordset
Global DY_A041 As Recordset
Global DY_A05 As Recordset
Global DY_A051 As Recordset
Global DY_A06 As Recordset
Global DY_A061 As Recordset
Global DY_A07 As Recordset
Global DY_A071 As Recordset
Global DY_A08 As Recordset
Global DY_A081 As Recordset
Global DY_A09 As Recordset
Global DY_A10 As Recordset
Global DY_A101 As Recordset
Global DY_A11 As Recordset
Global DY_A12 As Recordset
Global DY_A13 As Recordset
Global DY_A14 As Recordset
Global DY_A15 As Recordset
Global DY_A151 As Recordset
Global DY_A16 As Recordset
Global DY_A161 As Recordset
Global DY_A17 As Recordset
Global DY_A171 As Recordset
Global DY_A18 As Recordset
Global DY_A181 As Recordset
Global DY_A19 As Recordset
Global DY_A191 As Recordset
Global DY_A20 As Recordset
Global DY_A21 As Recordset
Global DY_A211 As Recordset
Global DY_A22 As Recordset
Global DY_A221 As Recordset
Global DY_A23 As Recordset
Global DY_A231 As Recordset
Global DY_A24 As Recordset
Global DY_A241 As Recordset
Global DY_A28 As Recordset
Global DY_A281 As Recordset
Global DY_A31 As Recordset
Global DY_A41 As Recordset
Global DY_A45 As Recordset
Global DY_A49 As Recordset
Global DY_B20 As Recordset
Global DY_LB13 As Recordset
Global DY_INI1 As Recordset
Global TB_INI As Recordset
Global DY_M1 As Recordset
Global DY_M11 As Recordset
Global DY_M2 As Recordset

Function GetCaption(ByVal Section$, ByVal Topic$, ByVal Default$) As String
'自資料庫取得辭庫
Dim A_STR$

    A_STR$ = GetSIniStr(Section$, Topic$)
    If Trim(A_STR$) = "" Then A_STR$ = Trim(Default$)
    GetCaption = A_STR$
End Function

Function CvrString2Character(ByVal Str$) As String
'將中文字串中沖碼符號為 ' or | 轉成字元相加,適用於SQL指令(for Access 2.0)
'例如:李四生='李?&Chr$(124)&'生'
Dim A_Temp$, A_Temp2$, I%

    CvrString2Character = ""
    Str$ = Trim$(Str$)
    If Str$ = "" Then Exit Function

    A_Temp$ = "'"
    For I% = 1 To Len(Str$)
    A_Temp2 = Mid(Str$, I%, 1)
    If A_Temp2$ = Chr(39) Or A_Temp2$ = Chr(124) Then
       A_Temp$ = A_Temp$ & "'&" & "Chr$(" & Trim(Asc(A_Temp2)) & ")&'"
    Else
       A_Temp$ = A_Temp$ & A_Temp2
    End If
    Next I%
    If A_Temp2$ = Chr(39) Or A_Temp2$ = Chr(124) Then   '沖碼符號為 ' or |
       A_Temp$ = Left(A_Temp$, Len(A_Temp$) - 2)
    Else
       A_Temp$ = A_Temp$ & "'"
    End If
    CvrString2Character = A_Temp$
End Function

Function DaysCount(ByVal FrmDate$, ByVal EndDate$) As Currency
'計算兩個日期的差異天數
Dim fy&, fm&, fd&, ty&, tm&, td&, d&
Dim a@, B@, C&

    fy& = Val(FrmDate$) / 10000
    fm& = (Val(FrmDate$) Mod 10000) / 100
    fd& = Val(FrmDate$) Mod 100
    ty& = Val(EndDate$) / 10000
    tm& = (Val(EndDate$) Mod 10000) / 100
    td& = Val(EndDate$) Mod 100

    ' From Date Days Caculate
    a@ = fd&
    C& = 1
    Do While C& < fm&
       Select Case C&
         Case 1, 3, 5, 7, 8, 10, 12
              a@ = a@ + 31
         Case 4, 6, 9, 11
              a@ = a@ + 30
         Case 2
              If fy& Mod 4 = 0 Or fy& Mod 400 = 0 Then
                 a@ = a@ + 29
              Else
                 a@ = a@ + 28
              End If
       End Select
       C& = C& + 1
    Loop
    d& = fy& / 4 - fy& / 100
    a@ = a@ + fy& * 365 + d&
    
    ' End Date Days Caculate
    B@ = td&
    C& = 1
    Do While C& < tm&
       Select Case C&
         Case 1, 3, 5, 7, 8, 10, 12
              B@ = B@ + 31
         Case 4, 6, 9, 11
              B@ = B@ + 30
         Case 2
              If ty& Mod 4 = 0 Or ty& Mod 400 = 0 Then
                 B@ = B@ + 29
              Else
                 B@ = B@ + 28
              End If
       End Select
       C& = C& + 1
    Loop
    d& = ty& / 4 - ty& / 100
    B@ = B@ + ty& * 365 + d&
    '
    DaysCount = B@ - a@
End Function

Function DBCSStrCheck(ByVal Temp$) As String
'傳回沖碼符號為 ' or | 字元前的字串(for Access 2.0)
Dim I%
    
    DBCSStrCheck = ""
    For I% = 1 To Len(Temp$)
        If Mid(Temp$, I%, 1) = Chr(39) Then GoTo DBCSStrCheckA
        If Mid(Temp$, I%, 1) = Chr(124) Then GoTo DBCSStrCheckA
    Next I%
    
DBCSStrCheckA:
    If I% > 1 Then DBCSStrCheck = Mid(Temp$, 1, I% - 1)
End Function

Function DelayDate(ByVal DateType$, ByVal Date1$, ByVal Date2$) As Long
'兩日期相減傳回天數 (Date2$ - Date1$)
'DateType$="1" ---> 國曆
'DateType$="2" ---> 西曆
'DateType$="3" ---> 西曆(西元年長短格式並存)
On Local Error GoTo MyError
Dim A_Date1, A_Date2

    Date1$ = DateIn(Date1$)
    Date2$ = DateIn(Date2$)
    A_Date1 = DateSerial(Val(Left$(Date1$, 4)), Val(Mid$(Date1$, 5, 2)), Val(Right$(Date1$, 2)))
    A_Date2 = DateSerial(Val(Left$(Date2$, 4)), Val(Mid$(Date2$, 5, 2)), Val(Right$(Date2$, 2)))
    '
    Select Case DateType$
      Case "1", "2", "3"
           DelayDate = DateDiff("d", A_Date1, A_Date2)
      Case Else
           DelayDate = 0
    End Select
    Exit Function

MyError:
   DelayDate = 0
   Exit Function
End Function


Sub GetSystemINIString()
'自GUI.INI取得系統設定Keep至Global變數

    Screen.MousePointer = HOURGLASS
   'Pick Server Path from Local INI
    G_INI_SerPath = GetIniStr("FilePath", "serverpath", "GUI.INI")
   'Pick Local INI DataPath String (GUI.MDB) & Connect String
    G_DB_PATH1 = GetIniStr("DBPath", "Path1", "GUI.INI")
    G_ConnectMethod1 = GetIniStr("DBPath", "Connect1", "GUI.INI")
   'Pick Local INI DataPath String ARTHGL.MDB) & Connect String
    G_DB_PATH2 = GetIniStr("DBPath", "Path2", "GUI.INI")
    G_ConnectMethod2 = GetIniStr("DBPath", "Connect2", "GUI.INI")
   'Pick Local INI DataPath String LGUI.MDB) & Connect String
    G_DB_PATH3 = GetIniStr("DBPath", "Path3", "GUI.INI")
    G_ConnectMethod3 = GetIniStr("DBPath", "Connect3", "GUI.INI")
   'pick User Name
    StrCut G_CmdStr3, "/", G_DUserId, G_UserName
    StrCut G_UserName, "/", G_UserName, G_UserGroup
    If G_UserName = "" Then
       G_UserName = GetIniStr("User_ID", "Name", "GUI.INI")
       If G_UserName = "" Then G_UserName = " "
    End If
    If G_DUserId = "" Then
       G_DUserId = GetIniStr("User_ID", "UID", "GUI.INI")
       If G_DUserId = "" Then G_DUserId = " "
    End If
    If G_UserGroup = "" Then
       G_UserGroup = GetIniStr("User_ID", "Group", "GUI.INI")
       If G_UserGroup = "" Then G_UserGroup = " "
    End If
   'Pick Program Path
    G_Program_Path = GetIniStr("FilePath", "ProgramPath", "GUI.INI")
   'Pick Help Path
    G_Help_Path = GetIniStr("FilePath", "HelpPath", "GUI.INI")
   'Pick System Path
    G_System_Path = GetIniStr("FilePath", "SYSTEMPATH", "GUI.INI")
   'Pick Report Path
    G_Report_Path = G_System_Path + "TMP\"
   '將資料庫連接字串中的密碼解密
    DecodingConnectStr ("GUI.INI")
    Screen.MousePointer = Default
End Sub

Sub OpenDB()
'開啟系統資料庫
On Local Error Resume Next

    Screen.MousePointer = HOURGLASS
    
    '*** Add For Vista 96/6/25 By Jennifer
    EnableVistaClient
    
    'Open GUI DataBase
    If Trim$(G_ConnectMethod1) = "" Then   'Access DataBase
       If G_WorkSpace1 Is Nothing Then
          Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
          GetEngine.Workspaces.Append G_WorkSpace1
       End If
       Set DB_ARTHGUI = G_WorkSpace1.OpenDatabase(G_DB_PATH1, False, False, G_ConnectMethod1)
    Else
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;", "ORACLE;"
              If G_WorkSpace1 Is Nothing Then
                 Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
                 GetEngine.Workspaces.Append G_WorkSpace1
              End If
              Set DB_ARTHGUI = G_WorkSpace1.OpenDatabase(G_DB_PATH1, False, False, G_ConnectMethod1)
         Case "DB2;"
              If G_WorkSpace2 Is Nothing Then
                 Set G_WorkSpace2 = GetEngine.CreateWorkspace("WK2", "admin", "", dbUseODBC)
                 GetEngine.Workspaces.Append G_WorkSpace2
              End If
              Set DB_ARTHGUI = G_WorkSpace2.OpenDatabase(G_DB_PATH1, , , G_ConnectMethod1)
       End Select
    End If
    If Err Then
       If Trim$(G_ConnectMethod1) = "" Then   'Access DataBase
          If Err = 3043 Then
             Err = 0
             DB_ARTHGUI.Close
             If G_WorkSpace1 Is Nothing Then
                Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
                GetEngine.Workspaces.Append G_WorkSpace1
             End If
             Set DB_ARTHGUI = G_WorkSpace1.OpenDatabase(G_DB_PATH1, False, False, G_ConnectMethod1)
          ElseIf Err = 3049 Then
             Err = 0
             GetEngine.RepairDatabase G_DB_PATH1
             If G_WorkSpace1 Is Nothing Then
                Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
                GetEngine.Workspaces.Append G_WorkSpace1
             End If
             Set DB_ARTHGUI = G_WorkSpace1.OpenDatabase(G_DB_PATH1, False, False, G_ConnectMethod1)
          End If
       End If
    End If
    If Err Then
       MsgBox Error(Err), MB_ICONEXCLAMATION, App.Title
       End
    End If
    DB_ARTHGUI.QueryTimeout = 0
    
    'Open ARTHGL DataBase
    If Not G_DBNotOpen2 Then
       If Trim$(G_ConnectMethod2) = "" Then   'Access DataBase
          If G_WorkSpace1 Is Nothing Then
             Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
             GetEngine.Workspaces.Append G_WorkSpace1
          End If
          Set DB_ARTHGL = G_WorkSpace1.OpenDatabase(G_DB_PATH2, False, False, G_ConnectMethod2)
       Else
          Select Case UCase$(Mid$(G_ConnectMethod2, InStr(1, G_ConnectMethod2, "DBTYPE=", 1) + 7))
            Case "SQL;", "ORACLE;"
                 If G_WorkSpace1 Is Nothing Then
                    Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
                    GetEngine.Workspaces.Append G_WorkSpace1
                 End If
                 Set DB_ARTHGL = G_WorkSpace1.OpenDatabase(G_DB_PATH2, False, False, G_ConnectMethod2)
            Case "DB2;"
                 If G_WorkSpace2 Is Nothing Then
                    Set G_WorkSpace2 = GetEngine.CreateWorkspace("WK2", "admin", "", dbUseODBC)
                    GetEngine.Workspaces.Append G_WorkSpace2
                 End If
                 Set DB_ARTHGL = G_WorkSpace2.OpenDatabase(G_DB_PATH2, , , G_ConnectMethod2)
          End Select
       End If
       If Err Then
          If Trim$(G_ConnectMethod2) = "" Then   'Access DataBase
             If Err = 3043 Then
                Err = 0
                DB_ARTHGL.Close
                If G_WorkSpace1 Is Nothing Then
                   Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
                   GetEngine.Workspaces.Append G_WorkSpace1
                End If
                Set DB_ARTHGL = G_WorkSpace1.OpenDatabase(G_DB_PATH2, False, False, G_ConnectMethod2)
             ElseIf Err = 3049 Then
                Err = 0
                GetEngine.RepairDatabase G_DB_PATH2
                If G_WorkSpace1 Is Nothing Then
                   Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
                   GetEngine.Workspaces.Append G_WorkSpace1
                End If
                Set DB_ARTHGL = G_WorkSpace1.OpenDatabase(G_DB_PATH2, False, False, G_ConnectMethod2)
             End If
          End If
       End If
       If Err Then
          MsgBox Error(Err), MB_ICONEXCLAMATION, App.Title
          End
       End If
       DB_ARTHGL.QueryTimeout = 0
    End If
    
    'Open LGUI DataBase
    If G_WorkSpace1 Is Nothing Then
       Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
       GetEngine.Workspaces.Append G_WorkSpace1
    End If
    Set DB_LOCAL = G_WorkSpace1.OpenDatabase(G_DB_PATH3, False, False, G_ConnectMethod3)
    If Err = 3043 Then
       Err = 0
       DB_LOCAL.Close
       If G_WorkSpace1 Is Nothing Then
          Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
          GetEngine.Workspaces.Append G_WorkSpace1
       End If
       Set DB_LOCAL = G_WorkSpace1.OpenDatabase(G_DB_PATH3, False, False, G_ConnectMethod3)
    ElseIf Err = 3049 Then
       Err = 0
       GetEngine.RepairDatabase G_DB_PATH3
       If G_WorkSpace1 Is Nothing Then
          Set G_WorkSpace1 = GetEngine.CreateWorkspace("WK1", "admin", "", dbUseJet)
          GetEngine.Workspaces.Append G_WorkSpace1
       End If
       Set DB_LOCAL = G_WorkSpace1.OpenDatabase(G_DB_PATH3, False, False, G_ConnectMethod3)
    End If
    If Err Then
       MsgBox Error(Err), MB_ICONEXCLAMATION, App.Title
       End
    End If
    
    'Open Table
    If Trim(DB_LOCAL.Connect) = "" Then
        Set TB_INI = DB_LOCAL.OpenRecordset("INI", dbOpenTable)
        TB_INI.index = "INI"
    End If
    
    'Pick User Name
    If G_UserName = "" Then G_UserName = " "
    If G_DUserId = "" Then G_DUserId = " "
    If G_UserGroup = "" Then G_UserGroup = " "
    
    'Get Program Name
    'GetProgramName
    
    'Write Log File
    'WriteJournalLog DB_ARTHGUI, G_Program_Start, UCase$(App.EXEName), G_ProgramName
            
    If G_SecurityPgm = False Then
        Select Case UCase$(App.EXEName)
          Case "SECMENU", "SECMENU1", "SECMENUE"
          
          Case "MCF10", "SCP01", "SCR01", "SCR02", "SCR03"
          
          Case "MCF10A", "SCR02A", "SCR03A", "MCF12", "MCF15"
          
          Case "NEWSEC", "MCF16", "MCF17", "UNLICENS"
          
          Case Else
               If Not (UCase$(App.EXEName) = "MCF08" And Left$(G_CmdStr2, 1) = "1") Then
                  WriteJournalLog DB_ARTHGUI, G_Program_Start, UCase$(App.EXEName), ""
               End If
        End Select
    Else
        WriteJournalLog_Security DB_ARTHGUI, G_Program_Start, UCase$(App.EXEName), ""
    End If
    '
    Screen.MousePointer = Default
End Sub

Sub PickWINPath()
'取得OS路徑
Const MAX_PATH = 260

    G_WinDir = Space(MAX_PATH)
    If GetWindowsDirectory(G_WinDir, MAX_PATH) > 0 Then
        G_WinDir = StripTerminator(Trim$(G_WinDir))
    Else
        G_WinDir = "C:\Windows"
    End If
End Sub

Function UpdateIniValue(ByVal Section$, ByVal Topic$, ByVal UpdateValue$, ByVal IniValue$) As Boolean
'Update INI File中的Value

    If OSWritePrivateProfileString%(Section$, Topic$, UpdateValue$, IniValue$) = False Then
       UpdateIniValue = False
    Else
       UpdateIniValue = True
    End If
End Function


Function GetGLRptFont(ByVal ft$) As Double
'自GUI.INI取得報表字型大小
Dim ff$

    ff$ = Space(6)
    If OSGetPrivateProfileString%("FontType", ft$, "", ff$, 6, "GUI.INI") Then
       GetGLRptFont = Val(ff$)
    Else
       GetGLRptFont = "9.6"
    End If
End Function

Function GetGLRptPageLine(ByVal ft$) As Double
'自GUI.INI取得報表一頁可列印列數
Dim ff$

    ff$ = Space(5)
    If Not OSGetPrivateProfileString%("PageSize", ft$, "", ff$, 5, "GUI.INI") Then
       ff$ = "66"
    End If
    GetGLRptPageLine = Val(ff$)
End Function

Function GetGLRptPageSize(ByVal ft$) As Double
'自GUI.INI取得報表一頁允許可列印的列數
Dim ff$

    ff$ = Space(5)
    If Not OSGetPrivateProfileString%("Overflow", ft$, "", ff$, 5, "GUI.INI") Then
       ff$ = "66"
    End If
    GetGLRptPageSize = Val(ff$)
End Function

Sub GetSvrDefault()
'取得系統預設值Keep至Global變數

    Screen.MousePointer = HOURGLASS
    'Pick DateFlag
    G_DateFlag = Val(GetSvrIniStr("Customer", "CalanDate"))
    'Pick LeadYear
    G_LeadYear$ = ""
    If G_DateFlag = 2 Then
        G_LeadYear$ = Trim(GetSvrIniStr("Customer", "LeadYear"))
        If G_LeadYear$ = "" Then G_LeadYear$ = Left$(CStr(Year(Now)), 2)
    End If
    'Pick Picture File
    G_PICTURE_NAME = GetSvrIniStr("Customer", "PictureName")
    Screen.MousePointer = Default
End Sub

Function GetSvrIniStr(ByVal A_Section$, ByVal A_Topic$) As String
'取得資料庫ARTHGUI中,SINI-TABLE中的TOPICVALUE值
Dim A_Sql$

    GetSvrIniStr = " "
    A_Sql$ = "Select TOPICVALUE From SINI Where"
    A_Sql$ = A_Sql$ & " SECTION='" & A_Section$ & "'"
    A_Sql$ = A_Sql$ & " AND TOPIC='" & A_Topic$ & "'"
    A_Sql$ = A_Sql$ & " Order by SECTION,TOPIC"
    CreateDynasetODBC DB_ARTHGUI, DY_SINI, A_Sql$, "DY_SINI", True
    If Not (DY_SINI.BOF And DY_SINI.EOF) Then
       GetSvrIniStr = DY_SINI.Fields("TOPICVALUE") & ""
    End If
End Function








