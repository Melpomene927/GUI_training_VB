Attribute VB_Name = "mod_Report"
Option Explicit

'�w�qExcel�ܼ�
Global G_ExcelWkb As Excel.Application
Global G_XLSRptName As String
Global G_XlsHRows%      'Excel �������Y�C��
Global G_XlsStartCol%   '���w�C�L���_�l���

'*** Add for New Report Pattern 2001/11/14 ***
Global G_ExcelIndex#    'Keep�ثe�@�ΦC
Global G_ExcelMaxCols%  '�]�wExcel �����̤j����

'*** Add New Variable at 93/3/16 ***
Global G_HaveDataPrint% '�P�_����O�_�w����ƦC�L

'*** Add New Variable at 93/4/1 ***
'Global G_WordDoc As Word.Application
'Global G_DocSelection As Word.Selection
Global G_WordDoc As Object
Global G_DocSelection As Object
Global G_DocRptName As String
Global G_DocFontSize() As String

'==================================================================================================================
'�]�w�w�]�L��� 93/10/1 (Start)
'==================================================================================================================
Public G_PrinterName As String
Public G_SetDefaultPrinter As Integer

Public Const HWND_BROADCAST = &HFFFF
Public Const WM_WININICHANGE = &H1A

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

' constants for DesiredAccess member of PRINTER_DEFAULTS
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

' constant that goes into PRINTER_INFO_5 Attributes member
' to set it as default
Public Const PRINTER_ATTRIBUTE_DEFAULT = 4

' Constant for OSVERSIONINFO.dwPlatformId
Public Const VER_PLATFORM_WIN32_WINDOWS = 1

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type DEVMODE
     dmDeviceName As String * CCHDEVICENAME
     dmSpecVersion As Integer
     dmDriverVersion As Integer
     dmSize As Integer
     dmDriverExtra As Integer
     dmFields As Long
     dmOrientation As Integer
     dmPaperSize As Integer
     dmPaperLength As Integer
     dmPaperWidth As Integer
     dmScale As Integer
     dmCopies As Integer
     dmDefaultSource As Integer
     dmPrintQuality As Integer
     dmColor As Integer
     dmDuplex As Integer
     dmYResolution As Integer
     dmTTOption As Integer
     dmCollate As Integer
     dmFormName As String * CCHFORMNAME
     dmLogPixels As Integer
     dmBitsPerPel As Long
     dmPelsWidth As Long
     dmPelsHeight As Long
     dmDisplayFlags As Long
     dmDisplayFrequency As Long
     dmICMMethod As Long        ' // Windows 95 only
     dmICMIntent As Long        ' // Windows 95 only
     dmMediaType As Long        ' // Windows 95 only
     dmDitherType As Long       ' // Windows 95 only
     dmReserved1 As Long        ' // Windows 95 only
     dmReserved2 As Long        ' // Windows 95 only
End Type

Public Type PRINTER_INFO_5
     pPrinterName As String
     pPortName As String
     Attributes As Long
     DeviceNotSelectedTimeout As Long
     TransmissionRetryTimeout As Long
End Type

Public Type PRINTER_DEFAULTS
     pDatatype As Long
     pDevMode As Long
     DesiredAccess As Long
End Type

Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As String) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Public Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Any) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
'==================================================================================================================
'�]�w�w�]�L��� 93/10/1 (End)
'==================================================================================================================

Function GetRptLineFormat(ByVal Char$) As String
'�^�Ǧr��榡,�A�Ω���U���ΰϹj�C���榡
Dim A_Length%, A_STR$
    
    A_Length% = G_RptNeedWidth - Len(G_LineLeft) - Len(G_LineRight)
    A_STR$ = G_LineLeft & String(A_Length%, Char$) & G_LineRight
    GetRptLineFormat = A_STR$
End Function

Sub SetExcelAllocate(ByVal TitleRowRange$, Optional ByVal Orientation% = xlPortrait, _
Optional ByVal PrtGrdLine% = False, Optional ByVal CenterHorizontally% = True, _
Optional ByVal Zoom% = False, Optional ByVal Fit2PageWide% = 1, _
Optional ByVal Fit2PageTall% = False, Optional ByVal DisplayGridLines% = False)
'�]�wExcel�����]�w

    With G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet.PageSetup
        '�]�w�������C�L�����Y��
        .PrintTitleRows = TitleRowRange$
        .PrintGridlines = PrtGrdLine%
        .CenterHorizontally = CenterHorizontally%
        '���L(xlPortrait:���L  xlLandscape:��L)
        .Orientation = Orientation%
        '��ܭ���
        .CenterFooter = GetCaption("PanelDescpt", "page_format", "�� &P �� , �@ &N ��")
        '��ܦL��H
        .RightFooter = GetCaption("PanelDescpt", "print_person", "�L��H") & " : " & G_UserName
        .Zoom = CBool(Zoom%)
        .FitToPagesWide = Fit2PageWide%
        .FitToPagesTall = CBool(Fit2PageTall%)
    End With
    '����ܮ�u
    G_ExcelWkb.Windows(G_XLSRptName).DisplayGridLines = CBool(DisplayGridLines%)
End Sub

Sub SetExcelDataType(Spd As vaSpread, tSPD As Spread)
'�]�w�C�L��Excel������ƫ��A�ι���Ҧ�
Dim I%, A_Col%, A_CellType%, A_Places%, A_Align%, A_Format$, A_RIndex%

    For I% = 1 To UBound(tSPD.Columns)
        A_RIndex% = tSPD.Columns(I%).ReportIndex
        If A_RIndex% > 0 Then
           Spd.Row = -1
           Spd.Col = tSPD.Columns(I%).ScreenIndex
           Select Case Spd.CellType
             Case SS_CELL_TYPE_INTEGER, SS_CELL_TYPE_FLOAT
                  A_CellType% = SS_CELL_TYPE_FLOAT
                  A_Places% = Spd.TypeFloatDecimalPlaces
                  A_Format$ = "#,##0"
                  If A_Places% > 0 Then A_Format$ = A_Format$ & "." & String(A_Places%, "0")
             Case Else
                  A_CellType% = SS_CELL_TYPE_EDIT
                  A_Format$ = "#,##0"
           End Select
            Select Case Left(tSPD.Columns(I%).dFormat, 1)
                Case "#"
                    A_Align% = xlLeft
                Case "^"
                    A_Align% = xlCenter
                Case "~"
                    A_Align% = xlRight
            End Select
           SetColumnFormat2 GetExcelColName(A_RIndex%), A_CellType%, A_Format$, A_Align%
        End If
    Next I%
End Sub

Sub SetExcelNewPage(Optional ByVal Row# = 0)
'��ʳ]�w�����Ÿ�����m

    If Row# = 0 Then Row# = G_ExcelIndex# + G_XlsHRows% + 1
    G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet.Rows(Row#). _
    PageBreak = xlPageBreakManual
End Sub

Sub SetExcelSize(ByVal ColRange$, Optional ByVal ColWidth% = 12, Optional ByVal RowHeight% = 20)
'�]�wExcel��ΦC���w�]�j�p,�ýվ����d�򪺳̾A�e��

    '�N�d�򤤪���e�M�C���վ㬰�̾A����
    SelectExcelCells ColRange$
    With G_ExcelWkb.Windows(G_XLSRptName).Selection
         .ColumnWidth = ColWidth%
         .RowHeight = RowHeight%
    End With

    '�վ������e�ܳ̾A�e��
    G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet. _
    Columns(ColRange$).EntireColumn.AutoFit
End Sub

Sub SetExcelTextToColumns(ByVal StartCol As Integer, ByVal StartRow As Currency, ByVal EndRow As Currency, Optional ByVal fieldtype As Variant = Null)
'�]�w�r���R,�NCell������r�HG_G1�r������,���h��Cell��
Dim A_CellRange$
    
    If G_PrintSelect <> G_Print2Excel Then Exit Sub
    If StartRow <= 0 Or EndRow <= 0 Then Exit Sub
    If StartRow > EndRow Then Exit Sub

    A_CellRange$ = GetExcelColName(StartCol) & CStr(StartRow) & ":" & _
                   GetExcelColName(StartCol) & CStr(EndRow)
    
    If StartRow = EndRow Then
       If Trim(G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet. _
       Range(A_CellRange$).Value) = "" Then Exit Sub
    End If
                   
    If IsNull(fieldtype) Then
        G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet.Range(A_CellRange$).TextToColumns _
            DataType:=xlDelimited, TEXTQUALIFIER:=xlTextQualifierNone, _
            OTHER:=True, OTHERCHAR:=G_G1
    Else
        G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet.Range(A_CellRange$).TextToColumns _
            DataType:=xlDelimited, TEXTQUALIFIER:=xlTextQualifierNone, _
            OTHER:=True, OTHERCHAR:=G_G1, FieldInfo:=fieldtype
    End If
End Sub
    
Sub SetRptAllocate(Optional ByVal Left% = 2, Optional ByVal ColSpace% = 1, Optional ByVal Right% = 0)
'�]�w�����k���Ŷ����涡�j�ܦ@���ܼ�

    G_LineLeft = Space(Left%)
    G_ColSpace = Space(ColSpace%)
    G_LineRight = Space(Right%)
End Sub

Function SetXlsFldDataType(tSPD As Spread) As Variant
'�]�wExcel��쪺��ƫ��A��Array
Dim I%, A_Cols%, A_Index%, A_DataType(), A_Max%

    If G_PrintSelect <> G_Print2Excel Then Exit Function

    A_Max% = G_ExcelMaxCols% - 1
    ReDim A_DataType(A_Max%, 1)

    '���o��������`���
    A_Cols% = UBound(tSPD.Columns)

    '�]�w����ƫ��A
    For I% = 1 To A_Cols%
        If tSPD.Columns(I%).ReportIndex > 0 Then
           A_DataType(A_Index%, 0) = A_Index% + 1
           If Left(tSPD.Columns(I%).dFormat, 1) <> "~" Then
              If tSPD.Columns(I%).DateFormat = True Then
                 Select Case G_DateFlag
                   Case 0, 2
                        A_DataType(A_Index%, 1) = 5  'yyyy/m/d
                   Case 1
                        '�Y��OS������(�x�W)���B�]�w�ҥΰ�����榡(EMD)��,�ϥΰ�����榡,�_�h�]����r�榡.
                        If IsWinForTaiwan = True And XlsFldUseChinaDate = True Then
                            A_DataType(A_Index%, 1) = 10 'yy/m/d
                        Else
                            A_DataType(A_Index%, 1) = 2
                        End If
                 End Select
              Else
                 A_DataType(A_Index%, 1) = 2
              End If
           Else
              A_DataType(A_Index%, 1) = 1
           End If
           A_Index% = A_Index% + 1
        End If
    Next I%
    
    '�^��Array
    SetXlsFldDataType = A_DataType
End Function

Function GetRptHeaderFormat(ByVal FStr$, Optional ByVal FDate$ = "") As String
'�^�ǳ�����Y���r��榡
Dim A_Length%, A_STR$
    
    A_Length% = G_RptNeedWidth - Len(G_LineLeft) - lstrlen(FStr$) _
                - lstrlen(FDate$) - Len(G_LineRight)
    If A_Length% < 0 Then A_Length% = 1
    A_STR$ = G_LineLeft & FStr$ & String(A_Length%, Space(1)) & FDate$ & G_LineRight
    GetRptHeaderFormat = A_STR$
End Function

Function GetRptTitleFormat() As String
'�^�ǳ�����Y���榡
Dim A_Tmp$, A_Format$, A_Length%

    A_Tmp$ = "<" & UCase$(App.EXEName) & "> "
    A_Length% = G_RptNeedWidth - Len(G_LineLeft) - Len(A_Tmp$) * 2 - Len(G_LineRight)
    If G_RptNeedWidth - A_Length% < 0 Then A_Length% = 40
    A_Format$ = G_LineLeft & A_Tmp$ & String(A_Length%, "^") & G_LineRight
    GetRptTitleFormat = A_Format$
End Function

Function GetRptFootFormat(RightFoot$) As String
'�^�ǳ�����U�����r��榡
Dim A_Length%, A_STR$
    
    RightFoot$ = PrintUse(RightFoot$, GetCaption("PanelDescpt", "print_person", "�L �� �H") & G_G1 & G_UserName)
    A_Length% = G_RptNeedWidth - Len(G_LineLeft) - lstrlen(RightFoot$) * 2 _
                - Len(G_LineRight)
    If A_Length% < 0 Then A_Length% = 1
    A_STR$ = G_LineLeft & Space(lstrlen(RightFoot$)) & _
             PrintUse(String(A_Length%, "^"), G_Print_NextPage) & _
             RightFoot$ & G_LineRight
    GetRptFootFormat = A_STR$
End Function

Sub GetRptMinWidth(ByVal Str$)
'���o�����̤p�e��

    G_RptMinWidth = Len(G_LineLeft) + Len(Str$) + Len(G_LineRight)
End Sub

Sub PrintEnd2(Spd As vaSpread, tSPD As Spread)
'����C�L�������B�z�ʧ@,New Report Pattern Use
On Local Error GoTo MY_Error
    
    If G_PrintSelect = G_Print2Printer Then
       Printer.EndDoc
       If G_SetDefaultPrinter = Unchecked Then RestoreDefaultPrinter G_PrinterName
    ElseIf G_PrintSelect = G_Print2File Then
       Close
       If G_HaveDataPrint% Then
          retcode = Shell("Notepad " + G_OutFile, 1)
       End If
    ElseIf G_PrintSelect = G_Print2Screen Then
       On Error Resume Next
       Spd.SetFocus
       On Error GoTo 0
       Spd.TopRow = 1
       If tSPD.SortEnable Then SpreadColsSort Spd, tSPD
       DoEvents
    ElseIf G_PrintSelect = G_Print2Excel Then
       If G_ExcelWkb Is Nothing Then
          Close
       Else
          If G_HaveDataPrint% Then
             G_ExcelWkb.Calculation = xlCalculationAutomatic
             G_ExcelWkb.Parent.Visible = True
             SelectExcelCells "A1"
             G_ExcelWkb.Workbooks(G_XLSRptName).Save
             G_ExcelWkb.Parent.DisplayAlerts = True
             G_ExcelWkb.WindowState = xlMaximized
          Else
             CloseExcelFile
          End If
       End If
    End If
    Exit Sub
    
MY_Error:
    If Err = 1004 Then
       Err = 0
       Exit Sub
    End If
End Sub

Function OpenExcelFile(ByVal FileName$, Optional ByVal SheetName$ = "") As Boolean
'For�@�����,����榡��Run Time����
On Local Error GoTo MY_Error
Dim I%, A_Msg$

    OpenExcelFile = True
    '
    CloseExcelFile
    Set G_ExcelWkb = CreateObject("Excel.Application")
    
    If Dir(FileName$) <> "" Then Kill FileName$
    G_ExcelWkb.Workbooks.Add.SaveAs FileName$, xlNormal

'    If Dir(FileName$) = "" Then   '�ɮפ��s�b
'        G_ExcelWkb.Workbooks.Add.SaveAs FileName$
'    Else
'        G_ExcelWkb.Workbooks.Open Trim(FileName$), 0, False
'    End If
'    G_XLSRptName = Dir(FileName$)
    
    G_XLSRptName = Dir(FileName$)
    With G_ExcelWkb.Workbooks(G_XLSRptName)
         .RunAutoMacros xlAutoOpen
         G_ExcelWkb.Calculation = xlManual
         G_ExcelWkb.Parent.DisplayAlerts = False    '�����ܥ���ĵ�i
         For I% = .Worksheets.Count To 2 Step -1
             .Worksheets(I%).Delete
         Next I%
         If Trim(SheetName$) <> "" Then .Worksheets(1).Name = SheetName$
         If G_ExcelWkb.Windows(G_XLSRptName).View = xlPageBreakPreview Then   '�N�˵��]�w���зǼҦ�
            G_ExcelWkb.Windows(G_XLSRptName).View = xlNormalView
         End If
         .Worksheets(1).Cells.PageBreak = xlNone
         .Worksheets(1).Cells.Clear
    End With
    Exit Function

MY_Error:
    OpenExcelFile = False
    Select Case Err
    'PgmMsg  excel_file_inuse    �ɮץ��b�ϥΤ�,�Эק��ɦW��,�A����C�L!
      Case 70   'Permission denied
           A_Msg$ = GetCaption("PgmMsg", "excel_file_inuse", _
           "�ɮץ��b�ϥΤ�,�Эק��ɦW��,�A����C�L!")
           MsgBox A_Msg$, vbExclamation, App.Title
      Case Else
           MsgBox Error$, vbExclamation, App.Title
    End Select
    CloseExcelFile
End Function

Function OpenExcelFile_ReadOnly(ByVal FileName$) As Boolean
'For �S������(����榡�w����Excel File ���q�n)
On Local Error GoTo MY_Error

    OpenExcelFile_ReadOnly = True
    '
    CloseExcelFile
    Set G_ExcelWkb = CreateObject("Excel.Application")
    G_ExcelWkb.Workbooks.Open Trim(FileName$), 0, True
    G_XLSRptName = Dir(FileName$)
    G_ExcelWkb.Workbooks(G_XLSRptName).RunAutoMacros xlAutoOpen
    G_ExcelWkb.Calculation = xlManual
    G_ExcelWkb.Parent.DisplayAlerts = False    '�����ܥ���ĵ�i
    Exit Function
    
MY_Error:
    OpenExcelFile_ReadOnly = False
    MsgBox Error$
End Function

Function OpenExcelFile_Import(ByVal FileName$) As Boolean
'For ��Excel�ɤ��w�����,���{���B�z�L�{����|�^�gExcel File
On Local Error GoTo MY_Error

    OpenExcelFile_Import = True
    '
    CloseExcelFile
    Set G_ExcelWkb = CreateObject("Excel.Application")
    G_ExcelWkb.Workbooks.Open Trim(FileName$), 0, False
    G_XLSRptName = Dir(FileName$)
    G_ExcelWkb.Workbooks(G_XLSRptName).RunAutoMacros xlAutoOpen
    G_ExcelWkb.Calculation = xlManual
    G_ExcelWkb.Parent.DisplayAlerts = False    '�����ܥ���ĵ�i
    Exit Function
    
MY_Error:
    OpenExcelFile_Import = False
    MsgBox Error$
End Function

Sub SetColumnFormat(ByVal Col$, ByVal DType%, Optional ByVal dFormat$ = "#,##0")
'�]�w�Y�S�w��쪺�Ʀr�榡(��C�L��ƫe���]�w)
    
    With G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet.Columns(Col$)
         Select Case DType%
           Case 1  '��r
                .NumberFormat = "@"                   '��r�榡
           Case 2  '�f��
                .NumberFormat = dFormat$              '�f��
                .HorizontalAlignment = xlRight
         End Select
    End With
End Sub

Sub SetColumnFormat2(ByVal Col$, ByVal DType%, Optional ByVal dFormat$ = "#,##0", Optional ByVal Align% = xlLeft)
'�]�w�Y�S�w��쪺�Ʀr�榡�ι���Ҧ�(��C�L��ƫe���]�w)
    
    With G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet.Columns(Col$)
         Select Case DType%
           Case 1  '��r
                .NumberFormat = "@"              '��r�榡
           Case 2  '�f��
                .NumberFormat = dFormat$         '�f��
         End Select
         .HorizontalAlignment = Align%
    End With
End Sub

Sub Copy2NewSheet(ByVal SourceSheet$, ByVal NewSheet$)
'�ƻs������Ʀܥt�@������
'SourceSheet$.....�ӷ������W��,Example: Sheet1
'NewSheet$........�s�W�����W��,Example: Sheet2
Dim A_SheetCounts%

    With G_ExcelWkb.Workbooks(G_XLSRptName)
         A_SheetCounts% = .Worksheets.Count
         .Worksheets(SourceSheet$).Copy After:=.Worksheets(A_SheetCounts%)
         .Worksheets(A_SheetCounts%).Name = NewSheet$
         '
         G_ExcelWkb.CutCopyMode = False '�����ŤU�νƻs�Ҧ��ò������ʮؽu
         .Worksheets(NewSheet$).Select
         .Worksheets(NewSheet$).Range("A1").Select
    End With
End Sub

Sub SelectExcelCells(Optional ByVal Range$ = "")
'������w�d��Cells �� �]�wActive Cell
' Range$ - �ťեN�������Sheet

    With G_ExcelWkb.Workbooks(G_XLSRptName)
         .Activate
         If Trim(Range$) = "" Then
            .ActiveSheet.Cells.Select
         Else
            .ActiveSheet.Range(Range$).Select
         End If
    End With
End Sub

Sub xlsDrawLine(WKB As Excel.Application, ByVal Cells$, Optional A_OutlineOnly As Boolean = False)
'�b�x�s��d��|�g�[�W�ؽu
On Error Resume Next
    
    WKB.Workbooks(G_XLSRptName).Activate
    WKB.Workbooks(G_XLSRptName).ActiveSheet.Range(Cells$).Select
    
    StrCut Cells$, ":", "", Cells$
    
    With WKB.Windows(G_XLSRptName).Selection
         .Borders(xlDiagonalDown).LineStyle = xlNone
         .Borders(xlDiagonalUp).LineStyle = xlNone
    End With
    
    With WKB.Windows(G_XLSRptName).Selection.Borders(xlEdgeLeft)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    
    With WKB.Windows(G_XLSRptName).Selection.Borders(xlEdgeTop)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    
    With WKB.Windows(G_XLSRptName).Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    
    With WKB.Windows(G_XLSRptName).Selection.Borders(xlEdgeRight)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
    End With
    
    If A_OutlineOnly = False Then
        With WKB.Windows(G_XLSRptName).Selection.Borders(xlInsideVertical)
             .LineStyle = xlContinuous
             .Weight = xlHairline
             .ColorIndex = xlAutomatic
        End With
        
        With WKB.Windows(G_XLSRptName).Selection.Borders(xlInsideHorizontal)
             .LineStyle = xlContinuous
             .Weight = xlHairline
             .ColorIndex = xlAutomatic
        End With
    End If
End Sub

Sub AddExcelSheet(SheetIndex&, Optional ByVal SheetName$ = "")
'�bExcel���[�J�s������
Dim A_Counts&, A_SheetCounts&

    If G_PrintSelect <> G_Print2Excel Then Exit Sub
    
    'SheetName���i��\/?*[] ���r��
    SheetName$ = Replace(SheetName$, "/", "_")
    SheetName$ = Replace(SheetName$, "\", "_")
    SheetName$ = Replace(SheetName$, "?", "_")
    SheetName$ = Replace(SheetName$, "*", "_")
    SheetName$ = Replace(SheetName$, "[", "_")
    SheetName$ = Replace(SheetName$, "]", "_")
    
    SheetIndex& = SheetIndex& + 1
    A_SheetCounts& = G_ExcelWkb.Workbooks(G_XLSRptName).Sheets.Count
    '�YG_ExcelWkb.Workbooks(G_XLSRptName).Worksheets.Count<SheetIndex&, ��Sheet
    With G_ExcelWkb.Workbooks(G_XLSRptName)
         For A_Counts& = 1 To SheetIndex& - A_SheetCounts&
             .Worksheets.Add
             A_SheetCounts& = A_SheetCounts& + 1
             .Worksheets(A_SheetCounts&).Select
             .Sheets(A_SheetCounts& - 1).Move After:=.Sheets(A_SheetCounts&)
         Next
         .Worksheets(SheetIndex&).Select
         If Trim(SheetName$) <> "" Then
            .Sheets(SheetIndex&).Name = Trim(SheetName$)
         End If
    End With
End Sub

Sub SetExcelRangeColor(ByVal A_Row#, ByVal A_Row2#, ByVal A_Col#, ByVal A_Col2#, Optional ByVal A_BColor# = COLOR_WHITE, Optional ByVal A_FColor# = COLOR_BLACK)
'�]�wExcel���Y�@��Cell��Range���I���Ϋe���C��
'A_Row# : �_�l�C    A_Row2# : �I��C
'A_Col# : �_�l��    A_Col2# : �I����
'A_BColor# : �I���C�� (�Ѽƥi���ǻ�,�w�]���զ�)
'A_FColor# : �e���C�� (�Ѽƥi���ǻ�,�w�]���¦�)

    If G_PrintSelect <> G_Print2Excel Then Exit Sub
        
    Dim A_Range$
    A_Range$ = GetExcelColName(A_Col#) & CStr(A_Row#) & ":" & _
               GetExcelColName(A_Col2#) & CStr(A_Row2#)
    
    With G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet.Range(A_Range$)
         .Interior.Color = A_BColor
         .Font.Color = A_FColor
    End With
End Sub

'===============================================================================
' Add New Function at 93/3/8
'===============================================================================
Sub ShowExcelFile()
'�������,����Excel����,�A�HShell Function�}�ҳ����ɮ�.
'�קK�bExcel 2003�H�U������,�Y�}�Ҫ�Excel File�Q�ϥΪ̥���
'������B�{��������,�A���ɮ��`�ޤ��}�ҧO��Excel File��,�|
'�L�k����ɮפ��e�����p�o��.

    If G_PrintSelect <> G_Print2Excel Then Exit Sub
    CloseExcelFile
    If Not G_HaveDataPrint% Then Exit Sub
    retcode = Shell(GetExcelAppPath & " """ & G_OutFile & """", vbMaximizedFocus)
End Sub

'===============================================================================
' Edit Function at 93/4/1
'===============================================================================
Sub PrintEnd(Tmp As Object)
'����C�L�������B�z�ʧ@
On Local Error GoTo MY_Error
Const wdWindowStateMaximize = 1

    If G_PrintSelect = G_Print2Printer Then
       Printer.EndDoc
       If G_SetDefaultPrinter = Unchecked Then RestoreDefaultPrinter G_PrinterName
    ElseIf G_PrintSelect = G_Print2File Then
       Close
    ElseIf G_PrintSelect = G_Print2Screen Then
       On Error Resume Next
       Tmp.SetFocus
       On Error GoTo 0
       Tmp.TopRow = 1
       DoEvents
    ElseIf G_PrintSelect = G_Print2Excel Then
       If G_ExcelWkb Is Nothing Then
          Close
       Else
          If G_HaveDataPrint% Then
             G_ExcelWkb.Calculation = xlCalculationAutomatic
             G_ExcelWkb.Parent.Visible = True
             SelectExcelCells "A1"
             G_ExcelWkb.Workbooks(G_XLSRptName).Save
             G_ExcelWkb.Parent.DisplayAlerts = True
             G_ExcelWkb.WindowState = xlMaximized
          Else
             CloseExcelFile
          End If
       End If
    ElseIf G_PrintSelect = G_Print2Word Then
       If G_HaveDataPrint% Then
          G_WordDoc.WindowState = wdWindowStateMaximize
          G_WordDoc.Parent.DisplayAlerts = True
          G_WordDoc.Documents(G_DocRptName).Save
          ShowWordFile
       Else
          CloseWordFile
       End If
    End If
    Exit Sub
    
MY_Error:
    If Err = 1004 Then
       Err = 0
       Exit Sub
    End If
End Sub

Sub PrintEnd3(Tmp As Object)
'����C�L�������B�z�ʧ@
On Local Error GoTo MY_Error
Const wdWindowStateMaximize = 1

    If G_PrintSelect = G_Print2Printer Then
       Printer.EndDoc
       If G_SetDefaultPrinter = Unchecked Then RestoreDefaultPrinter G_PrinterName
    ElseIf G_PrintSelect = G_Print2File Then
       Close
    ElseIf G_PrintSelect = G_Print2Screen Then
       On Error Resume Next
       Tmp.SetFocus
       On Error GoTo 0
       Tmp.TopRow = 1
       DoEvents
    ElseIf G_PrintSelect = G_Print2Excel Then
       If G_ExcelWkb Is Nothing Then
          Close
       Else
          If G_HaveDataPrint% Then
             G_ExcelWkb.Calculation = xlCalculationAutomatic
             SelectExcelCells "A1"
             G_ExcelWkb.Parent.DisplayAlerts = True
             G_ExcelWkb.WindowState = xlMaximized
             G_ExcelWkb.Workbooks(G_XLSRptName).Save
             ShowExcelFile
          Else
             CloseExcelFile
          End If
       End If
    ElseIf G_PrintSelect = G_Print2Word Then
       If G_HaveDataPrint% Then
          G_WordDoc.WindowState = wdWindowStateMaximize
          G_WordDoc.Parent.DisplayAlerts = True
          G_WordDoc.Documents(G_DocRptName).Save
          ShowWordFile
       Else
          CloseWordFile
       End If
    End If
    Exit Sub
    
MY_Error:
    If Err = 1004 Then
       Err = 0
       Exit Sub
    End If
End Sub

Function GetRptFormatStr(tSPD As Spread, ByVal FType%) As String
'��s������D�θ����ܪ��r��榡
'FType% = 1,�Ǧ^���D���r��榡(Single Line Use)
'FType% = 2,�Ǧ^�����ܪ��r��榡
'FType% = 3,�Ǧ^���D���r��榡(Multi Line Use)
Dim I%, A_Cols%, A_STR$, A_Connect$

    '���o��������`���
    A_Cols% = UBound(tSPD.Columns)
    
    If G_PrintSelect = G_Print2Excel Or G_PrintSelect = G_Print2Word Then
       A_Connect$ = G_G1
    Else
       A_Connect$ = G_ColSpace
    End If
    
    '�զ�Format(�����w�d�Ŷ�+������+�涡���j����+�k���w�d�Ŷ�)
    If G_PrintSelect <> G_Print2Excel And G_PrintSelect <> G_Print2Word Then
       A_STR$ = G_LineLeft
    End If
       
    For I% = 1 To A_Cols%
        If tSPD.Columns(I%).ReportIndex > 0 Then
           Select Case FType%
             Case 1
                  If G_PrintSelect = G_Print2Excel Or G_PrintSelect = G_Print2Word Then
                     A_STR = A_STR & tSPD.Columns(I%).Caption
                  Else
                     A_STR$ = A_STR$ & PrintUse(tSPD.Columns(I%).CFormat, tSPD.Columns(I%).Caption)
                  End If
             Case 2
                  A_STR$ = A_STR$ & tSPD.Columns(I%).dFormat
             Case 3
                  If G_PrintSelect = G_Print2Excel Or G_PrintSelect = G_Print2Word Then
                     A_STR = A_STR & tSPD.Columns(I%).Caption
                  Else
                     A_STR$ = A_STR$ & PrintUse(tSPD.Columns(I%).CFormat, tSPD.Columns(I%).mCaption)
                  End If
          End Select
           If I% <> A_Cols% Then A_STR$ = A_STR$ & A_Connect$
        End If
    Next I%
    A_STR$ = A_STR$ & G_LineRight
    
    '�]�w����Ҷ��e��
    G_RptNeedWidth = Len(A_STR$)
    If G_RptNeedWidth < G_RptMinWidth Then G_RptNeedWidth = G_RptMinWidth
    
    '�^��Format
    GetRptFormatStr = A_STR$
End Function

Sub PrintOut(Tmp As Object, ByVal f$, ByVal v$, Optional ByVal Row#)
'�C�L�ɨϥ�,�|�Ϥ����P���C�L�覡
'�YRow# = -1,��G_PrintSelect = G_Print2Screen,G_Print2Excel��,G_Print2Word��,
'���B�z�C�L�ʧ@
Dim A_G1Pos%
Dim A_Str1$, A_Str2$, I%, A_Start As Boolean
Dim A_STR$(), A_CellStart$, A_CellEnd$, A_Index#

    If G_PrintSelect = G_Print2Printer Then
       Printer.Print PrintUse(f$, v$)
    ElseIf G_PrintSelect = G_Print2File Then
       Print #1, PrintUse(f$, v$)
    ElseIf G_PrintSelect = G_Print2Screen Then
       If Row# = -1 Then Exit Sub
       GoSub PrintOutA
       Tmp.MaxRows = Row#
       Tmp.Row = Row#: Tmp.Col = 1
       Tmp.Row2 = Row#: Tmp.Col2 = Tmp.MaxCols
       Tmp.Clip = f$
       On Error Resume Next
       Tmp.SetFocus
       On Error GoTo 0
       Tmp.TopRow = SetSpreadTopRow(Tmp)
       DoEvents
    ElseIf G_PrintSelect = G_Print2Excel Then
       If Row# = -1 Then Exit Sub
       If G_ExcelWkb Is Nothing Then
          GoSub PrintOutA
          Print #1, f$
       Else
          GoSub PrintOutB
       End If
    ElseIf G_PrintSelect = G_Print2Word Then
       If Row# = -1 Then Exit Sub
          GoSub PrintOutC
    End If
    If Not G_HaveDataPrint% Then G_HaveDataPrint% = True
    Exit Sub

PrintOutA:
    f$ = ""
    Do While True
       A_G1Pos% = InStr(v$, G_G1)
       If A_G1Pos% <> 0 Then
          If G_PrintSelect = G_Print2Excel Then
             f$ = f$ & Chr$(34) & Left$(v$, A_G1Pos% - 1) & Chr$(34) & Chr$(44)
          Else
             f$ = f$ & Left$(v$, A_G1Pos% - 1) & Chr$(KEY_TAB)
          End If
          v$ = Mid(v$, A_G1Pos% + 1)
       Else
          If G_PrintSelect = G_Print2Excel Then
             f$ = f$ & Chr$(34) & v$ & Chr$(34)
          Else
             f$ = f$ & v$
          End If
          Exit Do
       End If
    Loop
    Return

PrintOutB:
    If Trim(v$) <> "" Then
       A_Index# = G_XlsHRows% + Row#
       A_CellStart$ = GetExcelColName(G_XlsStartCol%) & Trim(A_Index#)
       On Error Resume Next
       A_STR$ = Split(v$, G_G1, -1, vbTextCompare)
       With G_ExcelWkb.Workbooks(G_XLSRptName).ActiveSheet
            If Err <> 0 Then
               .Range(A_CellStart$ & ":" & A_CellStart$).Value = v$
            Else
               A_CellEnd$ = GetExcelColName(G_XlsStartCol% + UBound(A_STR$)) & _
                            Trim(A_Index#)
               .Range(A_CellStart$ & ":" & A_CellEnd$).Value = A_STR$
            End If
       End With
       On Error GoTo 0
    End If
    Return

PrintOutC:
    A_STR$ = Split(v$, G_G1)
    A_Index# = G_XlsHRows% + Row#
    With G_DocSelection.Tables(1)
         If A_Index# > .Rows.Count Then
            .Rows.Add
         End If
    End With
        
    If UBound(A_STR$) <= 0 Then
       SelectWordCells "A" & CStr(A_Index#)
    Else
       SelectWordCells "A" & CStr(A_Index#) & ":" & GetExcelColName(UBound(A_STR$) + 1) & CStr(A_Index#)
    End If
    
    Dim A_Cell As Object
    For Each A_Cell In G_DocSelection.Cells
        If I% > UBound(A_STR$) Then Exit For
        If Trim(A_STR$(I%)) <> "" Then
            A_Cell.WordWrap = False
            A_Cell.Range.text = A_STR$(I%)
        End If
        I% = I% + 1
    Next
    Return
End Sub

Sub PrintOut2(Spd As vaSpread, ByVal f$, ByVal v$, Optional ByVal Row#)
'�C�L�ɨϥ�,�|�Ϥ����P���C�L�覡,New Report Pattern Use
'�YRow# = -1,��G_PrintSelect = G_Print2Screen,G_Print2Excel��,G_Print2Word��,
'���B�z�C�L�ʧ@
Dim I%, A_G1Pos%
Dim A_STR$(), A_CellStart$, A_CellEnd$, A_Index#

    If G_PrintSelect = G_Print2Printer Then
       Printer.Print PrintUse(f$, v$)
    ElseIf G_PrintSelect = G_Print2File Then
       Print #1, PrintUse(f$, v$)
    ElseIf G_PrintSelect = G_Print2Screen Then
       If Row# = -1 Then Exit Sub
       GoSub PrintOut2A
       Spd.MaxRows = Row#
       Spd.Row = Row#: Spd.Col = 1
       Spd.Row2 = Row#: Spd.Col2 = Spd.MaxCols
       Spd.Clip = f$
       On Error Resume Next
       Spd.SetFocus
       On Error GoTo 0
       Spd.TopRow = SetSpreadTopRow(Spd)
       DoEvents
    ElseIf G_PrintSelect = G_Print2Excel Then
       If Row# = -1 Then Exit Sub
       If G_ExcelWkb Is Nothing Then
          GoSub PrintOut2A
          Print #1, f$
       Else
          GoSub PrintOut2B
       End If
    ElseIf G_PrintSelect = G_Print2Word Then
       If Row# = -1 Then Exit Sub
          GoSub PrintOutC
    End If
    If Not G_HaveDataPrint% Then G_HaveDataPrint% = True
    Exit Sub

PrintOut2A:
    f$ = ""
    Do While True
       A_G1Pos% = InStr(v$, G_G1)
       If A_G1Pos% <> 0 Then
          If G_PrintSelect = G_Print2Excel Then
             f$ = f$ & Chr$(34) & Left$(v$, A_G1Pos% - 1) & Chr$(34) & Chr$(44)
          Else
             f$ = f$ & Left$(v$, A_G1Pos% - 1) & Chr$(KEY_TAB)
          End If
          v$ = Mid(v$, A_G1Pos% + 1)
       Else
          If G_PrintSelect = G_Print2Excel Then
             f$ = f$ & Chr$(34) & v$ & Chr$(34)
          Else
             f$ = f$ & v$
          End If
          Exit Do
       End If
    Loop
    Return

PrintOut2B:
    If Trim(v$) <> "" Then
       A_Index# = G_XlsHRows% + Row#
       A_CellStart$ = GetExcelColName(G_XlsStartCol%) & Trim(A_Index#)
       On Error Resume Next
       A_STR$ = Split(v$, G_G1, -1, vbTextCompare)
       With G_ExcelWkb.Windows(G_XLSRptName).ActiveSheet
            If Err <> 0 Then
               .Range(A_CellStart$ & ":" & A_CellStart$).Value = v$
            Else
               A_CellEnd$ = GetExcelColName(G_XlsStartCol% + UBound(A_STR$)) & _
                            Trim(A_Index#)
               .Range(A_CellStart$ & ":" & A_CellEnd$).Value = A_STR$
            End If
       End With
       On Error GoTo 0
    End If
    Return

PrintOutC:
    A_STR$ = Split(v$, G_G1)
    A_Index# = G_XlsHRows% + Row#
    With G_DocSelection.Tables(1)
         If A_Index# > .Rows.Count Then
            .Rows.Add
         End If
    End With
        
    If UBound(A_STR$) <= 0 Then
       SelectWordCells "A" & CStr(A_Index#)
    Else
       SelectWordCells "A" & CStr(A_Index#) & ":" & GetExcelColName(UBound(A_STR$) + 1) & CStr(A_Index#)
    End If
    
    Dim A_Cell As Object
    For Each A_Cell In G_DocSelection.Cells
        If I% > UBound(A_STR$) Then Exit For
        If Trim(A_STR$(I%)) <> "" Then
            A_Cell.WordWrap = False
            A_Cell.Range.text = A_STR$(I%)
        End If
        I% = I% + 1
    Next
    Return
End Sub

Sub PrintOut3(Tmp As Object, ByVal f$, ByVal v$, Optional ByVal Row#)
'�C�L�ɨϥ�,�|�Ϥ����P���C�L�覡
'�YRow# = -1,��G_PrintSelect = G_Print2Screen,G_Print2Excel,G_Print2Word��,
'���B�z�C�L�ʧ@
Dim I%, A_G1Pos%
Dim A_Str1$, A_Str2$, A_Start As Boolean
Dim A_STR$(), A_CellStart$, A_CellEnd$, A_Index#

    If G_PrintSelect = G_Print2Printer Then
       Printer.Print PrintUse(f$, v$)
    ElseIf G_PrintSelect = G_Print2File Then
       Print #1, PrintUse(f$, v$)
    ElseIf G_PrintSelect = G_Print2Screen Then
       If Row# = -1 Then Exit Sub
       GoSub PrintOutA
       Tmp.MaxRows = Row#
       Tmp.Row = Row#: Tmp.Col = 1
       Tmp.Row2 = Row#: Tmp.Col2 = Tmp.MaxCols
       Tmp.Clip = f$
       On Error Resume Next
       Tmp.SetFocus
       On Error GoTo 0
       Tmp.TopRow = SetSpreadTopRow(Tmp)
       DoEvents
    ElseIf G_PrintSelect = G_Print2Excel Then
       If Row# = -1 Then Exit Sub
       If G_ExcelWkb Is Nothing Then
          GoSub PrintOutA
          Print #1, f$
       Else
          GoSub PrintOutB
       End If
    ElseIf G_PrintSelect = G_Print2Word Then
       If Row# = -1 Then Exit Sub
          GoSub PrintOutC
    End If
    If Not G_HaveDataPrint% Then G_HaveDataPrint% = True
    Exit Sub

PrintOutA:
    f$ = ""
    Do While True
       A_G1Pos% = InStr(v$, G_G1)
       If A_G1Pos% <> 0 Then
          If G_PrintSelect = G_Print2Excel Then
             f$ = f$ & Chr$(34) & Left$(v$, A_G1Pos% - 1) & Chr$(34) & Chr$(44)
          Else
             f$ = f$ & Left$(v$, A_G1Pos% - 1) & Chr$(KEY_TAB)
          End If
          v$ = Mid(v$, A_G1Pos% + 1)
       Else
          If G_PrintSelect = G_Print2Excel Then
             f$ = f$ & Chr$(34) & v$ & Chr$(34)
          Else
             f$ = f$ & v$
          End If
          Exit Do
       End If
    Loop
    Return

PrintOutB:
    A_Index# = G_XlsHRows% + Row#
    A_CellStart$ = GetExcelColName(G_XlsStartCol%) & Trim(A_Index#)
    G_ExcelWkb.Windows(G_XLSRptName).ActiveSheet.Range(A_CellStart$).Value = v$
    Return

PrintOutC:
    A_STR$ = Split(v$, G_G1)
    A_Index# = G_XlsHRows% + Row#
    With G_DocSelection.Tables(1)
         If A_Index# > .Rows.Count Then
            .Rows.Add
         End If
    End With
        
    If UBound(A_STR$) <= 0 Then
       SelectWordCells "A" & CStr(A_Index#)
    Else
       SelectWordCells "A" & CStr(A_Index#) & ":" & GetExcelColName(UBound(A_STR$) + 1) & CStr(A_Index#)
    End If
    
    Dim A_Cell As Object
    For Each A_Cell In G_DocSelection.Cells
        If I% > UBound(A_STR$) Then Exit For
        If Trim(A_STR$(I%)) <> "" Then
            A_Cell.WordWrap = False
            A_Cell.Range.text = A_STR$(I%)
        End If
        I% = I% + 1
    Next
    Return
End Sub

Sub PrintEnd4(Spd As vaSpread, tSPD As Spread, Optional ByVal ShowNotePad As Boolean = True)
'����C�L�������B�z�ʧ@,New Report Pattern Use
On Local Error GoTo MY_Error
Const wdWindowStateMaximize = 1

    If G_PrintSelect = G_Print2Printer Then
       Printer.EndDoc
       If G_SetDefaultPrinter = Unchecked Then RestoreDefaultPrinter G_PrinterName
    ElseIf G_PrintSelect = G_Print2File Then
       Close
       If G_HaveDataPrint% And ShowNotePad Then
          retcode = Shell("Notepad " + G_OutFile, 1)
       End If
    ElseIf G_PrintSelect = G_Print2Screen Then
       On Error Resume Next
       Spd.SetFocus
       On Error GoTo 0
       Spd.TopRow = 1
       If tSPD.SortEnable Then SpreadColsSort Spd, tSPD
       DoEvents
    ElseIf G_PrintSelect = G_Print2Excel Then
       If G_ExcelWkb Is Nothing Then
          Close
       Else
          If G_HaveDataPrint% Then
             G_ExcelWkb.Calculation = xlCalculationAutomatic
             SelectExcelCells "A1"
             G_ExcelWkb.Parent.DisplayAlerts = True
             G_ExcelWkb.WindowState = xlMaximized
             G_ExcelWkb.Workbooks(G_XLSRptName).Save
             ShowExcelFile
          Else
             CloseExcelFile
          End If
       End If
    ElseIf G_PrintSelect = G_Print2Word Then
       If G_HaveDataPrint% Then
          G_WordDoc.WindowState = wdWindowStateMaximize
          G_WordDoc.Visible = False
          G_WordDoc.Parent.DisplayAlerts = True
          G_WordDoc.Documents(G_DocRptName).Save
          ShowWordFile
       Else
          CloseWordFile
       End If
    End If
    Exit Sub
    
MY_Error:
    If Err = 1004 Then
       Err = 0
       Exit Sub
    End If
End Sub

Sub CloseExcelFile()
'����Excel�ɮ�
On Local Error Resume Next
    
    Select Case G_ExcelWkb.Workbooks.Count
      Case 0
           G_ExcelWkb.Quit
      Case 1
           G_ExcelWkb.ActiveWorkbook.Close savechanges:=False
           G_ExcelWkb.Quit
      Case Else
           G_ExcelWkb.Workbooks(G_XLSRptName).Close savechanges:=False
    End Select
    Set G_ExcelWkb = Nothing
End Sub

Function PrintUse(ByVal f$, ByVal v$) As String
'�^�ǮM�ή榡�᪺�r��
Dim a_fp%, a_fl%, a_vp%, a_vl%, A_Tmp%
Dim a_1%, a_2%, a_3%, a_4%
Dim a_f$, a_v$, a_out$, a_SCharLen%
On Error GoTo PrintUse_Error
'�������׺�@�Ӧ줸���r��
Const A_SChar$ = "�X��"

    a_out$ = ""
    If Trim$(f$) = "" Then
       PrintUse = f$
       Exit Function
    End If
    a_fp% = 1: a_vp% = 1
    If Trim$(v$) = "" Then GoTo RptPrint
'
    Do While a_vp% <= Len(v$)
       A_Tmp% = InStr(a_vp%, v$, G_G1)
       If A_Tmp% = 0 Then
          a_vl% = Len(v$) - a_vp% + 1
       Else
          a_vl% = A_Tmp% - a_vp%
       End If
'
       a_v$ = Mid$(v$, a_vp%, a_vl%)
       GoSub PrintUse_A
       a_vp% = a_vp% + a_vl% + 1
    Loop
    GoTo RptPrint
    
PrintUse_A:
    Do Until a_fp% > Len(f$)
       If Mid$(f$, a_fp%, 1) = "#" Then
          For a_fl% = a_fp% To Len(f$)
              If Mid$(f$, a_fl%, 1) <> "#" Then
                 a_fl% = a_fl% - a_fp%
                 Exit For
              End If
          Next a_fl%
          If a_fl% > Len(f$) Then
             a_fl% = a_fl% - a_fp%
          ElseIf a_fl% = Len(f$) Then
             a_fl% = a_fl% - a_fp% + 1
          End If
          GoSub PrintUse_Left
          a_fp% = a_fp% + a_fl%
          Return
       End If
       
       If Mid$(f$, a_fp%, 1) = "~" Then
          For a_fl% = a_fp% To Len(f$)
              If Mid$(f$, a_fl%, 1) <> "~" Then
                 a_fl% = a_fl% - a_fp%
                 Exit For
              End If
          Next a_fl%
          If a_fl% > Len(f$) Then
             a_fl% = a_fl% - a_fp%
          ElseIf a_fl% = Len(f$) Then
             a_fl% = a_fl% - a_fp% + 1
          End If
          GoSub PrintUse_Right
          a_fp% = a_fp% + a_fl%
          Return
       End If
       
       If Mid$(f$, a_fp%, 1) = "^" Then
          For a_fl% = a_fp% To Len(f$)
              If Mid$(f$, a_fl%, 1) <> "^" Then
                 a_fl% = a_fl% - a_fp%
                 Exit For
              End If
          Next a_fl%
          If a_fl% > Len(f$) Then
             a_fl% = a_fl% - a_fp%
          ElseIf a_fl% = Len(f$) Then
             a_fl% = a_fl% - a_fp% + 1
          End If
          GoSub PrintUse_Middle
          a_fp% = a_fp% + a_fl%
          Return
       End If
       a_out$ = a_out$ + Mid$(f$, a_fp%, 1)
       a_fp% = a_fp% + 1
       If a_fp% > Len(f$) Then Exit Do
    Loop
    Return
    
PrintUse_Left:
    a_2% = 0: a_3% = 0: a_SCharLen% = 0
    a_f$ = Mid$(f$, a_fp%, a_fl%)
    For a_1% = 1 To a_vl%
        If Asc(Mid$(a_v$, a_1%, 1)) > 0 Then
           a_2% = a_2% + 1
        Else
           If InStr(1, A_SChar$, Mid$(a_v$, a_1%, 1), vbTextCompare) > 0 Then
              'KEEP�S��r��������
              a_SCharLen% = a_SCharLen% + 1
              a_2% = a_2% + 1
           Else
              a_2% = a_2% + 2
           End If
        End If
        If a_2% <= a_fl% Then
           a_3% = a_3% + 1
           Mid$(a_f$, a_3%, 1) = Mid$(a_v$, a_1%, 1)
        Else
           Exit For
        End If
    Next a_1%
    If a_fl% > a_3% Then
       a_f$ = Mid$(a_f$, 1, a_3%)
       If lstrlen(a_f$) < a_fl% Then
          a_f$ = a_f$ & Space$(a_fl% - (lstrlen(a_f$) - IIf(G_PrintSelect = G_Print2Printer, a_SCharLen%, 0)))
       End If
    End If
       
    '�ѨMTab�r�����榡���D
    Dim M%, A_Start%, A_Tab$(), A_Complete$
    A_Start% = 0
    A_Tab$ = Split(a_f$, Chr(9))
    A_Complete$ = ""
    If UBound(A_Tab$) > 0 Then
       For M% = 0 To UBound(A_Tab$)
           A_Complete$ = A_Complete$ & RTrim(A_Tab$(M%))
           '�C�@Tab�䶡�j8�Ӧ줸(1,9,17,35,43....)
           A_Start% = 8 * (Int(lstrlen(A_Complete$) / 8) + 1)
           A_Complete$ = A_Complete$ & Space(A_Start% - lstrlen(A_Complete$))
       Next M%
       If lstrlen(a_f$) < lstrlen(A_Complete$) Then
          A_Complete$ = GetLenStr(A_Complete$, 1, lstrlen(a_f$))
       End If
       A_Complete$ = A_Complete$ & Space(lstrlen(a_f$) - lstrlen(A_Complete$))
    Else
       A_Complete$ = a_f$
    End If

    a_out$ = a_out$ + A_Complete$  'a_f$

    Return
    
PrintUse_Right:
    a_2% = 0: a_3% = 0: a_SCharLen% = 0
    a_f$ = Mid$(f$, a_fp%, a_fl%)
    For a_1% = 1 To a_vl%
        If Asc(Mid$(a_v$, a_1%, 1)) > 0 Then
           a_2% = a_2% + 1
        Else
           If InStr(1, A_SChar$, Mid$(a_v$, a_1%, 1), vbTextCompare) > 0 Then
              'KEEP�S��r��������
              a_SCharLen% = a_SCharLen% + 1
              a_2% = a_2% + 1
           Else
              a_2% = a_2% + 2
           End If
        End If
        If a_2% <= a_fl% Then
           a_3% = a_3% + 1
           Mid$(a_f$, a_3%, 1) = Mid$(a_v$, a_1%, 1)
        Else
           Exit For
        End If
    Next a_1%
    If a_fl% > a_3% Then
       a_f$ = Mid$(a_f$, 1, a_3%)
       If lstrlen(a_f$) < a_fl% Then
          a_f$ = Space$(a_fl% - (lstrlen(a_f$) - IIf(G_PrintSelect = G_Print2Printer, a_SCharLen%, 0))) & a_f$
       End If
    End If
    a_out$ = a_out$ + a_f$
    Return
    
PrintUse_Middle:
    a_2% = 0: a_3% = 0
    a_f$ = Mid$(f$, a_fp%, a_fl%)
    For a_1% = 1 To a_vl%
        If Asc(Mid$(a_v$, a_1%, 1)) > 0 Then
           a_2% = a_2% + 1
        Else
           If InStr(1, A_SChar$, Mid$(a_v$, a_1%, 1), vbTextCompare) > 0 Then
              a_2% = a_2% + 1
           Else
              a_2% = a_2% + 2
           End If
        End If
        If a_2% <= a_fl% Then
           a_3% = a_3% + 1
           Mid$(a_f$, a_3%, 1) = Mid$(a_v$, a_1%, 1)
        Else
           Exit For
        End If
    Next a_1%
    If a_fl% > a_3% Then
       If a_fl% > a_2% Then
          a_4% = (a_fl% - a_2%) / 2
          If a_4% > 0 Then
             a_f$ = Space$((a_fl% - a_2%) - a_4%) + Mid$(a_f$, 1, a_3%) + Space$(a_4%)
          Else
             a_f$ = Space$((a_fl% - a_2%) - a_4%) + Mid$(a_f$, 1, a_3%)
          End If
       Else
          a_f$ = Mid$(a_f$, 1, a_3%)
       End If
    End If
    a_out$ = a_out$ + a_f$
    Return
    
RptPrint:
    Do While a_fp% <= Len(f$)
       If Mid$(f$, a_fp%, 1) = "#" Or Mid$(f$, a_fp%, 1) = "~" Or Mid$(f$, a_fp%, 1) = "^" Then
          Mid$(f$, a_fp%, 1) = " "
       End If
       a_out$ = a_out$ + Mid$(f$, a_fp%, 1)
       a_fp% = a_fp% + 1
    Loop
    PrintUse = a_out$
    Exit Function
    
PrintUse_Error:
    MsgBox Error(Err)
    Resume Next
End Function

Sub SetCellAlignment(ByVal Range$, ByVal Hpos&, ByVal Vpos&, ByVal Flag As Boolean, Optional ByVal WrapText As Boolean = False)
'�]�w�x�s�檺����Ҧ��ΦX���x�s��

    If G_PrintSelect <> G_Print2Excel And G_PrintSelect <> G_Print2Word Then Exit Sub
    
    If G_PrintSelect = G_Print2Excel Then
       SelectExcelCells Range$
       With G_ExcelWkb.Windows(G_XLSRptName)
            .Selection.HorizontalAlignment = Hpos&         '��������Ҧ�
            .Selection.VerticalAlignment = Vpos&
            .Selection.WrapText = WrapText                'True:��r�۰ʴ��C
            .Selection.Orientation = 0
            If Flag Then .Selection.Merge
       End With
    Else
       With G_DocSelection
            '����d��
            SelectWordCells Range$
            
            '�X���x�s��
            If Flag Then .Cells.Merge
            
            '�m�k(wdAlignParagraphRight=2) �m��(wdAlignParagraphLeft=0) �m��(wdAlignParagraphCenter=1)
            Select Case Hpos&
                Case xlLeft
                    .ParagraphFormat.Alignment = 0
                Case xlCenter
                    .ParagraphFormat.Alignment = 1
                Case xlRight
                    .ParagraphFormat.Alignment = 2
            End Select
            
            '�m�U(wdCellAlignVerticalBottom=3) �m�W(wdCellAlignVerticalTop=0) �m��(wdCellAlignVerticalCenter=1)
            Select Case Vpos&
                Case xlTop
                    .Cells.VerticalAlignment = 0
                Case xlCenter
                    .Cells.VerticalAlignment = 1
                Case xlBottom
                    .Cells.VerticalAlignment = 3
            End Select
       End With
    End If
End Sub

'===============================================================================
' Add New Function at 93/4/1
'===============================================================================
Sub CreateWordDocTable(ByVal Rows#, ByVal Cols#, Optional ByVal ColFmt$ = "", Optional ByVal SplitChar$ = " ", Optional ByVal InPageHeader% = False)
Const wdSeekCurrentPageHeader = 9
Const wdWord9TableBehavior = 1
Const wdAutoFitWindow = 2
Const wdLineStyleNone = 0
Const wdBorderTop = -1
Const wdBorderLeft = -2
Const wdBorderBottom = -3
Const wdBorderRight = -4
Const wdBorderHorizontal = -5
Const wdBorderVertical = -6
Const wdBorderDiagonalDown = -7
Const wdBorderDiagonalUp = -8
Const wdLineStyleSingle = 1
Const wdLineWidth050pt = 4
Const wdColorAutomatic = -16777216
Const wdSeekMainDocument = 0
Const wdRowHeightExactly = 2

    If InPageHeader% Then
       G_WordDoc.Documents(G_DocRptName).ActiveWindow.ActivePane. _
       View.SeekView = wdSeekCurrentPageHeader
    End If

    G_WordDoc.Documents(G_DocRptName).Tables.Add _
    Range:=G_DocSelection.Range, _
    NumRows:=Rows#, NumColumns:=Cols#, _
    DefaultTableBehavior:=wdWord9TableBehavior, _
    AutoFitBehavior:=wdAutoFitWindow
       
    With G_DocSelection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
        .TopPadding = G_WordDoc.CentimetersToPoints(0)
        .BottomPadding = G_WordDoc.CentimetersToPoints(0)
        .LeftPadding = G_WordDoc.CentimetersToPoints(0)
        .RightPadding = G_WordDoc.CentimetersToPoints(0.2)
        .Spacing = G_WordDoc.CentimetersToPoints(0)
        
        .AllowPageBreaks = True
        .AllowAutoFit = False
        If Trim(ColFmt$) <> "" Then
            .Select
            .Rows.HeightRule = wdRowHeightExactly
            .Rows.Height = GetWordTextHeight(G_FontSize)
            SetWordColWidth Cols#, ColFmt$, SplitChar$
        End If
    End With
    
    With G_WordDoc.Documents(G_DocRptName).Application.Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = wdColorAutomatic
    End With
    
'    G_WordDoc.Documents(G_DocRptName).ActiveWindow.View. _
'         TableGridlines = False
    
    If InPageHeader% Then
       G_WordDoc.Documents(G_DocRptName).ActiveWindow.ActivePane. _
            View.SeekView = wdSeekMainDocument
    End If
End Sub

Sub SetWordDocStyle(ByRef FontName$, ByRef FontSize#, Optional ByVal Orientation% = 0, _
Optional ByVal TMargin# = 1.4, Optional ByVal BMargin# = 2, Optional ByVal LMargin# = 1.5, _
Optional ByVal HeaderDistance# = 1.5, Optional ByVal FooterDistance# = 1.75, _
Optional ByVal RMargin# = 1.5, Optional ByVal PageWidth# = 21, Optional ByVal PageHeight# = 29.7, _
Optional ByVal BmpHeight# = 1)
'��INI FileŪ�����e�X���]�w,����ܩ�Word�������t�m��ܤ��,
'�ϥΪ̥i��C�L�e�ۦ�ק�C�L�ȱi,��V,�r����,���������t�m��ܤ����,
'�Y�N�����t�m�]�w�s�JINI File��(���:����)
Dim A_IniPath$, A_Section$, A_Topic$
Dim A_Orientation$, A_TopMargin$, A_BottomMargin$, A_LeftMargin$
Dim A_RightMargin$, A_PageWidth$, A_PageHeight$
Dim A_HeaderDistance$, A_FooterDistance$
Dim A_FontName$, A_FontSize$, A_FontBold$, A_EnglishName$
Const wdAlignVerticalTop = 0
Const wdLayoutModeLineGrid = 2
Const wdStyleNormal = -1

    A_IniPath$ = G_INI_SerPath & "Data\" & App.EXEName & ".INI"
    A_Section$ = GetUserId()
    A_Topic$ = "DOC/" & App.EXEName & "/"
    
    A_Orientation$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "Orientation", CStr(Orientation%), A_Orientation$, 1000, A_IniPath$
    
    A_PageWidth$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "PageWidth", CStr(PageWidth#), A_PageWidth$, 1000, A_IniPath$
    
    A_PageHeight$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "PageHeight", CStr(PageHeight#), A_PageHeight$, 1000, A_IniPath$
    
    A_TopMargin$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "TopMargin", CStr(TMargin#), A_TopMargin$, 1000, A_IniPath$
    
    A_BottomMargin$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "BottomMargin", CStr(BMargin#), A_BottomMargin$, 1000, A_IniPath$
    
    A_LeftMargin$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "LeftMargin", CStr(LMargin#), A_LeftMargin$, 1000, A_IniPath$
    
    A_RightMargin$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "RightMargin", CStr(RMargin#), A_RightMargin$, 1000, A_IniPath$
    
    A_HeaderDistance$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "HeaderDistance", CStr(HeaderDistance#), A_HeaderDistance$, 1000, A_IniPath$
      
    A_FooterDistance$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "FooterDistance", CStr(FooterDistance#), A_FooterDistance$, 1000, A_IniPath$
  
    A_FontName$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "FontName", FontName$, A_FontName$, 1000, A_IniPath$
    
    A_FontSize$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "FontSize", CStr(FontSize#), A_FontSize$, 1000, A_IniPath$
    
    A_FontBold$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "FontBold", "0", A_FontBold$, 1000, A_IniPath$
    
    A_EnglishName$ = Space(1000)
    OSGetPrivateProfileString% A_Section$, A_Topic$ & "EnglishName", "", A_EnglishName$, 1000, A_IniPath$
    
    With G_WordDoc.Documents(G_DocRptName).Styles(wdStyleNormal).Font
         .NameFarEast = A_FontName$
         If Replace(Trim(A_EnglishName$), Chr(0), "") <> "" Then
            .NameAscii = A_EnglishName$
            .NameOther = A_EnglishName$
            .Name = A_EnglishName$
         End If
         .Size = CInt(A_FontSize$)
         .Bold = CBool(A_FontBold$)
         .Italic = False
         FontName$ = .Name
         FontSize# = .Size
    End With
     
    G_WordDoc.Documents(G_DocRptName).GridDistanceVertical = G_WordDoc.CentimetersToPoints(0.01)
    
    With G_WordDoc.Documents(G_DocRptName).PageSetup
         If BmpHeight# > CDbl(A_TopMargin$) - CDbl(A_HeaderDistance$) Then
            A_TopMargin$ = BmpHeight# + CDbl(A_HeaderDistance$) + 0.5
         End If
            
         .Orientation = CLng(A_Orientation$)
         .PageWidth = G_WordDoc.CentimetersToPoints(CSng(A_PageWidth$))
         .PageHeight = G_WordDoc.CentimetersToPoints(CSng(A_PageHeight$))
         .TopMargin = G_WordDoc.CentimetersToPoints(CSng(A_TopMargin$))
         .BottomMargin = G_WordDoc.CentimetersToPoints(CSng(A_BottomMargin$))
         .LeftMargin = G_WordDoc.CentimetersToPoints(CSng(A_LeftMargin$))
         .RightMargin = G_WordDoc.CentimetersToPoints(CSng(A_RightMargin$))
         .HeaderDistance = G_WordDoc.CentimetersToPoints(CSng(A_HeaderDistance$))
         .FooterDistance = G_WordDoc.CentimetersToPoints(CSng(A_FooterDistance$))

         Dim I%, A_Lines%
         A_Lines% = .LinesPage
         For I% = 1 To 50
            On Error Resume Next
            A_Lines% = A_Lines% + 1
            .LinesPage = A_Lines%
            If Err > 0 Then Exit For
         Next I%
         On Error GoTo 0

         .LayoutMode = wdLayoutModeLineGrid
    End With
    
    SaveDocStyle BmpHeight#
    
    G_DocSelection.TypeParagraph
End Sub

Sub ShowWordFile()
'�������,����Word����,�A�HShell Function�}�ҳ����ɮ�.
    If G_PrintSelect <> G_Print2Word Then Exit Sub
    CloseWordFile
    If Not G_HaveDataPrint% Then Exit Sub
    retcode = Shell(GetWordAppPath & " " & G_OutFile, vbMaximizedFocus)
End Sub

Sub SetWordNewPage(ByVal Continue$, ByVal StyleCopy%)
'��ʳ]�w�����Ÿ�����m
Const wdStory = 6
Const wdPageBreak = 7
Const wdAlignParagraphCenter = 1
Const wdAlignParagraphLeft = 0

'    If Not StyleCopy% Then SetDocTableAutoFit2Window
    
    With G_DocSelection
         .EndKey unit:=wdStory
         .TypeParagraph
         .TypeParagraph
         .ParagraphFormat.Alignment = wdAlignParagraphCenter
         .TypeText Continue$
         .InsertBreak Type:=wdPageBreak
         .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With
End Sub

Sub SetWordBodyHAlign(tSPD As Spread, ByVal TitleRows%)
'�]�w�C�L��Word��������Ҧ�
Dim I%, A_Align%, A_RIndex%, A_Rows%
Dim A_Range$

    With G_DocSelection
         A_Rows% = .Tables(1).Rows.Count
    End With
    
    For I% = 1 To UBound(tSPD.Columns)
        A_RIndex% = tSPD.Columns(I%).ReportIndex
        If A_RIndex% > 0 Then
           Select Case Left(tSPD.Columns(I%).dFormat, 1)
             Case "#"
                  A_Align% = xlLeft
             Case "~"
                  A_Align% = xlRight
             Case "^"
                  A_Align% = xlCenter
           End Select
           
           A_Range$ = Chr(A_RIndex% + 64) & CStr(TitleRows%) & ":" & _
                      Chr(A_RIndex% + 64) & CStr(A_Rows%)
           SetCellAlignment A_Range$, A_Align%, xlCenter, False
        End If
    Next I%
End Sub

Sub SetWordLineStyle(ByVal Range$, Optional ByVal LBorderStyle& = -1, _
Optional ByVal RBorderStyle& = -1, Optional ByVal TBorderStyle& = -1, _
Optional ByVal BBorderStyle& = -1, Optional ByVal HBorderStyle& = -1, _
Optional ByVal VBorderStyle& = -1, Optional ByVal LBorderWidth& = 4, _
Optional ByVal RBorderWidth& = 4, Optional ByVal TBorderWidth& = 4, _
Optional ByVal BBorderWidth& = 4, Optional ByVal HBorderWidth& = 4, _
Optional ByVal VBorderWidth& = 4, Optional ByVal BorderColor& = 0)
'�]�wWord�x�s�檺�ؽu�˦�

    If G_PrintSelect <> G_Print2Word Then Exit Sub

    SelectWordCells Range$
    
    With G_DocSelection.Cells
         'wdBorderLeft:���ؽu
         If LBorderStyle& <> -1 Then
            .Borders(-2).LineStyle = LBorderStyle&
            If LBorderStyle& <> 0 Then .Borders(-2).LineWidth = LBorderWidth&
            .Borders(-2).ColorIndex = BorderColor&
         End If
         
        'wdBorderRight:�k�ؽu
         If RBorderStyle& <> -1 Then
            .Borders(-4).LineStyle = RBorderStyle&
            If RBorderStyle& <> 0 Then .Borders(-4).LineWidth = RBorderWidth&
            .Borders(-4).ColorIndex = BorderColor&
         End If
         
         'wdBorderTop:�W�ؽu
         If TBorderStyle& <> -1 Then
            .Borders(-1).LineStyle = TBorderStyle&
            If TBorderStyle& <> 0 Then .Borders(-1).LineWidth = TBorderWidth&
            .Borders(-1).ColorIndex = BorderColor&
         End If
         
         'wdBorderBottom:�U�ؽu
         If BBorderStyle& <> -1 Then
            .Borders(-3).LineStyle = BBorderStyle&
            If BBorderStyle& <> 0 Then .Borders(-3).LineWidth = BBorderWidth&
            .Borders(-3).ColorIndex = BorderColor&
         End If
         
         'wdBorderHorizontal:�������ؽu
         If HBorderStyle& <> -1 Then
            .Borders(-5).LineStyle = HBorderStyle&
            If HBorderStyle& <> 0 Then .Borders(-5).LineWidth = HBorderWidth&
            .Borders(-5).ColorIndex = BorderColor&
         End If
         
         'wdBorderVertical:�������ؽu
         If VBorderStyle& <> -1 Then
            .Borders(-6).LineStyle = VBorderStyle&
            If VBorderStyle& <> 0 Then .Borders(-6).LineWidth = VBorderWidth&
            .Borders(-6).ColorIndex = BorderColor&
         End If
    End With
End Sub

Sub SaveDocStyle(ByVal BmpHeight#)
'�x�s����C�L��Word�������t�m�Ȩ�Data�U��AppEXEName.INI��
Dim A_IniPath$, A_Section$, A_Topic$
Dim A_TopMargin#, A_HeaderDistance#
Const wdWindowStateMinimize = 2
Const wdDialogFilePageSetup = 178
Const wdDialogFilePageSetupTabPaperSize = 150001
Const wdStyleNormal = -1

    With G_WordDoc
         .WindowState = wdWindowStateMinimize
         .Visible = True
         .Activate
         With .Dialogs(wdDialogFilePageSetup)
'��ܪ����t�m��ܤ��,�N�w�]�������d�b�ȱi�����W
              .DefaultTab = wdDialogFilePageSetupTabPaperSize

'��ܪ����t�m��ܤ��,�åH�Ǧ^�ȧP�_������ܤ���ɫ��@�U�����s
'0:�������s  -1:�T�w���s  -2:�������s
              retcode = .Show
         End With
    End With
    
    A_IniPath$ = G_INI_SerPath & "Data\" & App.EXEName & ".INI"
    A_Section$ = GetUserId()
    A_Topic$ = "DOC/" & App.EXEName & "/"
    
    With G_WordDoc.Documents(G_DocRptName).PageSetup
         A_TopMargin# = G_WordDoc.PointsToCentimeters(.TopMargin)
         A_HeaderDistance# = G_WordDoc.PointsToCentimeters(.HeaderDistance)
         
         If BmpHeight# > A_TopMargin# - A_HeaderDistance# Then
            A_TopMargin# = BmpHeight# + A_HeaderDistance# + 0.5
            .TopMargin = G_WordDoc.CentimetersToPoints(A_TopMargin#)
         End If
         
         UpdateIniValue A_Section$, A_Topic$ & "Orientation", CStr(.Orientation), A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "TopMargin", CStr(A_TopMargin#), A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "BottomMargin", CStr(G_WordDoc.PointsToCentimeters(.BottomMargin)), A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "LeftMargin", CStr(G_WordDoc.PointsToCentimeters(.LeftMargin)), A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "RightMargin", CStr(G_WordDoc.PointsToCentimeters(.RightMargin)), A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "HeaderDistance", CStr(A_HeaderDistance#), A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "FooterDistance", CStr(G_WordDoc.PointsToCentimeters(.FooterDistance)), A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "PageWidth", CStr(G_WordDoc.PointsToCentimeters(.PageWidth)), A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "PageHeight", CStr(G_WordDoc.PointsToCentimeters(.PageHeight)), A_IniPath$
    End With
     
    With G_WordDoc.Documents(G_DocRptName).Styles(wdStyleNormal).Font
         UpdateIniValue A_Section$, A_Topic$ & "FontName", .NameFarEast, A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "EnglishName", .Name, A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "FontSize", CStr(.Size), A_IniPath$
         UpdateIniValue A_Section$, A_Topic$ & "FontBold", CStr(.Bold), A_IniPath$
    End With
End Sub

Sub CopyDocPageStyle()
'�ƻsWord��󭶭����t�m�˦�
Const wdGoToPage = 1
Const wdGoToLast = -1
Const wdStory = 6
Const wdExtend = 1
Const wdCharacter = 1
Const wdGoToTable = 2
Const wdAutoFitContent = 1
Const wdAutoFitWindow = 2
Const wdWindowStateNormal = 0
    
    With G_DocSelection
         .Goto what:=wdGoToPage, which:=wdGoToLast
         .EndKey unit:=wdStory, Extend:=wdExtend
         .MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
         .Copy
         .Goto what:=wdGoToTable, which:=wdGoToLast
    End With
End Sub

Sub SelectWordCell(ByVal Row%, ByVal Col%)
'���Word�ثe������椤���Y�@���x�s��
Const wdLine = 5
Const wdCharacter = 1
Const wdCell = 12
Const wdRow = 10
Const wdMove = 0
Dim I%

    With G_DocSelection
'        'for Word 2000���Φ��覡
'        .Tables(1).Select
'        .MoveLeft Unit:=wdCharacter, Count:=1
'         .Application.ScreenRefresh
'
'        For I% = 1 To Row% - 1
'            .SelectRow
'            .MoveDown Unit:=wdLine, Count:=1
'        Next I%
'        '.MoveDown Unit:=wdLine, Count:=Row% - 1
'
'        .HomeKey Unit:=wdRow, Extend:=wdMove
'        If Col% <> 1 Then
'           .MoveRight Unit:=wdCell, Count:=Col% - 1
'        End If
'        .SelectCell
        
        'Word 2003�i�Ϊ��覡
         .Tables(1).Cell(Row%, Col%).Select
    End With
End Sub

Sub SelectWordCells(ByVal Range$)
'���Word���w�d��Cells
Dim A_Str1$, A_Str2$, A_Cols%, A_Cole%, A_Rows%, A_Rowe%
Const wdLine = 5
Const wdExtend = 1
Const wdSentence = 3
Const wdCharacter = 1
Const wdCell = 12

    StrCut Range$, ":", A_Str1$, A_Str2$
    If A_Str1$ <> "" Then A_Cols% = Asc(Left(A_Str1$, 1)) - 64
    If A_Str2$ <> "" Then A_Cole% = Asc(Left(A_Str2$, 1)) - 64
    A_Rows% = CDbl(Mid(A_Str1$, 2))
    If A_Str2$ <> "" Then A_Rowe% = CDbl(Mid(A_Str2$, 2))

    With G_DocSelection
         '����d��
         If A_Str2$ = "" Then
            SelectWordCell A_Rows%, A_Cols%
         Else
            If A_Rows% <> A_Rowe% Then
               SelectWordCell A_Rows%, A_Cols%
               .MoveDown unit:=wdLine, Count:=A_Rowe% - A_Rows%, Extend:=wdExtend
               If A_Cols% <> A_Cole% Then
                  .MoveRight unit:=wdSentence, Count:=A_Cole% - A_Cols%, Extend:=wdExtend
               End If
            Else
               If A_Cols% = 1 And A_Cole% = G_ExcelMaxCols Then
                  SelectWordCell A_Rows%, A_Cols%
                  If A_Cols% <> A_Cole% Then
                     .SelectRow
                  End If
               Else
                  SelectWordCell A_Rows%, A_Cols%
                  If A_Str2$ <> "" Then
                     If A_Cols% <> A_Cole% Then
                        .MoveRight unit:=wdCharacter, Count:=A_Cole% - A_Cols%, Extend:=wdExtend
                     End If
                  End If
               End If
            End If
         End If
    End With
End Sub

Function GetEndColofWordTableRow(ByVal Row%) As Integer
Dim A_EndCol%
Const wdGoToTable = 2
Const wdGoToLast = -1
Const wdLine = 5
Const wdMaximumNumberOfColumns = 18

    With G_DocSelection
''�N��в����檺�Ĥ@���x�s��
'         .GoTo what:=wdGoToTable, Which:=wdGoToLast
'
''??? �N��в�������w�C���Ĥ@���x�s��
'         .MoveDown Unit:=wdLine, Count:=Row% - 1
         
         SelectWordCell Row%, 1
         
'���o�ӦC���̫�@����쪺�Ǹ�
         A_EndCol% = .Information(wdMaximumNumberOfColumns)
    End With
    GetEndColofWordTableRow = A_EndCol%
End Function

Sub SetDocTableAutoFit2Window()
'�NWord�������۰ʽվ㦨�����j�p
Const wdGoToTable = 2
Const wdGoToLast = -1
Const wdCharacter = 1
Const wdAutoFitContent = 1
Const wdAutoFitWindow = 2

    With G_DocSelection
         .Goto what:=wdGoToTable, which:=wdGoToLast
         .Tables(1).Select
         .Tables(1).AutoFitBehavior (wdAutoFitContent)
         .Application.ScreenRefresh
         .Tables(1).AutoFitBehavior (wdAutoFitWindow)
         .MoveLeft unit:=wdCharacter, Count:=1
    End With
End Sub

Sub SetDocPageHeader(ByVal BmpFileName$, ByVal Height#, ByVal Width#, ByVal Align%)
'�bWord��󭶭�����J�Ϥ�
Const wdPaneNone = 0
Const wdNormalView = 1
Const wdOutlineView = 2
Const wdSeekCurrentPageHeader = 9
Const wdCharacter = 1
Const wdExtend = 1
Const wdPrintView = 3
Const wdSeekMainDocument = 0

    With G_WordDoc.Documents(G_DocRptName).ActiveWindow
        If .View.SplitSpecial <> wdPaneNone Then
            .Panes(2).Close
         End If
         
         With .ActivePane
              If .View.Type = wdNormalView Or .View.Type = wdOutlineView Then
                 .View.Type = wdPrintView
              End If
              .View.SeekView = wdSeekCurrentPageHeader
         End With
         
         With .Selection
              .InlineShapes.AddPicture FileName:=BmpFileName$, LinkToFile:=False, SaveWithDocument:=True
              .MoveLeft unit:=wdCharacter, Count:=1, Extend:=wdExtend
              .InlineShapes(1).Fill.Visible = False
              .InlineShapes(1).Fill.Transparency = 0
              .InlineShapes(1).Line.Weight = 0.75
              .InlineShapes(1).Line.Transparency = 0
              .InlineShapes(1).Line.Visible = False
              .InlineShapes(1).LockAspectRatio = False
              .InlineShapes(1).Height = .Application.CentimetersToPoints(Height#)
              .InlineShapes(1).Width = .Application.CentimetersToPoints(Width#)
              .InlineShapes(1).PictureFormat.Brightness = 0.5
              .InlineShapes(1).PictureFormat.Contrast = 0.5
              .InlineShapes(1).PictureFormat.ColorType = 1 '�۰ʿ��
              .InlineShapes(1).PictureFormat.CropLeft = 0
              .InlineShapes(1).PictureFormat.CropRight = 0
              .InlineShapes(1).PictureFormat.CropTop = 0
              .InlineShapes(1).PictureFormat.CropBottom = 0
              .ParagraphFormat.Alignment = Align%
         End With
         
         .ActivePane.View.SeekView = wdSeekMainDocument
    End With
End Sub

Sub SetWordCellText(ByVal Row%, ByVal Col%, ByVal text$)
'�]�wWord����x�s�椺����r
    
    With G_DocSelection
         .Tables(1).Cell(Row%, Col%).WordWrap = False
         .Tables(1).Cell(Row%, Col%).Range.text = text$
    End With
End Sub

Sub CloseWordFile()
'����Word�ɮ�
On Local Error Resume Next
    
    G_WordDoc.ScreenUpdating = True
    Select Case G_WordDoc.Documents.Count
      Case 0
           G_WordDoc.Quit
      Case 1
           G_WordDoc.ActiveDocument.Close savechanges:=False
           G_WordDoc.Quit
      Case Else
           G_WordDoc.Documents(G_DocRptName).Close savechanges:=False
    End Select
    Set G_WordDoc = Nothing
End Sub

Function OpenWordFile(ByVal FileName$, Optional ByVal Password$ = "", _
Optional ByVal SaveAs2003Format As Boolean = False) As Boolean
'�إ�Word���
On Local Error GoTo MY_Error
Const wdAlertsNone = 0
Const wdAutoOpen = 2
Const wdPrintView = 3
Const wdWindowStateMinimize = 2
Const wdDialogFilePageSetup = 178
Const wdDialogFilePageSetupTabPaperSize = 150001
Dim A_Msg$
    
    OpenWordFile = True
    '
    CloseWordFile
    Set G_WordDoc = CreateObject("Word.Application")
    G_WordDoc.DisplayAlerts = wdAlertsNone
    
    If Dir(FileName$) <> "" Then Kill FileName$
    
    '��Word������2007�H�W��, �N�ɮץt�s�� 97-2003 Word Format (FileFormat=0)
    If Trim(Password$) <> "" Then
       If SaveAs2003Format And Val(G_WordDoc.Version) >= 12 Then
          G_WordDoc.Documents.Add.SaveAs FileName$, 0, , Password$
       Else
          G_WordDoc.Documents.Add.SaveAs FileName$, , , Password$
       End If
    Else
       If SaveAs2003Format And Val(G_WordDoc.Version) >= 12 Then
          G_WordDoc.Documents.Add.SaveAs FileName$, 0
       Else
          G_WordDoc.Documents.Add.SaveAs FileName$
       End If
    End If

    G_DocRptName = Dir(FileName$)
    Set G_DocSelection = G_WordDoc.Documents(G_DocRptName). _
    ActiveWindow.Selection
    
    With G_WordDoc
         .Documents(G_DocRptName).RunAutoMacro wdAutoOpen
         .Parent.DisplayAlerts = False    '�����ܥ���ĵ�i
         .Documents(G_DocRptName).ActiveWindow.View = wdPrintView
         .ScreenUpdating = False
    End With
    Exit Function

MY_Error:
    OpenWordFile = False
    Select Case Err
    'PgmMsg  file_inuse    �ɮץ��b�ϥΤ�,�Эק��ɦW��,�A����C�L!
      Case 70   'Permission denied
           A_Msg$ = GetCaption("PgmMsg", "file_inuse", _
           "�ɮץ��b�ϥΤ�,�Эק��ɦW��,�A����C�L!")
           MsgBox A_Msg$, vbExclamation, App.Title
      Case Else
           MsgBox Error$, vbExclamation, App.Title
    End Select
    CloseWordFile
End Function

Sub AddXlsFldDataType(FldType(), ByVal ColIndex%, ByVal ColType%)
'�ʺA�[�JExcel��쪺��ƫ��A��Array
Dim A_Max%

    If G_PrintSelect <> G_Print2Excel Then Exit Sub
    
    FldType(ColIndex% - 1, 0) = ColIndex%
    '�]�w����ƫ��A
     Select Case ColType%
       Case G_Data_Numeric
            FldType(ColIndex% - 1, 1) = 1
       Case G_Data_String
            FldType(ColIndex% - 1, 1) = 2
       Case G_Data_Date
            Select Case G_DateFlag
              Case 0, 2
                    FldType(ColIndex% - 1, 1) = 5  'yyyy/m/d
              Case 1
                    '�Y��OS������(�x�W)���B�]�w�ҥΰ�����榡(EMD)��,�ϥΰ�����榡,�_�h�]����r�榡.
                    If IsWinForTaiwan = True And XlsFldUseChinaDate = True Then
                        FldType(ColIndex% - 1, 1) = 10 'yy/m/d
                    Else
                        FldType(ColIndex% - 1, 1) = 2
                    End If
            End Select
    End Select
End Sub

'===============================================================================
' Add New Function at 93/6/25
'===============================================================================
Function GetWordTextHeight(ByVal FontSize%) As Single
'���oWORD��󤤫��w�r���j�p�Ҧ����C��
Dim I%, A_LineHeight!

    A_LineHeight! = 0
    For I% = 1 To UBound(G_DocFontSize)
        If CInt(G_DocFontSize(I%, 1)) = FontSize% Then
            A_LineHeight! = CSng(G_DocFontSize(I%, 2))
            Exit For
        End If
    Next I%
    GetWordTextHeight = G_WordDoc.CentimetersToPoints(A_LineHeight!)
End Function

Function GetWordTextLines(ByVal rptfontsize%, ByVal LargeFontSize%) As Integer
'���oWORD����j�r���j�p��Table Row Height,������r���j�p���B�~���Ϊ��C��
Dim I%, A_RptLineHeight!, A_LargeLineHeight!

    A_RptLineHeight! = GetWordTextHeight(rptfontsize%)
    A_LargeLineHeight! = GetWordTextHeight(LargeFontSize%)
    GetWordTextLines = Abs(Int((A_LargeLineHeight! - A_RptLineHeight!) / A_RptLineHeight! * -1))
End Function

Sub SetDOCUseFontSize(ByVal ary As Variant)
'�NWORD��󤤷|�ϥΨ쪺�r���j�p�Ψ�Ҧ����C��KEEP��G_DocFontSize Array��
Dim I%, A_Size, A_LineStart!, A_LineStart2!
Const wdVerticalPositionRelativeToPage = 6
Const wdLine = 5
Const wdCharacter = 1
Const wdExtend = 1

    Erase G_DocFontSize
    ReDim G_DocFontSize(1 To UBound(ary) - LBound(ary) + 1, 1 To 2)
    
    With G_DocSelection
        .TypeParagraph
        A_LineStart! = G_WordDoc.PointsToCentimeters(.Information(wdVerticalPositionRelativeToPage))
        .TypeText text:="1"
        .TypeParagraph
        For Each A_Size In ary
            I% = I% + 1
            .MoveUp unit:=wdLine, Count:=1
            .HomeKey unit:=wdLine
            .MoveRight unit:=wdCharacter, Count:=1, Extend:=wdExtend
            .Font.Size = A_Size
            .MoveDown unit:=wdLine, Count:=1
            A_LineStart2! = G_WordDoc.PointsToCentimeters(.Information(wdVerticalPositionRelativeToPage))
            G_DocFontSize(I%, 1) = A_Size
            G_DocFontSize(I%, 2) = Format(A_LineStart2! - A_LineStart!, "0.0")
        Next
        .MoveUp unit:=wdLine, Count:=1
        .HomeKey unit:=wdLine
        .MoveDown unit:=wdLine, Count:=2, Extend:=wdExtend
        .Delete unit:=wdCharacter, Count:=1
    End With
End Sub

Function GetDOCPageLines() As Integer
'���oWORD���@���i�H�ϥΪ��`�C��

    With G_DocSelection.PageSetup
         GetDOCPageLines = Int((.PageHeight - .TopMargin - .BottomMargin) / GetWordTextHeight(G_FontSize))
    End With
End Function

Sub SetDocLineFont(ByVal FontSize%, Optional ByVal FontBold% = False, Optional ByVal FontName$ = "")
'�]�wWORD��󤤥ثe�C���r���j�p�ΦC��
    
    With G_DocSelection
         .Font.Size = FontSize%
         .Font.Bold = FontBold%
         If Trim(FontName$) <> "" Then .Font.Name = FontName$
         .Rows.Height = GetWordTextHeight(FontSize%)
    End With
End Sub

Sub SetWordColWidth(ByVal Cols#, ByVal ColFmt$, ByVal SplitChar$)
'�]�wWORD������e��
Dim I#, A_Share%, a_percent%, A_SPercent%, A_Cols$()
Dim A_TotalLen%, A_CurPercent%
Const wdCharacter = 1
Const wdColumn = 9
Const wdPreferredWidthPercent = 2

    ColFmt$ = Trim(ColFmt$)
    If Right(ColFmt$, 1) = SplitChar$ Then
        ColFmt$ = Left(ColFmt$, Len(ColFmt$) - 1)
    End If
    A_Cols$ = Split(Trim(ColFmt$), SplitChar$)
    A_Share% = (UBound(A_Cols$) + 1 <> Cols#)
    If Not A_Share% Then
        For I# = 1 To Cols#
            A_TotalLen% = A_TotalLen% + Len(Trim(A_Cols$(I# - 1)))
        Next I#
    Else
        A_SPercent% = CInt(Format(100 / Cols#, "0"))
    End If
    
    With G_DocSelection
         .Tables(1).Select
         For I# = 1 To Cols#
             If I# = 1 Then
                .MoveLeft unit:=wdCharacter, Count:=1
             Else
                .Move unit:=wdColumn, Count:=1
             End If
             .SelectColumn
             .Columns.PreferredWidthType = wdPreferredWidthPercent
             If I# = Cols# Then
                a_percent% = 100 - A_CurPercent%
             Else
             
                If A_Share% Then
                    a_percent% = A_SPercent%
                Else
                    a_percent% = CInt(Format(Len(Trim(A_Cols$(I# - 1))) / A_TotalLen% * 100, "0"))
                End If
             End If
             .Columns.PreferredWidth = a_percent%
             A_CurPercent% = A_CurPercent% + a_percent%
         Next I#
    End With
End Sub

Sub SetDocPrintInfoFixedColWidth(ByVal Fmt$)
'�]�wWord����,�C�L����ϰ��檺�T�w���e��
Const wdHorizontalPositionRelativeToTextBoundary = 7
Const wdLine = 5
Const wdAutoFitFixed = 0
Const wdPreferredWidthPoints = 3
Dim A_TextLen%, A_TextWidth!

    With G_DocSelection
         SelectWordCell 1, 1
         A_TextLen% = lstrlen(Replace(.text, vbCr + Chr(7), ""))
         .EndKey unit:=wdLine
         A_TextWidth! = .Information(wdHorizontalPositionRelativeToTextBoundary)
         
         .SelectColumn
         .Tables(1).AutoFitBehavior (wdAutoFitFixed)
         .Columns.PreferredWidthType = wdPreferredWidthPoints
         .Columns.PreferredWidth = Len(Fmt$) * A_TextWidth! / A_TextLen%
    End With
End Sub

'==================================================================================================================
'�]�w�w�]�L��� 93/10/1 (Start)
'==================================================================================================================
Public Sub RestoreDefaultPrinter(ByVal PrinterName As String)
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer

    If Trim(PrinterName) = "" Then Exit Sub
    
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
        Call Win95SetDefaultPrinter(PrinterName)
    Else
    ' This assumes that future versions of Windows use the NT method
        Call WinNTSetDefaultPrinter(PrinterName)
    End If
End Sub

Private Sub SelectPrinter(NewPrinter As String)
Dim Prt As Printer
    
    For Each Prt In Printers
        If Prt.DeviceName = NewPrinter Then
            Set Printer = Prt
            Exit For
        End If
    Next
End Sub

Private Function PtrCtoVbString(Add As Long) As String
Dim sTemp As String * 512, x As Long

    x = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Private Sub SetDefaultPrinter(ByVal PrinterName As String, ByVal DriverName As String, ByVal PrinterPort As String)
Dim DeviceLine As String
Dim r As Long
Dim L As Long

    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

Private Sub Win95SetDefaultPrinter(ByVal PrinterName As String)
Dim Handle As Long          'handle to printer
Dim pd As PRINTER_DEFAULTS
Dim x As Long
Dim need As Long            ' bytes needed
Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
Dim LastError As Long
Dim t() As Long

    
    ' none - exit
    If PrinterName = "" Then Exit Sub

    ' set the PRINTER_DEFAULTS members
    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

    ' Get a handle to the printer
    x = OpenPrinter(PrinterName, Handle, pd)
    ' failed the open
    If x = False Then Exit Sub

    ' Make an initial call to GetPrinter, requesting Level 5
    ' (PRINTER_INFO_5) information, to determine how many bytes
    ' you need
    x = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    ' don't want to check Err.LastDllError here - it's supposed
    ' to fail
    ' with a 122 - ERROR_INSUFFICIENT_BUFFER
    ' redim t as large as you need
    ReDim t((need \ 4)) As Long

    ' and call GetPrinter for keepers this time
    x = GetPrinter(Handle, 5, t(0), need, need)
    ' failed the GetPrinter
    If x = False Then Exit Sub

    ' set the members of the pi5 structure for use with SetPrinter.
    ' PtrCtoVbString copies the memory pointed at by the two string
    ' pointers contained in the t() array into a Visual Basic string.
    ' The other three elements are just DWORDS (long integers) and
    ' don't require any conversion
    pi5.pPrinterName = PtrCtoVbString(t(0))
    pi5.pPortName = PtrCtoVbString(t(1))
    pi5.Attributes = t(2)
    pi5.DeviceNotSelectedTimeout = t(3)
    pi5.TransmissionRetryTimeout = t(4)

    ' this is the critical flag that makes it the default printer
    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

    ' call SetPrinter to set it
    x = SetPrinter(Handle, 5, pi5, 0)

    If x = False Then   ' SetPrinter failed
        MsgBox "SetPrinter Failed. Error code: " & Err.LastDllError
        Exit Sub
    Else
        If Printer.DeviceName <> PrinterName Then
        ' Make sure Printer object is set to the new printer
             SelectPrinter (PrinterName)
        End If
    End If

    ' and close the handle
    ClosePrinter (Handle)
End Sub

Private Sub WinNTSetDefaultPrinter(ByVal PrinterName As String)
Dim Buffer As String
Dim DeviceName As String
Dim DriverName As String
Dim PrinterPort As String
Dim r As Long

    ' Get the printer information for the currently selected
    ' printer in the list. The information is taken from the
    ' WIN.INI file.
    Buffer = Space(1024)
    r = GetProfileString("PrinterPorts", PrinterName, "", Buffer, Len(Buffer))

    ' Parse the driver name and port name out of the buffer
    GetDriverAndPort Buffer, DriverName, PrinterPort

    If DriverName <> "" And PrinterPort <> "" Then
        SetDefaultPrinter PrinterName, DriverName, PrinterPort
        If Printer.DeviceName <> PrinterName Then
        ' Make sure Printer object is set to the new printer
           SelectPrinter (PrinterName)
        End If
    End If
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As String, PrinterPort As String)
Dim iDriver As Integer
Dim iPort As Integer

    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, iPort - iDriver - 1)
        End If
    End If
End Sub
'==================================================================================================================
'�]�w�w�]�L��� 93/10/1 (End)
'==================================================================================================================





