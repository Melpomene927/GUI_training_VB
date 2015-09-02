Attribute VB_Name = "mod_Common"
Option Explicit
Option Compare Text
  
Declare Function OSGetPrivateProfileString% Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnString$, ByVal NumBytes As Integer, ByVal FileName$)
Declare Function OSWritePrivateProfileString% Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Declare Function SpreadGetText Lib "Spread25.OCX" (SS As Control, ByVal Col As Long, ByVal Row As Long, Var As Variant) As Integer
Declare Function SpreadSetText Lib "Spread25.OCX" (ByVal Col As Long, ByVal Row As Long, lpVar As Variant)
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lprect As RECT) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'Get OS Language Id
Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer

'Open Folder Browser and Get Path  ---20091112 Add By Yvonne---
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

' Registry access API
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

' Registry constants
Public Const ERROR_SUCCESS = 0
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const KEY_ALL_ACCESS = &HF003F     ' ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const REG_SZ = 1
Public Const SYNCHRONIZE = &H100000
Public Const KEY_QUERY_VALUE = &H1
Public Const READ_CONTROL = &H20000
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Public Const gsODBCINST_INI_REG_KEY = "Software\ODBC\ODBCINST.INI"
Public Const gsODBC_INI_REG_KEY = "Software\ODBC\ODBC.INI"      ' Registry path to DSNs
Public Const glMAX_NAME_LENGTH As Long = 250   ' Max length for a server name
Public Const E_UNEXPECTED = &H8000FFFF

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type APPBARDATA
        cbSize As Long
        hwnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lparam As Long
End Type

Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    FType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

'Open Folder Browser ---20091112 Add By Yvonne---
Public Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lparam As Long
   iImage As Long
End Type
 
Global Const SC_CLOSE = &HF060&
Global Const SC_RESTORE = &HF120&
Global Const xMenuID = 10&    ' �ΨӴ��N SC_CLOSE,SC_SIZE.. �� Menu ID
Global Const SC_MOVE = &HF010&
Global Const SC_SIZE = &HF000&
Global Const SC_MINIMIZE = &HF020&
Global Const SC_MAXIMIZE = &HF030&
Global Const MF_BYCOMMAND = &H0&
Global Const MIIM_STATE = &H1&
Global Const MIIM_ID = &H2&
Global Const MFS_GRAYED = &H3&
Global Const MFS_CHECKED = &H8&
Global Const WM_NCACTIVATE = &H86

Global Const INFINITE = &HFFFF      '  Infinite timeout
Global Const EM_GETLINE = 196
Global Const EM_GETLINECOUNT = 186
Global Const STILL_ACTIVE = &H103
Global Const PROCESS_QUERY_INFORMATION = &H400

'Fields Data Type
Global Const G_Data_Numeric = 1
Global Const G_Data_String = 2
Global Const G_Data_Date = 3
Global Const G_Data_Float = 4 'Add By Lidia (S021024037)
Global Const G_Data_VarBinary = 5 'Add By Lidia (S021024037 For SQL��Ʈw�~�i�H��)
Global Const G_Data_uniqueidentifier = 6 'Add By Lidia (S021024037 For SQL��Ʈw�~�i�H��)

' Show parameters
Global Const MODAL = 1
Global Const MODELESS = 0

' Key Codes
Global Const KEY_LBUTTON = &H1
Global Const KEY_RBUTTON = &H2
Global Const KEY_CANCEL = &H3
Global Const KEY_MBUTTON = &H4
Global Const KEY_BACK = &H8
Global Const KEY_TAB = &H9
Global Const KEY_CLEAR = &HC
Global Const KEY_RETURN = &HD
Global Const KEY_SHIFT = &H10
Global Const KEY_CONTROL = &H11
Global Const KEY_MENU = &H12
Global Const KEY_PAUSE = &H13
Global Const KEY_CAPITAL = &H14
Global Const KEY_ESCAPE = &H1B
Global Const KEY_SPACE = &H20
Global Const KEY_PRIOR = &H21
Global Const KEY_NEXT = &H22
Global Const KEY_END = &H23
Global Const KEY_HOME = &H24
Global Const KEY_LEFT = &H25
Global Const KEY_UP = &H26
Global Const KEY_RIGHT = &H27
Global Const KEY_DOWN = &H28
Global Const KEY_SELECT = &H29
Global Const KEY_PRINT = &H2A
Global Const KEY_EXECUTE = &H2B
Global Const KEY_SNAPSHOT = &H2C
Global Const KEY_INSERT = &H2D
Global Const KEY_DELETE = &H2E
Global Const KEY_HELP = &H2F
Global Const KEY_NUMPAD0 = &H60
Global Const KEY_NUMPAD1 = &H61
Global Const KEY_NUMPAD2 = &H62
Global Const KEY_NUMPAD3 = &H63
Global Const KEY_NUMPAD4 = &H64
Global Const KEY_NUMPAD5 = &H65
Global Const KEY_NUMPAD6 = &H66
Global Const KEY_NUMPAD7 = &H67
Global Const KEY_NUMPAD8 = &H68
Global Const KEY_NUMPAD9 = &H69
Global Const KEY_MULTIPLY = &H6A
Global Const KEY_ADD = &H6B
Global Const KEY_SEPARATOR = &H6C
Global Const KEY_SUBTRACT = &H6D
Global Const KEY_DECIMAL = &H6E
Global Const KEY_DIVIDE = &H6F
Global Const KEY_F1 = &H70
Global Const KEY_F2 = &H71
Global Const KEY_F3 = &H72
Global Const KEY_F4 = &H73
Global Const KEY_F5 = &H74
Global Const KEY_F6 = &H75
Global Const KEY_F7 = &H76
Global Const KEY_F8 = &H77
Global Const KEY_F9 = &H78
Global Const KEY_F10 = &H79
Global Const KEY_F11 = &H7A
Global Const KEY_F12 = &H7B
Global Const KEY_F13 = &H7C
Global Const KEY_F14 = &H7D
Global Const KEY_F15 = &H7E
Global Const KEY_F16 = &H7F
Global Const KEY_NUMLOCK = &H90

' Colors
Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF
                  
' System Colors
Global Const SCROLL_BARS = &H80000000           ' Scroll-bars gray area.
Global Const DESKTOP = &H80000001               ' Desktop.
Global Const ACTIVE_TITLE_BAR = &H80000002      ' Active window caption.
Global Const INACTIVE_TITLE_BAR = &H80000003    ' Inactive window caption.
Global Const MENU_BAR = &H80000004              ' Menu background.
Global Const WINDOW_BACKGROUND = &H80000005     ' Window background.
Global Const WINDOW_FRAME = &H80000006          ' Window frame.
Global Const MENU_TEXT = &H80000007             ' Text in menus.
Global Const WINDOW_TEXT = &H80000008           ' Text in windows.
Global Const TITLE_BAR_TEXT = &H80000009        ' Text in caption, size box, scroll-bar arrow box..
Global Const ACTIVE_BORDER = &H8000000A         ' Active window border.
Global Const INACTIVE_BORDER = &H8000000B       ' Inactive window border.
Global Const APPLICATION_WORKSPACE = &H8000000C ' Background color of multiple document interface (MDI) applications.
Global Const HIGHLIGHT = &H8000000D             ' Items selected item in a control.
Global Const HIGHLIGHT_TEXT = &H8000000E        ' Text of item selected in a control.
Global Const BUTTON_FACE = &H8000000F           ' Face shading on command buttons.
Global Const BUTTON_SHADOW = &H80000010         ' Edge shading on command buttons.
Global Const GRAY_TEXT = &H80000011             ' Grayed (disabled) text.  This color is set to 0 if the current display driver does not support a solid gray color.
Global Const BUTTON_TEXT = &H80000012           ' Text on push buttons.

' MousePointer
Global Const Default = 0        ' 0 - Default
Global Const ARROW = 1          ' 1 - Arrow
Global Const CROSSHAIR = 2      ' 2 - Cross
Global Const IBEAM = 3          ' 3 - I-Beam
Global Const ICON_POINTER = 4   ' 4 - Icon
Global Const SIZE_POINTER = 5   ' 5 - Size
Global Const SIZE_NE_SW = 6     ' 6 - Size NE SW
Global Const SIZE_N_S = 7       ' 7 - Size N S
Global Const SIZE_NW_SE = 8     ' 8 - Size NW SE
Global Const SIZE_W_E = 9       ' 9 - Size W E
Global Const UP_ARROW = 10      ' 10 - Up Arrow
Global Const HOURGLASS = 11     ' 11 - Hourglass
Global Const NO_DROP = 12       ' 12 - No drop

' WindowState
Global Const NORMAL = 0    ' 0 - Normal
Global Const MINIMIZED = 1 ' 1 - Minimized
Global Const MAXIMIZED = 2 ' 2 - Maximized

' Shift parameter masks
Global Const SHIFT_MASK = 1
Global Const CTRL_MASK = 2
Global Const ALT_MASK = 4

' Button parameter masks
Global Const LEFT_BUTTON = 1
Global Const RIGHT_BUTTON = 2
Global Const MIDDLE_BUTTON = 4

' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

Global Const MB_APPLMODAL = 0          ' Application Modal Message Box
Global Const MB_DEFBUTTON1 = 0         ' First button is default
Global Const MB_DEFBUTTON2 = 256       ' Second button is default
Global Const MB_DEFBUTTON3 = 512       ' Third button is default
Global Const MB_SYSTEMMODAL = 4096      'System Modal

' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed

' SetAttr, Dir, GetAttr functions
Global Const ATTR_NORMAL = 0
Global Const ATTR_READONLY = 1
Global Const ATTR_HIDDEN = 2
Global Const ATTR_SYSTEM = 4
Global Const ATTR_VOLUME = 8
Global Const ATTR_DIRECTORY = 16
Global Const ATTR_ARCHIVE = 32

'Common Dialog Control
Global Const DLG_FILE_OPEN = 1
Global Const DLG_FILE_SAVE = 2
Global Const DLG_COLOR = 3
Global Const DLG_FONT = 4
Global Const DLG_PRINT = 5
Global Const DLG_HELP = 6

'Fonts Dialog Flags
Global Const CF_SCREENFONTS = &H1&
Global Const CF_PRINTERFONTS = &H2&
Global Const CF_BOTH = &H3&
Global Const CF_SHOWHELP = &H4&
Global Const CF_INITTOLOGFONTSTRUCT = &H40&
Global Const CF_USESTYLE = &H80&
Global Const CF_EFFECTS = &H100&
Global Const CF_APPLY = &H200&
Global Const CF_ANSIONLY = &H400&
Global Const CF_NOVECTORFONTS = &H800&
Global Const CF_NOSIMULATIONS = &H1000&
Global Const CF_LIMITSIZE = &H2000&
Global Const CF_FIXEDPITCHONLY = &H4000&
Global Const CF_WYSIWYG = &H8000&
Global Const CF_FORCEFONTEXIST = &H10000
Global Const CF_SCALABLEONLY = &H20000
Global Const CF_TTONLY = &H40000
Global Const CF_NOFACESEL = &H80000
Global Const CF_NOSTYLESEL = &H100000
Global Const CF_NOSIZESEL = &H200000

'Printer Dialog Flags
Global Const PD_ALLPAGES = &H0&
Global Const PD_SELECTION = &H1&
Global Const PD_PAGENUMS = &H2&
Global Const PD_NOSELECTION = &H4&
Global Const PD_NOPAGENUMS = &H8&
Global Const PD_COLLATE = &H10&
Global Const PD_PRINTTOFILE = &H20&
Global Const PD_PRINTSETUP = &H40&
Global Const PD_NOWARNING = &H80&
Global Const PD_RETURNDC = &H100&
Global Const PD_RETURNIC = &H200&
Global Const PD_RETURNDEFAULT = &H400&
Global Const PD_SHOWHELP = &H800&
Global Const PD_USEDEVMODECOPIES = &H40000
Global Const PD_DISABLEPRINTTOFILE = &H80000
Global Const PD_HIDEPRINTTOFILE = &H100000

'Colors
Global Const G_BLACK = 0
Global Const G_BLUE = 1
Global Const G_GREEN = 2
Global Const G_CYAN = 3
Global Const G_RED = 4
Global Const G_MAGENTA = 5
Global Const G_BROWN = 6
Global Const G_LIGHT_GRAY = 7
Global Const G_DARK_GRAY = 8
Global Const G_LIGHT_BLUE = 9
Global Const G_LIGHT_GREEN = 10
Global Const G_LIGHT_CYAN = 11
Global Const G_LIGHT_RED = 12
Global Const G_LIGHT_MAGENTA = 13
Global Const G_YELLOW = 14
Global Const G_WHITE = 15
Global Const G_AUTOBW = 16

'Key Status Control
Global Const KEYSTAT_CAPSLOCK = 0
Global Const KEYSTAT_NUMLOCK = 1
Global Const KEYSTAT_INSERT = 2
Global Const KEYSTAT_SCROLLLOCK = 3

' Field Data Types
Global Const DB_BOOLEAN = 1
Global Const DB_BYTE = 2
Global Const DB_INTEGER = 3
Global Const DB_LONG = 4
Global Const DB_CURRENCY = 5
Global Const DB_SINGLE = 6
Global Const DB_DOUBLE = 7
Global Const DB_DATE = 8
Global Const DB_TEXT = 10
Global Const DB_LONGBINARY = 11
Global Const DB_MEMO = 12

' Option argument values (OpenRecordset, etc)
Global Const DB_DENYWRITE = &H1
Global Const DB_DENYREAD = &H2
Global Const DB_READONLY = &H4
Global Const DB_APPENDONLY = &H8
Global Const DB_INCONSISTENT = &H10
Global Const DB_CONSISTENT = &H20
Global Const DB_SQLPASSTHROUGH = &H40

'spreadsheet actions
Global Const SS_ACTION_ACTIVE_CELL = 0
Global Const SS_ACTION_GOTO_CELL = 1
Global Const SS_ACTION_SELECT_BLOCK = 2
Global Const SS_ACTION_CLEAR = 3
Global Const SS_ACTION_DELETE_COL = 4
Global Const SS_ACTION_DELETE_ROW = 5
Global Const SS_ACTION_INSERT_COL = 6
Global Const SS_ACTION_INSERT_ROW = 7
Global Const SS_ACTION_LOAD_SPREAD_SHEET = 8
Global Const SS_ACTION_SAVE_ALL = 9
Global Const SS_ACTION_SAVE_VALUES = 10
Global Const SS_ACTION_RECALC = 11
Global Const SS_ACTION_CLEAR_TEXT = 12
Global Const SS_ACTION_PRINT = 29
Global Const SS_ACTION_DESELECT_BLOCK = 14
Global Const SS_ACTION_DSAVE = 15
Global Const SS_ACTION_SET_CELL_BORDER = 16
Global Const SS_ACTION_ADD_MULTISELBLOCK = 17
Global Const SS_ACTION_GET_MULTI_SELECTION = 18
Global Const SS_ACTION_COPY_RANGE = 19
Global Const SS_ACTION_MOVE_RANGE = 20
Global Const SS_ACTION_SWAP_RANGE = 21
Global Const SS_ACTION_CLIPBOARD_COPY = 22
Global Const SS_ACTION_CLIPBOARD_CUT = 23
Global Const SS_ACTION_CLIPBOARD_PASTE = 24
Global Const SS_ACTION_SORT = 25
Global Const SS_ACTION_COMBO_CLEAR = 26
Global Const SS_ACTION_COMBO_REMOVE = 27
Global Const SS_ACTION_RESET = 28
Global Const SS_ACTION_SS_ACTION_SEL_MODE_CLEAR = 29
Global Const SS_ACTION_VMODE_REFRESH = 30
Global Const SS_ACTION_REFRESH_BOUND = 31
Global Const SS_ACTION_SMARTPRINT = 32

'cell type
Global Const SS_CELL_TYPE_DATE = 0
Global Const SS_CELL_TYPE_EDIT = 1
Global Const SS_CELL_TYPE_FLOAT = 2
Global Const SS_CELL_TYPE_INTEGER = 3
Global Const SS_CELL_TYPE_PIC = 4
Global Const SS_CELL_TYPE_STATIC_TEXT = 5
Global Const SS_CELL_TYPE_TIME = 6
Global Const SS_CELL_TYPE_BUTTON = 7
Global Const SS_CELL_TYPE_COMBOBOX = 8
Global Const SS_CELL_TYPE_PICTURE = 9
Global Const SS_CELL_TYPE_CHECKBOX = 10
Global Const SS_CELL_TYPE_OWNER_DRAWN = 11

'Spread Sort
Global Const SS_SORT_BY_ROW = 0
Global Const SS_SORT_BY_COL = 1
Global Const SS_SORT_ORDER_NONE = 0
Global Const SS_SORT_ORDER_ASCENDING = 1
Global Const SS_SORT_ORDER_DESCENDING = 2

'EditmodeAction
Global Const SS_CELL_EDITMODE_EXIT_NONE = 0
Global Const SS_CELL_EDITMODE_EXIT_UP = 1
Global Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Global Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Global Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Global Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Global Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6

'Cell Text Align
Global Const SS_CELL_H_ALIGN_LEFT = 0
Global Const SS_CELL_H_ALIGN_RIGHT = 1
Global Const SS_CELL_H_ALIGN_CENTER = 2

'TabOrientation (ssdesignerwidgetstabs.ssindextab)
Global Const SS_TABS_TOP = 0                         '0   (Default) Tabs on Top
Global Const SS_TABS_BOTTOM = 1                      '1   Tabs on Bottom
Global Const SS_TABS_LEFT = 2                        '2   Tabs on Left
Global Const SS_TABS_RIGHT = 3                       '3   Tabs on Right

Global Const G_G1 = ""                         'Field Seperator for Report Use
Global G_NUM As String
Global G_CHR As String
Global Const G_Program_Start = 1
Global Const G_Program_End = 2
Global Const G_Program_Add = 3
Global Const G_Program_Delete = 4
Global Const G_Program_Update = 5

'Process Status
Global Const G_AP_STATE_NORMAL = 0
Global Const G_AP_STATE_ADD = 1
Global Const G_AP_STATE_DELETE = 2
Global Const G_AP_STATE_UPDATE = 3
Global Const G_AP_STATE_QUERY = 4
Global Const G_AP_STATE_NEW = 5
Global Const G_AP_STATE_PRINT = 6
Global Const G_AP_STATE_COPY = 7
Global Const G_AP_STATE_NODATA = 9
Global Const G_AP_STATE_TABLE = 10

'Color
Global Const COLOR_GRAY = &HC0C0C0              'Gray
Global Const COLOR_SKY = &HFFFF00               'sky (for Vista : �쬰&HFFFF80,�D�зǦ�)
Global Const COLOR_WHITE = &HFFFFFF             'white
Global Const COLOR_GREEN = &H808000             'green
Global Const COLOR_YELLOW = &HFFFF&             'Yellow
Global Const COLOR_BLUE = &H800000              'blue
Global Const COLOR_BLACK = &H0&                 'Black
Global Const COLOR_RED = &HFF&                  'Red
Global Const COLOR_MILK = &HE0FFFF              'Milk
Global Const COLOR_DARKGREEN = &H808000         'dark green
Global Const COLOR_LIGHTGREEN = &H80FF80        'light green

'User Define Values
Global G_WinDir As String
Global G_System_Title As String                 '�t�Ϊ��Y
Global G_PICTURE_NAME As String                 '�t�ιϧ�(Icon Mark)���|
Global G_Customer_NAME As String                '�Ȥ�W��
Global G_DB_PATH1 As String
Global G_DB_PATH2 As String
Global G_DB_PATH3 As String
Global G_DB_PATH4 As String
Global G_DB_PATH5 As String
Global G_DB_PATH6 As String
Global G_DB_PATH7 As String
Global G_DB_PATH8 As String
Global G_DB_PATH9 As String
Global G_DB_PATH10 As String
Global G_ConnectMethod1 As String
Global G_ConnectMethod2 As String
Global G_ConnectMethod3 As String
Global G_ConnectMethod4 As String
Global G_ConnectMethod5 As String
Global G_ConnectMethod6 As String
Global G_ConnectMethod7 As String
Global G_ConnectMethod8 As String
Global G_ConnectMethod9 As String
Global G_ConnectMethod10 As String
Global G_System_Path As String
Global G_Report_Path As String
Global G_INI_SerPath As String
Global G_Help_Path As String
Global G_Program_Path  As String
Global G_DUserId As String
Global G_UserName As String
Global G_UserGroup As String
Global G_DateFlag As Integer                    '0:English 1:Chinese
Global G_Form_Color As String
Global G_Title_Color As String
Global G_Label_Color As String
Global G_TabBack_Color As String
Global G_TabFore_Color As String
Global G_TextHelpBack_Color As String
Global G_TextGotBack_Color As String
Global G_TextLostBack_Color As String
Global G_TextGotFore_Color As String
Global G_TextLostFore_Color As String
Global G_Msgline_Color As String                'Message Line Color
Global G_Today_Color As String                  'Date Label Color
Global G_Font_Name As String                    '�r�ΦW��
Global G_Font_Size As String                    '�r�Τj�p
Global G_FixFont_Name As String
Global G_FixFont_Size As String
Global G_CheckCompany As String                 '�O�_�ˮ֤��q�v��

Global G_FontName As String
Global G_FontSize As Double
Global G_ReportCopies As Integer
Global G_CloseDate As String
Global G_List_Flag As Integer
Global G_RptNeedWidth As Integer
Global G_RptWidth As Integer
Global G_PageSize As Integer
Global G_OverFlow As Integer
Global G_RptSet As Integer
Global G_Str As String
Global G_ProgramName As String
Global G_CmdStr1 As String
Global G_CmdStr2 As String
Global G_CmdStr3 As String
Global G_dtDateError As Date
Global G_dtDateMax As Date
Global G_dtDateMin As Date
Global G_ExecuteErr As Variant
Global G_Print_NextPage As String
Global G_Print_Date As String
Global G_Print_Time As String
Global G_Print_Page As String
Global G_Terminal_Check As Boolean

'common message variable
Global retcode As Variant       '�Ǧ^��
Global G_AP_STATE As Integer    '�@�~���A
Global G_AP_ADD As String       '�s�W
Global G_AP_DELETE As String    '�R���T�{
Global G_AP_NORMAL As String    '���~��@�~
Global G_AP_NODATA As String    '��Ƥ��s�b
Global G_AP_NOPRVS As String    '�L�e��
Global G_AP_NONEXT As String    '�L����
Global G_AP_PRINT As String     '�C�L
Global G_AP_QUERY As String     '�d��
Global G_AP_SEARCH As String    'Search
Global G_AP_UPDATE As String    '��
Global G_AP_COPY As String      '�ƻs
Global G_AP_TABLE As String     '���
Global G_CmdSet As String       '���]�wF9

'def command button caption
Global G_CmdHelp  As String                     '����F1
Global G_CmdSort  As String                     '�Ƨ�F2
Global G_CmdQuery As String                     '�d��F2
Global G_CmdDel As String                       '�R��F3
Global G_CmdAdd As String                       '�s�WF4
Global G_CmdUpdate As String                    '�ק�F5
Global G_CmdCopy As String                      '�ƻsF5
Global G_CmdPrint As String                     '�L��F6
Global G_CmdPrevious As String                  '�W��F7
Global G_CmdNext As String                      '�U��F8
Global G_CmdPrvPage As String                   '�e��F7
Global G_CmdNxtPage As String                   '����F8
Global G_CmdRecordSet As String                 '���F9
Global G_CmdTable As String                     '���F9
Global G_CmdOk As String                        '�T�{F11
Global G_CmdSearch As String                    '��MF11
Global G_CmdExit As String                      '����ESC
Global G_CmdPause As String
Global G_CmdInsert As String
Global G_CmdHistory As String

'common message
Global G_Add_Check As String               'add new check
Global G_Add_Ok As String                  'add ok
Global G_Delete_Check As String            'delete check
Global G_Delete_Ok As String               'delete ok
Global G_NoMoreData As String              'No More Data
Global G_Save_Check As String              'save check
Global G_OverDate As String                'Over Date
Global G_RecordExist As String             'Record Exist
Global G_NoReference As String
Global G_NoQueryData As String
Global G_Printing As String
Global G_DataErr As String
Global G_FieldErr As String
Global G_MustInput As String
Global G_Process As String
Global G_DateError As String
Global G_NumericErr As String
Global G_Range_Error As String
Global G_Update_Ok As String
Global G_Query_Ok As String
Global G_PrintOk As String
Global G_DataLockErr As String             'S020911050 ��Ƥw�QPCName\UserID[UserName]�ϥΤ�,�L�k�Q��w,�е��ݩάO�q���ӨϥΪ̰h�X!

'Definition Database & Dynaset Name
Global G_DBNotOpen1 As Boolean
Global G_DBNotOpen2 As Boolean
Global G_DBNotOpen3 As Boolean
Global G_DBNotOpen4 As Boolean
Global G_DBNotOpen5 As Boolean
Global G_DBNotOpen6 As Boolean
Global G_DBNotOpen7 As Boolean
Global G_DBNotOpen8 As Boolean
Global G_DBNotOpen9 As Boolean
Global G_DBNotOpen10 As Boolean
Global G_WorkSpace1 As Workspace
Global G_WorkSpace2 As Workspace
Global G_WorkSpace3 As Workspace
Global G_WorkSpace4 As Workspace
Global G_WorkSpace5 As Workspace
Global G_WorkSpace6 As Workspace
Global G_WorkName(10) As Workspace
Global G_FileName(100) As Recordset
Global G_TableName(100) As Recordset
Global G_WorkFile(10) As String
Global G_File(100) As String
Global DY_TBLDCT As Recordset
Global DY_TBLDEF As Recordset
Global DY_TBLIDX As Recordset
Global DY_INI As Recordset
Global G_Err As Error

'Define System Menu vari
Global G_Authority As String
Global G_IllegalTerminal As String
Global Const G_GUIOpt$ = "wk_gui_"
Global Const G_GLOpt$ = "wk_gl_"
Global Const G_IVOpt$ = "wk_iv_"
Global Const G_TRDOpt$ = "wk_trd_"
Global Const G_LSMOpt$ = "wk_lsm_"
Global Const G_LSFOpt$ = "wk_lsf_"
Global Const G_ASOpt$ = "wk_as_"
Global Const G_PYOpt$ = "wk_py_"
Global Const G_ZINOpt$ = "wk_zin_"
Global Const G_MPOpt$ = "wk_mp_"


Global G_AUT_READ%
Global G_AUT_UPDATE%
Global G_AUT_DELETE%
Global G_AUT_ADD%

'S010605056 �Τ@�s���H��L�O�����u��
Global G_A1609uninumber$

'*** Add for New Report 2001/11/14 ***
Public Type SpreadCol
    BreakInLine As Boolean              'Break���O�_��ܦb��ƦC�W
    SelectIndex As Integer              '�{���w�]����춶��
    ReportIndex As Integer              '����C�L����춶��
    ScreenIndex As Integer              '�ثeSpread�W����춶��
    TempIndex As Integer                'Keep�ù���ܪ���l��춶��
    BreakIndex As Integer               'Break��쪺����(��1�}�l)
    Hidden As Integer                   '������ó]�w(0:��� 1:����� 2:����)
    ColWidth As Integer                 '���e��
    Name As String                      '���W��
    text As String                      '����
    Caption As String                   '�����D (Single Line�ϥ�)
    mCaption As String                  '����Header(Mutiline�ϥ�)
    CFormat As String                   '�����D���榡
    dFormat As String                   '����ƪ��榡
    DateFormat As Boolean               '����X��Excel��,�O�_�H����榡���
End Type
 
Public Type SpreadSort
    SortKey As String                   '�Ƨ���쪺�W��
    SortOrder As Integer                '�Ƨ���쪺��V(���W�λ���)
End Type

Public Type Spread
    SortEnable As Boolean               '�O�_���\���s�Ƨ�
    Refresh As Boolean
    RefreshCol As Boolean               '���O�_����
    RefreshSort As Boolean              '�Ƨ����O�_����
    Tag As String                       '�w�d
    Sorts() As SpreadSort
    Columns() As SpreadCol
End Type

Global tSpd_RptDef As Spread
Global G_ReportDataFrom As Integer
Global Const G_FromRecordSet = 1
Global Const G_FromScreen = 2

Global G_LineLeft As String
Global G_LineRight As String
Global G_ColSpace As String
Global G_RptMinWidth As Integer

Global G_SecLevel$
Global G_SecPwdMinLen$         '�ޱ��K�X�̤p����
Global G_SecPwdFixedLen$       '�ޱ��K�X�T�w�X�� S020308013
Global G_SecPwdComplexity$     '�ޱ��K�X�������� S020308013
Global G_SecPwdFailedWaitTime$ '�ޱ��K�X��J���~�ݵ��ݮɶ�(��),�~�i�A��J S020308013

'*** Add for Short Date 2002/8/19 ***
Global G_LeadYear$

'*** Add for ���Y�θ����Y�ɮ� 2003/1/3 ***
Public Const OFS_MAXPATHNAME = 128
Public Const STILL_ALIVE = &H103
Public Const OF_CREATE = &H1000
Public Const OF_READ = &H0

Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Declare Function LZOpenFile Lib "lz32.dll" Alias "LZOpenFileA" (ByVal lpszFile As String, lpOf As OFSTRUCT, ByVal style As Long) As Long
Public Declare Function LZCopy Lib "lz32.dll" (ByVal hfSource As Long, ByVal hfDest As Long) As Long
Public Declare Sub LZClose Lib "lz32.dll" (ByVal hfFile As Long)

' *** Add for �HMouse COPY & Paste��Ĳ�oChange
Global G_FieldText$
Global G_DataChange%

' *** Add For �ഫ���ˮ֥����r��
Global G_FullyChar%

' *** Add For
Global G_OldCrtDate$
Global G_OldCrtTime$
Global G_OldCrtUser$
Global G_OldCrtWrkStn$

'*** Add From System Module 93/4/1 ***
Global G_LineNo As Long             '�ثe�ӭ��w�C�L���
Global G_PageNo As Long             '�ثe�C�L����
Global G_PrintSelect As Integer     '�C�L�覡
Global Const G_Print2Screen = 1     '�ù��C�L
Global Const G_Print2Printer = 2    '�L����C�L
Global Const G_Print2File = 3       '�ɮצC�L
Global Const G_Print2Excel = 4      'Excel�ɮצC�L
Global Const G_Print2Word = 5       'Word�ɮצC�L
Global G_OutFile As String          '��r���ɦW

'***Add For Security Program Log (93/4/14 Add by cathy) *****
Global G_SecurityPgm As Boolean

'*** Add For Vista 95/12/20 By Jennifer
Global DY_INICommon As Recordset

'*** Add For Vista 96/6/25 By Jennifer
Global G_IsVistaClient As Boolean
Global DBEngine36 As DAO.DBEngine
Global G_VistaClientTitle As String

Function GetSvrINIStrA(DB As Database, ByVal Section$, ByVal Topic$) As String
'���o���w��Ʈw��,SINI-TABLE����TOPICVALUE��
Dim DY As Recordset
Dim A_Sql$

    GetSvrINIStrA = " "
    A_Sql$ = "Select TOPICVALUE From SINI Where"
    A_Sql$ = A_Sql$ & " SECTION='" & Section$ & "'"
    A_Sql$ = A_Sql$ & " AND TOPIC='" & Topic$ & "'"
    A_Sql$ = A_Sql$ & " Order by SECTION,TOPIC"
    CreateDynasetODBC DB, DY, A_Sql$, "DY", True
    If Not (DY.BOF And DY.EOF) Then
       GetSvrINIStrA = DY.Fields("TOPICVALUE") & ""
    End If
End Function

Sub ChangeReportHeader(tSPD As Spread, ByVal FieldName As String, ByVal Value As String)
'���ܳ�����쪺Caption
Dim A_Index%

    A_Index% = GetSpdColIndex(tSPD, FieldName)
    tSPD.Columns(A_Index%).mCaption = Value
End Sub

Sub ChangeReportHeaderAlign(tSPD As Spread, ByVal FieldName$, ByVal Align%)
'���ܳ�����쪺Caption
Dim A_Index%, A_Len%

    A_Index% = GetSpdColIndex(tSPD, FieldName$)
    A_Len% = Len(tSPD.Columns(A_Index%).CFormat)
    '
    Select Case Align%
      Case SS_CELL_H_ALIGN_LEFT
           tSPD.Columns(A_Index%).CFormat = String(A_Len%, "#")
      Case SS_CELL_H_ALIGN_CENTER
           tSPD.Columns(A_Index%).CFormat = String(A_Len%, "^")
      Case SS_CELL_H_ALIGN_RIGHT
           tSPD.Columns(A_Index%).CFormat = String(A_Len%, "~")
    End Select
End Sub


Function CompressFile(ByVal SourceFile$, ByVal DestFile$, Optional ErrMsg$ = "") As Boolean
'�ɮ����Y
On Error GoTo MyError
Dim A_nId&, A_hProcess&, A_nExitCode&

    CompressFile = True
            
    '�Y�ت��ɤw�s�b,���R��
    If Dir$(DestFile$) <> "" Then Kill DestFile$
    
    'Compress.exe����������ؿ��B����ª��ɦW
    A_nId& = Shell("Compress " & SourceFile$ & " " & DestFile$, vbHide)
    
    A_hProcess& = OpenProcess(PROCESS_QUERY_INFORMATION, 0, A_nId&)
    Do
        Call GetExitCodeProcess(A_hProcess&, A_nExitCode&)
        DoEvents
    Loop While A_nExitCode& = STILL_ALIVE
    Call CloseHandle(A_hProcess&)
    Exit Function
    
MyError:
    CompressFile = False
    ErrMsg$ = Error$
End Function


Function ExpandFile(ByVal SourceFile$, ByVal DestFile$, Optional ErrMsg$ = "") As Boolean
'�ɮ׸����Y
On Error GoTo MyError
Dim A_File1&, A_File2&, A_Retcode&
Dim ofFile1 As OFSTRUCT, ofFile2 As OFSTRUCT

    ExpandFile = True
            
    '�Y�ت��ɤw�s�b,���R��
    If Dir$(DestFile$) <> "" Then Kill DestFile$
            
    A_File1& = LZOpenFile(SourceFile$, ofFile1, OF_READ)
    A_File2& = LZOpenFile(DestFile$, ofFile2, OF_CREATE)
    A_Retcode& = LZCopy(A_File1&, A_File2&)
    LZClose A_File2&
    LZClose A_File1&
    Exit Function
    
MyError:
    ExpandFile = False
    ErrMsg$ = Error$
End Function


Function FillLine(ByVal A_Code$, ByVal A_Len!) As String
    Dim A_STR$, a!
    If A_Len! <= 0 Then
       FillLine = ""
       Exit Function
    End If
    A_STR$ = ""
    Do While a! <= A_Len!
       A_STR$ = A_STR$ + A_Code$
       a! = a! + 1
    Loop
    FillLine = A_STR$
End Function


Sub SQLInsert1(DB As Database, ByVal Table$, ErrCode)
'����SQL�s�W���O,�f�tInsertFields�{�Ǩϥ�
Dim A_Tmp$, A_Str1$, A_Str2$, A_Sql$
'S021114036 �]�ǲ�ñ�֮ɡA�ݲզ귥�����r��A�G�Ni%�ܼƩ��̤j(1021115 by Lidia)
Dim I As Currency

    A_Tmp$ = Chr(0) & Chr(128)
    I = InStr(1, G_Str, A_Tmp$)
    If I <> 0 Then
       A_Str1$ = Left(G_Str, I - 1)
       A_Str2$ = Right(G_Str, Len(G_Str) - (I + 1))
    End If
    A_Str1$ = "(" & A_Str1$ & ")"
    If Right(A_Str2$, 1) = "," Then
       A_Str2$ = Left(A_Str2$, Len(A_Str2$) - 1)
    End If
    A_Sql$ = "Insert into " & Table$ & Space(1) & A_Str1
    A_Sql$ = A_Sql$ & " values " & "(" & A_Str2$ & ")"
    ExecuteProcessReturnErr DB, A_Sql$, ErrCode
    G_Str = ""
End Sub

Function Value2Int(ByVal oNumber#, ByVal Fractional%) As Double
'�N�ƭȥ|�ˤ��J,�i�]�w�p�Ʀ���(���A��Single�����B�z�ƭ�, ���o�p�Ʀ��)
Dim A_Format$

    If Fractional% < 0 Then Value2Int = CCur(oNumber#): Exit Function
    
    A_Format$ = "0"
    If Fractional% > 0 Then
        A_Format$ = A_Format$ + "." + String(Fractional%, "0")
    End If
    Value2Int = Format(oNumber#, A_Format$)
End Function
Sub MoveData2Sini(DB As Database, ByVal A_Section$, ByVal A_Topic$, ByVal A_TopicValue$, Optional A_DeleteOnly As Boolean = False)
'�N�]�w�ȼg�JSINI
On Local Error GoTo MyError
Dim A_Sql$, DY As Recordset
Dim A_IsRecordExist As Boolean

    If A_DeleteOnly = True Then
        GoSub DeleteRecord
        Exit Sub
    End If
    
    GoSub IsRecordExist
    
    If A_IsRecordExist = True Then
        GoSub UpdateRecord
        Exit Sub
    End If
    
    GoSub InsertRecord
    
    
    Exit Sub
    
IsRecordExist:
    A_Sql$ = "Select * From SINI Where"
    A_Sql$ = A_Sql$ & " SECTION='" & A_Section$ & "'"
    A_Sql$ = A_Sql$ & " AND TOPIC='" & A_Topic$ & "'"
    A_Sql$ = A_Sql$ & " Order by SECTION,TOPIC"
    CreateDynasetODBC DB, DY, A_Sql$, "DY", True
    A_IsRecordExist = Not (DY.BOF And DY.EOF)
        
    Return
    
InsertRecord:
    G_Str = ""
    InsertFields "Section", Trim(A_Section$), G_Data_String
    InsertFields "Topic", Trim(A_Topic$), G_Data_String
    InsertFields "TopicValue", Trim(A_TopicValue$), G_Data_String
    SQLInsert DB, "SINI"
    Return
    
UpdateRecord:
    G_Str = "UPDATE SINI SET "
    UpdateString "TopicValue", A_TopicValue$, G_Data_String
    G_Str = Left$(G_Str, Len(G_Str) - 1)
    G_Str = G_Str & " where Section='" & A_Section$ & "'"
    G_Str = G_Str & " AND   Topic='" & A_Topic$ & "'"
    ExecuteProcess DB, G_Str
    Return
    
DeleteRecord:
    G_Str = "DELETE FROM Sini Where Section ='" & Trim(A_Section$) & "'"
    G_Str = G_Str & " And Topic='" & Trim(A_Topic$) & "'"
    ExecuteProcess DB, G_Str
    Return
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Sub KeepUserPwd2SINI(ByVal A0801$, ByVal DueDate$)
'Keep�C�ӨϥΪ̪����Ĥ��
'Section='Security'
'Topic='User_' & UserID
'TopicValue=�[�K��K�X���Ĵ�
Dim A_Topic$, A_Value$

    A_Topic$ = "User_" & Trim(A0801$)
    A_Value$ = Hex(Oct(Val(DueDate$)))
    UpdSecRow2SINI A_Topic$, A_Value$
End Sub


Function GetCurDueDate(ByVal A_DueDate$) As Date
Dim A_Date$

    A_Date$ = CStr(Val("&O" & CStr(Val("&H" & A_DueDate$))))
    GetCurDueDate = DateSerial(Left(A_Date$, 4), _
                    Mid(A_Date$, 5, 2), Right(A_Date$, 2))
End Function

Function Num(ByVal Key$, Optional Assign$ = "") As String
'���o�ϥΪ̱K�X�[�K�᪺�Ʀr
Dim Password#, length%, I%
    
    Num = ""
    If Key$ = "" Or Trim$(Key$) = "0" Then Exit Function
    '
    If Trim(Assign$) = "" Then Assign$ = "0"
    Select Case Assign$
      Case "0"
           Password# = Val(Key$)
           Password# = Password# * Password#
           Password# = Password# + 59
           Password# = Password# * 17
           Password# = Password# - 101
      Case "1"
           Password# = 0
           length% = Len(Trim$(Key$))
           For I% = 1 To length%
               Password# = Password# * 128 + Asc(Mid$(Trim$(Key$), I%, 1))
           Next I%
    End Select
    Num = CStr(Password#)
End Function

Function Word(ByVal Key$, Optional Assign$ = "") As String
'���o�ѱK�᪺�ϥΪ̱K�X
Dim Password, TmpKey
Dim I%, Pwd$, Key1, Key2

    Word = ""
    If Trim$(Key$) = "" Or Val(Key$) = 0 Then Exit Function
    '
    If Trim(Assign$) = "" Then Assign$ = "0"
    TmpKey = Val(Key$)
    Select Case Assign$
      Case "0"
           Password = TmpKey + 101
           Password = Password / 17
           Password = Password - 59
           Password = Sqr(Password)
           Word = Trim(Password)
      Case "1"
           Pwd$ = Space(7): I% = 7
           Do
              Key1 = Int(TmpKey / 128)
              Key2 = TmpKey - Key1 * 128
              If Key2 > 0 Then
                 Mid$(Pwd$, I%, 1) = Trim(Chr(Key2))
              ElseIf Key2 = 0 Then
                 Mid$(Pwd$, I%, 1) = Trim(Chr(128))
                 Key1 = (Key - 128) / 128
              End If
              TmpKey = Key1
              I% = I% - 1
           Loop Until TmpKey = 0
           Word = Trim(Pwd$)
    End Select
End Function


Sub KeepTryError(ByVal Topic$, ByVal Assign$, Optional Try%)
Dim A_Value$

    Select Case Assign$
      Case "0"
           Try% = Try% + 1
      Case "1"
           If ReferenceGUI_SINI("Security", Topic$) Then
              A_Value$ = CStr(Val(DY_SINI.Fields("TopicValue") & "") + 1)
              Try% = Val(A_Value$)
              UpdSecRow2SINI Topic$, A_Value$, G_AP_STATE_UPDATE
           Else
              A_Value$ = "1"
              Try% = 1
              UpdSecRow2SINI Topic$, A_Value$, G_AP_STATE_ADD
           End If
    End Select
End Sub

Sub UpdSecRow2SINI(ByVal Topic$, ByVal Value$, Optional State% = 0)
    G_Str = ""
    '
    Select Case State%
      Case 0
           If ReferenceGUI_SINI("Security", Topic$) Then
              GoSub UpdateSINI
           Else
              GoSub AddSINI
           End If
      
      Case G_AP_STATE_ADD
           GoSub AddSINI
           
      Case G_AP_STATE_UPDATE
           GoSub UpdateSINI
      
      Case G_AP_STATE_DELETE
           GoSub DeleteSINI
    End Select
    Exit Sub
    
AddSINI:
    InsertFields "Section", "Security", G_Data_String
    InsertFields "Topic", Topic$, G_Data_String
    InsertFields "TopicValue", Value$, G_Data_String
    SQLInsert DB_ARTHGUI, "SINI"
    Return
    
UpdateSINI:
    UpdateString "TopicValue", Value$, G_Data_String
    G_Str = G_Str & " where Section='Security'"
    G_Str = G_Str & " and Topic='" & Topic$ & "'"
    SQLUpdate DB_ARTHGUI, "SINI"
    Return
    
DeleteSINI:
    G_Str = "DELETE FROM SINI"
    G_Str = G_Str & " WHERE SECTION='Security'"
    G_Str = G_Str & " AND TOPIC='" & Topic$ & "'"
    ExecuteProcess DB_ARTHGUI, G_Str
    Return
End Sub

Function GetSecurityLevel() As String
Dim A_STR$

    A_STR$ = GetGUISvrIniStr("Security", "Level")
    GetSecurityLevel = IIf(Trim(A_STR$) = "", "0", A_STR$)
End Function

Function GetGUISvrIniStr(ByVal Section$, ByVal Topic$) As String
'�ۨt�θ�Ʈw���o��w���e
    GetGUISvrIniStr = " "
    If ReferenceGUI_SINI(Section$, Topic$) Then
       GetGUISvrIniStr = (DY_SINI.Fields("TOPICVALUE") & "")
    End If
End Function

Function ReferenceGUI_SINI(ByVal Section$, ByVal Topic$) As Boolean
Dim A_Sql$

    ReferenceGUI_SINI = False
    A_Sql$ = "Select TOPICVALUE From SINI Where"
    A_Sql$ = A_Sql$ & " SECTION='" & Section$ & "'"
    A_Sql$ = A_Sql$ & " AND TOPIC='" & Topic$ & "'"
    A_Sql$ = A_Sql$ & " ORDER BY SECTION,TOPIC"
    CreateDynasetODBC DB_ARTHGUI, DY_SINI, A_Sql$, "DY_SINI", True
    If Not (DY_SINI.BOF And DY_SINI.EOF) Then ReferenceGUI_SINI = True
End Function

Sub ShellEXEProcess(A_Form As Form, ByVal A_EXEName$)
'�ѥD�{���~Call�Ƶ{��,�����Ƶ{������,�D�{���~�i�~��
On Error Resume Next
Dim A_hProcess, A_RetVal, A_PId

    A_Form.Vse_Background.Enabled = False
    CloseSystemMenu A_Form, SC_RESTORE   '�٭�
    CloseSystemMenu A_Form, SC_MAXIMIZE  '�̤j��
    CloseSystemMenu A_Form, SC_MINIMIZE  '�̤p��
    CloseSystemMenu A_Form, SC_CLOSE     '����
    A_PId = Shell(A_EXEName$, vbNormalFocus)
    A_hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, A_PId)
    Do
        GetExitCodeProcess A_hProcess, A_RetVal
        DoEvents
        Sleep 100
    Loop While A_RetVal = STILL_ACTIVE
    OpenSystemMenu A_Form, SC_RESTORE    '�٭�
    OpenSystemMenu A_Form, SC_MAXIMIZE   '�̤j��
    OpenSystemMenu A_Form, SC_MINIMIZE   '�̤p��
    OpenSystemMenu A_Form, SC_CLOSE      '����
    A_Form.Vse_Background.Enabled = True
End Sub

Sub AddSpdComboBoxStr(Spd As vaSpread, tSPD As Spread, ByVal Row#, ByVal FldName$, ByVal Str$)
'�bSpread��Combobox���W�[�J���
Dim A_Index#

    '���o���Index
    A_Index# = GetSpdColIndex(tSPD, FldName$)
        
    '�[�J��ƦC
    Spd.Row = Row#
    Spd.Col = A_Index#
    Spd.TypeComboBoxIndex = -1
    Spd.TypeComboBoxString = Str$
End Sub

Function ConnectSemiColon(ByVal Str$) As String
'�b�r��e��[�J;�Ÿ�
    
    ConnectSemiColon = ";" & Str$ & ";"
End Function

Sub Control_Property(Tmp As Control, ByVal text$, Optional ByVal Visible As Boolean = True, Optional ByVal Color$, Optional ByVal FSize$, Optional ByVal FName$, Optional ByVal FColor$)
'�]�wPanel,Label,OptionButton,CheckBox,Frame���ݩ�
On Error Resume Next

    If Trim(Color$) = "" Then Color$ = G_Label_Color
    If Trim(FColor$) = "" Then FColor$ = G_TextLostFore_Color
    If Trim(FName$) = "" Then FName$ = G_Font_Name
    If Trim(FSize$) = "" Then FSize$ = G_Font_Size
    
    Tmp.BackColor = Val(Color$)
    Tmp.ForeColor = Val(FColor$)
    Tmp.Caption = text$
    Tmp.Visible = Visible
    Tmp.FontName = FName$
    Tmp.FontSize = FSize$
    Tmp.FontBold = False
    Tmp.FontItalic = False
End Sub

Function GetFullTableName(SourceDB As Database, MainDB As Database, ByVal Table$, Optional ByVal CrossDB As Boolean = True) As String
'���oTable JOIN������Table�W��,�Y��SQL Join Access,�����bSQL Server���إ�
'Access��Ʈw��Link Server,�s�������A������Ʈw��������|,���Ѫ̬�
'Microsoft Jet *.* OLE DB Provider,��ƨӷ�����Ʈw��������|
Dim A_HaveFull%, A_SDBName$

    Table$ = "[" & Table$ & "]"
    A_SDBName$ = "[" & SourceDB.Name & "]"
    '
    If Not CrossDB Then
       GetFullTableName = Table$
    Else
       If SourceDB.Connect <> "" Then
          A_HaveFull% = True
          If MainDB.Connect <> "" Then
             If StrComp(GetSQLServerName(MainDB.Connect), GetSQLServerName(SourceDB.Connect), vbTextCompare) = 0 Then
                A_HaveFull% = False
             End If
          End If
          If A_HaveFull% Then GetFullTableName = GetSQLServerName(SourceDB.Connect)
          GetFullTableName = GetFullTableName & A_SDBName$ & ".dbo." & Table$
       Else
          If MainDB.Connect <> "" Then
             GetFullTableName = A_SDBName$ & "..." & Table$
          Else
             GetFullTableName = A_SDBName$ & "." & Table$
          End If
       End If
    End If
    GetFullTableName = " " & GetFullTableName & " "
End Function

Function GetSpdComboBoxText(Spd As vaSpread, tSPD As Spread, ByVal FldName$, ByVal Row#, Optional ByVal Pos# = 1, Optional ByVal Separator$) As String
'���oSpread��Combobox��쪺�ثe�r���
Dim I#, A_Index#
Dim A_STR$, A_Str1$

    '���o���Index
    A_Index# = GetSpdColIndex(tSPD, FldName$)

    Spd.Row = Row#
    Spd.Col = A_Index#
    A_STR$ = Spd.text
    If Pos# = 0 Then
       'Pos#=0,�H��Ӧr��Ȱ����
       A_Str1$ = A_STR$
    Else
       'Pos#>0,���o���w���j���r��Ȱ����
       For I# = 1 To Pos#
           StrCut A_STR$, Separator$, A_Str1$, A_STR$
       Next I#
    End If
    GetSpdComboBoxText = A_Str1$
End Function

Function GetSpdComboBoxIndex(Spd As vaSpread, tSPD As Spread, ByVal Row#, ByVal FldName$, ByVal Str$, Optional ByVal Pos# = 1, Optional ByVal Separator$) As Double
'���oSpread��Combobox��쪺�ثeIndex
Dim I#, j#, A_Index#
Dim A_Source$, A_Str1$

    GetSpdComboBoxIndex = -1
    
    '���o���Index
    A_Index# = GetSpdColIndex(tSPD, FldName$)

    Spd.Row = Row#
    Spd.Col = A_Index#
    For I# = 0 To Spd.TypeComboBoxCount - 1
        Spd.TypeComboBoxIndex = I#
        A_Source$ = Spd.TypeComboBoxString
        If Pos# = 0 Then
           'Pos#=0,�H��Ӧr��Ȱ����
           A_Str1$ = A_Source$
        Else
           'Pos#>0,���o���w���j���r��Ȱ����
           For j# = 1 To Pos#
               StrCut A_Source$, Separator$, A_Str1$, A_Source$
           Next j#
        End If
        If StrComp(Str$, A_Str1$, vbTextCompare) = 0 Then
           GetSpdComboBoxIndex = I#
           Exit For
        End If
    Next I#
End Function
Sub SetSpreadColor(Spd As vaSpread, ByVal Row#, ByVal Col#, ByVal BColor$, ByVal FColor$)
'�]�wSpread�W�x�s�檺�I���Ϋe���C��

    Spd.Row = Row#
    Spd.Col = Col#
    Spd.BackColor = Val(BColor$)
    Spd.ForeColor = Val(FColor$)
End Sub

Sub SpdSortIndexReBuild(tSPD As Spread, ByVal Col#)
'Update Spread Type�����Ƨ����
Dim tSorts(1 To 3) As SpreadSort

    tSorts(1).SortKey = tSPD.Columns(Col#).Name
    If StrComp(tSPD.Sorts(1).SortKey, tSPD.Columns(Col#).Name, vbTextCompare) = 0 Then
       If tSPD.Sorts(1).SortOrder = SS_SORT_ORDER_ASCENDING Then
          tSorts(1).SortOrder = SS_SORT_ORDER_DESCENDING
       Else
          tSorts(1).SortOrder = SS_SORT_ORDER_ASCENDING
       End If
    Else
       tSorts(1).SortOrder = SS_SORT_ORDER_ASCENDING
    End If
    tSPD.Sorts = tSorts
End Sub

Sub Field_Property(Tmp As Control, ByVal length%, Optional Tmp2 As Control, Optional ByVal FCaption$, _
Optional FormatStr$, Optional Up$, Optional Down$, Optional ByVal DBName$, _
Optional ByVal TBName$, Optional ByVal TBField$)
'�]�wTextBox,ComboBox,ListBox��Label���ݩ�
On Error Resume Next
Dim A_Caption$

    '�Y���D�ѼƦ���,�h�j���]�wLabel�����D���Ѽƭ�,�Y�L�Ѽƭ�,Keep Design��Label�����D
    If FCaption$ <> "" Then
       A_Caption$ = FCaption$
    Else
       If Not Tmp2 Is Nothing Then A_Caption$ = Tmp2.Caption
       '�۸�Ʈw�����o��쪺�ݩʭ�(����,���D,�榡,�W�U����)
       GetPropertyFromDB DBName$, TBName$, TBField$, length%, A_Caption$, FormatStr$, Up$, Down$
    End If
    
    '�]�wTextBox���ݩ�
    Tmp.BackColor = Val(G_TextLostBack_Color)
    Tmp.ForeColor = Val(G_TextLostFore_Color)
    Tmp.MaxLength = length%
    Tmp.FontName = G_Font_Name
    Tmp.FontSize = G_Font_Size
    Tmp.FontBold = False
    Tmp.FontItalic = False
    
    '�]�wLabel���ݩ�
    If Not Tmp2 Is Nothing Then
       Tmp2.BackColor = Val(G_Label_Color)
       Tmp2.ForeColor = Val(G_TextLostFore_Color)
       Tmp2.Caption = A_Caption$
       Tmp2.FontName = G_Font_Name
       Tmp2.FontSize = G_Font_Size
       Tmp2.FontBold = False
       Tmp2.FontItalic = False
    End If
End Sub

Sub SpreadColumnMove(Spd As vaSpread, tSPD As Spread, ByVal Col#, ByVal newcol#, ByVal NewRow#, Cancel As Boolean)
'�B�zSpread��쪺����
    
    Screen.MousePointer = HOURGLASS
    
    Cancel = True
    If Col# = newcol# Then Screen.MousePointer = Default: Exit Sub
        
    '�N��첾�ܷs��m
    ChangeSpdCols Spd, Col#, newcol#
    
    '�]�w��Ц�m
    Spd.Row = NewRow#
    Spd.Col = newcol#
    Spd.Action = SS_ACTION_ACTIVE_CELL
    
    '������춶�ǤΦW��
    SwapSpreadColName tSPD, Col#, newcol#
    
    Screen.MousePointer = Default
End Sub

Sub SwapSpreadColName(tSPD As Spread, ByVal Col#, ByVal newcol#)
'����Spread Type��Keep�����W�٤ζ��ǭ�
Dim I#, A_Start#, A_End#, A_Name$, A_Text$, A_Caption$, A_CFormat$, A_DFormat$
Dim A_ReportIndex#, A_SelectIndex#, A_ScreenIndex#, A_TempIndex#, A_BreakIndex#
Dim A_BreakInLine%, A_ColWidth#, A_Hidden%

    'Keep����쪺�W��
    A_BreakInLine% = tSPD.Columns(Col#).BreakInLine
    If tSPD.Columns(Col#).ReportIndex = 0 Then
       A_ReportIndex# = tSPD.Columns(Col#).ReportIndex
    Else
       A_ReportIndex# = newcol#
    End If
    A_SelectIndex# = tSPD.Columns(Col#).SelectIndex
    A_ScreenIndex# = newcol#
    A_TempIndex# = tSPD.Columns(Col#).TempIndex
    A_BreakIndex# = tSPD.Columns(Col#).BreakIndex
    A_Hidden% = tSPD.Columns(Col#).Hidden
    A_ColWidth# = tSPD.Columns(Col#).ColWidth
    A_Name$ = tSPD.Columns(Col#).Name
    A_Text$ = tSPD.Columns(Col#).text
    A_Caption$ = tSPD.Columns(Col#).Caption
    A_CFormat$ = tSPD.Columns(Col#).CFormat
    A_DFormat$ = tSPD.Columns(Col#).dFormat
    
    '�������N���ʪ��W�٤ζ��ǭ�
    A_Start# = Col#
    If Col# > newcol# Then
        A_End# = newcol# + 1
        For I# = A_Start# To A_End# Step -1
            tSPD.Columns(I#).BreakInLine = tSPD.Columns(I# - 1).BreakInLine
            If tSPD.Columns(I# - 1).ReportIndex > 0 Then
               tSPD.Columns(I#).ReportIndex = tSPD.Columns(I# - 1).ReportIndex + 1
            Else
               tSPD.Columns(I#).ReportIndex = tSPD.Columns(I# - 1).ReportIndex
            End If
            tSPD.Columns(I#).SelectIndex = tSPD.Columns(I# - 1).SelectIndex
            tSPD.Columns(I#).ScreenIndex = tSPD.Columns(I# - 1).ScreenIndex + 1
            tSPD.Columns(I#).TempIndex = tSPD.Columns(I# - 1).TempIndex
            tSPD.Columns(I#).BreakIndex = tSPD.Columns(I# - 1).BreakIndex
            tSPD.Columns(I#).Hidden = tSPD.Columns(I# - 1).Hidden
            tSPD.Columns(I#).ColWidth = tSPD.Columns(I# - 1).ColWidth
            tSPD.Columns(I#).Name = tSPD.Columns(I# - 1).Name
            tSPD.Columns(I#).text = tSPD.Columns(I# - 1).text
            tSPD.Columns(I#).Caption = tSPD.Columns(I# - 1).Caption
            tSPD.Columns(I#).CFormat = tSPD.Columns(I# - 1).CFormat
            tSPD.Columns(I#).dFormat = tSPD.Columns(I# - 1).dFormat
        Next I#
    Else
        A_End# = newcol# - 1
        For I# = A_Start# To A_End#
            tSPD.Columns(I#).BreakInLine = tSPD.Columns(I# + 1).BreakInLine
            If tSPD.Columns(I# + 1).ReportIndex > 0 Then
               tSPD.Columns(I#).ReportIndex = tSPD.Columns(I# + 1).ReportIndex - 1
            Else
               tSPD.Columns(I#).ReportIndex = tSPD.Columns(I# + 1).ReportIndex
            End If
            tSPD.Columns(I#).SelectIndex = tSPD.Columns(I# + 1).SelectIndex
            tSPD.Columns(I#).ScreenIndex = tSPD.Columns(I# + 1).ScreenIndex - 1
            tSPD.Columns(I#).TempIndex = tSPD.Columns(I# + 1).TempIndex
            tSPD.Columns(I#).BreakIndex = tSPD.Columns(I# + 1).BreakIndex
            tSPD.Columns(I#).Hidden = tSPD.Columns(I# + 1).Hidden
            tSPD.Columns(I#).ColWidth = tSPD.Columns(I# + 1).ColWidth
            tSPD.Columns(I#).Name = tSPD.Columns(I# + 1).Name
            tSPD.Columns(I#).text = tSPD.Columns(I# + 1).text
            tSPD.Columns(I#).Caption = tSPD.Columns(I# + 1).Caption
            tSPD.Columns(I#).CFormat = tSPD.Columns(I# + 1).CFormat
            tSPD.Columns(I#).dFormat = tSPD.Columns(I# + 1).dFormat
        Next I#
    End If
    
    '�]�w�s��쪺�W�٤ζ��ǭ�=�����
    tSPD.Columns(newcol#).BreakInLine = A_BreakInLine%
    tSPD.Columns(newcol#).ReportIndex = A_ReportIndex#
    tSPD.Columns(newcol#).SelectIndex = A_SelectIndex#
    tSPD.Columns(newcol#).ScreenIndex = A_ScreenIndex#
    tSPD.Columns(newcol#).TempIndex = A_TempIndex#
    tSPD.Columns(newcol#).BreakIndex = A_BreakIndex#
    tSPD.Columns(newcol#).Hidden = A_Hidden%
    tSPD.Columns(newcol#).ColWidth = A_ColWidth#
    tSPD.Columns(newcol#).Name = A_Name$
    tSPD.Columns(newcol#).text = A_Text$
    tSPD.Columns(newcol#).Caption = A_Caption$
    tSPD.Columns(newcol#).CFormat = A_CFormat$
    tSPD.Columns(newcol#).dFormat = A_DFormat$

    '�NSpread Type������춶�Ǩ�ScreenIndex���s�إ�
    RebuildByDefIndex tSPD, 2
End Sub

Sub AdjustColWidth(Spd As vaSpread, tSPD As Spread, ByVal ColName$, ByVal FmtStr$)
'�Y������Break��,���K�������פ�����ܾɭP��ƿ�,�����s�]�w������
Dim A_Col#, A_Len&, A_Len2&

    '�H���W�٨��o�����
    A_Col# = GetSpdColIndex(tSPD, ColName$)
    
    '���o��l�e��
    Spd.Row = 1
    Spd.Col = A_Col#
    A_Len& = Spd.TypeEditLen
    
    '���oBreak�榡�r�ꪺ�e��
    A_Len2& = lstrlen(FmtStr$)
    
    '���s�]�w��e
    If A_Len2& > A_Len& Then
       Spd.Row = -1
       Spd.Col = A_Col#
       Spd.TypeEditLen = A_Len2&
    End If
End Sub

Sub ShowRptDefForm(Frm As Form, tSPD As Spread, Optional ByVal RefreshSpd As Boolean = False)
'�B�z���]�w��檺���
    
    tSPD.Refresh = RefreshSpd
    tSpd_RptDef = tSPD
    Frm.Show MODAL
    tSPD = tSpd_RptDef
End Sub
Sub ChangeSpdCols(Spd As vaSpread, ByVal Col#, ByVal newcol#)
'�NvaSpread�W���Y�@��첾��t�@�����
Dim A_Width%, A_Align$, A_Len%, A_Type%, A_Min$, A_Max$, A_DPlaces%

'---------------------------------------------------------------------------------
'�O�d����쪺�ݩʭ�
'---------------------------------------------------------------------------------
    '�O�d���e��
    A_Width% = Spd.ColWidth(Col#)
    '�O�d����ƫ��A
    Spd.Row = -1: Spd.Col = Col#
    A_Type% = Spd.CellType
    '�O�d��L�ݩʭ�
    Select Case A_Type%
      Case 1
           A_Len% = Spd.TypeEditLen
           A_Align$ = Spd.TypeHAlign
      Case 2
           A_Min$ = Spd.TypeFloatMin
           A_Max$ = Spd.TypeFloatMax
           A_DPlaces% = Spd.TypeFloatDecimalPlaces
      Case 3
           A_Min$ = CStr(Spd.TypeIntegerMin)
           A_Max$ = CStr(Spd.TypeIntegerMax)
    End Select

    
'---------------------------------------------------------------------------------
'�N������ƽƻs��s���
'---------------------------------------------------------------------------------
    '���N���̤j�ȥ[�@,�A���J�@������ئa���
    Spd.MaxCols = Spd.MaxCols + 1
    If newcol + 1 <> Spd.MaxCols Then
       Spd.Col = IIf(Col# < newcol, newcol + 1, newcol#)
       Spd.Action = SS_ACTION_INSERT_COL
    End If
    '---------------------------------------------------------------------------------
    '�]�w�s��쪺�ݩʭ�
    '---------------------------------------------------------------------------------
    '�]�w���e��
    Spd.ColWidth(IIf(Col# < newcol#, newcol# + 1, newcol#)) = A_Width%
    '�]�w����ƫ��A
    Spd.Row = -1: Spd.Col = IIf(Col# < newcol#, newcol# + 1, newcol#)
    Spd.CellType = A_Type%
    '�]�w��L�ݩʭ�
    Select Case A_Type%
      Case 1
           Spd.TypeEditLen = A_Len%
           Spd.TypeHAlign = A_Align$
      Case 2
           Spd.TypeFloatMin = A_Min$
           Spd.TypeFloatMax = A_Max$
           Spd.TypeFloatDecimalPlaces = A_DPlaces%
           Spd.TypeFloatDecimalChar = Asc(".")
           Spd.TypeFloatSeparator = True
      Case 3
           Spd.TypeIntegerMin = CInt(A_Min$)
           Spd.TypeIntegerMax = CInt(A_Max$)
    End Select
    '---------------------------------------------------------------------------------
    '�]�w�ӷ����ƻs����ƽd��,�A�ƻs��ئa���
    Dim A_STR$
    Spd.Col = IIf(Col# > newcol#, Col# + 1, Col#)
    Spd.Row = 0
    Spd.Col2 = IIf(Col# > newcol#, Col# + 1, Col#)
    Spd.Row2 = Spd.MaxRows
'    Spd.DestRow = 0
'    Spd.DestCol = IIf(Col# < NewCol#, NewCol# + 1, NewCol#)
'    Spd.Action = SS_ACTION_COPY_RANGE
    A_STR$ = Spd.Clip
    Spd.Row = 0
    Spd.Col = IIf(Col# < newcol#, newcol# + 1, newcol#)
    Spd.Row2 = Spd.MaxRows
    Spd.Col2 = IIf(Col# < newcol#, newcol# + 1, newcol#)
    Spd.Clip = A_STR$
    '�R�������
    Spd.Col = IIf(Col# > newcol#, Col# + 1, Col#)
    Spd.Action = SS_ACTION_DELETE_COL
    '�N���̤j�ȴ�@
    Spd.MaxCols = Spd.MaxCols - 1
End Sub
Function GetSpdText(Spd As vaSpread, tSPD As Spread, ByVal FldName$, ByVal Row#, Optional ByVal HaveComma As Boolean = True, _
Optional ByVal Pos# = 1, Optional ByVal Separator$, Optional ByVal ValueType% = 2, Optional ByVal DateFld As Boolean = False) As String
'��vaSpread�W���o�Y��Cell����
'ValueType = 1, ComboBox�HIndex���
'ValueType = 2, ComboBox�HText���
'DateFld=true,��ܬ�������
Dim A_Col#, A_Value$

    GetSpdText = ""
    
    '�H���W�٨��o��쪺Index
    A_Col# = GetSpdColIndex(tSPD, FldName$)
    If A_Col# = 0 Then Exit Function
    
    '���oCell����
    Spd.Row = Row#
    Spd.Col = A_Col#
    Select Case Spd.CellType
      Case SS_CELL_TYPE_COMBOBOX
           If ValueType% = 1 Then
              A_Value$ = Spd.TypeComboBoxCurSel
           Else
              A_Value$ = GetSpdComboBoxText(Spd, tSPD, FldName$, Row#, Pos#, Separator$)
           End If
      Case SS_CELL_TYPE_FLOAT
           A_Value$ = IIf(HaveComma, Spd.text, Spd.Value)
      Case Else
           A_Value$ = IIf(HaveComma, Spd.text, Trim(CvrTxt2Num(Spd.text)))
           A_Value$ = IIf(DateFld, DateFormat(RejectSlash(Trim(Spd.text))), A_Value$)
    End Select
    
    'Keep Cell���Ȧ�Spread Type
    tSPD.Columns(A_Col#).text = A_Value$
    
    GetSpdText = A_Value$
End Function

Sub SetSpdText(Spd As vaSpread, tSPD As Spread, ByVal FldName$, ByVal Row#, ByVal Value$, Optional ByVal Pos# = 1, Optional ByVal Separator$, Optional ByVal ValueType% = 2)
'�]�wvaSpread�W���Y��Cell��
'ValueType = 1, ComboBox�HIndex���
'ValueType = 2, ComboBox�HText���
Dim A_Col#

    '�H���W�٨��o��쪺Index
    A_Col# = GetSpdColIndex(tSPD, FldName$)
    If A_Col# = 0 Then Exit Sub
    
    '�]�wCell����
    Spd.Row = Row#
    Spd.Col = A_Col#
    Select Case Spd.CellType
      Case SS_CELL_TYPE_COMBOBOX
            If ValueType% = 1 Then
                Spd.TypeComboBoxCurSel = Value$
            Else
                Spd.TypeComboBoxCurSel = GetSpdComboBoxIndex(Spd, tSPD, Row#, FldName$, Value$, Pos#, Separator$)
            End If
      Case Else
           Spd.text = Value$
    End Select
End Sub
Sub SetColPosition(tSPD As Spread, tDefault As Spread)
'�]�w����Y�N��ܪ���춶�ǤαƧ����
Dim A_Flag%, I%, j%, k%, A_Index%, A_RptIndex%
Dim A_CUBound%, A_CUBound2%
Dim A_SUBound%, A_SUBound2%

    A_Flag% = False
    
    '���oUser�ۭq��������αƧ�����
    On Error Resume Next
    A_CUBound% = UBound(tDefault.Columns)
    A_SUBound% = UBound(tDefault.Sorts)
    On Error GoTo 0
    
    '���o�{���w�]��������αƧ�����
    A_CUBound2% = UBound(tSPD.Columns)
    A_SUBound2% = UBound(tSPD.Sorts)

    '�NUser�ۭq����춶��,Update��Spread Type��ReportIndex�ݩʤ�
    If A_CUBound% > 0 Then
       For I% = 1 To A_CUBound%
           For j% = 1 To A_CUBound2%
               If StrComp(tSPD.Columns(j%).Name, tDefault.Columns(I%).Name, vbTextCompare) = 0 Then
                  A_Flag% = True
                  A_Index% = A_Index% + 1
                  tSPD.Columns(j%).ScreenIndex = A_Index%
                  'If tSpd.Columns(J%).Hidden = 0 And Not (tSpd.Columns(J%).BreakIndex > 0 And Not tSpd.Columns(J%).BreakInLine) Then
                  If Not (tSPD.Columns(j%).BreakIndex > 0 And Not tSPD.Columns(j%).BreakInLine) Then
                     A_RptIndex% = A_RptIndex% + 1
                     tSPD.Columns(j%).ReportIndex = A_RptIndex%
                  End If
                  If tSPD.Columns(j%).Hidden = 1 Then tSPD.Columns(j%).Hidden = 0
                  Exit For
               End If
           Next j%
       Next I%
    End If
    
    '�Y�LUser�ۭq�������,�h�H�{���w�]����춶�����
    If Not A_Flag% Then
       For I% = 1 To A_CUBound2%
           If tSPD.Columns(I%).Hidden <> 2 And Not (tSPD.Columns(I%).BreakIndex > 0 And Not tSPD.Columns(I%).BreakInLine) Then
              A_RptIndex% = A_RptIndex% + 1
              tSPD.Columns(I%).ReportIndex = A_RptIndex%
           End If
           tSPD.Columns(I%).ScreenIndex = tSPD.Columns(I%).SelectIndex
           If tSPD.Columns(I%).Hidden = 1 Then tSPD.Columns(I%).Hidden = 0
       Next I%
    Else
       For I% = 1 To A_CUBound2%
           If tSPD.Columns(I%).ScreenIndex = 0 Then
              A_Index% = A_Index% + 1
              tSPD.Columns(I%).ScreenIndex = A_Index%
           End If
           If tSPD.Columns(I%).ReportIndex = 0 Then
              If tSPD.Columns(I%).Hidden = 0 And Not (tSPD.Columns(I%).BreakIndex > 0 And Not tSPD.Columns(I%).BreakInLine) Then
                 A_RptIndex% = A_RptIndex% + 1
                 tSPD.Columns(I%).ReportIndex = A_RptIndex%
              End If
           End If
       Next I%
       '�NSpread Type������춶�Ǩ�ScreenIndex���s�إ�
       RebuildByDefIndex tSPD, 2
    End If
    
    '�Y�������\User�ۭq�Ƨ����,�h���L�U�����ƧǳB�z�޿�
    If tSPD.SortEnable = False Then Exit Sub
    
    '�NUser�ۭq���Ƨ����,Update��Spread Type��
    A_Flag% = False: A_Index% = 0
    If A_SUBound% > 0 Then
       For I% = 1 To A_SUBound%
           If tDefault.Sorts(I%).SortKey = "" Then Exit For
           For j% = 1 To A_CUBound2%
               If StrComp(tSPD.Columns(j%).Name, tDefault.Sorts(I%).SortKey, vbTextCompare) = 0 Then
                  If Not A_Flag% Then
                     A_Flag% = True
                     '�Y��User�ۭq���Ƨ����,�h���NSpread Type�}�C�ȲM��
                     For k% = 1 To A_SUBound2%
                         tSPD.Sorts(k%).SortKey = ""
                         tSPD.Sorts(k%).SortOrder = 0
                     Next k%
                  End If
                  A_Index% = A_Index% + 1
                  tSPD.Sorts(A_Index%).SortKey = tDefault.Sorts(I%).SortKey
                  tSPD.Sorts(A_Index%).SortOrder = tDefault.Sorts(I%).SortOrder
                  Exit For
               End If
           Next j%
       Next I%
    End If
End Sub

Sub RebuildByDefIndex(tSPD As Spread, ByVal style%)
'��ScreenIndex���s�إ�Spread Type������춶��
'Style%=1,��ReportIndex���ǱƦC
'Style%=2,��ScreenIndex���ǱƦC
'Style%=3,��SelectIndex���ǱƦC
'Style%=4,��TempIndex���ǱƦC
'Style%=5,��BreakIndex���ǱƦC
Dim I%, A_Index%, A_Cols%

    A_Cols% = UBound(tSPD.Columns)

    ReDim tCols(1 To A_Cols%) As SpreadCol
    For I% = 1 To A_Cols%
        Select Case style%
          Case 1
               A_Index% = tSPD.Columns(I%).ReportIndex
          Case 2
               A_Index% = tSPD.Columns(I%).ScreenIndex
          Case 3
               A_Index% = tSPD.Columns(I%).SelectIndex
          Case 4
               A_Index% = tSPD.Columns(I%).TempIndex
          Case 5
               A_Index% = tSPD.Columns(I%).BreakIndex
        End Select
        tCols(A_Index%).BreakInLine = tSPD.Columns(I%).BreakInLine
        tCols(A_Index%).ReportIndex = tSPD.Columns(I%).ReportIndex
        tCols(A_Index%).ScreenIndex = tSPD.Columns(I%).ScreenIndex
        tCols(A_Index%).SelectIndex = tSPD.Columns(I%).SelectIndex
        tCols(A_Index%).TempIndex = tSPD.Columns(I%).TempIndex
        tCols(A_Index%).BreakIndex = tSPD.Columns(I%).BreakIndex
        tCols(A_Index%).Hidden = tSPD.Columns(I%).Hidden
        tCols(A_Index%).ColWidth = tSPD.Columns(I%).ColWidth
        tCols(A_Index%).Name = tSPD.Columns(I%).Name
        tCols(A_Index%).text = tSPD.Columns(I%).text
        tCols(A_Index%).Caption = tSPD.Columns(I%).Caption
        tCols(A_Index%).CFormat = tSPD.Columns(I%).CFormat
        tCols(A_Index%).dFormat = tSPD.Columns(I%).dFormat
    Next I%
    tSPD.Columns = tCols
End Sub
Sub GetPropertyFromDB(ByVal DBName$, ByVal TBName$, ByVal TBField$, length%, FCaption$, FormatStr$, Up$, Down$)
'�۸�Ʈw�����o��쪺�ݩʭ�
Dim A_Sql$
    
    If Trim(DBName$) <> "" And Trim(TBName$) <> "" And Trim(TBField$) <> "" Then
       A_Sql$ = "Select DEF07,DEF09,DEF11,DEF12,DEF13 From TBLDEF"
       A_Sql$ = A_Sql$ & " where DEF01='" & DBName$ & "'"
       A_Sql$ = A_Sql$ & " and DEF02='" & TBName$ & "'"
       A_Sql$ = A_Sql$ & " and DEF05='" & TBField$ & "'"
       CreateDynasetODBC DB_ARTHGUI, DY_TBLDEF, A_Sql$, "DY_TBLDEF", True
       If Not (DY_TBLDEF.BOF And DY_TBLDEF.EOF) Then
          '���oTextBox������
          If CDbl(DY_TBLDEF.Fields("DEF07") & "") <> 0 Then
             length% = CDbl(DY_TBLDEF.Fields("DEF07") & "")
          End If
          '���oLabel�����D
          If Trim(DY_TBLDEF.Fields("DEF09") & "") <> "" Then
             FCaption$ = Trim(DY_TBLDEF.Fields("DEF09") & "")
          End If
          '���oTextBox����J�榡
          FormatStr$ = Trim(DY_TBLDEF.Fields("DEF13") & "")
          '���o�ƭȿ�J���W����
          If Trim(DY_TBLDEF.Fields("DEF11") & "") <> "" Then
             Up$ = Trim(DY_TBLDEF.Fields("DEF11") & "")
          End If
          '���o�ƭȿ�J���U����
          If Trim(DY_TBLDEF.Fields("DEF12") & "") <> "" Then
             Down$ = Trim(DY_TBLDEF.Fields("DEF12") & "")
          End If
       End If
    End If
End Sub
Function GetReportCols(tSPD As Spread) As Double
'���o����C�L�`���
Dim I#, Count#

    For I# = 1 To UBound(tSPD.Columns)
        If tSPD.Columns(I#).ReportIndex > 0 Then
           Count# = Count# + 1
        End If
    Next I#
    GetReportCols = Count#
End Function
Sub AddSpreadMaxRows(Spd As vaSpread, Row#)
'�bvaSpread�W�W�[�@�C

    Spd.MaxRows = Spd.MaxRows + 1
    Row# = Spd.MaxRows
End Sub
Function GetMergeCols(ByVal Col#, ByVal Row#, ByVal ColsCount#, ByVal MergeCols#, ByVal KeepCols#) As String
'���oExcel�W�̨Ϊ��X������}
Dim A_Start$, A_End$
    
    'A_Start$ = Chr(Col# + 64) & Trim(Row#)
    A_Start$ = GetExcelColName(Col#) & Trim(Row#)
    If Col# + MergeCols# + KeepCols# >= ColsCount# Then
       If ColsCount# - Col# - KeepCols# >= MergeCols# Then
          'A_End$ = Chr(Col# + MergeCols# - 1 + 64)
          A_End$ = GetExcelColName(Col# + MergeCols# - 1)
       Else
            If ColsCount# - Col# - KeepCols# < 0 Then
                'A_End$ = Chr(Col# + 64)
                A_End$ = GetExcelColName(Col#)
            Else
                'A_End$ = Chr(Col# + (ColsCount# - Col# - KeepCols#) + 64)
                A_End$ = GetExcelColName(Col# + (ColsCount# - Col# - KeepCols#))
            End If
       End If
    Else
       'A_End$ = Chr(Col# + MergeCols# - 1 + 64)
       A_End$ = GetExcelColName(Col# + MergeCols# - 1)
    End If
    GetMergeCols = A_Start$ & ":" & A_End$ & Trim(Row#)
End Function

Function PrintStrConnect(tSPD As Spread, ByVal FType%) As String
'��Spread Type�걵���C�L���r��
'FType% = 1,�Ǧ^�C�L��Screen����Ʀr��
'FType% = 2,�Ǧ^�C�L��Report����Ʀr��
'FType% = 3,�Ǧ^�C�L��Report�����D�r��
Dim I#, A_PrtStr$

    A_PrtStr$ = ""
    For I# = 1 To UBound(tSPD.Columns)
        Select Case FType%
          Case 1
               A_PrtStr$ = A_PrtStr$ & tSPD.Columns(I#).text & G_G1
               tSPD.Columns(I#).text = ""
          Case 2
               If tSPD.Columns(I#).ReportIndex > 0 Then
                  A_PrtStr$ = A_PrtStr$ & tSPD.Columns(I#).text & G_G1
               End If
               tSPD.Columns(I#).text = ""
          Case 3
               If tSPD.Columns(I#).ReportIndex > 0 Then
                  A_PrtStr$ = A_PrtStr$ & tSPD.Columns(I#).Caption & G_G1
               End If
        End Select
    Next I#
    If A_PrtStr$ <> "" Then A_PrtStr$ = Left(A_PrtStr$, Len(A_PrtStr$) - 1)
    
    PrintStrConnect = A_PrtStr$
End Function


Function GetRptColName(tSPD As Spread, ByVal Col#) As String
'�H���Ǹ����o�����ҥN�����W��
Dim I#, A_ColName$

    GetRptColName = ""
    For I# = 1 To UBound(tSPD.Columns)
        If tSPD.Columns(I#).ReportIndex = Col# Then
           GetRptColName = tSPD.Columns(I#).Name
           Exit For
        End If
    Next I#
End Function

Sub GetSpreadDefault(tSPD As Spread, ByVal FormName$, ByVal SpreadName$)
'��Data���|�U��EXEName.INI,���o�w�]��Spread��춶�ǤαƧ����
'Section : [User ID]
'Topic   : Form Name/Spread Name/Column=Field 1;Field 2; ..... ;Field N
'Topic   : Form Name/Spread Name/Sort=Field 1;Field 2; ..... ;Field N
Dim tSpd_Temp As Spread
Dim A_IniPath$, A_Section$, A_Topic$, A_Value$
Dim A_Cols$(), A_Sorts$(), I#, A_CUBound#, A_SUBound#

    A_IniPath$ = G_INI_SerPath & "Data\" & App.EXEName & ".INI"
    A_Section$ = GetUserId()
    
    '��EXEName.INI�����o�ϥΪ̦ۭq����춶�Ǧr���
    A_Topic$ = FormName$ & "/" & SpreadName$ & "/Column"
    A_Value$ = GetIniStr(A_Section$, A_Topic$, A_IniPath$)
    If A_Value$ = "" Then GoTo GetSpreadDefaultA
    A_Cols$ = Split(A_Value$, ";", , vbTextCompare)
    
    '��EXEName.INI�����o�ϥΪ̦ۭq���Ƨ���춶�Ǧr���
    A_Topic$ = FormName$ & "/" & SpreadName$ & "/Sort"
    A_Value$ = GetIniStr(A_Section$, A_Topic$, A_IniPath$)
    If A_Value$ <> "" Then
       A_Sorts$ = Split(A_Value$, ";", , vbTextCompare)
    End If
    
    '���o������αƧ���쪺�W����
    A_CUBound# = UBound(A_Cols$) + 1
    A_SUBound# = UBound(A_Sorts$) + 1
    
    '�ŧi���A�}�C
    ReDim tSpdCol(1 To A_CUBound#) As SpreadCol
    ReDim tSpdSort(1 To 3) As SpreadSort

    '�N�ϥΪ̦ۭq����춶��,��JSpread Type���A��
    tSpd_Temp.Columns = tSpdCol
    For I# = 1 To A_CUBound#
        tSpd_Temp.Columns(I#).Name = A_Cols$(I# - 1)
    Next I#

    '�N�ϥΪ̦ۭq���Ƨ���춶��,��JSpread Type���A��
    tSpd_Temp.Sorts = tSpdSort
    For I# = 1 To A_SUBound#
        If I# > 3 Then Exit For
        tSpd_Temp.Sorts(I#).SortKey = Replace(A_Sorts$(I# - 1), "-", "", 1, 1, vbTextCompare)
        tSpd_Temp.Sorts(I#).SortOrder = IIf(InStr(1, A_Sorts$(I# - 1), "-", vbTextCompare) > 0, SS_SORT_ORDER_DESCENDING, SS_SORT_ORDER_ASCENDING)
    Next I#
    
GetSpreadDefaultA:
    '�]�w����N��ܪ����ζ���
    SetColPosition tSPD, tSpd_Temp
End Sub
Sub AddReportCol(tSPD As Spread, ByVal ColName$, Optional ByVal Hidden% = 1, _
Optional ByVal SortIndex% = 0, Optional ByVal SortOrder% = SS_SORT_ORDER_ASCENDING, _
Optional ByVal BreakCol% = 0, Optional ByVal BreakInLine As Boolean = True)
'�]�w������ܪ����αƧ�����Spread Type��
Dim I%

    '�]�w�������Ҧ��i��ܪ����
    For I% = 1 To UBound(tSPD.Columns)
        If Trim(tSPD.Columns(I%).Name) = "" Then
           tSPD.Columns(I%).SelectIndex = I%
           tSPD.Columns(I%).Name = ColName$
           tSPD.Columns(I%).BreakIndex = BreakCol%
           tSPD.Columns(I%).BreakInLine = BreakInLine
           tSPD.Columns(I%).Hidden = Hidden%
           Exit For
        End If
    Next I%
    
    '�]�w�Ƨ����
    If SortIndex% <> 0 Then
       tSPD.Sorts(SortIndex%).SortKey = ColName$
       tSPD.Sorts(SortIndex%).SortOrder = SortOrder%
    End If
End Sub
Sub InitialCols(tSPD As Spread, ByVal Cols#, ByVal SortEnable As Boolean)
'�ŧiSpread���A����Columns��Sorts���A
ReDim tCols(1 To Cols#) As SpreadCol
ReDim tSorts(1 To Cols#) As SpreadSort

    tSPD.SortEnable = SortEnable
    tSPD.Columns = tCols
    tSPD.Sorts = tSorts
End Sub


Sub SpdFldProperty(Spd As vaSpread, tSPD As Spread, ByVal FldName$, ByVal Width%, _
ByVal Caption$, ByVal CellType%, Optional ByVal MIN$, Optional ByVal MAX$, _
Optional ByVal length%, Optional ByVal HAlign% = SS_CELL_H_ALIGN_LEFT, _
Optional ByVal RAlign% = SS_CELL_H_ALIGN_LEFT, Optional ByVal RDateFormat% = False, _
Optional ByVal DBName$, Optional ByVal TBName$, Optional ByVal TBField$, Optional Multi% = False)
'�]�wSpread��쪺�ݩ�
Dim A_Pos%, A_SNo#, A_CLen%, A_MaxLen%, A_FChar$
Dim Hide As Boolean
    
    
    '�H���W�٨��o������ܪ����
    A_SNo# = GetSpdColIndex(tSPD, FldName$)
    
'    '�]�w���欰�������,���}��User���
'    If Hide Then tSpd.Columns(A_SNo#).Hidden = 2
    
    If tSPD.Columns(A_SNo#).Hidden > 0 Then Hide = True
    
    '�۸�Ʈw�����o��쪺�ݩʭ�(����,���D,�榡,�W�U����)
    GetPropertyFromDB DBName$, TBName$, TBField$, length%, Caption$, "", MAX$, MIN$
    
    '�N�����D����eKeep��Spread Type
    tSPD.Columns(A_SNo#).Caption = Caption$
    tSPD.Columns(A_SNo#).mCaption = Caption$
    tSPD.Columns(A_SNo#).ColWidth = Width%

    '�]�w���������D�θ�ƦC�L��Format
    A_CLen% = lstrlen(Caption$)
    '20110429 Add�Y���D���h����ܮɨ̳]�w�����שw�q(Yvonne)
    If Multi% = True Then
        A_MaxLen% = length%
    Else
        A_MaxLen% = IIf(A_CLen% >= length%, A_CLen%, length%)
    End If
    Select Case RAlign%
      Case SS_CELL_H_ALIGN_LEFT
           tSPD.Columns(A_SNo#).CFormat = String(A_MaxLen%, "#")
           tSPD.Columns(A_SNo#).dFormat = String(A_MaxLen%, "#")
      Case SS_CELL_H_ALIGN_CENTER
           If A_CLen% >= length% Then
              tSPD.Columns(A_SNo#).CFormat = String(A_MaxLen%, "^")
              tSPD.Columns(A_SNo#).dFormat = String(A_MaxLen%, "^")
           Else
              tSPD.Columns(A_SNo#).CFormat = String(A_MaxLen%, "#")
              tSPD.Columns(A_SNo#).dFormat = String(A_MaxLen%, "^")
           End If
      Case SS_CELL_H_ALIGN_RIGHT
           tSPD.Columns(A_SNo#).CFormat = String(A_MaxLen%, "~")
           tSPD.Columns(A_SNo#).dFormat = String(A_MaxLen%, "~")
    End Select
    
    '�����X��Excel��,�O�_�N������榡�Ʀ�����榡
    tSPD.Columns(A_SNo#).DateFormat = RDateFormat%
    
    '�]�wColumn�e��
    Spd.ColWidth(A_SNo#) = Width%
    
    '�]�w�����D
    Spd.Row = 0
    Spd.Col = A_SNo#
    Spd.text = Caption$

    Spd.Row = -1
    Spd.Col = A_SNo#
    
    '�]�w���O�_����
    Spd.ColHidden = Hide
    
    '�]�w�C�檺��ƫ��A
    If Spd.MaxRows > 0 Then
        Spd.Row = 1
        Spd.Row2 = Spd.MaxRows
        Spd.Col = A_SNo#
        Spd.Col2 = A_SNo#
    End If
    Spd.CellType = CellType%
    Select Case CellType%
      Case 1, 5
           Spd.TypeHAlign% = HAlign%
           Spd.TypeEditLen = length%                                                    '��r��Ƥ�����
      Case 2
           Spd.TypeFloatMin = MIN$                                                      '�B�I�Ƥ��̤p��
           Spd.TypeFloatMax = MAX$                                                      '�B�I�Ƥ��̤j��
           Spd.TypeFloatDecimalChar = Asc(".")                                          '�]�w�p���I����ܫ��A
           A_Pos% = InStr(1, MAX$, ".", vbTextCompare)
           Spd.TypeFloatDecimalPlaces = IIf(A_Pos% > 0, Len(MAX$) - A_Pos%, 0)          '�]�w�p�Ʀ����
           Spd.TypeFloatSeparator = True                                                '�]�w�T��@ ,
      Case 3
           Spd.TypeIntegerMin = MIN$                                                    '��Ƥ��̤p��
           Spd.TypeIntegerMax = MAX$                                                    '��Ƥ��̤j��
    End Select
End Sub
Sub VSElastic_Property2(vsEtc As VideoSoftElastic)
'�]�wElastic���ݩ�,�S���p�ϥ�

    'General Defined
    vsEtc.Template = 0          'tpNone
    vsEtc.style = esClassic
    'Panels Defined
    vsEtc.Align = asFill
    vsEtc.AutoSizeChildren = azProportional
    vsEtc.BackColor = Val(G_Label_Color)
    vsEtc.ForeColor = Val(G_TextLostFore_Color)
    vsEtc.BevelOuter = bsGroove
    vsEtc.BevelInner = 0        'bsNone
    vsEtc.BevelOuterDir = bdBoth
    vsEtc.BevelChildren = bcAll
    vsEtc.BevelOuterWidth = 2
    vsEtc.BorderWidth = 2
End Sub



Sub RefreshSpreadData(Spd As vaSpread, tSPD As Spread)
'���}frm_RptDef����,�����s����Spread�W������T

'�Y��춶�Ǧ�����,�h���sPrepare Spread�W�����
    If tSPD.RefreshCol Then ProcessChangeCols Spd, tSPD
    
'�Y�Ƨ���즳����,�h��Spread���s�Ƨ�
    If tSPD.RefreshSort Then SpreadColsSort Spd, tSPD
End Sub


Sub SpreadColsSort(Spd As vaSpread, tSPD As Spread, Optional ByVal Row# = 1, Optional ByVal Col# = 1, _
Optional ByVal Row2# = -1, Optional ByVal Col2# = -1)
'�Q��Spread Type��Sort
Dim I#, A_Col#, A_Index#

    Screen.MousePointer = HOURGLASS
    
    With Spd
         .Row = Row#
         .Col = Col#
         If Row2# = -1 And Col2# = -1 Then
            .Row2 = .MaxRows
            .Col2 = .MaxCols
         Else
            .Row2 = Row2#
            .Col2 = Col2#
         End If
         .SortBy = SS_SORT_BY_ROW
         For I# = 1 To UBound(tSPD.Sorts)
             If Trim(tSPD.Sorts(I#).SortKey) <> "" Then
                '�H���W�٨��o��쪺Index
                A_Col# = GetSpdColIndex(tSPD, tSPD.Sorts(I#).SortKey)
                If A_Col# <> 0 Then
                   A_Index# = A_Index# + 1
                   .SortKey(A_Index#) = A_Col#
                   .SortKeyOrder(A_Index#) = tSPD.Sorts(I#).SortOrder
                End If
             End If
         Next I#
         .Action = SS_ACTION_SORT
    End With
    
    Screen.MousePointer = Default
End Sub
Function GetSpdColIndex(tSPD As Spread, ByVal FldName$) As Double
'�HvaSpread���ۭq���W�٨��o��쪺Index
Dim I#, A_Cols#

    GetSpdColIndex = 0
    
    A_Cols# = UBound(tSPD.Columns)
    For I# = 1 To A_Cols#
        If StrComp(tSPD.Columns(I#).Name, FldName$, vbTextCompare) = 0 Then
           GetSpdColIndex = I#
           Exit For
        End If
    Next I#
End Function


Sub ProcessChangeCols(Spd As vaSpread, tSPD As Spread)
'�B�z�b�ù���ܵe����������춶�Ǯ�,���sPrepare Spread�W��ƪ��{��
Dim I%, j%

    Screen.MousePointer = HOURGLASS
    
    '��Spread Type�HScreenIndex�����ǤU,�Q��TempIndex��Keep�U�Ӫ���Spread�W������m,
    '�i����,�HChange��춶�ǩ�����������,��粒��,�ñNTempIndex�k�s
    For I% = 1 To UBound(tSPD.Columns)
        '�ثeSpread�W������m�P�s���]�w���Ů�,�B�z�H�U�ʧ@
        If tSPD.Columns(I%).ScreenIndex <> tSPD.Columns(I%).TempIndex Then
           ChangeSpdCols Spd, tSPD.Columns(I%).TempIndex, tSPD.Columns(I%).ScreenIndex
           For j% = tSPD.Columns(I%).ScreenIndex + 1 To UBound(tSPD.Columns)
               If tSPD.Columns(j%).TempIndex >= tSPD.Columns(I%).ScreenIndex And _
               tSPD.Columns(j%).TempIndex < tSPD.Columns(I%).TempIndex Then
                  tSPD.Columns(j%).TempIndex = tSPD.Columns(j%).TempIndex + 1
               End If
           Next j%
        End If
        '�̷s���]�w�ȭ��]�����ܩ�����
        Spd.Col = tSPD.Columns(I%).ScreenIndex
        'Spd.ColHidden = (tspd.Columns(I%).ReportIndex = 0 Or tspd.Columns(I%).Hidden)
        Spd.ColHidden = (tSPD.Columns(I%).Hidden > 0)
        '���s�]�w��e(�]������ë�,��e�Y����)
        If Spd.ColHidden = 0 Then Spd.ColWidth(Spd.Col) = tSPD.Columns(I%).ColWidth
        '�NTempIndex�ݩʭ��k�s
        tSPD.Columns(I%).TempIndex = 0
    Next I%
    
    Screen.MousePointer = Default
End Sub


Sub SaveSpreadDefault(tSPD As Spread, ByVal FormName$, ByVal SpreadName$)
'�N�ثevaSpread�W����춶�ǤαƧ����,�s��Data���|�U��EXEName.INI
'Section : [User ID]
'Topic   : Form Name/Spread Name/Column=Field 1;Field 2; ..... ;Field N
'Topic   : Form Name/Spread Name/Sort=Field 1;Field 2; ..... ;Field N
Dim A_IniPath$, A_Section$, A_Topic$, A_Value$
Dim A_Cols$(), A_Sorts$(), I#, A_CUBound#, A_SUBound#

    A_IniPath$ = G_INI_SerPath & "Data\" & App.EXEName & ".INI"
    A_Section$ = GetUserId()
    
    '�N�ϥΪ̦ۭq����춶�Ǧr��Ȧs��Report.INI��
    A_Topic$ = FormName$ & "/" & SpreadName$ & "/Column"
    A_Value$ = ""
    For I# = 1 To UBound(tSPD.Columns)
        If tSPD.Columns(I#).Hidden > 0 Then Exit For
        A_Value$ = A_Value$ & tSPD.Columns(I#).Name & ";"
    Next I#
    A_Value$ = Left$(A_Value$, Len(A_Value$) - 1)
    UpdateIniValue A_Section$, A_Topic$, A_Value$, A_IniPath$
    
    '�N�ϥΪ̦ۭq���Ƨ���춶�Ǧr��Ȧs��Report.INI��
    A_Topic$ = FormName$ & "/" & SpreadName$ & "/Sort"
    A_Value$ = ""
    For I# = 1 To UBound(tSPD.Sorts)
        If Trim(tSPD.Sorts(I#).SortKey) = "" Then Exit For
        If tSPD.Sorts(I#).SortOrder = SS_SORT_ORDER_DESCENDING Then
           A_Value$ = A_Value$ & "-"
        End If
        A_Value$ = A_Value$ & tSPD.Sorts(I#).SortKey & ";"
    Next I#
    If Trim(A_Value$) <> "" Then A_Value$ = Left$(A_Value$, Len(A_Value$) - 1)
    UpdateIniValue A_Section$, A_Topic$, A_Value$, A_IniPath$
End Sub

Function GetOrderCols(tSPD As Spread, ByVal Order$) As String
'�Ǧ^SQL���O�����Ƨ����
Dim I#, A_STR$

    '�w�]���Ƨ���쬰�{���]�w�����,�Y�C�L�ܿù�,���b�}�ɮɨ�User�ۭq���Ƨ����Ƨ�
    GetOrderCols = " ORDER BY " & Order$
    If G_PrintSelect = G_Print2Screen Then Exit Function
    
    '���oUser�ۭq���Ƨ����
    For I# = 1 To UBound(tSPD.Sorts)
        If Trim(tSPD.Sorts(I#).SortKey) = "" Then Exit For
        A_STR$ = A_STR$ & tSPD.Sorts(I#).SortKey
        If tSPD.Sorts(I#).SortOrder = SS_SORT_ORDER_DESCENDING Then
           A_STR$ = A_STR$ & " DESC"
        End If
        A_STR$ = A_STR$ & ","
    Next I#
    If A_STR$ = "" Then Exit Function
    
    A_STR$ = Left$(A_STR$, Len(A_STR$) - 1)
    GetOrderCols = " ORDER BY " & A_STR$
End Function


Function CheckDirectoryExist(ByVal Str$) As Boolean
'�ˮָ��|�O�_�s�b
Dim A_Pos%

    CheckDirectoryExist = True
    
    A_Pos% = InStrRev(Str$, "\", -1, vbTextCompare)
    If A_Pos% <> 0 Then
       Str$ = Mid(Str$, 1, A_Pos% - 1)
       On Error Resume Next
       ChDir Str$
       If Err Then
          Err = 0
          CheckDirectoryExist = False
       End If
    End If
End Function



Function GetSubSystemID(ByVal APID$, ByVal PgID$, ByVal APOpt$) As String
'���o�{���k�ݪ��\��Ҳ�
Dim A_Sql$

    GetSubSystemID = ""
    
    A_Sql$ = "Select A1005 From A10"
    A_Sql$ = A_Sql$ & " where A1001='" & PgID$ & "'"
    A_Sql$ = A_Sql$ & " and A1003='" & APID$ & "'"
    CreateDynasetODBC DB_ARTHGUI, DY_A10, A_Sql$, "DY_A10", True
    If Not (DY_A10.BOF And DY_A10.EOF) Then
       GetSubSystemID = Trim(APOpt$) & Trim(DY_A10.Fields("A1005") & "")
    End If
End Function

Function GetTerminalID() As String
'���o�����W��
Dim A_ComputerName$
 
    GetTerminalID = ""
    '
    A_ComputerName$ = Space$(200)
    If GetComputerName(A_ComputerName$, 200) Then
       A_ComputerName$ = StripTerminator(Trim$(A_ComputerName$))
    End If
    GetTerminalID = A_ComputerName$
End Function

Sub HaveCheckTerminal()
'�P�_AP�O�_�����ˬd�׺ݾ������v�ϥ�
Dim A_Sql$

    G_Terminal_Check = False
    
    A_Sql$ = "Select TopicValue From SINI"
    A_Sql$ = A_Sql$ & " where Section='Terminal'"
    A_Sql$ = A_Sql$ & " and Topic='Check'"
    CreateDynasetODBC DB_ARTHGUI, DY_SINI, A_Sql$, "DY_SINI", True
    If Not (DY_SINI.BOF And DY_SINI.EOF) Then
       If UCase$(Trim(DY_SINI.Fields("TopicValue") & "")) = "Y" Then
          G_Terminal_Check = True
       End If
    End If
End Sub

Function HaveTerminalLicense(ByVal APOpt$) As Boolean
'�ˮֲ׺ݾ��O�_���v�ϥΥ\��s��
Dim A_Sql$

    HaveTerminalLicense = False
    
    A_Sql$ = "Select TopicValue From SINI"
    A_Sql$ = A_Sql$ & " where Section='" & APOpt$ & "'"
    A_Sql$ = A_Sql$ & " and Topic='" & ChangePCName(GetTerminalID()) & "'"
    CreateDynasetODBC DB_ARTHGUI, DY_SINI, A_Sql$, "DY_SINI", True
    If Not (DY_SINI.BOF And DY_SINI.EOF) Then
       HaveTerminalLicense = True
    End If
End Function

Function ChangePCName(ByVal A_PCName$) As String
Dim I%, A_Chr$, A_STR$

    For I% = 1 To Len(A_PCName$)
        A_Chr$ = Mid(A_PCName$, I%, 1)
        A_STR$ = A_STR$ & Hex(Oct(Asc(A_Chr$)))
    Next I%
    ChangePCName = A_STR$
End Function

Sub ExecuteProcessReturnErr(DB As Database, ByVal SQL$, Optional ErrCode)
'�����Ʈw���s�W,�ק�,�R���ʧ@,���ѮɶǦ^���~�T��
On Local Error GoTo ExecuteProcessReturnErr_Error

    ErrCode = 0
    
    If Trim$(DB.Connect) = "" Then           'Access DataBase
       DB.Execute SQL$, dbFailOnError
    Else
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;", "ORACLE;"
              DB.Execute SQL$, dbSQLPassThrough
         Case "DB2;"
              DB.Execute SQL$
       End Select
    End If
    Exit Sub
    
ExecuteProcessReturnErr_Error:
    ErrCode = Err
End Sub

Function ReturnPostExecuteState(DB As Database) As Integer
'�Ǧ^Process�{����,�s�W,�ק�,�R����SQL���O���檬�A
'���\=0;���s����=IDOK;�פ�=IDCANCEL
Dim A_Msg$

    ReturnPostExecuteState = 0
    ExecuteProcessReturnErr DB, G_Str, G_ExecuteErr
    If G_ExecuteErr <> 0 Then GoTo ReturnPostExecuteState_Error
    Exit Function
    
ReturnPostExecuteState_Error:
    Select Case G_ExecuteErr
      Case 3008, 3022, 3218, 3260, 3146  'Lock Error Code
           ReturnPostExecuteState = IDOK
      Case Else
           A_Msg$ = GetODBCErrorMessage()
           retcode = MsgBox(A_Msg$, MB_ICONSTOP, UCase$(App.Title))
           ReturnPostExecuteState = IDCANCEL
    End Select
End Function


Function ReturnSaveExecuteState(DB As Database, Optional ByVal Table$ = "") As Integer
'�Ǧ^�s�ɮɪ�SQL���O���檬�A
'���\=0;���s����=IDOK;�פ�=IDCANCEL;���s����=IDRETRY
Dim A_Msg$

    ReturnSaveExecuteState = 0
    If Trim(Table$) = "" Then
       ExecuteProcessReturnErr DB, G_Str, G_ExecuteErr
    Else
       SQLInsert1 DB, Table$, G_ExecuteErr
    End If
    If G_ExecuteErr <> 0 Then GoTo ReturnSaveExecuteState_Error
    Exit Function
    
ReturnSaveExecuteState_Error:
    Select Case G_ExecuteErr
      Case 3008, 3046, 3158, 3186, 3187, 3188, 3202, 3218, 3260     'Lock Error Code
           Idle
           MsgBox GetSIniStr("PanelDescpt", "unread"), vbOKOnly, UCase$(App.Title)
           ReturnSaveExecuteState = IDOK
      Case 3022                          'Duplicate (Access)
           ReturnSaveExecuteState = IDRETRY
      Case 3146                          'Duplicate (SQL Server)
           For Each G_Err In GetEngine.Errors
               If G_Err.Number = 2601 Or G_Err.Number = 2627 Then   'PKEY�]�����������,Duplicate��Number=2627
                  ReturnSaveExecuteState = IDRETRY
                  Exit For
               End If
           Next G_Err
           If ReturnSaveExecuteState <> IDRETRY Then GoTo Process_Start
      Case Else
Process_Start:
           A_Msg$ = GetODBCErrorMessage()
           retcode = MsgBox(A_Msg$, MB_ICONSTOP, UCase$(App.Title))
           ReturnSaveExecuteState = IDCANCEL
    End Select
End Function

Function GetODBCErrorMessage() As String
'���oSQL Server��Ʈw����SQL���O�o�ͪ��Ҧ����~�T��
Dim A_Msg$

    A_Msg$ = ""
    For Each G_Err In GetEngine.Errors
        A_Msg$ = A_Msg$ & G_Err.Number & ":"
        A_Msg$ = A_Msg$ & G_Err.Description
        A_Msg$ = A_Msg$ & Chr$(13) & Chr$(10)
    Next G_Err
    A_Msg$ = Left$(A_Msg$, Len(A_Msg$) - 2)
    Set G_Err = Nothing
    '
    GetODBCErrorMessage = A_Msg$
End Function

Function GetTextBoxStrArray(Txt As Control, ByVal MaxLen%) As String()
'�NTextBox�W���C�C���Keep��Array��,�s�ɮɨϥ�
'**********************************************************************
'Function �ޥΤ��d�ҵ{��,�ǤJ��ӰѼ�
'Txt : TextBox Control Name   MaxLen% : �C�C��ƪ��׳̤j��
'**********************************************************************
'�ŧiArray�ܼ�
'Dim A_Str$(), I%
'
'    �NTextBox�W���C�C���Keep��Array
'    A_Str$ = GetTextBoxStrArray(Text1, 120)
'
'    ��Array�����X�C�C��ƳB�z
'    I% = 0
'    Do While I% < UBound(A_Str$)
'       I% = I% + 1
'       MsgBox CStr(I%) & " : " & A_Str$(I%)
'    Loop
'**********************************************************************
Dim I&, A_Line&
    
    ReDim A_STR$(0)
    GetTextBoxStrArray = A_STR$
    If Trim(Txt.text) = "" Then Exit Function
    
    A_Line& = GetTextBoxLineCount(Txt)
    ReDim A_STR$(1 To A_Line&)
    
    For I& = 0 To A_Line& - 1
        A_STR$(I& + 1) = GetTextBoxLineStr(Txt, MaxLen%, I&)
        If Len(A_STR$(I& + 1)) > 0 Then
           A_STR$(I& + 1) = StripTerminator(A_STR$(I& + 1))
           A_STR$(I& + 1) = RTrim(A_STR$(I& + 1))
        End If
    Next I&
    
    GetTextBoxStrArray = A_STR$
End Function

Sub SetFramePosition(Fra As Control, Spd As vaSpread, ByVal Left%, ByVal Top%, ByVal Width%, ByVal Height%)
'�]�w���U������Frame��Spread����m
    
    Screen.ActiveForm!Vse_Background.AutoSizeChildren = azNone
    Fra.Move Left%, Top%, Width%, Height%
    Spd.Move 90, 180, Fra.Width - 200, Fra.Height - 300
    Screen.ActiveForm!Vse_Background.AutoSizeChildren = azProportional
End Sub
Sub SpreadWarnLine(Spd As vaSpread, ByVal Row#)
'�R��Spread Row��,�N�ӦC�C��H�©��զr���
On Local Error Resume Next

    If Row <= 0 Then Exit Sub
    Spd.Row = Row#
    Spd.Col = -1
    Spd.BackColor = COLOR_BLACK
    Spd.ForeColor = COLOR_WHITE
End Sub
Sub SpreadWarnLineCancel(Spd As vaSpread, ByVal Row#)
'�����R��Spread Row��,�٭�ӦC�C��
On Local Error Resume Next

    If Row <= 0 Then Exit Sub
    Spd.Row = Row#
    Spd.Col = -1
    Spd.BackColor = COLOR_WHITE
    Spd.ForeColor = COLOR_BLACK
End Sub

Function SetSpreadTopRow(Spd As vaSpread) As Double
'���oSpread��ܭ��W���Ĥ@�C�C��
Dim A_PageRows%
    
    SetSpreadTopRow = 1
    
    With Spd
         A_PageRows% = .Height / .RowHeight(0) - 2
         If A_PageRows% = 0 Then Exit Function
         If .MaxRows \ A_PageRows% < 1 Then Exit Function
         SetSpreadTopRow = (.MaxRows \ A_PageRows% - 1) * A_PageRows% + 1
    End With
End Function

Sub GetProgramName(Optional ByVal A0906$ = "")
'���o�{���W��
Dim A_Sql$

    A_Sql$ = "Select A1002 From A10"
    If Trim(A0906$) = "" Then
       A_Sql$ = A_Sql$ & " where A1001='" & G_CmdStr1$ & "'"
    Else
       A_Sql$ = A_Sql$ & " where A1001='" & A0906$ & "'"
    End If
    CreateDynasetODBC DB_ARTHGUI, DY_A10, A_Sql$, "DY_A10", True
    If Not (DY_A10.BOF And DY_A10.EOF) Then
       G_ProgramName = Trim(DY_A10.Fields("A1002") & "")
    End If
End Sub
Function GetPGName(ByVal SystemID$, ByVal PgID$) As String
'���o�t�Ϊ��{���W��
Dim A_Sql$

    GetPGName = ""
    
    A_Sql$ = "Select A1002 From A10"
    A_Sql$ = A_Sql$ & " where A1001='" & PgID$ & "'"
    A_Sql$ = A_Sql$ & " and A1003='" & SystemID$ & "'"
    CreateDynasetODBC DB_ARTHGUI, DY_A10, A_Sql$, "DY_A10", True
    If Not (DY_A10.BOF And DY_A10.EOF) Then
       GetPGName = Trim(DY_A10.Fields("A1002") & "")
    End If
    DY_A10.Close
End Function



Sub KeepOpenWorkSpace(File As Workspace, ByVal Name$)
'�N�}�Ҫ�WorkSpace,Keep��Array��
Dim A_Index%

    A_Index% = 0
    Do While A_Index% < 100
       If UCase$(Trim$(G_WorkFile(A_Index%))) = UCase$(Trim$(Name$)) Then
          Set G_WorkName(A_Index%) = File
          G_WorkFile(A_Index%) = Trim$(Name$)
          Exit Do
       End If
       If G_WorkName(A_Index%) Is Nothing Then
          Set G_WorkName(A_Index) = File
          G_WorkFile(A_Index%) = Trim$(Name$)
          Exit Do
       End If
       A_Index% = A_Index% + 1
    Loop
    
End Sub

Sub Checkbox_Property(Tmp As Control, ByVal text$, ByVal Size$, ByVal FName$)
'�]�wCheck Box���ݩ�
On Error Resume Next

    Tmp.Caption = text$
    If Trim$(FName$) <> "" Then Tmp.FontName = FName$
    Tmp.FontSize$ = Size$
    Tmp.BackColor = Val(G_Label_Color)
    Tmp.ForeColor = Val(G_TextLostFore_Color)
    Tmp.FontBold = False
    Tmp.FontItalic = False
End Sub
Sub CloseOpen(rs As Recordset, ByVal rsName$)
'�����w�}�Ҫ�RecordSet
On Error Resume Next

    If Not rs Is Nothing Then rs.Close: Set rs = Nothing
End Sub

Sub InsertFields(ByVal Field$, ByVal Str$, ByVal DType%, Optional Character% = False)
'�զ�s�W��ƪ�SQL���O
Dim A_Str1$, A_Str2$, A_Tmp$, A_Str3$
'S021114036 �]�ǲ�ñ�֮ɡA�ݲզ귥�����r��A�G�Ni%�ܼƩ��̤j(1021115 by Lidia)
Dim I As Currency

    A_Tmp$ = Chr(0) & Chr(128)
    I = InStr(1, G_Str, A_Tmp$)
    If I <> 0 Then
       A_Str1$ = Left(G_Str, I - 1)
       A_Str2$ = Right(G_Str, Len(G_Str) - (I + 1))
    End If
    If Trim(A_Str1$) <> "" Then
       A_Str1$ = A_Str1$ & "," & Field$
    Else
       A_Str1$ = Field$
    End If
    'Str$ = Trim(Str$)
    
    Select Case DType%
      Case G_Data_Numeric
           If Val(Str$) = 0 Then
              A_Str2$ = A_Str2$ & "0,"
           Else
              A_Str2$ = A_Str2$ & Str$ & ","
           End If
      Case G_Data_String
           If Str$ = "" Then
              A_Str2$ = A_Str2$ & "' ',"
           Else
                If Character = True Then
                    '�ѨM�u|�v�g�J��Ʈw���D
                    A_Str2$ = A_Str2$ & CvrString2Character(Str$) & ","
                Else
                    For I = 1 To Len(Str$)
                        If Mid$(Str$, I, 1) = "'" Then
                           A_Str3$ = A_Str3$ & "''"
                        Else
                           A_Str3$ = A_Str3$ & Mid$(Str$, I, 1)
                        End If
                    Next I
                    A_Str2$ = A_Str2$ & "'" & A_Str3$ & "',"
                End If
           End If
    End Select
    G_Str = A_Str1$ & A_Tmp$ & A_Str2$
End Sub
Sub SQLInsert(DB As Database, ByVal Table$)
'����SQL�s�W���O,�f�tInsertFields�{�Ǩϥ�
Dim A_Tmp$, A_Str1$, A_Str2$, A_Sql$
'S021114036 �]�ǲ�ñ�֮ɡA�ݲզ귥�����r��A�G�Ni%�ܼƩ��̤j(1021115 by Lidia)
Dim I As Currency

    A_Tmp$ = Chr(0) & Chr(128)
    I = InStr(1, G_Str, A_Tmp$)
    If I <> 0 Then
       A_Str1$ = Left(G_Str, I - 1)
       A_Str2$ = Right(G_Str, Len(G_Str) - (I + 1))
    End If
    A_Str1$ = "(" & A_Str1$ & ")"
    If Right(A_Str2$, 1) = "," Then
       A_Str2$ = Left(A_Str2$, Len(A_Str2$) - 1)
    End If
    A_Sql$ = "Insert into " & Table$ & Space(1) & A_Str1
    A_Sql$ = A_Sql$ & " values " & "(" & A_Str2$ & ")"
    ExecuteProcess DB, A_Sql$
    G_Str = ""
End Sub

Sub SQLUpdate(DB As Database, ByVal Table$)
'����SQL �ץ����O,�f�tUpdateString�{�Ǩϥ�
Dim A_Str1$, A_Str2$, A_Sql$
    
    StrCut G_Str, " where ", A_Str1$, A_Str2$
    '
    If Right$(A_Str1$, 1) = "," Then
       A_Str1$ = Left$(A_Str1$, Len(A_Str1$) - 1)
    End If
    If A_Str2$ <> "" Then
       A_Str2$ = " where " & A_Str2$
    End If
    '
    A_Sql$ = "Update " & Table$
    A_Sql$ = A_Sql$ & " SET " & A_Str1$ & A_Str2$
    ExecuteProcess DB, A_Sql$
    G_Str = ""
End Sub

Sub SQLUpdate1(DB As Database, ByVal Table$, ErrCode)
'����SQL �ץ����O,�f�tUpdateString�{�Ǩϥ�
Dim A_Str1$, A_Str2$, A_Sql$
    
    StrCut G_Str, " where ", A_Str1$, A_Str2$
    '
    If Right$(A_Str1$, 1) = "," Then
       A_Str1$ = Left$(A_Str1$, Len(A_Str1$) - 1)
    End If
    If A_Str2$ <> "" Then
       A_Str2$ = " where " & A_Str2$
    End If
    '
    A_Sql$ = "Update " & Table$
    A_Sql$ = A_Sql$ & " SET " & A_Str1$ & A_Str2$
    ExecuteProcessReturnErr DB, A_Sql$, ErrCode
    G_Str = ""
End Sub

Sub ExecuteProcess(DB As Database, ByVal SQL$)
'�����Ʈw���s�W,�ק�,�R���ʧ@,�������~�T��
On Local Error GoTo ExecuteProcess_Error

    G_Str = SQL$
    If Trim$(DB.Connect) = "" Then           'Access DataBase
       DB.Execute SQL$, dbSQLPassThrough
    Else
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;", "ORACLE;"
              DB.Execute SQL$, dbSQLPassThrough
         Case "DB2;"
              DB.Execute SQL$
       End Select
    End If
    Exit Sub
    
ExecuteProcess_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB True: End
End Sub

Function GetLikeStr(DB As Database, ByVal Options%) As String
'���o���Ʈw���ϥ�Like���O���S��r��(����t���s�өΧ�h�r�����r��)

    If Trim(DB.Connect) = "" Then   'Access Database
       GetLikeStr = "*"
    Else                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;", "ORACLE;"
              If Options% Then GetLikeStr = "%"
              If Not Options% Then GetLikeStr = "*"
         Case "DB2;"
              GetLikeStr = "%"
       End Select
    End If
End Function


Function GetSingleStr(DB As Database, ByVal Options%) As String
'���o���Ʈw���ϥ�Like���O���S��r��(�����@�r��)

    If Trim(DB.Connect) = "" Then   'Access Database
       GetSingleStr = "?"
    Else                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;", "ORACLE;"
              If Options% Then GetSingleStr = "_"
              If Not Options% Then GetSingleStr = "?"
         Case "DB2;"
              GetSingleStr = "_"
       End Select
    End If
End Function

Function GetDateStr(ByVal Field As Date, DB As Database) As String
'�N����ഫ����Ʈw�i�B�z����

    If Trim(DB.Connect) = "" Then   'Access Database
       GetDateStr = "#" + DateToString(Field) + "#"
    Else                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              GetDateStr = "'" + DateToString(Field) + "'"
              
'         Case "DB2;"
'              GetDateStr = "'" + DateToString(Field) + "'"
       End Select
    End If
End Function


Sub SetListBar_HScroll(Frm As Form, Lst As ListBox)
'�]�wListBox���������b
Const LB_SETHORIZONTALEXTENT = &H194
Dim A_FontName$, A_FontSize$, A_FontBold%
Dim I, A_MaxWidth

    If Lst.ListCount = 0 Then
       SendMessage Lst.hwnd, LB_SETHORIZONTALEXTENT, 0, 0
       Exit Sub
    End If
    '
    A_FontName$ = Frm.FontName
    A_FontSize$ = Frm.FontSize
    A_FontBold% = Frm.FontBold
    Frm.FontName = Lst.FontName
    Frm.FontSize = Lst.FontSize
    Frm.FontBold = Lst.FontBold
    '
    For I = 0 To Lst.ListCount - 1
        If Frm.TextWidth(Lst.List(I)) > A_MaxWidth Then
           A_MaxWidth = Frm.TextWidth(Lst.List(I))
        End If
    Next I
    If A_MaxWidth + 240 > Lst.Width Then
       SendMessage Lst.hwnd, LB_SETHORIZONTALEXTENT, A_MaxWidth, 0
    Else
       SendMessage Lst.hwnd, LB_SETHORIZONTALEXTENT, 0, 0
    End If
    '
    Frm.FontName = A_FontName$
    Frm.FontSize = A_FontSize$
    Frm.FontBold = A_FontBold%
End Sub

Function SubStrFunction(ByVal ConnectMethod$, ByVal Str$, ByVal Start%, ByVal length%, Optional ByPass% = True) As String
'���oSQL���O��,�^���r��Y�X�Ӧ줸��ƪ��y�k
    
    If Start% = 0 Then Start% = 1
    If length% = 0 Then length% = 1
    '
    Select Case UCase$(Trim$(Mid(ConnectMethod$, InStr(1, ConnectMethod$, "DBTYPE=", 1) + 7)))
      Case ""     'Access DataBase
           SubStrFunction = "Mid(" & Str$
      Case "SQL;"
           If Not ByPass% Then
              SubStrFunction = "Mid(" & Str$
           Else
              SubStrFunction = "SubString(" & Str$
           End If
      Case "DB2;", "ORACLE;"
           SubStrFunction = "SubStr(" & Str$
    End Select
    '
    SubStrFunction = SubStrFunction & "," & CStr(Start%)
    SubStrFunction = SubStrFunction & "," & CStr(length%)
    SubStrFunction = SubStrFunction & ")"
    
End Function

Sub UpdateString(ByVal FName$, ByVal Str$, ByVal DType%)
'�զ�ק��ƪ�SQL���O
Dim A_Str1$
'S021114036 �]�ǲ�ñ�֮ɡA�ݲզ귥�����r��A�G�Ni%�ܼƩ��̤j(1021115 by Lidia)
Dim I As Currency

    'Str$ = Trim(Str$)
    Select Case DType%
      Case G_Data_Numeric
           If Trim(Str$) = "" Then
              G_Str = G_Str & FName$ & "=" & "0,"
           Else
              G_Str = G_Str & FName$ & "=" & Str$ & ","
           End If
      Case G_Data_Date
           G_Str = G_Str & FName$ & "=" & Str$ & ","
           
      Case G_Data_String
           If Str$ = "" Then
              G_Str = G_Str & FName$ & "=' ',"
           Else
              For I = 1 To Len(Str$)
                  If Mid$(Str$, I, 1) = "'" Then
                     A_Str1$ = A_Str1$ & "''"
                  Else
                     A_Str1$ = A_Str1$ & Mid$(Str$, I, 1)
                  End If
              Next I
              G_Str = G_Str & FName$ & "='" & A_Str1$ & "',"
           End If
    End Select
End Sub

Sub InsertString(ByVal Str$, ByVal DType%)
'�զ�s�W��ƪ�SQL���O,�Ҧ���쥲���Ӷ��ǫ��w�ȷs�W
Dim A_Str1$
'S021114036 �]�ǲ�ñ�֮ɡA�ݲզ귥�����r��A�G�Ni%�ܼƩ��̤j(1021115 by Lidia)
Dim I As Currency

    'Str$ = Trim(Str$)
    Select Case DType%
      Case G_Data_Numeric
           If Val(Str$) = 0 Then
              G_Str = G_Str & "0,"
           Else
              G_Str = G_Str & Str$ & ","
           End If
      Case G_Data_Date
           G_Str = G_Str & Str$ & ","
           
      Case G_Data_String
           If Str$ = "" Then
              G_Str = G_Str & "' ',"
           Else
              For I = 1 To Len(Str$)
                  If Mid$(Str$, I, 1) = "'" Then
                     A_Str1$ = A_Str1$ & "''"
                  Else
                     A_Str1$ = A_Str1$ & Mid$(Str$, I, 1)
                  End If
              Next I
              G_Str = G_Str & "'" & A_Str1$ & "',"
           End If
    End Select
End Sub

Function CheckDataRange(Sts As StatusBar, ByVal Var1$, ByVal Var2$) As Boolean
'�ˮ֨�Ӥ�r��ƽd��O�_���T

    CheckDataRange = True
    Var1$ = UCase$(Trim$(Var1$))
    Var2$ = UCase$(Trim$(Var2$))
    If Var2$ = "" Then Exit Function
    '
    If Var1$ > Var2$ Then
       Sts.Panels(1) = G_Range_Error
       CheckDataRange = False
    End If
End Function


Sub GetSystemDefault()
'���o�t�ιw�]�Ȧs���Global�ܼƨѵ{���ϥ�
Dim A_Sql$

    Screen.MousePointer = HOURGLASS
   'Pick Label Background Color
    'G_Label_Color = GetColor(GetSIniStr("Customer", "Label"))
    G_Label_Color = Trim(GetSysColor(15))  'COLOR_BTNFACE = 15
   
   'Pick Form Background Color
    G_Form_Color = GetColor(GetSIniStr("Customer", "FormBackColor"))

   'Pick TabsIndex Background Color
    'G_TabBack_Color = GetColor(GetSIniStr("Customer", "TabBackColor"))
    G_TabBack_Color = G_Label_Color

   'Pick TabsIndex Fore Color
    G_TabFore_Color = GetColor(GetSIniStr("Customer", "TabForeColor"))

   'Pick Title Background Color
    G_Title_Color = GetColor(GetSIniStr("Customer", "Title"))

   'Pick TextGotBackColor Background Color
    G_TextGotBack_Color = GetColor(GetSIniStr("Customer", "TextGotBackColor"))

   'Pick TextLostBackColor Background Color
    G_TextLostBack_Color = GetColor(GetSIniStr("Customer", "TextLostBackColor"))

   'Pick TextGotForeColor Background Color
    G_TextGotFore_Color = GetColor(GetSIniStr("Customer", "TextGotForeColor"))

   'Pick TextLostForeColor Background Color
    G_TextLostFore_Color = GetColor(GetSIniStr("Customer", "TextLostForeColor"))

   'Pick Help Fields BackGround Color
    G_TextHelpBack_Color = GetColor(GetSIniStr("Customer", "TextHelpBackColor"))

   'Pick MessageLine Background Color
    G_Msgline_Color = GetColor(GetSIniStr("Customer", "Msgline"))

   'Pick Today Background Color
    G_Today_Color = GetColor(GetSIniStr("Customer", "Today"))

   'Pick FontName
    G_Font_Name = Trim(GetSIniStr("Customer", "FontName"))

   'Pick FontSize
    G_Font_Size = Trim(GetSIniStr("Customer", "Fontsize"))

   'Pick Fixed FontName
    G_FixFont_Name = Trim(GetSIniStr("Customer", "FixWidthFont"))
    If G_FixFont_Name = "" Then G_FixFont_Name = "Courier"

   'Pick Fixed FontSize
    G_FixFont_Size = Trim(GetSIniStr("Customer", "FixWidthFontSize"))
    If G_FixFont_Size = "" Then G_FixFont_Size = "10"
    
   'Pick Report Print Date
    G_Print_Date = GetSIniStr("PanelDescpt", "print_date")
    
   'Pick Report Print Time
    G_Print_Time = GetSIniStr("PanelDescpt", "print_time")
   
   'Pick Report Print PageNo
    G_Print_Page = GetSIniStr("PanelDescpt", "pageno")
   
   'Pick Report Print Next Page
    G_Print_NextPage = GetSIniStr("Customer", "NextPage")

   'Pick Command Key Value
    G_CmdHelp = Trim(GetSIniStr("CmdDescpt", "cmd_help"))
    G_CmdSort = Trim(GetSIniStr("CmdDescpt", "cmd_sort"))
    G_CmdQuery = Trim(GetSIniStr("CmdDescpt", "cmd_query"))
    G_CmdDel = Trim(GetSIniStr("CmdDescpt", "cmd_delete"))
    G_CmdAdd = Trim(GetSIniStr("CmdDescpt", "cmd_add"))
    G_CmdUpdate = Trim(GetSIniStr("CmdDescpt", "cmd_update"))
    G_CmdCopy = Trim$(GetSIniStr("CmdDescpt", "cmd_copy"))
    G_CmdPrint = Trim(GetSIniStr("CmdDescpt", "cmd_print"))
    G_CmdPrevious = Trim(GetSIniStr("CmdDescpt", "cmd_previous"))
    G_CmdNext = Trim(GetSIniStr("CmdDescpt", "cmd_next"))
    G_CmdPrvPage = Trim(GetSIniStr("CmdDescpt", "cmd_prvpage"))
    G_CmdNxtPage = Trim(GetSIniStr("CmdDescpt", "cmd_nxtpage"))
    G_CmdTable = Trim(GetSIniStr("CmdDescpt", "cmd_table"))
    G_CmdSet = Trim(GetSIniStr("CmdDescpt", "cmd_set"))
    G_CmdRecordSet = Trim$(GetSIniStr("CmdDescpt", "cmd_recordset"))
    G_CmdOk = Trim(GetSIniStr("CmdDescpt", "cmd_ok"))
    G_CmdSearch = Trim(GetSIniStr("CmdDescpt", "cmd_search"))
    G_CmdExit = Trim(GetSIniStr("CmdDescpt", "cmd_exit"))
    G_CmdPause = Trim(GetSIniStr("CmdDescpt", "cmd_pause"))
    G_CmdInsert = Trim$(GetSIniStr("CmdDescpt", "cmd_insert"))
    G_CmdHistory = Trim$(GetSIniStr("CmdDescpt", "cmd_history"))
    
   'Pick Common message
    G_AP_ADD = Trim(GetSIniStr("PgmMsg", "g_ap_add"))
    G_AP_DELETE = Trim(GetSIniStr("PgmMsg", "g_ap_delete"))
    G_AP_NORMAL = Trim(GetSIniStr("PgmMsg", "g_ap_normal"))
    G_AP_NODATA = Trim(GetSIniStr("PgmMsg", "g_ap_nodata"))
    G_AP_NOPRVS = Trim(GetSIniStr("PgmMsg", "g_ap_noprvs"))
    G_AP_NONEXT = Trim(GetSIniStr("PgmMsg", "g_ap_nonext"))
    G_AP_PRINT = Trim(GetSIniStr("PgmMsg", "g_ap_print"))
    G_AP_QUERY = Trim(GetSIniStr("PgmMsg", "g_ap_query"))
    G_AP_SEARCH = Trim(GetSIniStr("PgmMsg", "g_ap_search"))
    G_AP_UPDATE = Trim(GetSIniStr("PgmMsg", "g_ap_update"))
    G_AP_COPY = Trim$(GetSIniStr("PgmMsg", "g_ap_copy"))
    G_Add_Check = Trim(GetSIniStr("PgmMsg", "g_add_check"))
    G_Add_Ok = Trim(GetSIniStr("PgmMsg", "g_add_ok"))
    G_Delete_Check = Trim(GetSIniStr("PgmMsg", "g_delete_check"))
    G_Delete_Ok = Trim(GetSIniStr("PgmMsg", "g_delete_ok"))
    G_NoMoreData = Trim(GetSIniStr("PgmMsg", "g_nomore_data"))
    G_Save_Check = Trim(GetSIniStr("PgmMsg", "g_save_check"))
    G_OverDate = Trim(GetSIniStr("PgmMsg", "g_overdate"))
    G_RecordExist = Trim(GetSIniStr("PgmMsg", "g_recordexist"))
    G_NoReference = Trim(GetSIniStr("PgmMsg", "g_noreference"))
    G_NoQueryData = Trim(GetSIniStr("PgmMsg", "g_noquerydata"))
    G_Printing = Trim$(GetSIniStr("PgmMsg", "g_printing"))
    G_DataErr = Trim$(GetSIniStr("PgmMsg", "g_data_error"))
    G_FieldErr = Trim$(GetSIniStr("PgmMsg", "g_field_error"))
    G_Process = Trim$(GetSIniStr("PgmMsg", "g_data_process"))
    G_MustInput = Trim$(GetSIniStr("PgmMsg", "mustinput"))
    G_DateError = Trim$(GetSIniStr("PgmMsg", "g_date_error"))
    G_NumericErr = GetSIniStr("PgmMsg", "numeric_input_error")
    G_Range_Error = Trim$(GetSIniStr("PgmMsg", "g_range_error"))
    G_Update_Ok = Trim$(GetSIniStr("PgmMsg", "g_update_ok"))
    G_Query_Ok = Trim$(GetSIniStr("PgmMsg", "g_query_ok"))
    G_PrintOk = Trim$(GetSIniStr("PgmMsg", "printok"))
    
    'S010605056 �Τ@�s���H��L�O�����u��
    G_A1609uninumber$ = GetCaption("mcfgd", "AccountUniCode", "�״ڲΤ@�s��")
    
    '�O�_��UserID �����q�O���v
    A_Sql$ = "Select TOPICVALUE From SINI Where"
    A_Sql$ = A_Sql$ & " SECTION='Customer'"
    A_Sql$ = A_Sql$ & " AND TOPIC='check_company'"
    A_Sql$ = A_Sql$ & " ORDER BY SECTION,TOPIC"
    CreateDynasetODBC DB_ARTHGUI, DY_SINI, A_Sql$, "DY_SINI", True
    If Not (DY_SINI.BOF And DY_SINI.EOF) Then
        G_CheckCompany = Trim(DY_SINI.Fields("TOPICVALUE") & "")
    Else
        G_CheckCompany = ""
    End If
    
    'S020911050 ��Ƥw�QPCName\UserID[UserName]�ϥΤ�,�L�k�Q��w,�е��ݩάO�q���ӨϥΪ̰h�X!
    G_DataLockErr = GetCaption("PgmMsg", "DataLockErr", "��Ƥw�Q{0}\{1}{2}�ϥΤ�,�L�k�Q��w,�е��ݩάO�q���ӨϥΪ̰h�X!")
    '
    Screen.MousePointer = Default
End Sub


Function CheckNumericRange(Sts As StatusBar, ByVal Var1$, ByVal Var2$) As Boolean
'�ˮ֨�Ӽƭȸ�ƽd��O�_���T

    CheckNumericRange = True
    If Var2$ = "" Then Exit Function
    '
    If Val(Var1$) > Val(Var2$) Then
       Sts.Panels(1) = G_Range_Error
       CheckNumericRange = False
    End If
End Function

Function CheckDateRange(Sts As StatusBar, ByVal Var1$, ByVal Var2$) As Boolean
'�ˮ֨�Ӥ����ƽd��O�_���T

    CheckDateRange = CheckNumericRange(Sts, Var1$, Var2$)
End Function


Sub ProgressBar_Property(Prb As ProgressBar)
'�]�wProgressBar���ݩ�

    Prb.Align = vbAlignNone
    Prb.Appearance = cc3D
    Prb.BorderStyle = ccNone
    Prb.MIN = 0
    Prb.Value = 0
    Prb.Visible = False
End Sub

Sub SpreadGotFocus(ByVal Col#, ByVal Row#, Optional ByVal BColor$, Optional ByVal FColor$, Optional ByVal IgnoreBColor$, Optional ByVal IgnoreNegativeNumber As Boolean = False)
'�B�zSpread Gotfocus���C���ܤ�,�i�z�L�ѼƤ����B�z�C���ܤ�
On Local Error Resume Next
Dim A_OrgBColor$

    If Not TypeOf Screen.ActiveForm.ActiveControl Is FPSPREAD.vaSpread Then Exit Sub
    If IgnoreNegativeNumber Then
        If Row# = 0 Or Row# < -1 Then Exit Sub
        If Col# = 0 Or Col# < -1 Then Exit Sub
    Else
        If Row# <= 0 Then Exit Sub
    End If
    
    With Screen.ActiveForm.ActiveControl
         .Row = Row#
         .Col = Col#
         A_OrgBColor$ = ConnectSemiColon(CStr(.BackColor))
         If InStr(1, IgnoreBColor$, A_OrgBColor$, vbTextCompare) = 0 Then
            If BColor$ = "" Then BColor$ = G_TextGotBack_Color
            If FColor$ = "" Then FColor$ = G_TextGotFore_Color
            .Row = Row#
            .Col = Col#
            .BackColor = BColor$
            .ForeColor = FColor$
         End If
    End With
End Sub
Sub SpreadLostFocus2(Spd As vaSpread, ByVal Col#, ByVal Row#, Optional ByVal BColor$, Optional ByVal FColor$, Optional ByVal IgnoreBColor$, Optional ByVal IgnoreNegativeNumber As Boolean = False)
'�B�zLostfocus�x�s�檺�C���ܤ�,�i�z�L�ѼƤ����B�z�C���ܤ�
On Local Error Resume Next
Dim A_OrgBColor$
    
    If IgnoreNegativeNumber Then
        If Row# = 0 Or Row# < -1 Then Exit Sub
        If Col# = 0 Or Col# < -1 Then Exit Sub
    Else
        If Row# <= 0 Then Exit Sub
    End If
    
    With Spd
         .Row = Row#
         .Col = Col#
         A_OrgBColor$ = ConnectSemiColon(CStr(.BackColor))
         If InStr(1, IgnoreBColor$, A_OrgBColor$, vbTextCompare) = 0 Then
            If BColor$ = "" Then BColor$ = G_TextLostBack_Color
            If FColor$ = "" Then FColor$ = G_TextLostFore_Color
            .BackColor = BColor$
            .ForeColor = FColor$
         End If
    End With
End Sub

Sub SpreadLostFocus(ByVal Col#, ByVal Row#, Optional ByVal A_BackColor$ = "", Optional ByVal A_ForeColor$ = "", Optional ByVal IgnoreNegativeNumber As Boolean = False)
'�B�zLostfocus�x�s�檺�C���ܤ�
On Local Error Resume Next
Dim I%

    If IgnoreNegativeNumber Then
        If Row# = 0 Or Row# < -1 Then Exit Sub
        If Col# = 0 Or Col# < -1 Then Exit Sub
    Else
        If Row# <= 0 Then Exit Sub
    End If
    
    If A_BackColor$ = "" Then A_BackColor$ = G_TextLostBack_Color
    If A_ForeColor$ = "" Then A_ForeColor$ = G_TextLostFore_Color
    
    If Row# <= 0 Then Exit Sub
    For I% = 0 To Screen.ActiveForm.Count - 1
        If TypeOf Screen.ActiveForm.Controls(I%) Is FPSPREAD.vaSpread Then
           Screen.ActiveForm.Controls(I%).Row = Row#
           Screen.ActiveForm.Controls(I%).Col = Col#
           If Trim(Screen.ActiveForm.Controls(I%).BackColor) = G_TextGotBack_Color Or Trim(Screen.ActiveForm.Controls(I%).BackColor) = G_TextHelpBack_Color Then
              Screen.ActiveForm.Controls(I%).BackColor = A_BackColor$
              Screen.ActiveForm.Controls(I%).ForeColor = A_ForeColor$
              Screen.ActiveForm.Controls(I%).Row = Screen.ActiveForm.Controls(I%).ActiveRow
              Screen.ActiveForm.Controls(I%).Col = Screen.ActiveForm.Controls(I%).ActiveCol
              Exit For
           End If
        End If
    Next I%
End Sub

Sub TextGotFocus()
'��TextBox,ListBox,ComboBoxĲ�oGotFocus�ƥ��,�B�z�C�⪺�ܤ�
On Local Error Resume Next

    With Screen.ActiveForm
         If TypeOf .ActiveControl Is TextBox _
         Or TypeOf .ActiveControl Is MsMask.MaskEdBox _
         Or TypeOf .ActiveControl Is ListBox _
         Or TypeOf .ActiveControl Is ComboBox Then
            .ActiveControl.BackColor = G_TextGotBack_Color
            .ActiveControl.ForeColor = G_TextGotFore_Color
            G_FieldText$ = .ActiveControl.text
         End If
         If TypeOf .ActiveControl Is TextBox _
         Or TypeOf .ActiveControl Is MsMask.MaskEdBox Then
            If TypeOf .ActiveControl Is TextBox _
            And Not .ActiveControl.MultiLine Then
                .ActiveControl.text = Trim$(.ActiveControl.text)
                G_FieldText$ = .ActiveControl.text
            End If
            .ActiveControl.SelStart = 0
            .ActiveControl.SelLength = Len(.ActiveControl.text)
         End If
    End With
End Sub


Sub TextHelpGotFocus()
'���U���Ĳ�oGotFocus�ƥ��,�B�z�C�⪺�ܤ�
On Local Error Resume Next

    With Screen.ActiveForm
         If Not TypeOf .ActiveControl Is TextBox Then Exit Sub
         .ActiveControl.BackColor = G_TextHelpBack_Color
         .ActiveControl.ForeColor = G_TextGotFore_Color
         If Not .ActiveControl.MultiLine Then
            .ActiveControl.text = Trim$(.ActiveControl.text)
         End If
         G_FieldText$ = .ActiveControl.text
         .ActiveControl.SelStart = 0
         .ActiveControl.SelLength = Len(.ActiveControl.text)
    End With
End Sub

Sub TextLostFocus()
'��TextBox,ListBox,ComboBoxĲ�oLostFocus�ƥ��,�B�z�C�⪺�ܤ�
On Local Error Resume Next
Dim I%, j%, A_Pos%

    With Screen.ActiveForm
         For I% = 0 To .Count - 1
             If TypeOf .Controls(I%) Is TextBox _
             Or TypeOf .Controls(I%) Is MsMask.MaskEdBox Then
                If Trim(.Controls(I%).BackColor) = G_TextGotBack_Color _
                Or Trim(.Controls(I%).BackColor) = G_TextHelpBack_Color Then
                   .Controls(I%).BackColor = G_TextLostBack_Color
                   .Controls(I%).ForeColor = G_TextLostFore_Color
                   If TypeOf .Controls(I%) Is TextBox And Not .Controls(I%).MultiLine Then
                      .Controls(I%) = Trim$(.Controls(I%))
                      .Controls(I%) = Replace$(.Controls(I%), Chr$(13) & Chr$(10), "", 1, , vbTextCompare)
                      .Controls(I%) = Replace$(.Controls(I%), Chr$(13), "", 1, , vbTextCompare)
                      .Controls(I%) = Replace$(.Controls(I%), Chr$(9), "", 1, , vbTextCompare)
                   End If
                   For j% = .Controls(I%).MaxLength To 1 Step -1
                       If lstrlen(Mid$(.Controls(I%).text, 1, j%)) <= .Controls(I%).MaxLength Then
                          .Controls(I%) = Mid$(.Controls(I%).text, 1, j%)
                          Exit For
                       End If
                   Next j%
                   If Not G_DataChange% Then G_DataChange% = (G_FieldText$ <> .Controls(I%).text)
                   Exit For
                End If
             ElseIf TypeOf .Controls(I%) Is ComboBox Then
                If Trim(.Controls(I%).BackColor) = G_TextGotBack_Color Then
                   .Controls(I%).BackColor = G_TextLostBack_Color
                   .Controls(I%).ForeColor = G_TextLostFore_Color
                   If Not G_DataChange% Then G_DataChange% = (G_FieldText$ <> .Controls(I%).text)
                   Exit For
                End If
             ElseIf TypeOf .Controls(I%) Is ListBox Then
                If Trim(.Controls(I%).BackColor) = G_TextGotBack_Color Then
                   .Controls(I%).BackColor = G_TextLostBack_Color
                   .Controls(I%).ForeColor = G_TextLostFore_Color
                   .Controls(I%).Visible = False
                   Exit For
                End If
             End If
         Next I%
    End With
End Sub


Sub WriteErrorReport(ByVal MSG$, ByVal SqlStr$)
'�s����ƿ��~��,�N���~�T���g�J��r�ɤ�
Dim A_ErrPath$, A_STR$

    A_ErrPath$ = G_Report_Path & DateOut(GetCurrentDate()) & ".ERR"
    If Trim(Dir$(A_ErrPath$)) = "" Then
       Open A_ErrPath$ For Output As #99
    Else
       Open A_ErrPath$ For Append As #99
    End If
    A_STR$ = Format(Now, "HH:NN:SS") & Chr$(KEY_TAB)
    A_STR$ = A_STR$ & GetWorkStation() & Chr$(KEY_TAB)
    A_STR$ = A_STR$ & App.EXEName & Chr$(KEY_TAB)
    A_STR$ = A_STR$ & MSG$ & Chr$(KEY_TAB)
    A_STR$ = A_STR$ & SqlStr$
    Print #99, A_STR$
    Close #99
End Sub

Function GetWorkStation() As String
'���o�����W�٪��e10�X
Dim A_ComputerName$
 
    GetWorkStation = " "
    '
    A_ComputerName$ = Space$(200)
    If GetComputerName(A_ComputerName$, 200) Then
       A_ComputerName$ = StripTerminator(Trim$(A_ComputerName$))
    End If
    GetWorkStation = GetLenStr(A_ComputerName$, 1, 10)
End Function

Sub WriteJournalLog(DB As Database, ByVal State%, ByVal PgmId$, ByVal Memo$)
'�g�J�{���ϥΪ��p��A09
    'S020527055
'    G_Str = "INSERT INTO A09 VALUES ("
    G_Str = ""
    InsertFields "A0901", GetCurrentDate(), G_Data_String
    InsertFields "A0902", GetCurrentTime(), G_Data_String
    InsertFields "A0903", GetWorkStation(), G_Data_String
    InsertFields "A0904", GetUserId(), G_Data_String
    InsertFields "A0905", G_UserGroup, G_Data_String
    InsertFields "A0906", PgmId$, G_Data_String
    InsertFields "A0907", State%, G_Data_String
    InsertFields "A0908", " ", G_Data_String
    InsertFields "A0909", G_UserName, G_Data_String
    InsertFields "A0910", " ", G_Data_String
    InsertFields "A0911", G_SystemID, G_Data_String
    InsertFields "A0912", GetLenStr(Memo$, 1, 50), G_Data_String
'    G_Str = Left$(G_Str, Len(G_Str) - 1) & ")"
    SQLInsert DB, "A09"
End Sub


Sub CloseFileDB(Optional ByVal WriteFlag%)
'�����{���}�Ҫ��Ҧ���Ʈw
Dim A_Index%
Dim WK As Workspace
Dim DB As Database
Dim rs As Recordset
    
    If Not G_SecurityPgm Then
        If Not WriteFlag% Then
           WriteJournalLog DB_ARTHGUI, G_Program_End, UCase$(App.EXEName), ""
        End If
    Else
       WriteJournalLog_Security DB_ARTHGUI, G_Program_End, UCase$(App.EXEName), ""
    End If
    
    'Close RecordSets
    For Each WK In GetEngine.Workspaces
        For Each DB In WK.Databases
            For A_Index% = 0 To DB.Recordsets.Count - 1
                Set rs = DB.Recordsets(0)
                rs.Close
                Set rs = Nothing
            Next A_Index%
        Next
    Next
    
    'Close Databases
    For Each WK In GetEngine.Workspaces
        For A_Index% = 0 To WK.Databases.Count - 1
            Set DB = WK.Databases(0)
            DB.Close
            Set DB = Nothing
        Next
    Next
    
    'Close Workspace
    For A_Index% = 0 To GetEngine.Workspaces.Count - 1
        Set WK = Workspaces(0)
        WK.Close
        Set WK = Nothing
    Next
End Sub


Sub FormCenter(Frm As Form)
'�]�w��������ܦ�m
Dim A_Left&, A_Top&, A_Right&, A_Bottom&

    GetScreenPosition A_Left&, A_Top&, A_Right&, A_Bottom&
    If A_Right& > Frm.Width Then
       Frm.Left = A_Left& + (A_Right& - A_Left& - Frm.Width) \ 2
    Else
       Frm.Left = A_Left&
    End If
    If A_Bottom& > Frm.Height Then
       Frm.Top = A_Top& + (A_Bottom& - A_Top& - Frm.Height) \ 4
    Else
       Frm.Top = A_Top&
    End If
 End Sub

Sub GetScreenPosition(Left&, Top&, Right&, Bottom&)
'���o�ù��i�ΰϰ쪺�W�U���k��m
Const SPI_GETWORKAREA = 48
Dim A_Rect As RECT
Dim A_Hwnd, A_Multiple

    A_Hwnd = GetDesktopWindow()
    retcode = GetWindowRect(A_Hwnd, A_Rect)
    A_Multiple = Screen.Height / A_Rect.Bottom
    '
    retcode = SystemParametersInfo(SPI_GETWORKAREA, 0, A_Rect, 0)
    With A_Rect
         Left& = .Left * A_Multiple
         Top& = .Top * A_Multiple
         Right& = .Right * A_Multiple
         Bottom& = .Bottom * A_Multiple
    End With
End Sub

Sub StatusBar_ProPerty(StsBar As StatusBar)
'�]�wStatusBar���ݩ�

    'General Defined
    StsBar.Height = 375
    StsBar.Align = vbAlignBottom
    StsBar.style = sbrNormal
    StsBar.MousePointer = ccDefault
    StsBar.Enabled = True
    'Font Defined
    StsBar.Font.Name = G_Font_Name
    StsBar.Font.Size = G_Font_Size
    StsBar.Font.Bold = False
    StsBar.Font.Italic = False
    StsBar.Font.Strikethrough = False
    StsBar.Font.Underline = False
    'Panels Defined (Index = 1)
    StsBar.Panels(1).ToolTipText = "�T����"
    StsBar.Panels(1).Alignment = sbrLeft
    StsBar.Panels(1).style = sbrText
    StsBar.Panels(1).Bevel = sbrInset
    StsBar.Panels(1).AutoSize = sbrSpring
    StsBar.Panels(1).Enabled = True
    StsBar.Panels(1).Visible = True
    'Panels Defined (Index = 2)
    StsBar.Panels(2).ToolTipText = "���"
    StsBar.Panels(2).Width = 1200
    StsBar.Panels(2).Alignment = sbrCenter
    StsBar.Panels(2).style = sbrText
    StsBar.Panels(2).Bevel = sbrInset
    StsBar.Panels(2).AutoSize = sbrContents
    StsBar.Panels(2).Enabled = True
    StsBar.Panels(2).Visible = True
    'Picture Defined
    StsBar.MouseIcon = Nothing
End Sub

Sub InsertStatusBarPanel(StsBar As StatusBar, ByVal index As Integer, ByVal text As String, ByVal Width As Integer)
'�bStatus Bar�W�[�J�@��Panels
    If StsBar.Panels.Count >= index Then Exit Sub

    StsBar.Panels.Add index, , text, sbrText
    StsBar.Panels(index).ToolTipText = "Default"
    StsBar.Panels(index).Width = Width
    StsBar.Panels(index).Alignment = sbrCenter
    StsBar.Panels(index).style = sbrText
    StsBar.Panels(index).Bevel = sbrInset
    StsBar.Panels(index).AutoSize = sbrContents
End Sub



Sub TabIndex_Property(Tmp As Control, ByVal Size$, ByVal FName$, ByVal TabRows&, ByVal Position&)
'�]�wIndexTab���ݩ�

    Tmp.Font.Name = FName$
    Tmp.ActivePageFontName = FName$
    Tmp.Font.Size = Size$
    Tmp.ActivePageFontSize = Size$
    Tmp.TabCount = TabRows&          '�`�@ TAB ��
    Tmp.TabsPerRow = TabRows&    '�C�� N �� TAB ��
    Tmp.ActiveTab3DBackColor = G_TabBack_Color
    Tmp.ActivePageForeColor = G_TabFore_Color
    Tmp.ActiveTabBackColor = G_TabBack_Color
    Tmp.BackColor = G_TabBack_Color
    Tmp.TabForeColorDefault = G_TextLostFore_Color
    Tmp.ActiveTabForeColor = G_TabFore_Color
    Tmp.ActiveTabFont.Name = FName$
    Tmp.ActiveTabFont.Size = Size$
    Tmp.ActiveTabFont.Bold = True
    Tmp.BevelColorFace = G_Label_Color
    Tmp.BevelColorHighlight = COLOR_WHITE 'G_Label_Color
    Tmp.TabHeight = 360
    Tmp.Font.Bold = False
    Tmp.TabOrientation = Position&    '�������b�W,�U,��,�k
    'tmp.AlignmentCaption = SS_CAPTION_CENTER_MIDDLE   '��󤤶�
End Sub

Sub TabIndex_Caption_Property(Tmp As Control, ByVal n%, ByVal text$)
'�]�wIndexTab���������D

    Tmp.TabCaption(n%) = text$
End Sub

Sub DoubleRunCheck()
'�P�@��{���������⦸
Dim Temp$

    If App.PrevInstance Then
       Temp$ = App.Title
       App.Title = "KILLED"
       On Error Resume Next
       AppActivate Temp$
       On Error GoTo 0
       End
    End If
End Sub
Sub GetFunctionAction(ByVal ProgramID$, ByVal UserID$)
'���o�ϥΪ̩�{�����ϥ��v��(Ū��,�s�W,�ק�,�R��)
On Local Error Resume Next
Dim A_Sql$, DY_Tmp As Recordset

    G_AUT_READ% = True
    G_AUT_UPDATE% = True
    G_AUT_DELETE% = True
    G_AUT_ADD% = True
    '
    A_Sql$ = "Select A4703,A4704,A4705,A4706 from A47 "
    A_Sql$ = A_Sql$ & " Where A4701='" & Trim(UserID$) & "' "
    A_Sql$ = A_Sql$ & " and A4702='" & Trim(ProgramID$) & "' "
    CreateDynasetODBC DB_ARTHGUI, DY_Tmp, A_Sql$, "DY_TMP", True
    If Not (DY_Tmp.BOF And DY_Tmp.EOF) Then
       If UCase(Trim(DY_Tmp.Fields("A4703") & "")) <> "Y" Then G_AUT_READ = False
       If UCase(Trim(DY_Tmp.Fields("A4704") & "")) <> "Y" Then G_AUT_UPDATE = False
       If UCase(Trim(DY_Tmp.Fields("A4705") & "")) <> "Y" Then G_AUT_DELETE = False
       If UCase(Trim(DY_Tmp.Fields("A4706") & "")) <> "Y" Then G_AUT_ADD = False
    End If
End Sub
Sub IsAppropriateCheck()
'�P�_�{���O�_�Ѩt��Menu�Ұ�
'G_CmdStr1(�Ѽ�1)=ProgramName & "@"
'G_CmdStr2(�Ѽ�2)=Program Use Parameter
'G_CmdStr3(�Ѽ�3)=�ϥθӵ{�����n���ϥΪ̸�T(UID/Name/Group)
Dim A_String$
    
    StrCut Command$, ",", G_CmdStr1, G_CmdStr2
    StrCut G_CmdStr2, ",Inf=", G_CmdStr2, G_CmdStr3
    G_CmdStr1 = Trim(G_CmdStr1)
    G_CmdStr2 = Trim(G_CmdStr2)
    G_CmdStr3 = Trim(G_CmdStr3)
    '
    A_String$ = UCase(App.EXEName) + "@"
    '
    If UCase(G_CmdStr1) <> A_String$ Then
       Beep
       MsgBox "This Program can't execute, it's illegal way !", MB_ICONEXCLAMATION
       End
    End If
    '
    G_CmdStr1 = Left$(G_CmdStr1, Len(G_CmdStr1) - 1)
    If UCase$(Trim$(G_CmdStr1)) = UCase$("MCFGD") And G_CmdStr2 = "2" Then G_CmdStr1 = "MCFGDA"
    If UCase$(Trim$(G_CmdStr1)) = UCase$("MCFGU") And G_CmdStr2 = "2" Then G_CmdStr1 = "MCFGUA"
End Sub
Sub Command_Property(Tmp As Control, ByVal text$, ByVal FName$)
'�]�wCommand Button���ݩ�
On Error Resume Next
    
    Tmp.Caption = text$
    Tmp.FontName = FName$
    Tmp.FontSize = G_Font_Size
    Tmp.BackColor = G_Label_Color
    Tmp.ForeColor = Val(G_TextLostFore_Color)
    Tmp.FontBold = False
    Tmp.FontItalic = False
End Sub


Sub Form_Property(Frm As Form, ByVal Title$, ByVal FName$)
'�]�w�������ݩ�
Dim A_PName$, A_Title$
    
    If G_IsVistaClient Then Title$ = Title$ & "  " & G_VistaClientTitle
    Frm.Caption = Title$
    Frm.FontName = FName$
    Frm.FontSize = G_Font_Size
    Frm.BackColor = Val(G_Form_Color)
    Frm.ForeColor = Val(G_TextLostFore_Color)
    Frm.FontBold = False
    Frm.FontItalic = False
    A_PName$ = G_System_Path + G_PICTURE_NAME
    Frm.Icon = LoadPicture(A_PName$)
End Sub

Function GetCurrentDate() As String
'���o�{�b�������

     GetCurrentDate = Format$(Now, "YYYYMMDD")
End Function


Function DateIn(ByVal DateStr$) As String
'�N����ഫ�����,�s�ɮɨϥ�

    DateStr$ = Replace(Trim(DateStr$), "/", "")
    DateIn = " "
    If Val(DateStr$) = 0 Then Exit Function
    
    Select Case G_DateFlag
      Case 0
           DateIn = Trim(DateStr$)
      Case 1
           DateIn = Trim(Val(DateStr$) + 19110000)
      Case 2
           DateIn = IIf(Len(DateStr$) = 6, G_LeadYear$ & DateStr$, DateStr$)
    End Select
End Function
Function DateOut(ByVal DateStr$) As String
'�N����ഫ���t�γ]�w����ܫ��A(���Φ��),Output�ɨϥ�

    DateStr$ = Trim(DateStr$)
    DateOut = " "
    If Val(DateStr$) = 0 Then Exit Function
    
    Select Case G_DateFlag
      Case 0
           DateOut = Format$(DateStr$, "########")
      Case 1
           DateOut = Format$(Val(DateStr$) - 19110000, "#000000")
      Case 2
           DateOut = Format$(IIf(Left$(DateStr$, 2) = G_LeadYear$, _
                     Mid$(DateStr$, 3), DateStr$), "##000000")
    End Select
End Function

Public Function StringToDate(ByVal sDate$) As Date
'�N�����r���A�ഫ��������A
Dim nLen&, nYear%, nMonth%, nDay%, sTempDate$
Dim dtTemp As Date

    sDate$ = Trim(sDate$)
    nLen& = Len(sDate$)
    Select Case G_DateFlag
      Case 2
           If nLen& < 8 Then sDate$ = G_LeadYear$ & sDate$
           nYear% = Val(Mid$(sDate$, 1, 4))
           nMonth% = Val(Mid$(sDate$, 5, 2))
           nDay% = Val(Mid$(sDate$, 7, 2))
      Case 1
           Select Case nLen&
             Case 6          'yymmdd
                  nYear% = Val(Mid$(sDate$, 1, 2)) + 1911
                  nMonth% = Val(Mid$(sDate$, 3, 2))
                  nDay% = Val(Mid$(sDate$, 5, 2))
             Case 7          'yyymmdd
                  nYear% = Val(Mid$(sDate$, 1, 3)) + 1911
                  nMonth% = Val(Mid$(sDate$, 4, 2))
                  nDay% = Val(Mid$(sDate$, 6, 2))
             Case 8          'yy/mm/dd
                  nYear% = Val(Mid$(sDate$, 1, 2)) + 1911
                  nMonth% = Val(Mid$(sDate$, 4, 2))
                  nDay% = Val(Mid$(sDate$, 7, 2))
             Case 9          'yyy/mm/dd
                  nYear% = Val(Mid$(sDate$, 1, 3)) + 1911
                  nMonth% = Val(Mid$(sDate$, 5, 2))
                  nDay% = Val(Mid$(sDate$, 8, 2))
           End Select
      Case 0
           Select Case nLen&
             Case 6          'yymmdd
                  nYear% = Val(Mid$(sDate$, 1, 2))
                  Select Case nYear%
                    Case 0 To 29
                         nYear% = nYear% + 2000
                    Case 30 To 99
                         nYear% = nYear% + 1900
                  End Select
                  nMonth% = Val(Mid$(sDate$, 3, 2))
                  nDay% = Val(Mid$(sDate$, 5, 2))
             Case 8          'yy/mm/dd or yyyymmdd
                  nYear% = Val(Mid$(sDate$, 1, 4))
                  If nYear% < 100 Then
                     Select Case nYear%
                       Case 0 To 29
                            nYear% = nYear% + 2000
                       Case 30 To 99
                            nYear% = nYear% + 1900
                     End Select
                     nMonth% = Val(Mid$(sDate$, 4, 2))
                     nDay% = Val(Mid$(sDate$, 7, 2))
                  Else
                     nMonth% = Val(Mid$(sDate$, 5, 2))
                     nDay% = Val(Mid$(sDate$, 7, 2))
                  End If
             Case 10         'yyyy/mm/dd
                  nYear% = Val(Mid$(sDate$, 1, 4))
                  nMonth% = Val(Mid$(sDate$, 6, 2))
                  nDay% = Val(Mid$(sDate$, 9, 2))
           End Select
    End Select
    Select Case G_DateFlag
      Case 2, 1, 0
           sTempDate$ = Format$(nYear%, "0000") & "/" & Format$(nMonth%, "00") & "/" & Format$(nDay%, "00")
      Case Else
           sTempDate$ = sDate$
    End Select
    If IsDate(sTempDate$) Then
       dtTemp = CDate(sTempDate$)
    Else
       dtTemp = G_dtDateError
    End If
    StringToDate = dtTemp
End Function

Public Function StrStr(ByVal tStr$) As String
'�N�r��e��[�W��޸��Ǧ^
Dim tTemp$, tTempWork$
Dim n&, nStart&, nLen&, nValue&
Dim bNotFind As Boolean

    bNotFind = False
    tTempWork$ = ""
    nStart& = 1
    tTemp$ = tStr$
    nLen& = Len(tTemp$)
    If nLen& = 0 Then
       StrStr = "''"
       Exit Function
    End If
    For n& = 1 To nLen&
        nValue& = Asc(Mid(tTemp$, n&, 1))
        Select Case nValue&
          Case 39, 124
               If nStart& < n& Then
                  tTempWork$ = tTempWork$ & "'" & Mid(tTemp$, nStart&, n& - nStart&) & "' +"
               End If
               tTempWork$ = tTempWork$ & " char(" & Format$(nValue&) + ")"
               nStart& = n& + 1
               If nStart& <= nLen& Then
                  tTempWork$ = tTempWork$ & " + "
               End If
        End Select
    Next
    If nStart& <= nLen& Then
       tTempWork$ = tTempWork$ & "'" & Mid(tTemp$, nStart&, nLen& - nStart& + 1) & "'"
    End If
    StrStr = tTempWork$
End Function

Public Function StrDate(ByVal dt As Date) As String
'�N����e��[�W��޸��Ǧ^
Dim tTemp$

    If dt = G_dtDateError Then
       tTemp$ = "'1899/12/30 12:00:00 AM'"
    Else
       tTemp$ = StrStr(dt)
    End If
    StrDate = tTemp$
End Function


Public Function DateToString(ByVal dtDate As Date) As String
'�N������A��Ʈ榡�Ƭ��~/��/��
Dim nYear%, nMonth%, nDay%, sDate$

    Select Case dtDate
      Case G_dtDateError
           'sDate$ = "-"
           sDate$ = " "
      Case G_dtDateMax, G_dtDateMin
           'sDate$ = "."
           sDate$ = " "
      Case Else
           nYear% = Year(dtDate)
           nMonth% = Month(dtDate)
           nDay% = Day(dtDate)
           Select Case G_DateFlag
             Case 1
                  nYear% = nYear% - 1911
           End Select
           Select Case G_DateFlag
             Case 2
                  sDate$ = Format$(IIf(Left$(CStr(nYear%), 2) = G_LeadYear$, _
                           Right$(CStr(nYear%), 2), CStr(nYear%)), "##00") _
                           & "/" & Format$(nMonth%, "00") & "/" & Format$(nDay%, "00")
             Case 1
                  sDate$ = Format$(nYear%, "00") & "/" & Format$(nMonth%, "00") & "/" & Format$(nDay%, "00")
             Case 0
                  sDate$ = Format$(nYear%, "0000") & "/" & Format$(nMonth%, "00") & "/" & Format$(nDay%, "00")
             Case Else
                  sDate$ = Format$(dtDate, "Short Date")
           End Select
    End Select
    DateToString = sDate$
End Function

Function GetCurrentDay(ByVal FormatFlag%) As String
'���o�榡�ƫ᪺�{�b���
Dim A_CDate$, yy$, MM$, dd$

    A_CDate$ = GetCurrentDate()
    yy$ = Mid$(A_CDate$, 1, 4)
    MM$ = Mid$(A_CDate$, 5, 2)
    dd$ = Mid$(A_CDate$, 7, 2)
    
    Select Case G_DateFlag
      Case 0    'ENGLISH
           Select Case FormatFlag%
             Case 0  'YYYYMMDD
                  GetCurrentDay = A_CDate$
             Case 1  'YYYY/MM/DD
                  GetCurrentDay = Format(A_CDate$, "0000/00/00")
             Case 2  'YYYYMM
                  GetCurrentDay = Mid$(A_CDate$, 1, 6)
             Case 3  'MMDD
                  GetCurrentDay = Mid$(A_CDate$, 5, 4)
             Case 4  'MM/DD
                  GetCurrentDay = Format(Mid$(A_CDate$, 5, 4), "00/00")
           End Select
      
      Case 1, 2   '1:Chinese 2:ENGLISH(Support Log & Short Year Date)
           If G_DateFlag = 1 Then
              yy$ = CStr(Val(yy$) - 1911)
           Else
              If Left$(yy$, 2) = G_LeadYear$ Then yy$ = Right$(yy$, 2)
           End If
           Select Case FormatFlag%
             Case 0  'YYMMDD
                  GetCurrentDay = yy$ & MM$ & dd$
             Case 1  'YY/MM/DD
                  GetCurrentDay = yy$ & "/" & MM$ & "/" & dd$
             Case 2  'YYMM
                  GetCurrentDay = yy$ & MM$
             Case 3  'MMDD
                  GetCurrentDay = MM$ & dd$
             Case 4  'MM/DD
                  GetCurrentDay = MM$ & "/" & dd$
           End Select
    End Select
End Function

Function GetCurrentTime() As String
'���o�{�b���ɶ�

    GetCurrentTime = Format$(Now, "hhmmss") + "00"
End Function


Function GetIniStr(ByVal Section$, ByVal Topic$, ByVal inipath$) As String
'���oINI File�������
Dim A_RetStr$

     A_RetStr$ = Space(1000)
     GetIniStr = ""
     If OSGetPrivateProfileString%(Section$, Topic$, "", A_RetStr$, 1000, inipath$) Then
        A_RetStr$ = StripTerminator(Trim$(A_RetStr$))
        GetIniStr = Trim$(A_RetStr$)
     End If
End Function

Function GetSIniStr(ByVal Section$, ByVal Topic$) As String
'��Local MDB�����o��ܦb�e���W�����J
    GetSIniStr = " "
    If Trim(DB_LOCAL.Connect) <> "" Then
        Dim A_Sql$
        A_Sql$ = "SELECT TOPICVALUE FROM INI"
        A_Sql$ = A_Sql$ & " WHERE SECTION='" & Section$ & "'"
        A_Sql$ = A_Sql$ & " AND TOPIC='" & Topic$ & "'"
        Set DY_INICommon = DB_LOCAL.OpenRecordset(A_Sql$, dbOpenSnapshot, dbSQLPassThrough)
        If Not (DY_INICommon.BOF And DY_INICommon.EOF) Then
            GetSIniStr = Trim(DY_INICommon.Fields("TOPICVALUE") & "")
        End If
        DY_INICommon.Close
    Else
        TB_INI.Seek "=", Section$, Topic$
        If Not TB_INI.NoMatch Then
           GetSIniStr = Trim(TB_INI.Fields("TOPICVALUE") & "")
        End If
    End If
End Function

Sub SaveGUIINIValue(ByVal A_Section$, ByVal A_Topic$, ByVal A_TopicValue$)
    Dim A_Sql$, DY_Tmp As Recordset
    A_Sql$ = "SELECT * FROM INI"
    A_Sql$ = A_Sql$ + " WHERE SECTION='" & Trim(A_Section$) & "'"
    A_Sql$ = A_Sql$ + " AND TOPIC='" & Trim(A_Topic$) & "'"
    CreateDynasetODBC DB_LOCAL, DY_Tmp, A_Sql$, "DY_TMP", True
    If DY_Tmp.BOF And DY_Tmp.EOF Then
        G_Str = "'"
        InsertFields "SECTION", A_Section$, G_Data_String
        InsertFields "TOPIC", A_Topic$, G_Data_String
        InsertFields "TOPICVALUE", A_TopicValue$, G_Data_String
        SQLInsert DB_LOCAL, "INI"
    Else
        A_Sql$ = "UPDATE INI SET TOPICVALUE='" & Trim(A_TopicValue$) & "' "
        A_Sql$ = A_Sql$ + "WHERE SECTION='" & Trim(A_Section$) & "' AND TOPIC='" & Trim(A_Topic$) & "'"
        DB_LOCAL.Execute A_Sql$
    End If
End Sub

Function IsDateValidate(ByVal DateStr$) As Boolean
'�ˮ֤���O�_�X�k
Dim Temp$, leapYear%, DateLen%
Dim I%, yy&, MM%, dd%
    
    IsDateValidate = False
    DateStr$ = Trim(DateStr$)
    If DateStr$ = "" Then IsDateValidate = True: Exit Function
    
    '�����ˮ�
    DateLen% = Len(DateStr$)
    Select Case G_DateFlag
      Case 0
           If DateLen% <> 8 Then Exit Function
      Case 1
           If DateLen% < 6 Or DateLen% > 8 Then Exit Function
      Case 2
           If DateLen% <> 6 And DateLen% <> 8 Then Exit Function
    End Select
    
    '�Ʀr���T���ˮ�
    For I% = 1 To DateLen%
        If InStr(1, "0123456789", Mid(DateStr$, I%, 1), vbTextCompare) = 0 Then
            Exit Function
        End If
    Next I%
    
    '������T���ˮ�
    Temp$ = DateIn(DateStr$)
    If Len(Temp$) > 8 Then Exit Function
    '
    yy& = Val(Left$(Temp$, 4))
    MM% = Val(Mid$(Temp$, 5, 2))
    dd% = Val(Mid$(Temp$, 7, 2))
    '20110308�W�[�~�פ��o�p�󵥩�1911�P�_(Yvonne)-------S
    If yy& <= 1911 Then Exit Function
    '---------------------------------------------------E
    If MM% < 1 Or MM% > 12 Then Exit Function
    If dd% < 1 Or dd% > 31 Then Exit Function
    '�P�_�Ӧ~�O�_���|�~
    leapYear% = False
    If yy& Mod 4 = 0 Then
       If yy& Mod 100 = 0 Then
          If yy& Mod 400 = 0 Then leapYear% = True
       Else
          leapYear% = True
       End If
    End If
    '
    Select Case MM%
      Case 4, 6, 9, 11
           If dd% > 30 Then Exit Function
      Case 2
           If dd% > 29 Then Exit Function
           If leapYear% = False And dd% > 28 Then Exit Function
    End Select
    '
    IsDateValidate = True
End Function

Sub KeyPress(KeyAscii%)
'����Enter�����U�@��Control

    If KeyAscii% = KEY_RETURN Then
       KeyAscii% = 0
       SetActiveControlFocus
    Else
       If KeyAscii% = KEY_BACK Then Exit Sub
       If TypeOf Screen.ActiveForm.ActiveControl Is TextBox Then
          Dim A_CharLen%
          On Error Resume Next
          '�YOS��O�]�w��English(US),�U�@��{���N����Error
          A_CharLen% = lstrlen(Trim$(Chr$(KeyAscii)))
          If Err = 0 Then
             If lstrlen(Screen.ActiveForm.ActiveControl.text) - Screen.ActiveForm.ActiveControl.SelLength + A_CharLen% > Screen.ActiveForm.ActiveControl.MaxLength Then KeyAscii = 0
          End If
          On Error GoTo 0
       End If
    End If
End Sub

Sub Label_Property(Tmp As Control, ByVal text$, ByVal Color$, ByVal Size$, ByVal FName$, Optional FColor$)
'�]�wLabel���ݩ�

    Tmp.BackColor = Val(Color$)
    If Trim(FColor$) = "" Then FColor$ = G_TextLostFore_Color
    Tmp.ForeColor = Val(FColor$)
    Tmp.Caption = text$
    Tmp.FontName = FName$
    Tmp.FontSize = Size$
    Tmp.FontBold = False
    Tmp.FontItalic = False
End Sub

Function NoEdit(ByVal No$, ByVal length%) As String
'�N�ƭȸ�Ʈ榡�Ʀ��T�w���ץB�m�k
Dim A1%, A2%, a$
  
  A1% = Len(No$)
  If A1% <= 0 Then Exit Function
  
  A2% = length% - A1% + 1
  If A2% <= 0 Then A2% = 1
  a$ = String$(length%, " ")
  Mid$(a$, A2%, A1%) = LTrim$(No$)
  NoEdit = a$
End Function
Sub Mask_Property(Tmp As Control, ByVal MaskStr$, ByVal length%)
'�]�wMaskEdBox���ݩ�

    Tmp.AutoTab = False
    Tmp.BackColor = Val(G_TextLostBack_Color)
    Tmp.ForeColor = Val(G_TextLostFore_Color)
    Tmp.Format = MaskStr$
    Tmp.MaxLength = length%
    Tmp.PromptChar = " "
    Tmp.PromptInclude = False
    Tmp.FontName = G_Font_Name
    Tmp.FontSize = G_Font_Size
    Tmp.FontBold = False
    Tmp.FontItalic = False
End Sub

Sub Option_Property(ByVal Tmp As Control, ByVal text$, ByVal FSize$, ByVal FName$)
'�]�wOption Button���ݩ�
On Error Resume Next

    Tmp.Caption = text$
    Tmp.FontSize = FSize$
    Tmp.FontName = FName$
    Tmp.BackColor = G_Label_Color
    Tmp.ForeColor = Val(G_TextLostFore_Color)
    Tmp.FontBold = False
    Tmp.FontItalic = False
End Sub

Function SetMessage(ByVal StateId%) As String
'�]�w�@�~���A���T������

    Select Case StateId%
           Case G_AP_STATE_NORMAL
                 SetMessage = G_AP_NORMAL
           Case G_AP_STATE_ADD
                SetMessage = G_AP_ADD
           Case G_AP_STATE_UPDATE
                SetMessage = G_AP_UPDATE
           Case G_AP_STATE_DELETE
                SetMessage = G_AP_DELETE
           Case G_AP_STATE_QUERY
                SetMessage = G_AP_QUERY
           Case G_AP_STATE_PRINT
                SetMessage = G_AP_PRINT
           Case G_AP_STATE_TABLE
                SetMessage = G_AP_TABLE
           Case G_AP_STATE_NODATA
                SetMessage = G_AP_NODATA
           Case G_AP_STATE_COPY
                SetMessage = G_AP_COPY
           Case Else
                SetMessage = " "
       End Select
End Function

Sub Spread_Property(Spd As vaSpread, ByVal Rows#, ByVal Cols#, ByVal Color&, ByVal Size$, ByVal FName$)
'�]�wSpread���ݩ�
    
    Spd.ProcessTab = False
    Spd.EditEnterAction = 7
    Spd.MaxRows = Rows#              '�`�C��
    Spd.MaxCols = Cols#              '�`���
    Spd.Row = -1: Spd.Col = -1
    Spd.BackColor = Color&           '�]�w SPREAD ���I���C��
    Spd.FontSize = Size$             '�r�Τj�p
    Spd.FontName = FName$            '�r�κ���
    Spd.FontBold = False
    Spd.ShadowText = Val(G_TextLostFore_Color)
    Spd.GridColor = COLOR_GRAY
    Spd.ShadowDark = G_Label_Color
    Spd.GrayAreaBackColor = G_Label_Color
    Spd.ShadowColor = G_Label_Color
    Spd.BackColorStyle = 1
    Spd.AllowMultiBlocks = True
    Spd.EditModeReplace = True
    Spd.ZOrder 1
End Sub

Sub DelSpreadRows(Spd As vaSpread)
'�R��vaSpread�W���аO�C
Dim I#, A_MaxRows#, A_STR$, A_Col%

    With Spd
         If .IsBlockSelected Or .MultiSelCount Then
            .Action = SS_ACTION_GET_MULTI_SELECTION
            .BlockMode = True
            For I# = 0 To .MultiSelCount - 1
                .MultiSelIndex = I#
                .Action = SS_ACTION_CLEAR_TEXT
            Next I#
            .BlockMode = False
            .Action = SS_ACTION_DESELECT_BLOCK
            A_MaxRows# = .MaxRows
            For I# = 1 To A_MaxRows#
                If I# > .MaxRows Then Exit For
                A_STR$ = ""
                .Row = I#
                For A_Col% = 1 To .MaxCols
                    .Col = A_Col%
                    A_STR$ = A_STR$ & .text
                Next A_Col%
                If A_STR$ = "" Then
                   .Col = -1
                   .Action = SS_ACTION_DELETE_ROW
                   A_MaxRows# = A_MaxRows# - 1
                   .MaxRows = A_MaxRows#
                   I# = I# - 1
                End If
            Next I#
         Else
            .Col = -1
            .Row = .ActiveRow
            .Action = SS_ACTION_DELETE_ROW
         End If
         .MaxRows = .DataRowCnt + 1
         .SetFocus
    End With
End Sub

Sub StrCut(ByVal Source$, ByVal Separate$, Str1$, str2$)
'�N�r��H���j�r�����ά���r��
Dim Pointer%

    Pointer% = InStr(Source$, Separate$) 'A�r�ꤧ��m
    If Pointer% > 0 Then
        Str1$ = Trim$(Left(Source$, Pointer% - 1))
        str2$ = Trim$(Right(Source$, Len(Source$) - Pointer% - Len(Separate$) + 1))
    Else
        Str1$ = Trim$(Source$)
        str2$ = ""
    End If
End Sub

Sub Text_Property(Tmp As Control, ByVal length%, ByVal FName$)
'�]�wTextBox���ݩ�

    Tmp.BackColor = Val(G_TextLostBack_Color)
    Tmp.ForeColor = Val(G_TextLostFore_Color)
    Tmp.MaxLength = length%
    Tmp.FontName = FName$
    Tmp.FontSize = G_Font_Size
    Tmp.FontBold = False
    Tmp.FontItalic = False
End Sub

Sub TextFix_Property(Tmp As Control, ByVal FName$, ByVal FSize$)
'�]�wTextBox�r����Courier,Size=10���ݩ�,�S����J���ת�����

    Tmp.BackColor = Val(G_TextLostBack_Color)
    Tmp.ForeColor = Val(G_TextLostFore_Color)
    If Trim(FName) = "" Then FName = "Courier"
    If Trim(FSize) = "" Then FSize = "10"
    Tmp.FontName = FName$
    Tmp.FontSize = FSize$
    Tmp.FontBold = False
    Tmp.FontItalic = False
End Sub


Sub ListBox_Property(Tmp As Control, ByVal FName$, ByVal FSize$)
'�]�wListBox���ݩ�

    Tmp.Font.Name = FName$
    Tmp.Font.Size = FSize$
    Tmp.Font.Bold = False
    Tmp.FontItalic = False
End Sub


Function GetColor(ByVal Color$) As String
'���o�C���
    
    Select Case Color$
           Case "color_yellow"
                GetColor = Trim(COLOR_YELLOW)
           Case "color_blue"
                GetColor = Trim(COLOR_BLUE)
           Case "color_red"
                GetColor = Trim(COLOR_RED)
           Case "color_milk"
                GetColor = Trim(COLOR_MILK)
           Case "color_black"
                GetColor = Trim(COLOR_BLACK)
           Case "color_sky"
                GetColor = Trim(COLOR_SKY)
           Case "color_white"
                GetColor = Trim(COLOR_WHITE)
           Case "color_green"
                GetColor = Trim(COLOR_GREEN)
           Case "color_gray"
                GetColor = Trim(COLOR_GRAY)
           Case "color_darkgreen"
                GetColor = Trim(COLOR_DARKGREEN)
           Case "color_lightgreen"
                GetColor = Trim(COLOR_LIGHTGREEN)
    End Select
End Function
Sub FieldAssist(Lst As ListBox)
'��ܻ��U��ListBox

    If Lst.Visible Then Exit Sub
    If Lst.ListCount <= 0 Then Exit Sub
    If TypeOf Screen.ActiveControl Is ListBox Then Exit Sub
    
    Screen.ActiveForm!Vse_Background.AutoSizeChildren = azNone
    Lst.Visible = True
    Lst.ZOrder 0
    If Lst.ListCount <> 0 Then
       G_List_Flag = True
       Lst.Selected(0) = True
    End If
    Screen.ActiveForm!Vse_Background.AutoSizeChildren = azProportional
    Lst.SetFocus
End Sub
Sub Spread_Row_Property(Spd As vaSpread, ByVal Row#, ByVal text$)
'�]�wSpread�W,�C����r���e
    
    Spd.Col = 1
    Spd.Row = Row#
    Spd.text = text$
End Sub



Sub KeepOpen(File As Recordset, ByVal Name$)
'�N�}�Ҫ�Recordset,Keep��Array��
Dim A_Index%

    A_Index% = 0
    Do While A_Index% < 100
       If UCase$(Trim$(G_File(A_Index%))) = UCase$(Trim$(Name$)) Then
          Set G_FileName(A_Index%) = File
          G_File(A_Index%) = Trim$(Name$)
          Exit Do
       End If
       If G_FileName(A_Index%) Is Nothing Then
          Set G_FileName(A_Index%) = File
          G_File(A_Index%) = Trim$(Name$)
          Exit Do
       End If
       A_Index% = A_Index% + 1
    Loop
    
End Sub
Sub KeepOpenTable(File As Recordset, ByVal Name$)
'�N�}�Ҫ�Table,Keep��Array��
Dim A_Index%

    A_Index = 0
    Do While A_Index < 100
       If UCase$(Trim$(G_File(A_Index%))) = UCase$(Trim$(Name$)) Then
          Set G_TableName(A_Index%) = File
          G_File(A_Index%) = Trim$(Name$)
          Exit Do
       End If
       If G_TableName(A_Index%) Is Nothing Then
          Set G_TableName(A_Index%) = File
          G_File(A_Index%) = Trim$(Name$)
          Exit Do
       End If
       A_Index% = A_Index% + 1
    Loop
    
End Sub

Sub SpreadKeyPress(Spd As vaSpread, ByVal KeyCode%)
'�B�z�����J��,�Ĥ@�Ӧr���L�k��ܪ����D

    If KeyCode% >= 48 And KeyCode% <= 57 Then          '0 - 9
       Spd.EditMode = True
       Exit Sub
    End If
    If KeyCode% >= 65 And KeyCode% <= 90 Then          'a - z
       Spd.EditMode = True
       Exit Sub
    End If
    If KeyCode% >= 186 And KeyCode% <= 192 Then        '; , - . / `
       Spd.EditMode = True
       Exit Sub
    End If
    If KeyCode% >= 219 And KeyCode% <= 222 Then        '[ \ ] ,
       Spd.EditMode = True
    End If
    If KeyCode% = 229 Then                           'for Microsoft�s��J�k
       Spd.EditMode = True
    End If
End Sub

Sub VSElastic_Property(vsEtc As VideoSoftElastic)
'�]�wElastic���ݩ�,�C�Ӫ�橳�h���ϥ�Elastic

    'General Defined
    vsEtc.Template = 0          'tpNone
    vsEtc.style = esClassic
    'Panels Defined
    vsEtc.Align = asFill
    vsEtc.AutoSizeChildren = azProportional
    vsEtc.BackColor = G_Label_Color
    vsEtc.BevelOuter = bsGroove
    vsEtc.BevelInner = 0        'bsNone
    vsEtc.BevelOuterDir = bdHorz
    vsEtc.BevelChildren = bcAll
    vsEtc.BevelOuterWidth = 1
    vsEtc.BevelInnerWidth = 1
End Sub


Function CvrTxt2Num(ByVal text$) As Double
'�N��r���A�ഫ���ƭȫ��A
Dim A_STR$, A_Index%, A_Index1%, A_Minus%

    text$ = Trim$(text$)
    If text$ = "" Then CvrTxt2Num = 0: Exit Function
    
    A_STR$ = Space$(Len(text$))
    A_Minus% = False
    If Mid(text$, 1, 1) = "-" Or Mid(text$, Len(text$), 1) = "-" Or (Mid(text$, 1, 1) = "(" And Mid(text$, Len(text$), 1) = ")") Then A_Minus% = True
    
    A_Index1% = 1
    For A_Index% = 1 To Len(text$)
        If (Mid(text$, A_Index%, 1) >= "0" And Mid(text$, A_Index%, 1) <= "9") Or Mid(text$, A_Index%, 1) = "." Then
           Mid(A_STR$, A_Index1%, 1) = Mid(text$, A_Index%, 1)
           A_Index1% = A_Index1% + 1
           GoTo CvrTxt2NumA
        End If
        If A_Minus% = False Then
           If (Mid(text$, A_Index%, 1) = "C" Or Mid(text$, A_Index%, 1) = "c") And (Mid(text$, A_Index% + 1, 1) = "R" Or Mid(text$, A_Index% + 1, 1) = "r") Then
              A_Minus% = True
              GoTo CvrTxt2NumA
           End If
        End If
CvrTxt2NumA:
    Next A_Index%
    
    If A_Minus% = True Then
       CvrTxt2Num = Val(A_STR$) * (-1)
    Else
       CvrTxt2Num = Val(A_STR$)
    End If
End Function

Function CvrSumFields2Str(DB As Database, ByVal SumFields$) As String
'�N�J�`��ƪ���쫬�A�ഫ����r���A

    If Trim(DB.Connect) = "" Then       'Access Database
       CvrSumFields2Str = "CSTR(" & SumFields$ & ")"
    Else
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              CvrSumFields2Str = "Convert(varchar," & SumFields$ & ")"
       End Select
    End If
End Function

Sub ClearSpreadText(Spd As vaSpread)
'�M�����vaSpread�����e

    Spd.Row = -1
    Spd.Col = -1
    Spd.Action = SS_ACTION_CLEAR_TEXT
End Sub
Sub CreateDynasetProcess(DB As Database, DY As Recordset, ByVal SQL$, ByVal Str$)
On Local Error GoTo CreateDynasetProcess_Error
Dim A_Msg$, A_Msg1$, A_Msg2$, A_Msg3$, A_Msg4$, A_Msg5$
    
    A_Msg1$ = GetSIniStr("PanelDescpt", "unread")       '"��Ʈw�ثe�L�kŪ�g�A�еy�ݢ������U�T�w���~��,"
    A_Msg2$ = GetSIniStr("PanelDescpt", "cancel")       '"�Ϋ��U�����䵲�����\��!!"
    A_Msg3$ = GetSIniStr("PanelDescpt", "datachange")   '"��Ʈw���ʤ�,�ثe�L�kŪ�g�A�еy�ݢ������U�T�w���~��,"
    A_Msg4$ = GetSIniStr("PanelDescpt", "dataerror")    '"��ƮwŪ�g�o�Ϳ��~�A�{���N����!"
    A_Msg5$ = GetSIniStr("PanelDescpt", "writeerror")   '"�бN�����~�T���O�U�A�P�{���H���p��!"
    
    CloseOpen DY, Str$
    Set DY = DB.OpenRecordset(SQL$, dbOpenDynaset)
    DY.LockEdits = False
    Exit Sub

CreateDynasetProcess_Error:
    Select Case Err
      Case 3046, 3158, 3186, 3187, 3188, 3202, 3218, 3260  'Record Locked
           Idle
           A_Msg$ = Error(Err) & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg1$
           A_Msg$ = A_Msg$ & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg2$
           retcode = MsgBox(A_Msg$, MB_OKCANCEL + MB_ICONQUESTION, UCase$(App.Title))
           Err = 0
           Screen.ActiveForm.Refresh
           If retcode = IDOK Then Resume
           If retcode = IDCANCEL Then CloseFileDB: End
           
      Case 3167, 3197                                      'Record is deleted , changed.
           A_Msg$ = Error(Err) & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg3$
           A_Msg$ = A_Msg$ & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg2$
           retcode = MsgBox(A_Msg$, MB_OKCANCEL + MB_ICONQUESTION, UCase$(App.Title))
           Err = 0
           Screen.ActiveForm.Refresh
           If retcode = IDOK Then Resume
           If retcode = IDCANCEL Then CloseFileDB: End
           
      Case Else
           A_Msg$ = Str$ & Chr$(13) & Chr$(10)
           A_Msg$ = Error(Err) & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg4$
           A_Msg$ = A_Msg$ & A_Msg5$
           MsgBox A_Msg$, MB_ICONSTOP, UCase$(App.Title)
           
           Err = 0
           CloseFileDB
           End
    End Select
End Sub

Function AccessDBErrorMessage() As Integer
'�Φb�����Ʈw����Ʀs�����{�Ǥ�,�B�z���~�T�������
Dim A_Msg$
Dim A_Msg1$, A_Msg2$, A_Msg3$, A_Msg4$, A_Msg5$, A_Msg6$
    
    A_Msg1$ = GetSIniStr("PanelDescpt", "unread")       '"��Ʈw�ثe�L�kŪ�g�A�еy�ݢ������U�T�w���~��,"
    A_Msg2$ = GetSIniStr("PanelDescpt", "cancel")
    A_Msg3$ = GetSIniStr("PanelDescpt", "updating")     '"��ƥ��Q�L�H�ק襤�A�z�ثe�L�kŪ�g�A"
    A_Msg4$ = GetSIniStr("PanelDescpt", "wait")         '"�A�еy�ݢ������U�T�w���~��,�Ϋ��U�����䵲���ثe�@�~!!"
    A_Msg5$ = GetSIniStr("PanelDescpt", "dataerror")    '"��ƮwŪ�g�o�Ϳ��~�A�{���N����!"
    A_Msg6$ = GetSIniStr("PanelDescpt", "writeerror")   '"�бN�����~�T���O�U�A�P�{���H���p��!"

    Select Case Err
      Case 3046, 3158, 3186, 3187, 3188, 3202, 3218, 3260
           Idle
           A_Msg$ = Error(Err) & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg1$
           A_Msg$ = A_Msg$ & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg2$
           retcode = MsgBox(A_Msg$, MB_OKCANCEL + MB_ICONQUESTION, UCase$(App.Title))

      Case 3167, 3197
           A_Msg$ = Error(Err) & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg3$
           A_Msg$ = A_Msg$ & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg4$
           retcode = MsgBox(A_Msg$, MB_ICONQUESTION, UCase$(App.Title))
           
      Case 3146    'ODBC CALL FAIL
           A_Msg$ = GetODBCErrorMessage()
           retcode = MsgBox(A_Msg$, MB_ICONSTOP, UCase$(App.Title))
           retcode = IDCANCEL

      Case Else
           A_Msg$ = Error(Err) & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg5$
           A_Msg$ = A_Msg$ & Chr$(13) & Chr$(10)
           A_Msg$ = A_Msg$ & A_Msg6$
           retcode = MsgBox(A_Msg$, MB_ICONSTOP, UCase$(App.Title))
           retcode = IDCANCEL
    End Select
    WriteErrorReport A_Msg$, G_Str
    Err = 0
    
    On Error Resume Next
    Screen.ActiveForm.Refresh
    
    AccessDBErrorMessage = CInt(retcode)
End Function
Sub Frame_Property(Tmp As Control, ByVal text$, ByVal Size$, ByVal FName$)
'�]�wFrame���ݩ�
On Error Resume Next

    Tmp.Caption = text$
    Tmp.FontName = FName$
    Tmp.FontSize = Size$
    Tmp.FontBold = False
    Tmp.ForeColor = Val(G_TextLostFore_Color)
End Sub
Sub ComboBox_Property(Tmp As Control, ByVal Size$, ByVal FName$)
'�]�wComboBox���ݩ�

    Tmp.FontName = FName$
    Tmp.FontSize = Size$
    Tmp.FontBold = False
    Tmp.FontItalic = False
    Tmp.BackColor = G_TextLostBack_Color
    Tmp.ForeColor = Val(G_TextLostFore_Color)
End Sub
Sub SpreadHelpGotFocus(ByVal Col#, ByVal Row#)
'�B�zSpread�W���U�x�s��Gotfocus���C���ܤ�
On Local Error Resume Next

    If Not TypeOf Screen.ActiveForm.ActiveControl Is FPSPREAD.vaSpread Then Exit Sub
    If Row# <= 0 Then Exit Sub
    Screen.ActiveForm.ActiveControl.Row = Row#
    Screen.ActiveForm.ActiveControl.Col = Col#
    Screen.ActiveForm.ActiveControl.BackColor = G_TextHelpBack_Color
    Screen.ActiveForm.ActiveControl.ForeColor = G_TextGotFore_Color
End Sub
Function GetUserId() As String
'���o�n�J���ϥΪ̦W��

     GetUserId = G_DUserId
End Function

Function GetEmployeeID()
'���o�n�J�����u�s��
On Local Error GoTo MY_Error
Dim A_A0826$, A_Sql$, DY As Recordset

    A_A0826$ = G_DUserId
    '
    A_Sql$ = "Select A0801 From A08 Where A0826='" & ReplaceSingleSign(Trim(A_A0826$)) & "'"
    CreateDynasetODBC DB_ARTHGUI, DY, A_Sql$, "DY", True
    '
    If Not (DY.EOF And DY.BOF) Then
        GetEmployeeID = Trim(DY.Fields("A0801") & "")
    Else
        GetEmployeeID = ""
    End If
    '
    Exit Function
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Function DateFormat(ByVal DateStr$) As String
'�N����榡�Ƭ��~/��/��

    DateFormat = ""
    If Trim$(DateStr$) = "" Then Exit Function
    DateFormat = Format$(DateStr$, "##00/##/##")
End Function

Function GetLenStr(ByVal Source$, ByVal Start%, ByVal sLen%) As String
'�ۦr����o�S�w���ת����
Dim I%, A_Position%

    GetLenStr = ""
    A_Position% = 1
    If Start% > 1 Then
       For I% = Start% - 1 To 1 Step -1
           If lstrlen(Mid$(Source$, 1, I%)) < Start% Then
              A_Position% = I% + 1
              Exit For
           End If
       Next I%
    End If
    '
    For I% = sLen% To 1 Step -1
        If lstrlen(Mid$(Source$, A_Position%, I%)) <= sLen% Then
           GetLenStr = Mid$(Source$, A_Position%, I%)
           Exit For
        End If
    Next I%
End Function

Public Function Get_DateString(ByVal DateStr$, ByVal DateCnt%, ByVal Opt%) As String
'��������[��B��,�Ǧ^���榡
'      DateStr$���n�p�⪺����r��(6�Xor8�X)
'      DateCnt%���n�[��Ʀr
'      Opt%=1���p��"�~"
'      Opt%=2���p��"��"
'      Opt%=3���p��"��"
Dim A_Date, A_DateStr$, I%, A_Pos1%, A_Pos2%
Dim a_Year%, A_Month%, A_Day%
Dim A_YY$, A_MM$, A_DD$

    Get_DateString = ""
    DateStr$ = Trim(DateStr$)
    If DateStr$ = "" Then Exit Function
    
    ''�N����ର���榡
    DateStr$ = DateIn(DateStr$)
    
    ''�N����r���~,��,��
    a_Year% = Val(Mid(DateStr$, 1, 4))
    A_Month% = Val(Mid(DateStr$, 5, 2))
    A_Day% = Val(Mid(DateStr$, 7, 2))
    
    ''���o�g�L�p��᪺���
    Select Case Opt%
           Case 1
                A_Date = DateSerial(a_Year% + DateCnt%, A_Month%, A_Day%)
           Case 2
                A_Date = DateSerial(a_Year%, A_Month% + DateCnt%, A_Day%)
           Case 3
                A_Date = DateSerial(a_Year%, A_Month%, A_Day% + DateCnt%)
    End Select
    
    ''�ഫ����榡���r��(1997/11/1 �� 19971101)
    A_DateStr$ = Format$(A_Date, "YYYYMMDD")
    
    ''�N�ഫ�᪺����r��Ǧ^
    Get_DateString = A_DateStr$
End Function


Function ReplaceSingleSign(ByVal Str$) As String
'�B�z�d�ߪ�SQL���O��,���ȧt���R�X�Ÿ�" ' "
Dim I%, A_RStr$

    Str$ = Trim$(Str$)
    ReplaceSingleSign = Str$
    If Str$ = "" Then Exit Function
    '
    A_RStr$ = ""
    For I% = 1 To Len(Str$)
        A_RStr$ = A_RStr$ & Mid$(Str$, I%, 1)
        If Mid$(Str$, I%, 1) = "'" Then A_RStr$ = A_RStr$ & "'"
    Next I%
    ReplaceSingleSign = A_RStr$
End Function


Function GetFieldPos(DY As Recordset, ByVal FldName$) As Integer
'���o�����Recordset��Position
Dim Fld As Field
Dim I%

    GetFieldPos = 0
    I% = 0
    For Each Fld In DY.Fields
        If UCase$(Trim$(Fld.Name)) = UCase$(Trim$(FldName$)) Then
           GetFieldPos = I%
           Exit For
        End If
        I% = I% + 1
    Next
End Function
Function GetRowsOK(DY As Recordset, ByVal intRows#, varRecords As Variant) As Boolean
'��Recordset���@�����o�h�����

    GetRowsOK = False
    If DY.EOF Then Exit Function
    '
    varRecords = DY.GetRows(intRows#)
    GetRowsOK = True
End Function


Public Function GetWinPlatform() As Long
'���o�@�~�t�Ϊ��N��
Dim osvi As OSVERSIONINFO
Dim strCSDVersion As String

    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then Exit Function
    
    GetWinPlatform = osvi.dwPlatformId
End Function



Function IsWindows95() As Boolean
'�P�_OS�O�_��WIN95
Const dwMask95 = &H1&

    IsWindows95 = (GetWinPlatform() And dwMask95)
End Function




Function IsWindowsNT() As Boolean
'�P�_OS�O�_��WINNT
Const dwMaskNT = &H2&

    IsWindowsNT = (GetWinPlatform() And dwMaskNT)
End Function




Function IsWindowsNT4WithoutSP5() As Boolean
'�P�_OS��NT4.0��Service Pack�O�_��SP5�H�W
    
    IsWindowsNT4WithoutSP5 = False
    
    If Not IsWindowsNT() Then
       Exit Function
    End If
    
    Dim osvi As OSVERSIONINFO
    Dim strCSDVersion As String
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
       Exit Function
    End If
    strCSDVersion = StripTerminator(osvi.szCSDVersion)
    
    'Is this Windows NT 4.0?
    Const NT4MajorVersion = 4
    Const NT4MinorVersion = 0
    If (osvi.dwMajorVersion <> NT4MajorVersion) Or (osvi.dwMinorVersion <> NT4MinorVersion) Then
       'No.  Return True. Version Upper 4.0
       IsWindowsNT4WithoutSP5 = True
       Exit Function
    End If
    
    'If no service pack is installed, or if Service Pack 1 is
    'installed, then return True.
    Const strSP5 = "SERVICE PACK 5"
    If strCSDVersion = "" Then
       IsWindowsNT4WithoutSP5 = True 'No service pack installed
    ElseIf strCSDVersion = strSP5 Then
       IsWindowsNT4WithoutSP5 = True 'Only SP1 installed
    End If
End Function

Function StripTerminator(ByVal strText$) As String
'�^�ǥh�����XAscii Code=0���r��
Dim intZeroPos%

    intZeroPos% = InStr(strText$, Chr$(0))
    If intZeroPos% > 0 Then
        StripTerminator = Left$(strText$, intZeroPos% - 1)
    Else
        StripTerminator = strText$
    End If
End Function


Function GetTextBoxLineStr(Txt As Control, ByVal MaxLen%, ByVal LineNo&) As String
'���oTextBox�W�Y�C�����
Dim byteLo%, byteHi%, x, Buffer$

    byteLo% = MaxLen% And (255)
    byteHi% = Int(MaxLen% / 256)
    Buffer$ = Chr$(byteLo%) + Chr$(byteHi%) + Space$(MaxLen% - 2)
    
    x = SendMessageAsString(Txt.hwnd, EM_GETLINE, LineNo&, Buffer$)
    Buffer$ = RTrim(Buffer$)
    
    GetTextBoxLineStr = GetLenStr(Buffer$, 1, x)
End Function

Function GetTextBoxLineCount(Txt As Control) As Long
'���oTextBox�W����Ʀ��
Dim lcount

    lcount = SendMessageAsLong(Txt.hwnd, EM_GETLINECOUNT, 0, 0)
    GetTextBoxLineCount = lcount
End Function

Sub CboStrCut(ByVal cbo As Control, ByVal Str1$, ByVal CutStr$)
'�]�wComboBox�ثe��ܦC�ܰѼƤG�Ҧb�C
'�Ѽ�:ComboBox Name,��쪺��,���j�r��
Dim I%, A_Pos&

    cbo.ListIndex = -1
    If CutStr$ = "" Then Exit Sub
    
    For I% = 0 To cbo.ListCount - 1
        A_Pos& = InStr(cbo.List(I%), CutStr$)
        If A_Pos& = 0 Then A_Pos& = Len(cbo.List(I%)) + 1
        If UCase$(Trim$(Left$(cbo.List(I%), A_Pos& - 1))) = UCase$(Trim$(Str1$)) Then
           cbo.ListIndex = I%
           Exit For
        End If
    Next I%
End Sub

Function GetTextMultiOutput(ByVal Source$, ByVal MaxLen%) As String()
'�N�ǤJ��r�̨C�C�i�s�񤧳̤j��,��h�CKeep��Array��,��ܦ�TextBox���ϥ�
'**********************************************************************
'Function �ޥΤ��d�ҵ{��,�ǤJ��ӰѼ�
'Source$ : �ǤJ��r   MaxLen% : �C�C��ƪ��׳̤j��
'**********************************************************************
'�ŧiArray�ܼ�
'Dim I%, A_Str$()
'
'    �N�ǤJ��r�̨C�C�i�s�񤧳̤j��,��h�CKeep��Array
'    A_Str$ = GetTextMultiOutput(Trim(txt_input), 40)
'
'    ��Array�����X�C�C��ƳB�z
'    I% = 0
'    Do While I% < UBound(A_Str$)
'       I% = I% + 1
'       MsgBox CStr(I%) & " : " & A_Str$(I%)
'    Loop
'**********************************************************************
Dim A_Pos!, A_Tmp$, A_Char$, A_Buffer$
ReDim A_STR$(0)

    GetTextMultiOutput = A_STR$
    If Source$ = "" Then Exit Function
    If Trim(MaxLen%) = "" Then Exit Function
    '
    If Right$(Source$, 2) = Chr(13) & Chr(10) Then
        A_Buffer$ = Left$(Source$, Len(Source$) - 2)
    Else
        A_Buffer$ = Source$
    End If
    Do While Len(A_Buffer$) > 0
        A_Char$ = Left(A_Buffer$, 1)
        If A_Char$ = Chr(13) Then
           If Left(A_Buffer$, 2) = Chr(13) & Chr(10) Then
              GoSub ChangeLine
              A_Buffer$ = Right(A_Buffer$, Len(A_Buffer$) - 2)
           End If
        Else
           If A_Pos! >= MaxLen% Then
              GoSub ChangeLine_A
           End If
           'If A_Char$ > Chr(128) Then
           If lstrlen(A_Char$) = 2 Then
              If A_Pos! + 2 > MaxLen% Then
                 GoSub ChangeLine_A
              End If
              A_Pos! = A_Pos! + 2
           Else
              If A_Pos! + 1 > MaxLen% Then
                 GoSub ChangeLine_A
              End If
              A_Pos! = A_Pos! + 1
           End If
           A_Tmp$ = A_Tmp$ & A_Char$
           A_Buffer$ = Right(A_Buffer$, Len(A_Buffer$) - 1)
        End If
    Loop
    GoSub ChangeLine
    GetTextMultiOutput = A_STR$
    Exit Function
    
ChangeLine:
    ReDim Preserve A_STR$(0 To UBound(A_STR$) + 1)
    A_STR$(UBound(A_STR$)) = A_Tmp$
    A_Tmp$ = "": A_Pos! = 0
    Return
    
ChangeLine_A:
    If A_Char$ = Space(1) Then
       GoSub ChangeLine
       Return
    End If
    
    Dim A_Len!
    A_Len! = Len(A_Tmp$)
    Do Until A_Len! = 1
        If Mid(A_Tmp$, A_Len!, 1) = Space(1) Then
           A_Buffer$ = Right(A_Tmp$, Len(A_Tmp$) - A_Len!) + A_Buffer$
           A_Tmp$ = Left(A_Tmp$, A_Len! - 1)
           A_Char$ = Left(A_Buffer$, 1)
           Exit Do
        End If
        A_Len! = A_Len! - 1
    Loop
    GoSub ChangeLine
    Return
End Function

Function GetEngSingleLineText2Multi(ByVal Source$, ByVal MaxLen%) As String()
'for �^��r����
'�N�ǤJ��r�̨C�C�i�s�񤧳̤j��,��h�CKeep��Array��,��ܦ�TextBox���ϥ�
'**********************************************************************
'Function �ޥΤ��d�ҵ{��,�ǤJ��ӰѼ�
'Source$ : �ǤJ��r   MaxLen% : �C�C��ƪ��׳̤j��
'**********************************************************************
'�ŧiArray�ܼ�
'Dim I%, A_Str$()
'
'    �N�ǤJ��r�̨C�C�i�s�񤧳̤j��,��h�CKeep��Array
'    A_Str$ = GetTextMultiOutput(Trim(txt_input), 40)
'
'    ��Array�����X�C�C��ƳB�z
'    I% = 0
'    Do While I% < UBound(A_Str$)
'       I% = I% + 1
'       MsgBox CStr(I%) & " : " & A_Str$(I%)
'    Loop
'**********************************************************************
Dim I%, A_Len%, A_Tmp$, A_StrTmp$()
ReDim A_STR$(0)

    GetEngSingleLineText2Multi = A_STR$
    If Source$ = "" Then Exit Function
    If Trim(MaxLen%) = "" Then Exit Function
    '
    A_StrTmp$ = Split(Source$, Space(1))
    For I% = 0 To UBound(A_StrTmp$)
        If A_Len% + lstrlen(A_StrTmp$(I%)) > MaxLen% Then
            ReDim Preserve A_STR$(0 To UBound(A_STR$) + 1)
            A_STR$(UBound(A_STR$)) = A_Tmp$
            A_Len% = 0
            A_Tmp$ = ""
        End If
        A_Len% = A_Len% + lstrlen(A_StrTmp$(I%))
        A_Tmp$ = A_Tmp$ & A_StrTmp(I%)
        If I% <> UBound(A_StrTmp$) Then
            A_Len% = A_Len% + 1
            A_Tmp$ = A_Tmp$ & Space(1)
        End If
    Next I%
    ReDim Preserve A_STR$(0 To UBound(A_STR$) + 1)
    A_STR$(UBound(A_STR$)) = A_Tmp$
    '
    GetEngSingleLineText2Multi = A_STR$
End Function


Sub WriteLogForAUD(ByVal State%, ByVal LogStr$)
'Keep�ϥΪ̷s�W,�ק�,�R���O����A09
    
    State% = State% + 2
    If State% < 3 Then Exit Sub
    If State% > 6 Then Exit Sub
    '
    LogStr$ = Trim(LogStr$)
    LogStr$ = GetLenStr(LogStr$, 1, 50)
    'Write Log File
    '--------------Edit By Cathy 2004/4/14-----------------------
    If G_SecurityPgm = False Then
        WriteJournalLog DB_ARTHGUI, State%, UCase$(App.EXEName), LogStr$
    Else
        WriteJournalLog_Security DB_ARTHGUI, State%, UCase$(App.EXEName), LogStr$
    End If
End Sub

Function Ref_LB13(DY As Recordset, ByVal LB1301$, ByVal LB1302$) As Boolean
'Reference�U����ƫȤ�ۭq��
On Error GoTo Ref_LB13_Error
Dim A_Sql$

    Ref_LB13 = False
    A_Sql$ = "Select * From LB13"
    A_Sql$ = A_Sql$ & " WHERE LB1301='" & ReplaceSingleSign(LB1301$) & "'"
    If Trim(LB1302$) <> "" Then
       A_Sql$ = A_Sql$ & " AND LB1302='" & ReplaceSingleSign(LB1302$) & "'"
    End If
    A_Sql$ = A_Sql$ & " ORDER BY LB1301,LB1302"
    CreateDynasetODBC DB_ARTHGUI, DY, A_Sql$, "DY", True
    If Not (DY.BOF And DY.EOF) Then Ref_LB13 = True
    Exit Function
    
Ref_LB13_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Sub CloseSystemMenu(Frm As Form, ByVal MenuID&)
'�N�����W����j,�٭�,�j�p,�������\��Disable,�ѰѼƤG�]�wDisable���\��
Dim hMenu&, MII As MENUITEMINFO
    
    hMenu& = GetSystemMenu(Frm.hwnd, 0)
    
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    MII.wID = MenuID&
    
    GetMenuItemInfo hMenu&, MenuID&, False, MII
    
    MII.wID = xMenuID
    MII.fMask = MIIM_ID
    SetMenuItemInfo hMenu&, MenuID&, False, MII
    
    MII.fState = MII.fState Or MFS_GRAYED
    MII.fMask = MIIM_STATE
    SetMenuItemInfo hMenu&, MII.wID, False, MII
    
    SendMessage Frm.hwnd, WM_NCACTIVATE, True, ByVal 0&
End Sub

Sub OpenSystemMenu(Frm As Form, ByVal MenuID&)
'�N�����W����j,�٭�,�j�p,�������\��Enable,�ѰѼƤG�]�wEnable���\��
Dim hMenu As Long, MII As MENUITEMINFO
    
    hMenu = GetSystemMenu(Frm.hwnd, 0)
    
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    MII.wID = MenuID&
    
    GetMenuItemInfo hMenu, xMenuID, False, MII
    
    MII.wID = MenuID&
    MII.fMask = MIIM_ID
    SetMenuItemInfo hMenu, xMenuID, False, MII
    
    MII.fState = MII.fState And (Not MFS_GRAYED)
    MII.fMask = MIIM_STATE
    SetMenuItemInfo hMenu, MII.wID, False, MII
    
    SendMessage Frm.hwnd, WM_NCACTIVATE, True, ByVal 0&
End Sub

Sub SetHelpWindowPos(Fra As Control, Spd As vaSpread, ByVal Left%, ByVal Top%, ByVal Width%, ByVal Height%)
'�]�w���U������Frame��Spread����m
    
    Screen.ActiveForm!Vse_Background.AutoSizeChildren = azNone
    Fra.Move Left%, Top%, Width%, Height%
    Spd.Move 120, 240, Fra.Width - 270, Fra.Height - 360
    Fra.ZOrder 0
    Fra.Visible = True
    Screen.ActiveForm!Vse_Background.AutoSizeChildren = azProportional
End Sub

Sub SpreadSort(Spd As vaSpread, ByVal Col#, Optional ByVal SortWay% = SS_SORT_ORDER_ASCENDING, Optional ByVal Col1# = 1, Optional ByVal Row1# = 1, Optional ByVal Col2# = -1, Optional ByVal Row2# = -1)
'�H�Y�@���B�zSpread�W����ƭ��s�Ƨ�,�̦h�@�����

    With Spd
         If Row2# = -1 Then Row2# = .MaxRows
         If Col2# = -1 Then Col2# = .MaxCols
         .Row = Row1#
         .Col = Col1#
         .Row2 = Row2#
         .Col2 = Col2#
         .SortBy = SS_SORT_BY_ROW
         .SortKey(1) = Col#
         .SortKeyOrder(1) = SortWay%
         .Action = SS_ACTION_SORT
    End With
End Sub

Sub ProgressBoxShow(Frm As Form, Spd As vaSpread)
'��ƳB�z�e,�Y�pSpread Height,�����ProgressBar
    
    With Frm
         .Prb_Percent.Left = Spd.Left
         .Prb_Percent.Height = 405
         .Prb_Percent.Top = Spd.Top + Spd.Height - 405
         .Prb_Percent.Width = Spd.Width
         
         .Vse_Background.AutoSizeChildren = azNone
         Spd.Height = Spd.Height - 450
         .Prb_Percent.Visible = True
    End With
    Frm.Refresh
End Sub

Sub ProgressBoxHide(Frm As Form, Spd As vaSpread)
'��ƳB�z����,�NProgressBar����,�é�jSpread Height
    
    With Frm
         .Prb_Percent.Visible = False
         Spd.Height = Spd.Height + 450
         .Vse_Background.AutoSizeChildren = azProportional
    End With
    Frm.Refresh
End Sub

Function ConvertNullStr(DB As Database, ByVal FldName$, ByVal Options%) As String
'�Ǧ^���ȬO�_��Null��SQL�y�k

    If Trim(DB.Connect) = "" Then   'Access Database
       ConvertNullStr = "ISNull(" & FldName$ & ")"
    Else                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;", "ORACLE;"
              If Options% Then ConvertNullStr = FldName$ & " IS NULL"
              If Not Options% Then ConvertNullStr = "ISNull(" & FldName$ & ")"
         Case "DB2;"
              ConvertNullStr = FldName$ & " IS NULL"
       End Select
    End If
End Function

Sub SAY_TOTAL_TRD(ByVal tmp1$, ByVal Tmp3$, G_Str1 As String, G_Str2 As String)
'�N���ԧB�Ʀr�����B,�ഫ���^��j�g
'Tmp1$:�ഫ�ƭ�
'Tmp3$:���O�W��
'G_STR1=�ഫ��Ĥ@��
'G_STR1=�ഫ��ĤG��
Dim left_2%, right_1%, left_2_c$, right_1_c$
Dim SRL$, I%, LEN_NUM%
Dim tmp21$, tmp22$
Dim tmp31#, tmp32#, tmp33%
Dim a#, A1&, B#, B1#, C#, d#
Dim aa$, AA1$, BB$, BB1$, CC$, dd$, EE$
Dim x_num#, y_num#, x_char$, y_char$
Dim n_num#, z_num#, n_char$, z_char$
Dim A_Char$, B_CHAR$, C_CHAR$, D_CHAR$, E_CHAR$
Dim G_NUM#, h_num#, g_char$, h_char$
Dim i_num#, j_num#, i_char$, j_char$
Dim M_Position%
Dim str_cut() As String
Dim g_number(100) As String
Dim j%, A_STR$()


    A_Char = ""
    B_CHAR = ""
    C_CHAR = ""
    D_CHAR = ""
    g_char = ""
    h_char = ""
    i_char = ""
    j_char = ""
    SRL$ = ""
    
    G_Str1 = ""
    G_Str2 = ""
    g_number(0) = ""
    g_number(1) = "ONE"
    g_number(2) = "TWO"
    g_number(3) = "THREE"
    g_number(4) = "FOUR"
    g_number(5) = "FIVE"
    g_number(6) = "SIX"
    g_number(7) = "SEVEN"
    g_number(8) = "EIGHT"
    g_number(9) = "NINE"
    g_number(10) = "TEN"
    g_number(11) = "ELEVEN"
    g_number(12) = "TWELVE"
    g_number(13) = "THIRTEEN"
    g_number(14) = "FOURTEEN"
    g_number(15) = "FIFTEEN"
    g_number(16) = "SIXTEEN"
    g_number(17) = "SEVENTEEN"
    g_number(18) = "EIGHTEEN"
    g_number(19) = "NINETEEN"
    
    g_number(20) = "TWENTY"
    g_number(30) = "THIRTY"
    g_number(40) = "FORTY"
    g_number(50) = "FIFTY"
    g_number(60) = "SIXTY"
    g_number(70) = "SEVENTY"
    g_number(80) = "EIGHTY"
    g_number(90) = "NINETY"
    
    left_2 = 0
    right_1 = 0
    tmp21 = ""
    tmp22 = ""
    tmp31 = 0
    tmp32 = 0
    
    StrCut tmp1$, ".", tmp21, tmp22      'TMP21=200
    
    left_2_c = Left(tmp22, 2)
    If Len(tmp22) > 2 Then
       right_1_c = Right(tmp22, 1)
    Else
       right_1_c = ""
    End If
    left_2 = Val(left_2_c)
    right_1 = Val(right_1_c)
    
    tmp31 = Val(tmp21)      '���       'TMP31=200
    tmp32 = left_2
    '��Ʀ��
    If tmp31 >= 1000000000 Then             '00040=10
       a# = tmp31 / 1000000000
       A1& = tmp31 Mod 1000000000
       StrCut a#, ".", aa$, AA1$
       a# = Val(aa$)
       A1& = Val(AA1$)
       A1& = tmp31 Mod 1000000000
    Else                                    '000410
       a# = 0                              '000410
       A1& = tmp31                         '000410
    End If                                  '000410
    
    If A1& >= 1000000 Then                '000410
       B# = A1& / 1000000
       StrCut B#, ".", BB$, BB1$
       B# = Val(BB$)
       B1# = Val(BB1$)
       B1# = A1& Mod 1000000
    Else                                    '000410
       B# = 0                              '000410
       B1# = A1&                         '000410
    End If                                  '000410
    
    If B1# >= 1000 Then                   '000410
       C# = B1# / 1000
       d# = B1# Mod 1000
       StrCut C#, ".", CC$, dd$
       C# = Val(CC$)
       'd# = Val(dd$)
    Else                                    '000410
       C# = 0                              '000410
       d# = B1#                          '000410
    End If
    '000410
    '---------�W�C A,B,C,D�̨D�o�ƭȾ�z�X��r���e
    For I = 1 To 4
    
        If I = 1 And a# >= 100 Then
           x_num = Int(a# / 100)
           y_num = a# Mod 100
        Else
           If I = 1 And a# < 100 Then
              x_num = 0
              y_num = a#
           End If
        End If
        
        If I = 2 And B# >= 100 Then
           x_num = Int(B# / 100)
           y_num = B# Mod 100
        Else
           If I = 2 And B# < 100 Then
              x_num = 0
              y_num = B#
           End If
        End If
        
        If I = 3 And C# >= 100 Then
           x_num = Int(C# / 100)
           y_num = C# Mod 100
        Else
           If I = 3 And C# < 100 Then
              x_num = 0
              y_num = C#
           End If
        End If
        
        If I = 4 And d# >= 100 Then
           x_num = Int(d# / 100)
           y_num = d# Mod 100
        Else
           If I = 4 And d# < 100 Then
              x_num = 0
              y_num = d#
           End If
        End If
    
        'StrCut x_num, ".", aa$, AA1$
        'x_num = Val(aa$)
        'y_num = Val(AA1$)
        
        
        x_char = g_number(x_num)
        If y_num <= 19 Then
           y_char = g_number(y_num)
        Else
           n_num = y_num / 10
           StrCut n_num, ".", aa$, AA1$
           n_num = Val(aa$)
           z_num = Val(AA1$)
           n_char = g_number(n_num * 10)
           z_char = g_number(z_num)
           y_char = n_char + " " + z_char
        End If
    
        If I = 1 Then
           If a# >= 100 Then
              If x_char <> "" Then A_Char$ = x_char + " " + "HUNDRED" + " " + y_char
              If x_char = "" Then A_Char$ = y_char
           Else
              If a# > 0 Then
                 If x_char <> "" Then A_Char$ = x_char + " " + y_char
                 If x_char = "" Then A_Char$ = y_char
              End If
           End If
    
    '        If a# >= 100 Then A_Char$ = x_char + " " + "HUNDRED" + " " + y_char
    '        If a# < 100 Then A_Char$ = x_char
        End If
        
        If I = 2 Then
           If B# >= 100 Then
              If x_char <> "" Then B_CHAR$ = x_char + " " + "HUNDRED" + " " + y_char
              If x_char = "" Then B_CHAR$ = y_char
           Else
              If B# > 0 Then
                 If x_char <> "" Then B_CHAR$ = x_char + " " + y_char
                 If x_char = "" Then B_CHAR$ = y_char
              End If
           End If
        End If
    
        If I = 3 Then
           If C# >= 100 Then
              If x_char <> "" Then C_CHAR$ = x_char + " " + "HUNDRED" + " " + y_char
              If x_char = "" Then C_CHAR$ = y_char
           Else
              If C# > 0 Then
                 If x_char <> "" Then C_CHAR$ = x_char + " " + y_char
                 If x_char = "" Then C_CHAR$ = y_char
              End If
           End If
        End If
        
        If I = 4 Then
           If d# >= 100 Then
              If x_char <> "" Then D_CHAR$ = x_char + " " + "HUNDRED" + " " + y_char
              If x_char = "" Then D_CHAR$ = y_char
           Else
              If d# > 0 Then
                 If x_char <> "" Then D_CHAR$ = x_char + " " + y_char
                 If x_char = "" Then D_CHAR$ = y_char
              End If
           End If
        End If
        
    Next I
    
    '�p�Ʀ��
    G_NUM = 0
    
    If tmp32 >= 10 And tmp32 <= 99 Then
       G_NUM = tmp32
    End If
    
    If tmp32 > 0 And tmp32 <= 9 Then
       G_NUM = tmp32
    End If
    
    
    If G_NUM <= 19 Then
       g_char = g_number(G_NUM)
    Else
       i_num = G_NUM / 10
       StrCut i_num, ".", aa$, AA1$
       i_num = Val(aa$)
       j_num = Val(AA1$)
       i_char = g_number(i_num * 10)
       j_char = g_number(j_num)
       g_char = i_char + " " + j_char
    End If
    
    h_char = g_number(right_1)
    If right_1 <> 0 Then
       If g_char <> " " Then E_CHAR = g_char + " " + "POINT" + " " + h_char
    Else
       If g_char <> " " Then E_CHAR = g_char
    End If
    
    '----�r�����
    SRL$ = Tmp3$
    
    If A_Char <> "" Then SRL$ = SRL$ + " " + A_Char$ + " " + "BILLION"
    If B_CHAR <> "" Then SRL$ = SRL$ + " " + B_CHAR$ + " " + "MILLION"
    If C_CHAR <> "" Then SRL$ = SRL$ + " " + C_CHAR$ + " " + "THOUSAND"
    If D_CHAR <> "" Then SRL$ = SRL$ + " " + D_CHAR$
    If E_CHAR <> "" Then
       If left_2 <> 0 Then SRL$ = SRL$ + " " + "AND CENTS" + " " + E_CHAR$
       If left_2 = 0 Then SRL$ = SRL$ + " " + "AND" + E_CHAR$
    End If
    SRL$ = SRL$ + " " + "ONLY"
    
    A_STR$ = GetTextMultiOutput(Trim(SRL), 100)
    For j% = 1 To UBound(A_STR$)
        If j% = 1 Then G_Str1 = A_STR$(j%)
        If j% = 2 Then G_Str2 = A_STR$(j%)
    Next j%

'    If Len(SRL$) > 100 Then
'       StrCut SRL$, " AND ", G_Str1, G_Str2
'       If Trim$(G_Str2) <> "" Then
'          G_Str2 = "AND " & G_Str2
'       End If
'       If Len(G_Str1) > 100 Then
'          A_Str$ = GetTextMultiOutput(Trim(SRL), 100)
'          G_Str1 = A_Str$(1)
'          G_Str2 = A_Str$(2) '& G_Str2
'       End If
'    Else
'       G_Str1 = SRL$
'    End If
    
    
    
    'A_Str$ = GetTextMultiOutput(Trim(SRL), 70)
    
    '��Array�����X�C�C��ƳB�z
    'j% = 0
    'Do While j% < UBound(A_Str$)
    '   j% = j% + 1
    '   If j% = 1 Then G_Str1 = A_Str$(j%)
    '   If j% = 2 Then G_Str2 = A_Str$(j%)
    'Loop
End Sub

Function CheckGUI(ByVal IT05$) As Boolean
'�ˬd�Τ@�s���O�_�ŦX�޿�
ReDim a_cnt$(8)                 '�N�Τ@�s����iARRAY
ReDim A_Multi$(8)               '�ﭼ�Ʀr
ReDim A_TRAN$(16)               '�ۭ���2��Ƹm�JARRAY
ReDim A_Add$(16)                '��[��2��Ƹm�JARRAY
Dim A_CheckNo%, n%
    
    CheckGUI = False
    
    If Len(Trim(IT05$)) = 0 Then
       CheckGUI = True
       Exit Function
    End If
    If Len(Trim(IT05$)) <> 8 Then Exit Function

    A_Multi$(1) = "1": A_Multi$(2) = "2": A_Multi$(3) = "1": A_Multi$(4) = "2"
    A_Multi$(5) = "1": A_Multi$(6) = "2": A_Multi$(7) = "4": A_Multi$(8) = "1"

    For n% = 1 To 8
        a_cnt$(n%) = Mid(IT05$, n%, 1)
    Next n%
    
    For n% = 1 To 8
        If Len(Trim(Val(a_cnt$(n%)) * Val(A_Multi$(n%)))) <> 1 Then
           A_TRAN$(n%) = Left(Trim(Val(a_cnt$(n%)) * Val(A_Multi$(n%))), 1)
           A_TRAN$(n% + 8) = Right(Trim(Val(a_cnt$(n%)) * Val(A_Multi$(n%))), 1)
        Else
           A_TRAN$(n%) = Val(a_cnt$(n%)) * Val(A_Multi$(n%))
        End If
    Next n%
    
    For n% = 1 To 8
        If Len(Trim(Val(A_TRAN$(n%)) + Val(A_TRAN$(n% + 8)))) <> 1 Then
           A_Add$(n%) = Left(Val(A_TRAN$(n%)) + Val(A_TRAN$(n% + 8)), 1)
           A_Add$(n% + 8) = Right(Val(A_TRAN$(n%)) + Val(A_TRAN$(n% + 8)), 1)
        Else
           A_Add$(n%) = Val(A_TRAN$(n%)) + Val(A_TRAN$(n% + 8))
        End If
    Next n%
    
    For n% = 1 To 8
        A_CheckNo% = A_CheckNo% + Val(A_Add$(n%))
    Next n%
    
    If A_CheckNo% Mod 10 = 0 Then
       CheckGUI = True
    Else
       A_CheckNo% = 0
       For n% = 1 To 8
           If Trim(A_Add$(n% + 8)) <> "" Then
              A_CheckNo% = A_CheckNo% + Val(A_Add$(n% + 8))
           Else
              A_CheckNo% = A_CheckNo% + Val(A_Add$(n%))
           End If
       Next n%
       If A_CheckNo% Mod 10 = 0 And A_CheckNo% <> 0 Then
          CheckGUI = True
       End If
    End If
End Function

Function CheckIdentityID(ByVal A_ID$, Optional ByVal A_Sex$ = "") As Boolean
'�ˬd�����Ҧr���O�_�ŦX�޿�
'�Ѽ�A_ID$ :�����Ҧr��
'�Ѽ�A_Sex$:�ťժ����ʧO,C���q�渹,F��k��,M��k��
Dim sA_aa$, sA_bb$, sA_NO$
Dim sA_xx$, iA_n%, sA_check$
Dim sA_number$, iA_for%, dA_CheckNo#
    
    CheckIdentityID = False
    
    If Len(Trim(A_ID$)) = 0 Then
        CheckIdentityID = True
        Exit Function
    End If
    If Len(Trim(A_ID$)) <> 10 Then Exit Function
    If Trim(A_Sex$) = "C" Then Exit Function '���q�渹��8�X���Τ@�s��
    
    sA_aa = UCase(Left(A_ID$, 1))
    If sA_aa >= "A" And sA_aa <= "Z" Then '�p�G�����Ҧr���Ĥ@�X����A~Z
        sA_bb = Mid$(A_ID$, 2, 1)
        If sA_bb <> "1" And sA_bb <> "2" Then '�p�G�����Ҧr���ĤG�X����1��2
            Exit Function
        ElseIf Trim(A_Sex$) <> "" Then '�ťժ����P�_�k�k
            '�k,�����Ҧr���a�G�X��"2"
            '�k,�����Ҧr���a�G�X��"1"
            If UCase(Trim(A_Sex$)) = "F" Then
               If Trim(sA_bb) <> 2 Then Exit Function
            Else
               If Trim(sA_bb) <> 1 Then Exit Function
            End If
        End If
        sA_number = "0123456789"
        For iA_for = 3 To 10
            sA_NO = Mid$(A_ID$, iA_for, 1)
            If InStr(sA_number, sA_NO) = 0 Then '�p�G�����Ҧr��3~10�X�����Ʀr
                Exit Function
            End If
        Next iA_for
        sA_xx = "ABCDEFGHJKLMNPQRSTUVXYWZIO"
        iA_n = InStr(sA_xx, sA_aa)
        iA_n = iA_n + 9
        sA_check = Format$(iA_n, "00") + Right(A_ID$, 9)
        dA_CheckNo = Val(Left$(sA_check, 1)) * 1
        For iA_for = 2 To 10
            dA_CheckNo = dA_CheckNo + Val(Mid$(sA_check, iA_for, 1)) * (11 - iA_for) '�֥[�ˬd����
        Next iA_for
        dA_CheckNo = dA_CheckNo + Val(Right$(sA_check, 1)) * 1
        dA_CheckNo = dA_CheckNo Mod 10
        If dA_CheckNo <> 0 Then '�p�G�ˬd����㰣
            Exit Function
        Else '�����Ҹ��X���T
            CheckIdentityID = True
            Exit Function
        End If
    Else '�����ҭ��X�����^��r��
        Exit Function
    End If
End Function

'S010515001�ܧ�^�ǫ��O,�H�קKOverFlow
'Function ACS(ByVal Number#) As Integer
Function ACS(ByVal Number#) As Double
'�L����i��ܾ��
    
    ACS = 0
    If Number# <> Int(Number#) Then
       Number# = Int(Number# + 1)
    End If
    ACS = Number#
End Function

Sub SetActiveControlFocus()
'�N��г]�w��U�@��Control
Dim I%, A_Flag%, a_count%
Dim A_Active%, A_MinIndex%, A_MaxIndex%

    'Form���L����Control,���X���{��
    a_count% = Screen.ActiveForm.Controls.Count
    If a_count% = 0 Then Exit Sub
    
    With Screen.ActiveForm
        
         '���oForm���Ҧ�Control Tabindex���̤p�γ̤j��,
         '�è�Control Index����,Keep Control Tabindex��A_ControlIndex%()
         ReDim A_ControlIndex%(0 To a_count% - 1)
         On Error Resume Next
         A_ControlIndex%(0) = .Controls(0).TabIndex
         A_MinIndex% = .Controls(0).TabIndex
         A_MaxIndex% = .Controls(0).TabIndex
         On Error GoTo 0
         For I% = 1 To a_count% - 1
             On Error Resume Next
             A_ControlIndex%(I%) = .Controls(I%).TabIndex
             If Err Then
                A_ControlIndex%(I%) = -1
             Else
                If .Controls(I%).TabIndex < A_MinIndex% Then
                   A_MinIndex% = .Controls(I%).TabIndex
                ElseIf .Controls(I%).TabIndex > A_MaxIndex% Then
                   A_MaxIndex% = .Controls(I%).TabIndex
                End If
             End If
             On Error GoTo 0
         Next I%
        
         '��Control TabIndex����,Keep Control Index��A_IndexControl%()
         ReDim A_IndexControl%(A_MinIndex% To A_MaxIndex%)
         For I% = 0 To a_count% - 1
             If A_ControlIndex%(I%) <> -1 Then
                A_IndexControl%(A_ControlIndex%(I%)) = I%
             End If
         Next I%
            
         '�]�w�U�@Control���oFocus
         A_Active% = .ActiveControl.TabIndex
         A_Flag% = False
         If A_Active% < A_MaxIndex% Then
            For I% = A_Active% + 1 To A_MaxIndex%
                On Error Resume Next
                If .Controls(A_IndexControl%(I%)).TabStop And _
                .Controls(A_IndexControl%(I%)).Visible And _
                .Controls(A_IndexControl%(I%)).Enabled Then
                   .Controls(A_IndexControl%(I%)).SetFocus
                   If Err = 0 Then A_Flag% = True: Exit For
                End If
            Next I%
            On Error GoTo 0
         End If
         If Not A_Flag% And A_Active% <> A_MinIndex% Then
            For I% = A_MinIndex% To A_Active% - 1
                On Error Resume Next
                If .Controls(A_IndexControl%(I%)).TabStop And _
                .Controls(A_IndexControl%(I%)).Visible And _
                .Controls(A_IndexControl%(I%)).Enabled Then
                   .Controls(A_IndexControl%(I%)).SetFocus
                   If Err = 0 Then A_Flag% = True: Exit For
                End If
            Next I%
            On Error GoTo 0
         End If
     
    End With
End Sub

Function Check_Company(ByVal Company$, ByVal UserID$, Optional ByVal ShowMessage$ = True) As Boolean
'�ˮ�User�O�_���ϥθӤ��q�O���v��
On Local Error GoTo MY_Error
Dim A_Sql$

    Check_Company = True
    '
    If UCase(Trim(G_CheckCompany)) <> "Y" Then Exit Function
    '
    A_Sql$ = "Select SECTION from Sini "
    A_Sql$ = A_Sql$ & " Where Section='check_company'"
    A_Sql$ = A_Sql$ & " And Topic='" & Trim(Company$) & "_" & Trim(UserID$) & "'"
    CreateDynasetODBC DB_ARTHGUI, DY_SINI, A_Sql$, "DY_SINI", True
    If Not (DY_SINI.EOF And DY_SINI.BOF) Then Exit Function
    '
    If ShowMessage$ Then MsgBox GetSIniStr("PgmMsg", "no_authority"), MB_OK
    '
    Check_Company = False
    Exit Function
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Function GetSQLServerName(ByVal CnnStr$) As String
'���oSQL Server���W��,��SQL Server��Join��,�����wSQL Server Name
Dim A_DSN$, A_STR$, A_Pos%, A_Pos2%

    A_Pos% = InStr(1, CnnStr$, "DSN=", vbTextCompare)
    A_Pos2% = InStr(A_Pos%, CnnStr$, ";", vbTextCompare)
    A_DSN$ = Mid(CnnStr$, A_Pos% + 4, A_Pos2% - A_Pos% - 4)
    A_STR$ = GetRegSetting(gsODBC_INI_REG_KEY, A_DSN$, "Server", "", HKEY_LOCAL_MACHINE)
    GetSQLServerName = "[" & A_STR$ & "]."
End Function

Public Function GetRegSetting(strMainKey$, strSubKey$, strValueName$, _
Optional strDefault$ = "", Optional hRootKey& = HKEY_LOCAL_MACHINE) As String
'���o���U�ɤ������
Dim hKey&, cbData&, lType&
Dim strKey$, strData As String * glMAX_NAME_LENGTH
    
    If strSubKey$ = "" Then
        strKey$ = strMainKey$
    Else
        strKey$ = strMainKey$ & "\" & strSubKey$
    End If
    
    If RegOpenKeyEx(hRootKey&, strKey$, 0, KEY_READ, hKey&) = ERROR_SUCCESS Then
        cbData& = LenB(StrConv(strData, vbFromUnicode))
        Dim lReserved As Long
        If RegQueryValueEx(hKey&, strValueName$, 0, lType&, ByVal strData, cbData&) = ERROR_SUCCESS Then
            GetRegSetting = StripTerminator(strData)
        Else
            GetRegSetting = strDefault$
        End If
        RegCloseKey hKey&
    Else
        GetRegSetting = strDefault$
    End If
End Function

Sub Menu_Property(men As Menu, ByVal Caption$)
'�]�wMenu���󪺼��D

    If Trim(Caption$) <> "" Then men.Caption = Caption$
End Sub

Function DisplayOverMaxLines(ByVal Records&, Optional ByVal A_MaxRecords& = 30) As Boolean
'��ܬd�ߵ��G�W�L30��,�O�_��ܪ��T��
Dim A_Msg$

    DisplayOverMaxLines = True
    If Records& > A_MaxRecords& Then
       A_Msg$ = GetSIniStr("PanelDescpt", "total") & Format(Records&, " # ") & GetSIniStr("PanelDescpt", "continue")
       retcode = MsgBox(A_Msg$, MB_YESNO + MB_ICONINFORMATION, Screen.ActiveForm.Caption)
       If retcode = IDNO Then DisplayOverMaxLines = False
    End If
End Function

Function CalPwdDueDate(ByVal Period%, ByVal Time As Date) As String
'�p��s�K�X�����Ĥ��
Dim A_DueDate As Date

    If Period% = 0 Then Period% = 1
    A_DueDate = DateAdd("m", Period%, Time)
    CalPwdDueDate = Year(A_DueDate) & Format(Month(A_DueDate), "00") & _
                    Format(Day(A_DueDate), "00")
End Function

Function GetInvoiceTitle(DB As Database, ByVal A_ID$, ByVal A_InvoiceTitle$) As String
'�ت�:�ѨM�o�����Y�L�k��J�W�L50�Ӧr�������D
'ADD BY ������ SRN:S910412020
'����o�����Y��
'1.�YARTHGUI..SINI���ӫȤ�o�����Y�h�H������
'2.�Y�L�h�H��A_InvoiceTitle$�^��
Dim DY As Recordset

    GetInvoiceTitle = ""
    If Ref_SINIA(DB, DY, "InvoiceTitle_" & A_ID$, "", "") = True Then
        Do While Not DY.EOF
            GetInvoiceTitle = GetInvoiceTitle & Trim(DY.Fields("TopicValue") & "")
            DY.MoveNext
        Loop
        
    End If
    If GetInvoiceTitle <> "" Then Exit Function
    GetInvoiceTitle = A_InvoiceTitle$
End Function

Sub WriteInvoiceTitle(DB As Database, ByVal A_ID$, ByVal A_InvoiceTitle$, Optional ByVal A_Limit% = 40)
'�ت�:�ѨM�o�����Y�L�k��J�W�L50�Ӧr�������D
'ADD BY ������ SRN:S910412020
'�g�J�o�����Y��
'1.�Y��J��A_InvoiceTitle$����<=A_Limit%�h���B�z
'2.�NA_InvoiceTitle$�C50�Ӧr�����@��g�JARTHGUI..SINI
'3.SECTION="InvoiceTitle_" & A_ID$
'4.TOPIC=��ƧǸ�
'5.TOPICVALUE=�Ҥ��Ϊ��r��
 
Dim A_STR$, A_Line%, A_Len%, A_Section$, A_Topic$, A_TopicValue$
    If lstrlen(A_InvoiceTitle$) <= A_Limit Then
        A_Section$ = "InvoiceTitle_" & A_ID$
        A_Topic$ = ""
        GoSub DeleteSINI
        Exit Sub
    End If
    A_Line% = 0
    
    A_STR$ = GetLenStr(A_InvoiceTitle$, 1, 50)
    A_Len% = 0
    Do While A_STR$ <> ""
        A_Line% = A_Line% + 1
        A_Section$ = "InvoiceTitle_" & A_ID$
        A_Topic$ = Format(A_Line%, "0")
        A_TopicValue$ = A_STR$
        GoSub MoveData2Sini
        A_Len% = A_Len% + lstrlen(A_STR$)
        A_STR$ = GetLenStr(A_InvoiceTitle$, A_Len% + 1, 50)
    Loop
    Exit Sub
    
MoveData2Sini:
    GoSub DeleteSINI
    
    G_Str = ""
    InsertFields "Section", Trim(A_Section$), G_Data_String
    InsertFields "Topic", Trim(A_Topic$), G_Data_String
    InsertFields "TopicValue", Trim(A_TopicValue$), G_Data_String
    SQLInsert DB, "SINI"
    Return
    
DeleteSINI:
    G_Str = "DELETE Sini Where Section ='" & Trim(A_Section$) & "'"
    If A_Topic$ <> "" Then
        G_Str = G_Str & " And Topic='" & Trim(A_Topic$) & "'"
    End If
    ExecuteProcess DB, G_Str
    Return
End Sub

Function Ref_SINIA(DB As Database, DY As Recordset, ByVal A_Section$, ByVal A_Topic$, ByVal A_TopicValue$) As Boolean
'��J  : A_SECTION   :  �ťեN�����
'��J  : A_TOPIC     :  �ťեN�����
'��J  : A_TOPICVALUE:  �ťեN�����
'��X  : DY
'RETURN: TRUE-�����  FALSE-�L���
On Error GoTo Ref_SINIA_Error
Dim A_Sql$

    Ref_SINIA = True
    A_Sql$ = "SELECT * FROM SINI WHERE 1=1"
    If Trim(A_Section) <> "" Then A_Sql$ = A_Sql$ + " AND SECTION='" & A_Section$ & "'"
    If Trim(A_Topic) <> "" Then A_Sql$ = A_Sql$ + " AND TOPIC='" & A_Topic$ & "'"
    If Trim(A_TopicValue) <> "" Then A_Sql$ = A_Sql$ + " AND TOPICVALUE='" & A_TopicValue$ & "'"
    A_Sql$ = A_Sql$ + " ORDER BY SECTION,TOPIC,TOPICVALUE"
    CreateDynasetODBC DB, DY, A_Sql$, "DY", True
    If Not (DY.BOF And DY.EOF) Then Exit Function
    Ref_SINIA = False
    Exit Function
    
Ref_SINIA_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Sub CvrHalfCharToFully(KeyAscii As Integer)
'�N�^��,�Ʀr�r���ѥb���ഫ������

    '�p�G���ҥΥ����]�w, ���B�z���禡���{���X
    If Not G_FullyChar% Then Exit Sub
    
    'S020925020 1021007����r�����b�����ഫ=kevin=
    If KeyAscii < 0 Or KeyAscii > 255 Then Exit Sub
    
    '�N�p�g�^���ഫ���j�g
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
    '�p�G�r�����b��,�ഫ������
    ' A   -  Z  �b���r����Ascii Code�� (65)     - (90)
    ' �ϡ@-�@�� �����r����Ascii Code�� (-23857) - (-23832)
    ' 0   -  9  �b���r����Ascii Code�� (65)     - (90)
    ' ���@-�@�� �����r����Ascii Code�� (-23889) - (-23880)
    ' -  �b���r����Ascii Code�� (45)
    ' �� �����r����Ascii Code�� (-24112)
    ' " "  �b�Φr����Ascii Code��(32)
    ' "�@" ����r����Ascii Code��(-24256)
    '(  �b���r����Ascii code�� (40)
    '�] �����r����Ascii code�� (-24227)
    ')  �b���r����Ascii code�� (41)
    '�^ �����r����Ascii code�� (-24226)
    '/  �b���r����Ascii code�� (47)
    '�� �����r����Ascii code�� (-24066)
    '\  �b���r����Ascii code�� (92)
    '�@ �����r����Ascii code�� (-24000)
    If lstrlen(Chr(KeyAscii)) = 1 Then
        Select Case Chr(KeyAscii)
            Case "A" To "Z"
                KeyAscii = KeyAscii - 23922
            Case "0" To "9"
                KeyAscii = KeyAscii - 23937
            Case "-"
                KeyAscii = KeyAscii - 24157
            Case " "
                KeyAscii = KeyAscii - 24288
            Case "(", ")"
                KeyAscii = KeyAscii - 24267
            Case "/"
                KeyAscii = KeyAscii - 24113
            Case "\"
                KeyAscii = KeyAscii - 24092
        End Select
    End If
End Sub

Function IsFullyText(ByVal Str$) As Boolean
'�ˮ֦r�ꤤ���C�Ӧr���O�_���O�����r��
Dim A_CharLen%, A_TextLen%

    IsFullyText = True
    If Trim(Str$) = "" Then Exit Function

    '�p�G���ҥΥ����]�w, ���B�z���禡���{���X
    If Not G_FullyChar% Then Exit Function
    
    A_CharLen% = Len(Str$)
    A_TextLen% = lstrlen(Str$)
    
    IsFullyText = (A_CharLen% * 2 = A_TextLen%)
End Function

Sub InitialtSpdTextValue(tSPD As Spread)
'��l��tSpd��������
Dim I#, A_Cols#

    A_Cols# = UBound(tSPD.Columns)
    For I# = 1 To A_Cols#
        tSPD.Columns(I#).text = ""
    Next I#
End Sub

'===============================================================================
'Edit Old Function at 92/7/7
'===============================================================================
Function GetNonValueSQL(DB As Database, ByVal A_FldName$, ByVal A_Operator$, Optional ByVal Options% = True) As String
'���o�U��Ʈw�U,����"="��"<>"��">"�ŭȪ�����
'�ŭȧt�ťխȤ�Null��
Dim A_STR$
    
    'Access DB : Fld < > ' ', ���G��  �|  �]�tNull�Ȫ�Record
    '            Fld  =  ' ', ���G�� ���| �]�tNull�Ȫ�Record
    'SQL    DB : Fld < > ' ', ���G�� ���| �]�tNull�Ȫ�Record
    '            Fld  =  ' ', ���G�� ���| �]�tNull�Ȫ�Record
    A_STR$ = GetSQLTransferNull(DB, A_FldName$, " ")
    GetNonValueSQL = A_STR$ & A_Operator$ & "' ' "
End Function

Sub Spread_DataType_Property(Spd As vaSpread, ByVal Col#, ByVal DType%, ByVal MIN$, ByVal MAX$, ByVal length%, Optional ByVal Alignment% = -1)
'�]�wSpread��쪺��ƫ��A

    Spd.Row = -1
    Spd.Col = Col#
    Spd.CellType = DType%                           'DATATYPE = INTEGER
    Select Case DType%
      Case SS_CELL_TYPE_EDIT
           Spd.TypeEditLen = length%                '��r��Ƥ�����
      Case SS_CELL_TYPE_FLOAT
           Spd.TypeFloatMin = MIN$                  '�B�I�Ƥ��̤p��
           Spd.TypeFloatMax = MAX$                  '�B�I�Ƥ��̤j��
           Spd.TypeFloatDecimalChar = Asc(".")      '�]�w�p���I����ܫ��A
           Spd.TypeFloatDecimalPlaces = length%
           Spd.TypeFloatSeparator = True            '�]�w�T��@ ,
      Case SS_CELL_TYPE_INTEGER
           Spd.TypeIntegerMin = MIN$                '��Ƥ��̤p��
           Spd.TypeIntegerMax = MAX$                '��Ƥ��̤j��
      Case SS_CELL_TYPE_CHECKBOX
    End Select
    
    '��줧����覡
    If Alignment% <> -1 Then
       If Alignment% = SS_CELL_H_ALIGN_CENTER And DType% = SS_CELL_TYPE_CHECKBOX Then
          Spd.TypeCheckCenter = True
       Else
          Spd.TypeHAlign = Alignment%
       End If
    End If
End Sub

Sub Spread_Col_Property(Spd As vaSpread, ByVal Col#, ByVal length%, ByVal text$, Optional ByVal ColHide% = False)
'�]�wSpread��쪺�ݩ�

    Spd.ColWidth(Col#) = length%    '�]�w�C�檺�e��
    Spd.Row = 0
    Spd.Col = Col#
    Spd.text = text$                '�]�w HEADING
    Spd.ColHidden = ColHide%        '���O�_����
End Sub

'===============================================================================
' Add New Function at 92/7/7
'===============================================================================
Function GetSQLRepeatChar(DB As Database, ByVal Character$, ByVal Count$, Optional ByVal AliasName$, Optional ByVal Options% = True) As String
'���o���w���ƭ��Ʀr����SQL���
Dim A_FmtStr$

    AliasName$ = Trim(AliasName$)
    Character$ = Replace(Character$, "'", "''", , , vbTextCompare)
    '
    If Trim(DB.Connect) = "" Or Not Options% Then   'Access Database
       A_FmtStr$ = " String(@count,'@char') "
    Else                                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              A_FmtStr$ = " Replicate('@char',@count) "
       End Select
    End If
    A_FmtStr$ = Replace(A_FmtStr$, "@count", CStr(Count$), 1, -1, vbTextCompare)
    A_FmtStr$ = Replace(A_FmtStr$, "@char", Character$, 1, -1, vbTextCompare)
    If AliasName$ <> "" Then A_FmtStr$ = A_FmtStr$ & " AS " & AliasName$ & " "
    '
    GetSQLRepeatChar = A_FmtStr$
End Function

Function GetSQLTransferNull(DB As Database, ByVal FldName$, ByVal ReplaceStr$, Optional ByVal IsTypeStr% = True, Optional ByVal AliasName$, Optional ByVal Options% = True) As String
'���o�ϥίS�w���N�ȨӴ���NULL�Ȫ�SQL���
'�Ѽ� : FldName$ - ���W��
'       ReplaceStr$ - ���NNULL�Ȫ��r��
'       IsTypeStr% - �O�_�૬����r
'       AliasName$ - ���O�W
Dim A_FmtStr$

    AliasName$ = Trim(AliasName$)
    ReplaceStr$ = Replace(ReplaceStr$, "'", "''", , , vbTextCompare)
    '
    If Trim(DB.Connect) = "" Or Not Options% Then   'Access Database
       A_FmtStr$ = " IIf(@FldName IS NULL, @ReplaceStr, @FldName) "
    Else                                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;", "DB2;"
              A_FmtStr$ = " ISNULL(@FldName, @ReplaceStr) "
         Case "ORACLE;"
              A_FmtStr$ = " NVL(@FldName, @ReplaceStr) "
       End Select
    End If
    If IsTypeStr% Then
       A_FmtStr$ = Replace(A_FmtStr$, "@ReplaceStr", "'@ReplaceStr'", 1, -1, vbTextCompare)
    End If
    A_FmtStr$ = Replace(A_FmtStr$, "@FldName", FldName$, 1, -1, vbTextCompare)
    A_FmtStr$ = Replace(A_FmtStr$, "@ReplaceStr", ReplaceStr$, 1, -1, vbTextCompare)
    If AliasName$ <> "" Then A_FmtStr$ = A_FmtStr$ & " AS " & AliasName$ & " "
    '
    GetSQLTransferNull = A_FmtStr$
End Function

Function GetSQLTopRows(DB As Database, ByVal SQL$, ByVal Rows%, Optional ByVal Options% = True) As String
'�bSQL Command���[�J���o�e�X����ƪ�SQL�B�⦡
Dim A_Sql$, A_Find$, A_Replace$

    A_Find$ = "SELECT "
    A_Replace$ = "SELECT TOP " & CStr(Rows%) & " "
    '
    If Trim(DB.Connect) = "" Or Not Options% Then   'Access Database
       A_Sql$ = Replace(SQL$, A_Find$, A_Replace$, 1, 1, vbTextCompare)
    Else                                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              A_Sql$ = Replace(SQL$, A_Find$, A_Replace$, 1, 1, vbTextCompare)
         Case "ORACLE;"
              A_Find$ = " WHERE "
              If InStr(1, SQL$, A_Find$, vbTextCompare) > 0 Then
                 A_Replace$ = " WHERE ROWNUM <= " & CStr(Rows%) & " AND "
                 A_Sql$ = Replace(SQL$, A_Find$, A_Replace$, 1, 1, vbTextCompare)
              Else
                 A_Sql$ = SQL$ & " WHERE ROWNUM <= " & CStr(Rows%)
              End If
       End Select
    End If
    '
    GetSQLTopRows = A_Sql$
End Function

Function GetSQLCvrFldType(DB As Database, ByVal FldStr$, ByVal FldType%, Optional ByVal AliasName$, Optional ByVal Options% = True) As String
'���o��쫬�A�ഫ��SQL���
'�Ѽ� : FldStr$ - ���ഫ����L��ƫ��A�����W�٩ιB�⦡�r��(�p:Sum(A1620))
'       FldType% - ���ഫ����ƫ��A
'                  G_Data_Numeric: �૬���ƭ�
'                  G_Data_String: �૬����r
'                  G_Data_Date: �૬�����
'                  G_Data_Float: �૬��Float (Add By Lidia-S010723048 �쬰String�૬��Numeric,�Y�쬰�ťզr��i��[�`�A�|�{���~�T��)
Dim A_FMT$

    AliasName$ = Trim(AliasName$)
    '
    If Trim(DB.Connect) = "" Or Not Options% Then   'Access Database
       Select Case FldType%
         Case G_Data_Numeric
              A_FMT$ = " CCur(@FldStr) "
         Case G_Data_String
              A_FMT$ = " CStr(@FldStr) "
         Case G_Data_Date
              A_FMT$ = " DateSerial(LEFT(@FldStr,4),MID(@FldStr,5,2),RIGHT(@FldStr,2))) "
         Case G_Data_Float
              A_FMT$ = " Val(@FldStr) "
       End Select
    Else                                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              Select Case FldType%
                Case G_Data_Numeric
                     A_FMT$ = " Convert(Numeric(25,4),@FldStr) "
                Case G_Data_String
                     A_FMT$ = " Convert(VarChar,@FldStr) "
                Case G_Data_Date
                     A_FMT$ = " Convert(DateTime,@FldStr) "
                Case G_Data_Float
                     A_FMT$ = " Convert(Float,@FldStr) "
              End Select
       End Select
    End If
    A_FMT$ = Replace(A_FMT$, "@FldStr", FldStr$, 1, -1, vbTextCompare)
    If AliasName$ <> "" Then A_FMT$ = A_FMT$ & " AS " & AliasName$ & " "
    '
    GetSQLCvrFldType = A_FMT$
End Function

Function GetSQLCvrFld2Date(DB As Database, ByVal FldStr$, Optional ByVal IsFldName% = True, Optional ByVal AliasName$, Optional ByVal Options% = True) As String
'���o���Τ�r���A�ഫ��Date��SQL���
'�Ѽ� : FldStr$ - ���ഫ������榡�����W�٩Τ�r��(�p:20030707)
'       IsFldName% - FldStr$�ܼƭȬO�_�����W��
Dim A_FMT$

    AliasName$ = Trim(AliasName$)
    '
    If Trim(DB.Connect) = "" Or Not Options% Then   'Access Database
       A_FMT$ = " DateSerial(LEFT(@FldStr,4),MID(@FldStr,5,2),RIGHT(@FldStr,2)) "
    Else                                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              A_FMT$ = " Convert(DateTime,@FldStr) "
       End Select
    End If
    If Not IsFldName% Then
       If IsDate(FldStr) Then
          A_FMT$ = " '@FldStr' "
       Else
          A_FMT$ = Replace(A_FMT$, "@FldStr", "'@FldStr'", 1, -1, vbTextCompare)
       End If
    End If
    A_FMT$ = Replace(A_FMT$, "@FldStr", FldStr$, 1, -1, vbTextCompare)
    If AliasName$ <> "" Then A_FMT$ = A_FMT$ & " AS " & AliasName$ & " "
    '
    GetSQLCvrFld2Date = A_FMT$
End Function

Function GetSQLDateDiff(DB As Database, ByVal Interval%, ByVal Date1$, ByVal Date2$, Optional ByVal AliasName$, Optional ByVal Options% = True) As String
'���o��Ӥ�������j���~�Τ�Τ�ƥت�SQL���
'�Ѽ� : Interval% - �]�w���o�ƥت����j��� (1:�~ 2:�� 3:��)
'       Date1$ - �Ĥ@�Ӥ���榡���r�� (�i�I�sGetSQLCvrFld2Date Function���o)
'       Date2$ - �ĤG�Ӥ���榡���r�� (�i�I�sGetSQLCvrFld2Date Function���o)
Dim A_FMT$, A_DatePart$

    AliasName$ = Trim(AliasName$)
    '
    If Trim(DB.Connect) = "" Or Not Options% Then   'Access Database
       A_FMT$ = " DateDiff('@DatePart',@Date1,@Date2) "
       A_DatePart$ = Choose(Interval%, "yyyy", "m", "d")
       If IsNull(A_DatePart$) Then A_DatePart = "d"
    Else                                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              A_FMT$ = " DateDiff(@DatePart,@Date1,@Date2) "
              A_DatePart$ = Choose(Interval%, "Year", "Month", "Day")
              If IsNull(A_DatePart$) Then A_DatePart = "Day"
       End Select
    End If
    A_FMT$ = Replace(A_FMT$, "@DatePart", A_DatePart$, 1, -1, vbTextCompare)
    A_FMT$ = Replace(A_FMT$, "@Date1", Date1$, 1, -1, vbTextCompare)
    A_FMT$ = Replace(A_FMT$, "@Date2", Date2$, 1, -1, vbTextCompare)
    If AliasName$ <> "" Then A_FMT$ = A_FMT$ & " AS " & AliasName$ & " "
    '
    GetSQLDateDiff = A_FMT$
End Function

Function GetSQLCase(DB As Database, ByVal ArgList, Optional ByVal AliasName$, Optional ByVal Options% = True) As String
'���o�h���i�ध�@���G��SQL�B�⦡
'�Ѽ� : ArgList - �s��Ҧ�����B�⦡�Ψ�������G�}�C���X��Array
'       AliasName$ - ���O�W
Dim I%, A_FMT$, A_Replace$, A_RetStr$, A_HaveElse%

    AliasName$ = Trim(AliasName$)
    '
    If Trim(DB.Connect) = "" Or Not Options% Then   'Access Database
       A_FMT$ = " IIf(@Condition,@TruePart,@FalsePart"
       For I% = 0 To UBound(ArgList)
           A_Replace$ = A_FMT$
           A_Replace$ = Replace(A_Replace$, "@Condition", ArgList(I%)(0), 1, 1, vbTextCompare)
           A_Replace$ = Replace(A_Replace$, "@TruePart", ArgList(I%)(1), 1, 1, vbTextCompare)
           If I% + 1 <= UBound(ArgList) Then
              If Trim(ArgList(I% + 1)(0)) = "" Then A_HaveElse% = True
           End If
           If A_HaveElse% Then
              A_Replace$ = Replace(A_Replace$, "@FalsePart", ArgList(I% + 1)(1), 1, 1, vbTextCompare)
              A_RetStr$ = A_RetStr$ & A_Replace$
              Exit For
           Else
              A_Replace$ = Left(A_Replace$, InStrRev(A_Replace$, ",", -1, vbTextCompare))
              A_RetStr$ = A_RetStr$ & A_Replace$
           End If
       Next I%
       If Not A_HaveElse% Then A_RetStr$ = A_RetStr$ & "'0'"
       A_RetStr$ = A_RetStr$ & String(UBound(ArgList), ")")
    Else                                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              A_FMT$ = " WHEN @Condition THEN @TruePart "
              For I% = 0 To UBound(ArgList)
                  If I% + 1 <= UBound(ArgList) Then
                     If Trim(ArgList(I% + 1)(0)) = "" Then A_HaveElse% = True
                  End If
                  A_Replace$ = A_FMT$
                  A_Replace$ = Replace(A_Replace$, "@Condition", ArgList(I%)(0), 1, 1, vbTextCompare)
                  A_Replace$ = Replace(A_Replace$, "@TruePart", ArgList(I%)(1), 1, 1, vbTextCompare)
                  A_RetStr$ = A_RetStr$ & A_Replace$
                  If A_HaveElse% Then
                     A_Replace$ = " ELSE @FalsePart "
                     A_Replace$ = Replace(A_Replace$, "@FalsePart", ArgList(I% + 1)(1), 1, 1, vbTextCompare)
                     A_RetStr$ = A_RetStr$ & A_Replace$
                     Exit For
                  End If
              Next I%
              A_RetStr$ = " CASE " & A_RetStr$ & " END "
       End Select
    End If
    If AliasName$ <> "" Then
       A_RetStr$ = A_RetStr$ & " AS " & AliasName$ & " "
    End If
    '
    GetSQLCase = A_RetStr$
End Function

Function GetSQLCharAscii(DB As Database, ByVal CharStr$, Optional ByVal IsString As Boolean = False, Optional ByVal AliasName$, Optional ByVal Options% = True) As String
'���o�r���B�⦡�̥���r��ASCII�Ȫ�SQL���
'�Ѽ� : CharStr$ - ���W�٩Φr���B�⦡
Dim A_FmtStr$

    AliasName$ = Trim(AliasName$)
    '
    If Trim(DB.Connect) = "" Or Not Options% Then   'Access Database
       A_FmtStr$ = " ASC(@CharStr) "
    Else                                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              A_FmtStr$ = " ASCII(@CharStr) "
       End Select
    End If
    If IsString Then A_FmtStr$ = Replace(A_FmtStr$, "@CharStr", "'@CharStr'", 1, -1, vbTextCompare)
    A_FmtStr$ = Replace(A_FmtStr$, "@CharStr", CharStr$, 1, -1, vbTextCompare)
    If AliasName$ <> "" Then A_FmtStr$ = A_FmtStr$ & " AS " & AliasName$ & " "
    '
    GetSQLCharAscii = A_FmtStr$
End Function

'===============================================================================
' Add New Function at 92/8/12
'===============================================================================
Function EncryptConnectStr(ByVal Connect$) As String
'���o�s����Ʈw�r��K�X�r��[�K�᪺�r��
Dim A_SPos%, A_EPos%, I%
Dim A_Pwd$, A_EPwd$

    EncryptConnectStr = Connect$
    
    If Connect$ = "" Then Exit Function

    A_Pwd$ = GetDBConnectPwd(Connect$, A_SPos%, A_EPos%)
    If A_SPos% = 0 Then Exit Function

    A_EPwd$ = StringEncrypt(A_Pwd$)
    
    EncryptConnectStr = Left(Connect$, A_SPos% - 1) & _
                        A_EPwd$ & Mid(Connect$, A_EPos%)
End Function

Function DecryptConnectStr(ByVal Connect$) As String
'���o�s����Ʈw�r��K�X�r��ѱK�᪺�r��
Dim A_SPos%, A_EPos%, I%
Dim A_Pwd$, A_EPwd$

    DecryptConnectStr = Connect$
    
    If Connect$ = "" Then Exit Function

    A_Pwd$ = GetDBConnectPwd(Connect$, A_SPos%, A_EPos%)
    If A_SPos% = 0 Then Exit Function

    A_EPwd$ = Trim(StringDecrypt(A_Pwd$))
    DecryptConnectStr = Left(Connect$, A_SPos% - 1) & _
                        A_EPwd$ & Mid(Connect$, A_EPos%)
End Function

Function GetDBConnectPwd(ByVal Connect$, SPos%, EPos%) As String
'���o�s����Ʈw�r�ꤤ���K�X�r��

    GetDBConnectPwd = ""
    
    SPos% = InStr(1, Connect$, "Pwd=", vbTextCompare)
    If SPos% > 0 Then
        SPos% = SPos% + Len("Pwd=")
    Else
        SPos% = InStr(1, Connect$, "Password=", vbTextCompare)
        If SPos% > 0 Then SPos% = SPos% + Len("Password=")
    End If
    If SPos% = 0 Then Exit Function
    
    EPos% = InStr(SPos%, Connect$, ";", vbTextCompare)
    If EPos% = 0 Then EPos% = Len(Connect$) + 1
    
    GetDBConnectPwd = Mid$(Connect$, SPos%, EPos% - SPos%)
End Function

Function StringEncrypt(ByVal strEncrypt$) As String
'�N�r��[�K
On Local Error GoTo MyError
Dim A_Len%, A_Loop%, I%, j%, A_Start%, A_End%
Dim A_RetStr$, A_Pwd$
Dim A_Value@

    StringEncrypt = strEncrypt$
    
    If strEncrypt$ = "" Then strEncrypt$ = Space(2)
    strEncrypt$ = StrReverse(strEncrypt$)
    For I% = 1 To Len(strEncrypt$)
        If I% = 1 Then
           A_Pwd$ = A_Pwd$ & Format(Asc(Mid(strEncrypt$, I%, 1)), "000")
        Else
           A_Pwd$ = A_Pwd$ & Format(Asc(Mid(strEncrypt$, I%, 1)) + Asc(Mid(strEncrypt$, 1, 1)), "000")
        End If
    Next I%
    A_Pwd$ = StrReverse(A_Pwd$)
        
    A_Len% = Len(A_Pwd$)
    A_Loop% = A_Len% \ 4
    If (A_Len% Mod 4) <> 0 Then A_Loop% = A_Loop% + 1
    For I% = 1 To A_Loop%
        A_Value@ = 0
        A_Start% = (I% - 1) * 4 + 1
        A_End% = I% * 4
        If A_End% > A_Len% Then A_End% = A_Len%
        For j% = A_Start% To A_End%
            A_Value@ = A_Value@ * 203 + (Asc(Mid(Trim(A_Pwd$), j%, 1)) + 23)
        Next j%
        A_RetStr$ = A_RetStr$ & Replace(Format(Hex(CCur(CStr(A_Value@))), "@@@@@@@@"), _
                    " ", "0", , , vbTextCompare)
    Next I%

    StringEncrypt = A_RetStr$
    
    Exit Function
    
MyError:
    StringEncrypt = strEncrypt$
    MsgBox Error, vbExclamation, App.Title
End Function

Function StringDecrypt(ByVal strDecrypt$) As String
'�N�r��ѱK
On Local Error GoTo MyError
Dim I%, A_Code2%
Dim A_Value@, A_Code1@
Dim A_RetStr$, A_TmpStr$, A_TmpStr2$

    For I% = 1 To Len(strDecrypt$) Step 8
        A_TmpStr$ = ""
        A_Value@ = CCur("&H" & Mid(strDecrypt$, I%, 8))
        Do
            A_Code1@ = A_Value@ \ 203
            A_Code2% = A_Value@ Mod 203
            If A_Code2% > 0 Then
                A_TmpStr$ = A_TmpStr$ & Trim(Chr(A_Code2% - 23))
            ElseIf A_Code2% = 0 Then
                A_TmpStr$ = A_TmpStr$ & Trim(Chr(203 - 23))
                A_Code1@ = (A_Value@ - 203) / 203
            End If
            A_Value@ = A_Code1@
        Loop Until A_Value@ = 0
        A_TmpStr2$ = A_TmpStr2$ & StrReverse(A_TmpStr$)
    Next I%
    
    A_TmpStr2$ = StrReverse(A_TmpStr2$)
    For I% = 1 To Len(A_TmpStr2$) Step 3
        If I% = 1 Then
           A_RetStr$ = A_RetStr$ & Chr(Mid(A_TmpStr2$, I%, 3))
        Else
           A_RetStr$ = A_RetStr$ & Chr(CCur(Mid(A_TmpStr2$, I%, 3)) - CCur(Mid(A_TmpStr2$, 1, 3)))
        End If
    Next I%

    StringDecrypt = StrReverse(A_RetStr$)
    
    Exit Function
    
MyError:
    StringDecrypt = strDecrypt$
End Function

Sub DecodingConnectStr(ByVal INIFile$)
'�N��Ʈw�s���r�ꤤ���K�X�ѱK

    'INI�����K�X�O�_�[�K
    If StrComp(GetIniStr("DBPath", "Encrypt", G_INI_SerPath & INIFile$), _
    "True", vbTextCompare) <> 0 Then Exit Sub
    
    '����ѱK
    G_ConnectMethod1 = DecryptConnectStr(G_ConnectMethod1)
    G_ConnectMethod2 = DecryptConnectStr(G_ConnectMethod2)
    G_ConnectMethod3 = DecryptConnectStr(G_ConnectMethod3)
    G_ConnectMethod4 = DecryptConnectStr(G_ConnectMethod4)
    G_ConnectMethod5 = DecryptConnectStr(G_ConnectMethod5)
    G_ConnectMethod6 = DecryptConnectStr(G_ConnectMethod6)
    G_ConnectMethod7 = DecryptConnectStr(G_ConnectMethod7)
    G_ConnectMethod8 = DecryptConnectStr(G_ConnectMethod8)
    G_ConnectMethod9 = DecryptConnectStr(G_ConnectMethod9)
    G_ConnectMethod10 = DecryptConnectStr(G_ConnectMethod10)
End Sub

Sub AutoEncryptINIPwd(ByVal INIFile$)
'�۰ʰ���INI File�����K�X�[�K�ʧ@
Dim A_WinDir$, A_Time$

    '���oWindows�����|
    A_WinDir$ = GetWinDir()

    A_Time$ = "." & Format(Now, "yymmddhhmmss")
    'Local INI�����K�X�[�K
    If StrComp(GetIniStr("DBPath", "Encrypt", INIFile$), "True", vbTextCompare) <> 0 Then
       FileCopy A_WinDir$ & INIFile$, A_WinDir$ & INIFile$ & A_Time$
       OSWritePrivateProfileString% "DBPath", "Connect1", EncryptConnectStr(G_ConnectMethod1), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect2", EncryptConnectStr(G_ConnectMethod2), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect3", EncryptConnectStr(G_ConnectMethod3), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect4", EncryptConnectStr(G_ConnectMethod4), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect5", EncryptConnectStr(G_ConnectMethod5), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect6", EncryptConnectStr(G_ConnectMethod6), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect7", EncryptConnectStr(G_ConnectMethod7), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect8", EncryptConnectStr(G_ConnectMethod8), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect9", EncryptConnectStr(G_ConnectMethod9), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect10", EncryptConnectStr(G_ConnectMethod10), INIFile$
       OSWritePrivateProfileString% "DBPath", "Encrypt", "True", INIFile$
    End If
    
    'Server INI�����K�X�[�K
    INIFile$ = G_INI_SerPath & INIFile$
    If StrComp(GetIniStr("DBPath", "Encrypt", INIFile$), "True", vbTextCompare) <> 0 Then
       FileCopy INIFile$, INIFile$ & A_Time$
       OSWritePrivateProfileString% "DBPath", "Connect1", EncryptConnectStr(G_ConnectMethod1), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect2", EncryptConnectStr(G_ConnectMethod2), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect3", EncryptConnectStr(G_ConnectMethod3), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect4", EncryptConnectStr(G_ConnectMethod4), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect5", EncryptConnectStr(G_ConnectMethod5), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect6", EncryptConnectStr(G_ConnectMethod6), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect7", EncryptConnectStr(G_ConnectMethod7), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect8", EncryptConnectStr(G_ConnectMethod8), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect9", EncryptConnectStr(G_ConnectMethod9), INIFile$
       OSWritePrivateProfileString% "DBPath", "Connect10", EncryptConnectStr(G_ConnectMethod10), INIFile$
       OSWritePrivateProfileString% "DBPath", "Encrypt", "True", INIFile$
    End If
End Sub

'���oWindows�t�θ��|
Function GetWinDir(Optional ByVal rejectBackSlash As Boolean = False) As String
Dim A_WinDir$
Const MAX_PATH = 260

    '���oWindows�����|
    A_WinDir$ = Space(MAX_PATH)
    If GetWindowsDirectory(A_WinDir$, MAX_PATH) > 0 Then
       A_WinDir$ = StripTerminator(Trim$(A_WinDir$)) & IIf(rejectBackSlash, "", "\")
    End If
    GetWinDir = A_WinDir$
End Function

'===============================================================================
' Add New Function at 93/1/15
'===============================================================================
Function CreateTable(DB As Database, ByVal A_TableName$, ByVal A_Flds, A_ErrMsg$) As Integer
'�b��Ʈw���إ߷s���
'Function�Ǧ^�� : Integer (1:���w�s�b True:�إߦ��\ False:�إߥ���)
'�Ѽ� : 1.DB - Database Object Name
'       2.A_TableName$ - �s�����W��
'       3.A_Flds - �s��檺�Ҧ����,�HArray�����A�ǤJ
'       4.A_ErrMsg$ - �^�ǫإߪ�楢�Ѯɪ����~�T��

Dim I%, A_Sql$
Dim A_NewTable As TableDef, A_Fld As Field

    CreateTable = 1
    A_ErrMsg$ = ""
    If IsTableExist(DB, A_TableName$) Then Exit Function
'    On Error Resume Next
'    If Trim(DB.Connect) = "" Then
'       Debug.Print DB.TableDefs(A_TableName$).Name
'    Else
'       Debug.Print DB.TableDefs("dbo." & A_TableName$).Name
'    End If
'    If Err = 0 Then On Error GoTo 0: Exit Function
'    On Error GoTo 0
    
    On Local Error GoTo MyError
    If Trim(DB.Connect) = "" Then   'Access Database
        Set A_NewTable = DB.CreateTableDef(A_TableName$)
    Else
        Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
            Case "SQL;"
                A_Sql$ = "CREATE TABLE " & A_TableName$ & "("
        End Select
    End If
    '
    For I% = 0 To UBound(A_Flds, 2)
        If Trim(DB.Connect) = "" Then   'Access Database
            Select Case CInt(A_Flds(1, I%))
                Case G_Data_String
                    Set A_Fld = A_NewTable.CreateField(A_Flds(0, I%), dbText, A_Flds(2, I%))
                    A_Fld.DefaultValue = """ """
                Case G_Data_Numeric
                    Set A_Fld = A_NewTable.CreateField(A_Flds(0, I%), dbDouble)
                    A_Fld.DefaultValue = 0
            End Select
            A_NewTable.Fields.Append A_Fld
        Else                                 'ODBC Database
            Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
                Case "SQL;"
                    Select Case CInt(A_Flds(1, I%))
                        Case G_Data_String
                            A_Sql$ = A_Sql$ & A_Flds(0, I%) & " VARCHAR(" & A_Flds(2, I%) & ") NOT NULL DEFAULT ' ',"
                        Case G_Data_Numeric
                            A_Sql$ = A_Sql$ & A_Flds(0, I%) & " Numeric(25,4) NOT NULL DEFAULT 0,"
                        Case G_Data_VarBinary 'by Lidia (S021024037)
                            A_Sql$ = A_Sql$ & A_Flds(0, I%) & " VarBinary(Max) NULL,"
                        Case G_Data_uniqueidentifier 'by Lidia (S021024037)
                            A_Sql$ = A_Sql$ & A_Flds(0, I%) & " [uniqueidentifier] NOT NULL,"
                    End Select
            End Select
        End If
    Next I%
    '
    If Trim(DB.Connect) = "" Then   'Access Database
        DB.TableDefs.Append A_NewTable
    Else
        Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
            Case "SQL;"
                A_Sql$ = Left(A_Sql$, Len(A_Sql$) - 1) & ")"
                DB.Execute A_Sql$, dbSQLPassThrough
        End Select
    End If
    '
    CreateTable = True
    Exit Function
    
MyError:
    A_ErrMsg$ = Error$
    CreateTable = False
End Function

Function CreateTableIndex(DB As Database, ByVal A_TableName$, ByVal A_IndexName$, ByVal A_Flds, ByVal A_Primary%, ByVal A_Unique%, ByVal A_Cluster%, A_ErrMsg$) As Integer
'�إ߸�Ʈw��椤������
'Function�Ǧ^�� : Integer (1:���ޤw�s�b True:�إߦ��\ False:�إߥ���)
'�Ѽ� : 1.DB - Database Object Name
'       2.A_TableName$ - �s�����W��
'       3.A_Flds - �s��檺�Ҧ����,�HArray�����A�ǤJ
'       4.A_ErrMsg$ - �^�ǫإߪ�楢�Ѯɪ����~�T��

Dim I%, A_Sql$
Dim A_Table As TableDef, A_Index As index, A_IndexFld As Field

    CreateTableIndex = 1
    A_ErrMsg$ = ""
    If IsIndexExist(DB, A_TableName$, A_IndexName$) Then Exit Function
'    On Error Resume Next
'    If Trim(DB.Connect) = "" Then
'       Debug.Print DB.TableDefs(A_TableName$).Indexes(A_IndexName$).Name
'    Else
'       Debug.Print DB.TableDefs("dbo." & A_TableName$).Indexes(A_IndexName$).Name
'    End If
'    If Err = 0 Then On Error GoTo 0: Exit Function
'    On Error GoTo 0
    '
    On Local Error GoTo MyError
    If Trim(DB.Connect) = "" Then   'Access Database
        Set A_Table = DB.TableDefs(A_TableName$)
    Else
        Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
            Case "SQL;"
                A_Sql$ = "CREATE"
                If A_Unique% Then A_Sql$ = A_Sql$ & " UNIQUE"
                If A_Cluster% Then A_Sql$ = A_Sql$ & " CLUSTERED"
                A_Sql$ = A_Sql$ & " INDEX " & A_IndexName$
                A_Sql$ = A_Sql$ & " ON " & A_TableName$ & " ("
        End Select
    End If
    '
    For I% = 0 To UBound(A_Flds)
        If Trim(DB.Connect) = "" Then   'Access Database
            If I% = 0 Then
                Set A_Index = A_Table.CreateIndex(A_IndexName$)
            End If
            Set A_IndexFld = A_Index.CreateField(A_Flds(I%))
            A_Index.Fields.Append A_IndexFld
        Else                                 'ODBC Database
            A_Sql$ = A_Sql$ & A_Flds(I%) & ","
        End If
    Next I%
    '
    If Trim(DB.Connect) = "" Then   'Access Database
        A_Index.Primary = A_Primary%
        A_Index.Unique = A_Unique%
        A_Table.Indexes.Append A_Index
    Else
        Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
            Case "SQL;"
                A_Sql$ = Left(A_Sql$, Len(A_Sql$) - 1) & ")"
                DB.Execute A_Sql$, dbSQLPassThrough
        End Select
    End If
    '
    CreateTableIndex = True
    Exit Function
    
MyError:
    A_ErrMsg$ = Error$
    CreateTableIndex = False
End Function

Sub AddArrayItem(A_Flds, ByVal A_FldName$, ByVal A_FldType%, Optional ByVal A_FldLen% = -1)
'�ʺA�[�J�G���}�C�������
'�Ѽ� : 1.A_Flds - �^�ǰʺA�}�C
'       2.A_FldName$ - ���W��
'       3.A_FldType% - ��쫬�A
'       4.A_FldLen% - ��r��쪺����,�ƭ����i����J
Dim A_MaxCol%

    On Local Error Resume Next
    A_MaxCol% = UBound(A_Flds, 2) + 1
    If Err <> 0 Then A_MaxCol% = 0
    On Error GoTo 0
    
    ReDim Preserve A_Flds(2, A_MaxCol%)
    A_Flds(0, A_MaxCol%) = A_FldName$
    A_Flds(1, A_MaxCol%) = A_FldType%
    A_Flds(2, A_MaxCol%) = A_FldLen%
End Sub


Function GetPrgFunAuth(ByVal A_EmployeeID$, ByVal A_PgmID$, A_Read As Boolean, A_Edit As Boolean, A_Delete As Boolean, A_Add As Boolean, A_PRINT As Boolean)
'���o�U�{���U�s�W,�R��,�ק�,Ū���ΦC�L���v��
'�ѼơG1.A_EmployeeID$ - �ǤJ�ثe�ϥΪ̪����u�s��
'      2.A_Read        - �^��Ū���v��(True/False)
'      3.A_Edit        - �^�ǭק��v��(True/False)
'      4.A_Delete      - �^�ǧR���v��(True/False)
'      5.A_Add         - �^�Ƿs�W�v��(True/False)
'      6.A_Print       - �^�ǦC�L�v��(True/False)
On Local Error GoTo MY_Error
Dim A_Sql$, DY As Recordset, A_A0801$

    A_Read = True: A_Edit = True: A_Delete = True: A_Add = True: A_PRINT = True
    '
    If UCase(GetSvrINIStrA(DB_ARTHGUI, "Button_Authorize", Trim(A_PgmID$))) <> "Y" Then Exit Function
    '
    A_Sql$ = "Select * from A47 "
    A_Sql$ = A_Sql$ & " Where A4701='" & Trim(A_EmployeeID$) & "'"
    A_Sql$ = A_Sql$ & " And A4702='" & Trim(A_PgmID$) & "'"
    CreateDynasetODBC DB_ARTHGUI, DY, A_Sql$, "DY", True
    '
    If Not (DY.EOF And DY.BOF) Then
        A_Read = IIf(UCase(Trim(DY.Fields("A4703") & "")) = "Y", True, False)
        A_Edit = IIf(UCase(Trim(DY.Fields("A4704") & "")) = "Y", True, False)
        A_Delete = IIf(UCase(Trim(DY.Fields("A4705") & "")) = "Y", True, False)
        A_Add = IIf(UCase(Trim(DY.Fields("A4706") & "")) = "Y", True, False)
        A_PRINT = IIf(UCase(Trim(DY.Fields("A4707") & "")) = "Y", True, False)
    End If
    '
    Exit Function
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

Sub GetPgmFieldAuth(ByVal A_EmployeeID$, ByVal A_PgmID$, A_PgmFld$())
'���o�{���U�U��쪺���v����
'�Ѽ� : 1.A_EmployeeID$ - �ǤJ�ثe�ϥΪ̪����u�s��
'       2.A_PgmID$      - �ǤJ�{���N�X
'       3.A_PgmFld$     - �^�����}�C,A_PgmFld$(1,?)=���W��,A_PgmFld$(2,?)=�ק��v��(Y/N),A_PgmFld$(3,?)=�O�_���(Y/N)
On Local Error GoTo MY_Error
Dim A_Sql$, DY As Recordset, a_count%
Dim A_ReadAuth As Boolean, A_EditAuth As Boolean, A_DelAuth As Boolean, A_AddAuth As Boolean, A_PrintAuth As Boolean
    
    ReDim A_PgmFld$(1 To 5, 0)
    '
    If UCase(GetSvrINIStrA(DB_ARTHGUI, "Button_Authorize", Trim(A_PgmID$))) <> "Y" Then Exit Sub
    GetPrgFunAuth A_EmployeeID$, A_PgmID$, A_ReadAuth, A_EditAuth, A_DelAuth, A_AddAuth, A_PrintAuth
    '
    A_Sql$ = "Select * from A52"
    A_Sql$ = A_Sql$ & " Where A5201='" & Trim(A_EmployeeID$) & "'"
    A_Sql$ = A_Sql$ & " And A5202='" & Trim(A_PgmID$) & "'"
    A_Sql$ = A_Sql$ & " Order by A5201,A5202,A5203"
    CreateDynasetODBC DB_ARTHGUI, DY, A_Sql$, "DY", True
    '
    a_count% = 0
    '
    Do While Not DY.EOF
        a_count% = a_count% + 1
        ReDim Preserve A_PgmFld$(1 To 5, a_count%)
        '
        A_PgmFld(1, a_count%) = Trim(DY.Fields("A5203") & "")
        If A_EditAuth = True Then
            A_PgmFld(2, a_count%) = Trim(DY.Fields("A5204") & "")
        Else
            A_PgmFld(2, a_count%) = "N"
        End If
        A_PgmFld(3, a_count%) = Trim(DY.Fields("A5205") & "")
        '
        DY.MoveNext
    Loop
    '
    Exit Sub
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Sub GetFieldAuth(A_PgmFld$(), ByVal A_FieldName$, A_Edit As Boolean, A_Show As Boolean)
'���o�ǤJ��쪺�ק������v��
'�ѼơG1.A_PgmFld$()  - �ǤJ���}�C,A_PgmFld$(1,?)=���W��,A_PgmFld$(2,?)=�ק��v��(Y/N),A_PgmFld$(3,?)=�O�_���(Y/N)
'      2.A_FieldName$ - �ǤJ��諸���W��
'      3.A_Edit       - �^�ǭק��v��(True/False)
'      4.A_Show       - �^������v��(True/False)
Dim I%
            
    A_Edit = True: A_Show = True
    '
    For I% = 1 To UBound(A_PgmFld$, 2)
        If UCase(Trim(A_PgmFld$(1, I%))) > UCase(Trim(A_FieldName$)) Then Exit For
        If UCase(Trim(A_PgmFld$(1, I%))) = UCase(Trim(A_FieldName$)) Then
            A_Edit = IIf(UCase(Trim(A_PgmFld$(2, I%))) = "Y", True, False)
            A_Show = IIf(UCase(Trim(A_PgmFld$(3, I%))) = "Y", True, False)
            Exit For
        End If
    Next I%
End Sub

Function SaveFieldCheck(A_PgmFld$(), ByVal A_FieldName$) As Boolean
'�ˮֶǤJ�����O�_�i��s��
'�ѼơG1.A_PgmFld$()  - �ǤJ���}�C,A_PgmFld$(1,?)=���W��,A_PgmFld$(2,?)=�ק��v��(Y/N),A_PgmFld$(3,?)=�O�_���(Y/N)
'      2.A_FieldName$ - �ǤJ��諸���W��
'      3.�^�ǬO�_�s��(True/False)
Dim A_Edit As Boolean, A_Show As Boolean
            
    GetFieldAuth A_PgmFld$, A_FieldName$, A_Edit, A_Show
    '
    If A_Edit = True And A_Show = True Then
        SaveFieldCheck = True
    Else
        SaveFieldCheck = False
    End If
End Function


Sub SetFieldStatus(A_PgmFld$(), ByVal A_FieldName$, ByVal A_PgmEditAuth As Boolean, Control As Control, A_Show As Boolean)
'�̶ǤJ��쪺���v����,�]�w���I���C��PEnable
'�ѼơG1.A_PgmFld$()  - �ǤJ���}�C,A_PgmFld$(1,?)=���W��,A_PgmFld$(2,?)=�ק��v��(Y/N),A_PgmFld$(3,?)=�O�_���(Y/N)
'      2.A_FieldName$ - �ǤJ�]�w�����W��
'      3.A_PgmEditAuth- �ǤJ�ק�\�ध�v��
'      4.A_Show       - �^�ǬO�_���(True/False)
Dim A_Edit As Boolean
Dim A_ReadAuth As Boolean, A_EditAuth As Boolean, A_DelAuth As Boolean, A_AddAuth As Boolean, A_PrintAuth As Boolean

    GetFieldAuth A_PgmFld$, A_FieldName$, A_Edit, A_Show
    '
    If A_Show = False Then
        If TypeOf Control Is TextBox Or TypeOf Control Is MaskEdBox Or _
           TypeOf Control Is ComboBox Then
            Control.BackColor = G_Label_Color
            Control.ForeColor = G_Label_Color
            Control.Enabled = False
        End If
    Else
        If TypeOf Control Is TextBox Or TypeOf Control Is MaskEdBox Or _
           TypeOf Control Is ComboBox Then
            Control.BackColor = G_TextLostBack_Color
            Control.ForeColor = G_TextLostFore_Color
            If G_AP_STATE = G_AP_STATE_DELETE Then
                Control.Enabled = False
            Else
                Control.Enabled = IIf(A_PgmEditAuth = False, False, A_Edit)
            End If
        End If
    End If
End Sub

Sub SetSpdFldStatus(A_PgmFld$(), ByVal A_FieldName$, Control As Control, ByVal A_Row As Long, ByVal A_Col As Long, A_Show As Boolean)
'�̶ǤJ��쪺���v����,�]�wSpread Cell �I���C��PLock
'�ѼơG1.A_PgmFld$()  - �ǤJ���}�C,A_PgmFld$(1,?)=���W��,A_PgmFld$(2,?)=�ק��v��(Y/N),A_PgmFld$(3,?)=�O�_���(Y/N)
'      2.A_FieldName$ - �ǤJ�]�w�����W��
'      3.A_Row        - �ǤJ�]�w���C��
'      4.A_Col        - �ǤJ�]�w������
'      3.A_Show       - �^�ǬO�_���(True/False)
Dim A_Edit As Boolean
        
    GetFieldAuth A_PgmFld$, A_FieldName$, A_Edit, A_Show
    '
    If A_Show = False Then
        If TypeOf Control Is FPSPREAD.vaSpread Then
            If Control.MaxRows > 0 Then
                Control.Row = A_Row
                Control.Col = A_Col
                Control.BackColor = G_Label_Color
                Control.ForeColor = G_Label_Color
            End If
            '
            Control.Row = A_Row
            Control.Col = A_Col
            Control.Lock = True
        End If
    Else
        If TypeOf Control Is FPSPREAD.vaSpread Then
            If Control.MaxRows > 0 Then
                Control.Row = A_Row
                Control.Col = A_Col
                Control.BackColor = G_TextLostBack_Color
                Control.ForeColor = G_TextLostFore_Color
            End If
            '
            Control.Row = A_Row
            Control.Col = A_Col
            If G_AP_STATE = G_AP_STATE_DELETE Then
                Control.Lock = True
            Else
                Control.Lock = Not A_Edit
            End If
        End If
    End If
End Sub

'===============================================================================
' Add New Function at 93/3/8
'===============================================================================
Function GetExcelAppPath() As String
'���oExcel Application�������ɦW
Const A_MainKey$ = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe"
Dim A_Path$

    A_Path$ = GetRegSetting(A_MainKey$, "", "", "")
    GetExcelAppPath = A_Path$
End Function

'===============================================================================
' Add New Function at 93/3/22 by Anita
'===============================================================================
 Function Check_CloseDate(ByVal DB_Source As Database, ByVal CompanyID$, ByVal SystemID$, ByVal A_Date$, ByRef A_ErrCompanyID$) As Boolean
'�ˮֻs��or�ǲ�����O�_�j�����b��
'DB_Source-�ӷ���Ʈw
'CompanyID-���q�O
'SystemID-�t�ΧO
'A_Date-�ǲ�or�s�����(��DateOut�᪺�榡)
'A_ErrCompanyID-�Ǧ^�ˮֿ��~�����q�O
On Local Error GoTo MY_Error
Dim A_Sql$, CloseDate$

    Check_CloseDate = False
    
    If Trim(A_Date$) = "" Then Check_CloseDate = True: Exit Function
    '���q�O
    CloseDate$ = GetSvrINIStrA(DB_Source, "CloseDate", "Date_" & Trim(CompanyID$))
    If Trim(CloseDate$) = "" Then CloseDate$ = GetSvrINIStrA(DB_Source, "CloseDate", "Date")
    If Trim(CloseDate$) <> "" Then
        If Val(DateIn(A_Date$)) <= Val(CloseDate$) Then
            A_ErrCompanyID$ = Trim(CompanyID$)
            Exit Function
        End If
    End If
    
    '�t�ΧO
    If Trim(CompanyID$) <> Trim(SystemID$) And Trim(SystemID$) <> "" Then
        CloseDate$ = GetSvrINIStrA(DB_Source, "CloseDate", "Date_" & Trim(SystemID$))
        If Trim(CloseDate$) = "" Then CloseDate$ = GetSvrINIStrA(DB_Source, "CloseDate", "Date")
        If Trim(CloseDate$) <> "" Then
            If Val(DateIn(A_Date$)) <= Val(CloseDate$) Then
                A_ErrCompanyID$ = Trim(SystemID$)
                Exit Function
            End If
        End If
    End If
    '
    Check_CloseDate = True
    Exit Function
    
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function



'===============================================================================
' Edit Function at 93/4/1
'===============================================================================
Sub SetTSpdText(tSPD As Spread, ByVal FldName$, ByVal Value$, Optional ByVal RTrimSpace% = False)
'�]�wSpread Type�W���Y�����W�٪���
Dim A_Col#

    '�H���W�٨��o��쪺Index
    A_Col# = GetSpdColIndex(tSPD, FldName$)
    If A_Col# = 0 Then Exit Sub
    
    '�]�w����
    If G_PrintSelect = G_Print2Excel Or G_PrintSelect = G_Print2Word Then
       If RTrimSpace% Then Value$ = RTrim(Value$)
    End If
    tSPD.Columns(A_Col#).text = Value$
End Sub

Sub CreateDynasetODBC(DB As Database, DY As Recordset, ByVal SQL$, ByVal Str$, ByVal Options%, Optional ByVal Ignore%)
'�}��Recordset
On Local Error GoTo CreateDynasetODBC_Error
Dim A_Connect$, A_Msg$, A_Msg1$, A_Msg2$, A_Msg3$, A_Msg4$, A_Msg5$
Dim A_TryCount%, A_KeepMsg$
    
    A_Msg1$ = GetSIniStr("PanelDescpt", "unread")       '"��Ʈw�ثe�L�kŪ�g�A�еy�ݢ������U�T�w���~��,"
    A_Msg2$ = GetSIniStr("PanelDescpt", "cancel")       '"�Ϋ��U�����䵲�����\��!!"
    A_Msg3$ = GetSIniStr("PanelDescpt", "datachange")   '"��Ʈw���ʤ�,�ثe�L�kŪ�g�A�еy�ݢ������U�T�w���~��,"
    A_Msg4$ = GetSIniStr("PanelDescpt", "dataerror")    '"��ƮwŪ�g�o�Ϳ��~�A�{���N����!"
    A_Msg5$ = GetSIniStr("PanelDescpt", "writeerror")   '"�бN�����~�T���O�U�A�P�{���H���p��!"
    
    CloseOpen DY, Str$
    '
    A_Connect$ = IIf(UCase$(G_SystemID) = "ARTHGUI", G_ConnectMethod2, G_ConnectMethod1)
    '
    If Trim$(DB.Connect) = "" Then
       If Options% Then
          Set DY = DB.OpenRecordset(SQL$, dbOpenSnapshot)
       Else
          Set DY = DB.OpenRecordset(SQL$, dbOpenDynaset)
          DY.LockEdits = False
       End If
    Else
       Select Case UCase$(Mid$(A_Connect$, InStr(1, A_Connect$, "DBTYPE=", 1) + 7))
         Case "SQL;", "ORACLE;"
              Select Case Options%
                Case True
                     Set DY = DB.OpenRecordset(SQL$, dbOpenSnapshot, dbSQLPassThrough)
                     If Not Ignore% Then
                        If Not (DY.BOF And DY.EOF) Then DY.MoveLast: DY.MoveFirst
                     End If
                Case False
                     Set DY = DB.OpenRecordset(SQL$, dbOpenDynaset)
                     DY.LockEdits = False
              End Select
         Case "DB2;"
              Select Case Options%
                Case True
                     Set DY = DB.OpenRecordset(SQL$, dbOpenSnapshot)
                Case False
                     Set DY = DB.OpenRecordset(SQL$, dbOpenDynaset)
              End Select
       End Select
       If A_TryCount% > 0 Then
          SetFormMsgLineText A_KeepMsg$
          A_Msg$ = vbCrLf & String(10, "*") & Chr(vbKeyTab)
          A_Msg$ = A_Msg$ & GetCaption("PgmMsg", "deadlock_success", "���հ���@�~���\!")
          WriteErrorReport A_Msg$, ""
       End If
    End If
    Exit Sub

CreateDynasetODBC_Error:
    
    Select Case Err
      Case 3046, 3158, 3186, 3187, 3188, 3202, 3218, 3260  'Record Locked
           Idle
           A_Msg$ = Error(Err) & vbCrLf
           A_Msg$ = A_Msg$ & A_Msg1$
           A_Msg$ = A_Msg$ & vbCrLf
           A_Msg$ = A_Msg$ & A_Msg2$
           retcode = MsgBox(A_Msg$, vbOKCancel + vbQuestion, UCase$(App.Title))
           Err = 0
           Screen.ActiveForm.Refresh
           WriteErrorReport A_Msg$, SQL$
           If retcode = IDOK Then Resume
           If retcode = IDCANCEL Then CloseFileDB: End
           
      Case 3167, 3197                                      'Record is deleted , changed.
           A_Msg$ = Error(Err) & vbCrLf
           A_Msg$ = A_Msg$ & A_Msg3$
           A_Msg$ = A_Msg$ & vbCrLf
           A_Msg$ = A_Msg$ & A_Msg2$
           retcode = MsgBox(A_Msg$, vbOKCancel + vbQuestion, UCase$(App.Title))
           Err = 0
           Screen.ActiveForm.Refresh
           WriteErrorReport A_Msg$, SQL$
           If retcode = IDOK Then Resume
           If retcode = IDCANCEL Then CloseFileDB: End
           
      Case 3146    'ODBC CALL FAIL
           A_Msg$ = GetODBCErrorMessage()
           If InStr(1, A_Msg$, "1205:") > 0 Then     'Dead Lock Process
              retcode = IDOK
              A_TryCount% = A_TryCount% + 1
              If A_TryCount% = 1 Then
                 A_KeepMsg$ = GetFormMsgLineText()
                 A_Msg$ = A_Msg$ & vbCrLf
                 A_Msg$ = A_Msg$ & GetCaption("PgmMsg", "deadlock_occur", "�{���N�~����հ��楢�Ѫ����O.")
                 WriteErrorReport A_Msg$, SQL$
              ElseIf A_TryCount% Mod 3 = 1 Then
                 A_Msg$ = A_Msg$ & vbCrLf
                 A_Msg$ = A_Msg$ & GetCaption("PgmMsg", "deadlock_tryagain", "�O�_�~����հ��楢�Ѫ����O?")
                 retcode = MsgBox(A_Msg$, vbOKCancel + vbQuestion, UCase$(App.Title))
                 Screen.ActiveForm.Refresh
              End If
              If retcode = IDOK Then
                 SetFormMsgLineText Replace(GetCaption("PgmMsg", "deadlock_occur2", "�s����Ƶo�ͽĬ�,���խ��s�����@��....."), "@", CStr(A_TryCount%))
                 A_Msg$ = vbCrLf & String(10, "*") & Chr(vbKeyTab)
                 A_Msg$ = A_Msg$ & Replace(GetCaption("PgmMsg", "deadlock_trytime", "���հ����@��....."), "@", CStr(A_TryCount%))
                 WriteErrorReport A_Msg$, ""
                 Err = 0
                 Sleep Int(5 * Rnd + 1) * 1000
                 Resume
              ElseIf retcode = IDCANCEL Then
                 A_Msg$ = vbCrLf & String(10, "*") & Chr(vbKeyTab)
                 A_Msg$ = A_Msg$ & GetCaption("PgmMsg", "deadlock_failed", "���հ���@�~�̲ץ���!")
                 WriteErrorReport A_Msg$, ""
                 CloseFileDB
                 End
              End If
           Else
              MsgBox A_Msg$, vbCritical, UCase$(App.Title)
              Err = 0
              WriteErrorReport A_Msg$, SQL$
              CloseFileDB
              End
           End If
           
      Case Else
           A_Msg$ = Error(Err) & vbCrLf
           A_Msg$ = A_Msg$ & A_Msg4$
           A_Msg$ = A_Msg$ & A_Msg5$
           MsgBox A_Msg$, vbCritical, UCase$(App.Title)
           Err = 0
           WriteErrorReport A_Msg$, SQL$
           CloseFileDB
           End
    End Select
End Sub

Sub SetDefaultFileName(Txt As TextBox, ByVal PrtType%)
'�]�w��r,Excel,Word�C�L���w�]�ɦW
Dim A_Value$, A_Str1$, A_ExtName$

    If Txt.Visible = False Then Txt.Visible = True
    
    A_Value$ = Trim(Txt)
    A_ExtName$ = IIf(PrtType% = G_Print2File, ".TXT", _
                 IIf(PrtType% = G_Print2Excel, ".XLS", ".DOC"))
    
    If A_Value$ = "" Then
        Txt = G_System_Path & "TMP\" & App.EXEName & A_ExtName$
    Else
        StrCut A_Value$, ".", A_Str1$, ""
        Txt = A_Str1$ & A_ExtName$
    End If
End Sub

'===============================================================================
' Add New Function at 93/4/1
'===============================================================================
Function GetWordAppPath() As String
'���oExcel Application�������ɦW
Const A_MainKey$ = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinWord.exe"
Dim A_Path$

    A_Path$ = GetRegSetting(A_MainKey$, "", "", "")
    GetWordAppPath = A_Path$
End Function

Private Function SetFormMsgLineText(ByVal MsgText$)
On Error Resume Next

    Screen.ActiveForm.Sts_MsgLine.Panels(1).text = MsgText$
    Screen.ActiveForm.Refresh
End Function

Private Function GetFormMsgLineText() As String
On Error Resume Next
Dim A_KeepMsg$

    A_KeepMsg$ = Screen.ActiveForm.Sts_MsgLine.Panels(1).text
    GetFormMsgLineText = A_KeepMsg$
End Function


'===============================================================================
' Add New Function at 93/4/14 by cathy
'===============================================================================
Sub WriteJournalLog_Security(DB As Database, ByVal State%, ByVal PgmId$, ByVal Memo$, Optional A_User$ = "SecManager")
'�g�J�{���ϥΪ��p��A09
    'S020527055
'    G_Str = "INSERT INTO A09 VALUES ("
    G_Str = ""
    InsertFields "A0901", GetCurrentDate(), G_Data_String
    InsertFields "A0902", GetCurrentTime(), G_Data_String
    InsertFields "A0903", GetWorkStation(), G_Data_String
    InsertFields "A0904", A_User$, G_Data_String
    InsertFields "A0905", " ", G_Data_String
    InsertFields "A0906", PgmId$, G_Data_String
    InsertFields "A0907", State%, G_Data_String
    InsertFields "A0908", " ", G_Data_String
    InsertFields "A0909", A_User$, G_Data_String
    InsertFields "A0910", " ", G_Data_String
    InsertFields "A0911", G_SystemID, G_Data_String
    InsertFields "A0912", GetLenStr(Memo$, 1, 50), G_Data_String
'    G_Str = Left$(G_Str, Len(G_Str) - 1) & ")"
    SQLInsert DB, "A09"
End Sub


'===============================================================================
' Add New Function at 93/6/1 by cathy
'===============================================================================
Sub WriteOverLenStr2SINI(DB As Database, ByVal A_Section$, ByVal A_FieldStr$, Optional ByVal A_Limit% = 40)
'�ت�:�ѨM�ƫ�}��User��J�����Ʈw���ת������,�b���ʸ�Ʈw�[�c��,�N��Ʀs�JSINI
'1.�Y��J��A_TopicValue$����<=A_Limit%�h���B�z
'2.�NA_TopicValue$�C50�Ӧr�����@��g�JSINI
'3.TOPIC=��ƧǸ�
'4.TOPICVALUE=�Ҥ��Ϊ��r��

Dim A_STR$, A_Line%, A_Len%, A_Topic$, A_TopicValue$, A_WordCnt&
    
    If lstrlen(A_FieldStr$) <= A_Limit% Then Exit Sub
    '
    A_STR$ = GetLenStr(A_FieldStr$, 1, 50)
    '�קK���ƶW�L10��,ORDER BY TOPIC��,�|���ƧǤW���D(EX:1,10,2,3,.....)�G�վ�̹�ڵ��Ʈ榡�ƧǸ�(EX:01,02,03....,10)
    A_WordCnt& = UBound(GetTextMultiOutput(A_FieldStr$, 50))
    If A_WordCnt& < 10 Then A_WordCnt& = 1
    If A_WordCnt& >= 10 And A_WordCnt& < 100 Then A_WordCnt& = 2
    If A_WordCnt& >= 100 And A_WordCnt& < 1000 Then A_WordCnt& = 3
    If A_WordCnt& >= 1000 And A_WordCnt& < 10000 Then A_WordCnt& = 4
    If A_WordCnt& >= 10000 And A_WordCnt& < 100000 Then A_WordCnt& = 5
    
    A_Len% = 0
    A_Line% = 0
    '
    Do While A_STR$ <> ""
        A_Line% = A_Line% + 1
        '�̹�ڵ��Ʈ榡�ƧǸ�
        'A_Topic$ = Format(A_Line%, "0")
        A_Topic$ = Format(A_Line%, String(A_WordCnt&, "0"))
        A_TopicValue$ = A_STR$
        GoSub MoveData2Sini
        '
        A_Len% = A_Len% + lstrlen(A_STR$)
        A_STR$ = GetLenStr(A_FieldStr$, A_Len% + 1, 50)
    Loop
    '
    Exit Sub
    
MoveData2Sini:
    G_Str = "DELETE FROM Sini Where Section ='" & Trim(A_Section$) & "'"
    G_Str = G_Str & " And Topic='" & Trim(A_Topic$) & "'"
    ExecuteProcess DB, G_Str
    '
    G_Str = ""
    InsertFields "Section", Trim(A_Section$), G_Data_String
    InsertFields "Topic", Trim(A_Topic$), G_Data_String
    InsertFields "TopicValue", Trim(A_TopicValue$), G_Data_String
    SQLInsert DB, "SINI"
    '
    Return
End Sub
'===============================================================================
' Add New Function at 93/12/27 by Yvonne
'===============================================================================
Sub Prepare_POTAXDeductType(DB As Database, cbo As ComboBox, ByVal A_Type$)
'�ت�:����N���쬰1.�i���褧�i�f�ζO��
'                 2.�i���褧�T�w�겣
'                ��.���i����i������
'             �W�[3.���i���褧�i�f�ζO��
'                 4.���i���褧�T�w�겣
'�ǤJ��Database�����OGENIE
On Local Error GoTo MY_Error
Dim A_Sql$, DY As Recordset

    cbo.Clear
    
    A_Sql$ = "Select * From SINI"
    A_Sql$ = A_Sql$ & " Where Section='POTAXDeductType'"
    A_Sql$ = A_Sql$ & " Order By Topic"
    CreateDynasetODBC DB, DY, A_Sql$, "DY", True
    If Not (DY.EOF And DY.BOF) Then
        Do While Not DY.EOF
            cbo.AddItem Trim(DY.Fields("Topic") & "") & Space(1) & Trim(DY.Fields("TopicValue") & "")
            DY.MoveNext
        Loop
    Else
        G_Str = ""
        InsertFields "Section", "POTAXDeductType", G_Data_String
        InsertFields "Topic", "1", G_Data_String
        InsertFields "TopicValue", GetCaption("POTAXDeductType", "1", "�i���褧�i�f�ζO��"), G_Data_String
        SQLInsert DB, "SINI"
        cbo.AddItem "1" & Space(1) & GetCaption("POTAXDeductType", "1", "�i���褧�i�f�ζO��")
        
        G_Str = ""
        InsertFields "Section", "POTAXDeductType", G_Data_String
        InsertFields "Topic", "2", G_Data_String
        InsertFields "TopicValue", GetCaption("POTAXDeductType", "2", "�i���褧�T�w�겣"), G_Data_String
        SQLInsert DB, "SINI"
        cbo.AddItem "2" & Space(1) & GetCaption("POTAXDeductType", "2", "�i���褧�T�w�겣")
        
        G_Str = ""
        InsertFields "Section", "POTAXDeductType", G_Data_String
        InsertFields "Topic", "3", G_Data_String
        InsertFields "TopicValue", GetCaption("POTAXDeductType", "3", "���i���褧�i�f�ζO��"), G_Data_String
        SQLInsert DB, "SINI"
        cbo.AddItem "3" & Space(1) & GetCaption("POTAXDeductType", "3", "���i���褧�i�f�ζO��")
        
        G_Str = ""
        InsertFields "Section", "POTAXDeductType", G_Data_String
        InsertFields "Topic", "4", G_Data_String
        InsertFields "TopicValue", GetCaption("POTAXDeductType", "4", "���i���褧�T�w�겣"), G_Data_String
        SQLInsert DB, "SINI"
        cbo.AddItem "4" & Space(1) & GetCaption("POTAXDeductType", "4", "���i���褧�T�w�겣")
    End If
    '�������O���ťծ�Default 3 ���i���褧�i�f�ζO��
    If Trim(A_Type$) = "" Then A_Type$ = "3"
    CboStrCut cbo, Trim(A_Type$), Space(1)
    Exit Sub
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub
'===============================================================================
' Add New Function at 93/12/27 by Yvonne
'===============================================================================
Function Get_ZeroTaxCustomerType(DB As Database) As String
'�ت�:���oIVSETUP���s�|�v�q���覡���]�w,�Y���]�w�w�]��'2:�D�g����'
'�ǤJ��Database�����OGENIE
On Local Error GoTo MY_Error
Dim A_Sql$, DY As Recordset

    A_Sql$ = "Select TopicValue From SINI"
    A_Sql$ = A_Sql$ & " Where Section='Customer'"
    A_Sql$ = A_Sql$ & " And Topic='ZeroTaxCustomerType'"
    CreateDynasetODBC DB, DY, A_Sql$, "DY", True
    If Not (DY.EOF And DY.BOF) Then
        Get_ZeroTaxCustomerType = Trim(DY.Fields("TopicValue") & "")
    Else
        Get_ZeroTaxCustomerType = "2"
    End If
    Exit Function
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

'===============================================================================
' Add New Function at 93/6/1 by cathy
'===============================================================================
'S010509042 �W�[A_PriorityC01%,A_C0101$,A_C0102$�ܼơAKeep�O�_�n��C01���
'Function GetFldStrFromSINI(DB As Database, ByVal A_Section$, ByVal A_FldStr$) As String
Function GetFldStrFromSINI(DB As Database, ByVal A_Section$, ByVal A_FldStr$, Optional ByVal A_PriorityC01% = False, Optional ByVal A_C0101$ = "", Optional ByVal A_C0102$ = "") As String

'�ت�:�f�tFunction:'WriteOverLenStr2SINI'�ϥ�
'1.�YSINI�������ȫh�H������
'2.�Y�L�h�H��A_FldStr$�^��
Dim DY As Recordset, A_Sql$

    GetFldStrFromSINI = ""

    'S010509042 Check���ҥ�C01, ���h��SINI-------------------------------------------------------------------------------------
    If UCase(GetSvrINIStrA(DB_ARTHGUI, "Customer", "C01")) = "Y" Then
        Dim A_Type$, A_C01Fldstr$
        A_Type$ = GetSvrINIStrA(DB_ARTHGUI, "A16_Property", Trim(A_C0102$))
        StrCut A_Type$, ",", A_Type$, ""
        If UCase(A_Type$) = "T" Then
            A_C01Fldstr$ = "C0103"
        Else
            A_C01Fldstr$ = "C0104"
        End If
    
        A_Sql$ = "SELECT * FROM C01"
        A_Sql$ = A_Sql$ + " Where C0101='" & Trim(A_C0101$) & "'"
        A_Sql$ = A_Sql$ + " And C0102='" & Trim(A_C0102$) & "'"
        CreateDynasetODBC DB_ARTHGUI, DY, A_Sql$, "DY", True
        
        If Not (DY.BOF And DY.EOF) Then GetFldStrFromSINI = GetFldStrFromSINI & Trim(DY.Fields(A_C01Fldstr$) & "")
        If GetFldStrFromSINI <> "" Then Exit Function
    End If
    '---------------------------------------------------------------------------------------------------------------------------
    '
    'S010605056 �Τ@�s���H��L�O�����u��
    If Trim(A_C0102$) = G_A1609uninumber$ Then
        GetFldStrFromSINI = A_FldStr$
        Exit Function
    End If
    
    A_Sql$ = "SELECT * FROM SINI"
    A_Sql$ = A_Sql$ + " Where SECTION='" & Trim(A_Section$) & "'"
    A_Sql$ = A_Sql$ + " ORDER BY SECTION,TOPIC,TOPICVALUE"
    CreateDynasetODBC DB, DY, A_Sql$, "DY", True
    '
    If Not (DY.BOF And DY.EOF) Then
        Do While Not DY.EOF
            GetFldStrFromSINI = GetFldStrFromSINI & Trim(DY.Fields("TopicValue") & "")
            DY.MoveNext
        Loop
        
    End If
    '
    If GetFldStrFromSINI <> "" Then Exit Function
    '
    GetFldStrFromSINI = A_FldStr$
End Function

'===============================================================================
' Add New Function at 93/6/1 by cathy
'===============================================================================
Function GetMultiLine2StrArray(ByVal A_EngStr$, ByVal A_MaxLen%) As String()
'�N�^��r��̧���r�����Φh��Keep��Array��,��ܩΦL���
'**********************************************************************
'Function �ޥΤ��d�ҵ{��,�ǤJ��ӰѼ�
'A_EngStr$ : �^��r��,�p�^����W,�^��a�}   A_MaxLen% : �C�C��ƪ��׳̤j��
'**********************************************************************
'�ŧiArray�ܼ�
'Dim A_Str$(), I%
'
'    �NTextBox�W���C�C���Keep��Array
'    A_Str$ = GetMultiLine2StrArray(A_A1641$, 30)
'
'    ��Array�����X�C�C��ƳB�z
'    I% = 0
'    Do While I% < UBound(A_Str$)
'       I% = I% + 1
'       MsgBox CStr(I%) & " : " & A_Str$(I%)
'    Loop
'**********************************************************************
Dim I&, A_Word$(), A_LineStr$(), A_STR$, A_Str1$, A_Str2$, A_Position%, A_Counts%
    
    ReDim A_LineStr$(0)
    '
    If Trim(A_EngStr$) = "" Then
        GetMultiLine2StrArray = A_LineStr$
        Exit Function
    End If
    '
    A_Word$() = Split(Trim(A_EngStr$), " ", -1, vbTextCompare)
    '
    A_STR$ = ""
    For I& = 0 To UBound(A_Word$)
        If lstrlen(A_STR$ & IIf(A_STR$ <> "", " ", "") & A_Word$(I&)) <= A_MaxLen% Then
            A_STR$ = A_STR$ & IIf(A_STR$ <> "", " ", "") & A_Word$(I&)
        Else
            If A_STR$ <> "" Then GoSub PrepareArray
            '
            '��@�^���r�j����ܪ���,�����H�̤j���פ���
            If lstrlen(A_Word$(I&)) > A_MaxLen% Then
                A_Str1$ = "": A_Str2$ = ""
                Do
                    A_STR$ = GetLenStr(A_Word$(I&), lstrlen(A_Str1$) + 1, A_MaxLen%)
                    A_Str1$ = A_Str1$ & A_STR$  'Keep�Ҧ��w�I��������
                    GoSub PrepareArray
                    '
                    A_Str2$ = GetLenStr(A_Word$(I&), lstrlen(A_Str1$) + 1, lstrlen(A_Word$(I&)) - lstrlen(A_Str1$))
                Loop Until lstrlen(A_Str2$) < A_MaxLen%
                '
                A_STR$ = IIf(Trim(A_Str2$) <> "", A_Str2$, "")
            Else
                A_STR$ = A_Word$(I&)
            End If
        End If
    Next I&
    '
    If A_STR$ <> "" Then GoSub PrepareArray
    '
    GetMultiLine2StrArray = A_LineStr$
    '
    Exit Function
    
PrepareArray:
    A_Counts% = UBound(A_LineStr$) + 1
    ReDim Preserve A_LineStr$(A_Counts%)
    A_LineStr$(A_Counts%) = A_STR$
    A_STR$ = ""
    '
    Return
End Function


'===============================================================================
' Add New Function at 93/6/7 by cathy
'===============================================================================
Public Sub SpreadCboStrCut(ByVal SS As FPSPREAD.vaSpread, ByVal Row&, ByVal Col&, ByVal A_Str1$, ByVal A_Cut, Optional Opt% = True)
''DESC:ComboBox Name,��쪺��,���j�Ÿ�
''     Opt%=True ,��ܥ��䳡�����N�X,�k����������
''          False,��ܥk�䳡�����N�X,������������
Dim I%, A_Pos

    SS.Row = Row&
    SS.Col = Col&
    SS.TypeComboBoxCurSel = -1
    If A_Str1$ = "" Then GoTo OutSub:
    For I% = 0 To SS.TypeComboBoxCount - 1
        SS.TypeComboBoxIndex = I%
        A_Pos = InStr(SS.TypeComboBoxString, A_Cut)
        If A_Pos = 0 Then A_Pos = Len(SS.TypeComboBoxString) + 1
        If Opt% Then
            If UCase$(Trim$(Left$(SS.TypeComboBoxString, A_Pos - 1))) = UCase$(Trim$(A_Str1$)) Then
                SS.TypeComboBoxCurSel = I%
                Exit For
            End If
        Else
            If UCase$(Trim$(Mid$(SS.TypeComboBoxString, A_Pos + 1))) = UCase$(Trim$(A_Str1$)) Then
                SS.TypeComboBoxCurSel = I%
                Exit For
            End If
        End If
    Next I%
    Exit Sub
    
OutSub:
    SS.TypeComboBoxCurSel = -1
End Sub

'===============================================================================
' Add New Function at 93/10/23
'===============================================================================
Sub TextFixWidth_Property(Frm As Form, Tmp As TextBox, ByVal MaxLen%, ByVal LineLen%)
'�]�wTextBox�r����Courier,Size=10���ݩ�,����@�C��J������
Dim A_ScaleMode%, A_AutoSizeChildren%
Dim A_FontName$, A_FontSize$, A_FontBold%, A_FontItalic%
Const SM_CXVSCROLL% = 2   'Width of arrow bitmap on vertical scroll bar
Const A_Twips% = 1

    With Frm
        'Keep���檺ScaleMode�Φr���ݩ�
        A_ScaleMode% = .ScaleMode
        A_FontName$ = .FontName
        A_FontSize$ = .FontSize
        A_FontBold% = .FontBold
        A_FontItalic% = .FontItalic
        
        'Keep���檺ScaleMode
        A_ScaleMode% = .ScaleMode
        
        '�]�w�����Font, MaxLength Property
        Tmp.FontName = IIf(G_FixFont_Name = "", "Courier", G_FixFont_Name)
        Tmp.FontSize = IIf(G_FixFont_Size = "", "10", G_FixFont_Size)
        Tmp.FontBold = False
        Tmp.FontItalic = False
        Tmp.MaxLength = MaxLen%
        
        '�]�w��檺�r��
        .FontName = Tmp.FontName
        .FontSize = Tmp.FontSize
        .FontBold = Tmp.FontBold
        .FontItalic = Tmp.FontItalic
          
        '�]�w��檺ScaleMode��Twips(��檺ScaleMode������Twips,�]Container��Elastic,
        '��Ȥ䴩Twips���䴩Pixels)
        .ScaleMode = A_Twips%

        '�p���r�����Width
        A_AutoSizeChildren% = .Vse_Background.AutoSizeChildren
        If A_AutoSizeChildren% <> azNone Then .Vse_Background.AutoSizeChildren = azNone
        Tmp.Width = IIf(Tmp.ScrollBars <> vbVertical, LineLen% * .TextWidth("a") + 125, _
                    LineLen% * .TextWidth("a") + 125 + GetSystemMetrics(SM_CXVSCROLL%) * Screen.TwipsPerPixelX)
        If A_AutoSizeChildren% <> azNone Then .Vse_Background.AutoSizeChildren = A_AutoSizeChildren%
  
        '�٭��檺ScaleMode�Φr���ݩ�
        .ScaleMode = A_ScaleMode%
        .FontName = A_FontName$
        .FontSize = A_FontSize$
        .FontBold = A_FontBold%
        .FontItalic = A_FontItalic%
    End With
End Sub


'===============================================================================
' Add New Function at 93/12/2
'===============================================================================
Function GetSQLRounding(DB As Database, ByVal FldStr$, ByVal DecimalNumber%, Optional ByVal Truncate% = False, Optional ByVal AliasName$, Optional ByVal Options% = True) As String
'���o���|�ˤ��J�εL����˥h��SQL���
'�Ѽ� : FldStr$ - ���|�ˤ��J�����W��
'       DecimalNumber% - ���|�ˤ��J�����
'       Truncate% - True:�L����˥h False:�|�ˤ��J
Dim A_FMT$, A_NumBer$

    AliasName$ = Trim(AliasName$)
    '
    If Trim(DB.Connect) = "" Or Not Options% Then   'Access Database
       If Truncate% Then
          If DecimalNumber% <= 0 Then
             A_FMT$ = " Fix(@FldStr) "
          Else
             A_NumBer$ = "1" & String(DecimalNumber%, "0")
             A_FMT$ = " Fix(@FldStr*" & A_NumBer$ & ")/" & A_NumBer$ & " "
          End If
       Else
          If DecimalNumber% <= 0 Then
             A_FMT$ = " Format(@FldStr,'0') "
          Else
             A_FMT$ = " Format(@FldStr,'0." & _
                   String(DecimalNumber%, "0") & "') "
          End If
       End If
    Else                            'ODBC Database
       Select Case UCase$(Mid$(G_ConnectMethod1, InStr(1, G_ConnectMethod1, "DBTYPE=", 1) + 7))
         Case "SQL;"
              If Truncate% Then
                 A_FMT$ = " Round(@FldStr, @DecimalNumber,1) "
              Else
                 A_FMT$ = " Round(@FldStr, @DecimalNumber) "
              End If
       End Select
    End If
    '
    A_FMT$ = Replace(A_FMT$, "@FldStr", FldStr$, 1, -1, vbTextCompare)
    A_FMT$ = Replace(A_FMT$, "@DecimalNumber", CStr(DecimalNumber%), 1, -1, vbTextCompare)
    If AliasName$ <> "" Then A_FMT$ = A_FMT$ & " AS " & AliasName$ & " "
    '
    GetSQLRounding = A_FMT$
End Function

'===============================================================================
' Add New Function at 94/5/19 by Cathy For ����100�~
'===============================================================================
Function DateFormat2(ByVal DateStr$, Optional ShowSlash As Boolean = True) As String
'��Function�D�n�ت�:��J���ƬO�C�L��Spread��,�J����~�פ����T�X��,����H'0'�ɨ�
'�ǤJ�Ѽ�ShowSlash���Ϊk:
'  1.�@��¿ù���ܩΦC�L�����������,ShowSlash�����ǤJ,����|�榡�Ƭ��~/��/��
'  2.�Y������i�P�ɿ�J�P�ƧǮ�,��Spread LeaveCell �w��Ӥ���ˮֵL�~��,�M�Φ�Function�H�ഫ������


    DateFormat2 = " "
    If Trim$(DateStr$) = "" Then Exit Function
    
    If G_PrintSelect = 0 Then G_PrintSelect = G_Print2Screen
    DateStr$ = Replace(DateStr$, "/", "")
    '
    If ShowSlash = True Then
        If G_PrintSelect = G_Print2Screen Then
            DateFormat2 = Format$(DateStr$, "#000/##/##")
        Else
            DateFormat2 = Format$(DateStr$, "##00/##/##")
        End If
    Else
        If G_DateFlag = 1 And G_PrintSelect = G_Print2Screen Then
            DateFormat2 = Format$(DateStr$, "0000000")
        Else
            DateFormat2 = DateStr$
        End If
    End If
End Function

Function DateOut2(ByVal DateStr$) As String
'�N����ഫ���t�γ]�w����ܫ��A(���Φ��),Output�ɨϥ�

    DateStr$ = Trim(DateStr$)
    DateOut2 = " "
    If Val(DateStr$) = 0 Then Exit Function
    
    Select Case G_DateFlag
      Case 0
           DateOut2 = Format$(DateStr$, "########")
      Case 1
           DateOut2 = Format$(Val(DateStr$) - 19110000, "0000000")
      Case 2
           DateOut2 = Format$(IIf(Left$(DateStr$, 2) = G_LeadYear$, _
                     Mid$(DateStr$, 3), DateStr$), "##000000")
    End Select
End Function

Function RejectSlash(ByVal A_Source$) As String
'�����r�ꤤ��"/"
'�p:�ǤJ"89/01/01",�h�Ǧ^"890101"
Dim I%, A_RStr$

    RejectSlash = Replace(A_Source$, "/", "")
End Function
'-------------------------------------------------------------------------------
Sub Get_SaleTaxDecimal(DB As Database, UnitSaleTaxDecimal&, SumSaleTaxDecimal&)
'940706 Add By Yvonne
'�ت�:���o�p��|�B���p�Ʀ�Ƥγ����|�B����ܪ��p�Ʀ��
'�ǤJ��Database�����OGENIE
'����:UnitSaleTaxDecimal#:���س������p�Ʀ�����
'����:SumSaleTaxDecimal# :�X�p�������p�Ʀ���㤣
On Local Error GoTo MY_Error
Dim A_Country$, A_TopicValue$

    A_Country$ = GetSvrINIStrA(DB_ARTHGUI, "Customer", "Country")
    If Trim(A_Country$) = "" Or UCase(Left(A_Country$, 2)) = "TW" Then A_Country$ = "TWN"
    
    '���|�B�p�Ʀ��
    If IsExistSvrINI(DB, "UnitSaleTaxDecimal", "TWN") = False Then
        MoveData2Sini DB, "UnitSaleTaxDecimal", "TWN", "2"
    End If
    If IsExistSvrINI(DB, "UnitSaleTaxDecimal", "CHN") = False Then
        MoveData2Sini DB, "UnitSaleTaxDecimal", "CHN", "4"
    End If

    UnitSaleTaxDecimal& = CvrTxt2Num(GetSvrINIStrA(DB, "UnitSaleTaxDecimal", "TWN"))

    If IsExistSvrINI(DB, "UnitSaleTaxDecimal", A_Country$, A_TopicValue$) = True Then
        UnitSaleTaxDecimal& = CvrTxt2Num(A_TopicValue$)
    End If
   
    '�X�p�|�B�p�Ʀ��
    If IsExistSvrINI(DB, "SumSaleTaxDecimal", "TWN") = False Then
        MoveData2Sini DB, "SumSaleTaxDecimal", "TWN", "0"
    End If
    If IsExistSvrINI(DB, "SumSaleTaxDecimal", "CHN") = False Then
        MoveData2Sini DB, "SumSaleTaxDecimal", "CHN", "2"
    End If

    SumSaleTaxDecimal& = CvrTxt2Num(GetSvrINIStrA(DB, "SumSaleTaxDecimal", "TWN"))

    If IsExistSvrINI(DB, "SumSaleTaxDecimal", A_Country$, A_TopicValue$) = True Then
        SumSaleTaxDecimal& = CvrTxt2Num(A_TopicValue$)
    End If
   
    Exit Sub
MY_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub
Function IsExistSvrINI(DB As Database, ByVal Section$, ByVal Topic$, Optional A_TopicValue$) As Boolean
'940706 Add By Yvonne
'���o���w��Ʈw��,SINI-TABLE����TOPICVALUE��
Dim DY As Recordset
Dim A_Sql$

    IsExistSvrINI = False: A_TopicValue$ = ""
    A_Sql$ = "Select TOPICVALUE From SINI Where"
    A_Sql$ = A_Sql$ & " SECTION='" & Section$ & "'"
    A_Sql$ = A_Sql$ & " AND TOPIC='" & Topic$ & "'"
    A_Sql$ = A_Sql$ & " Order by SECTION,TOPIC"
    CreateDynasetODBC DB, DY, A_Sql$, "DY", True
    If Not (DY.BOF And DY.EOF) Then
       A_TopicValue$ = Trim(DY.Fields("TOPICVALUE") & "")
       IsExistSvrINI = True
    End If
    DY.Close
    Set DY = Nothing
End Function

Function IsTableCollumExist(DB As Database, ByVal A_TableName$, ByVal A_ColName$) As Boolean

Dim A_Sql$
Dim A_ErrCode
    
    IsTableCollumExist = False
    A_Sql$ = "Select " & A_ColName$ & " From " & A_TableName$ & " WHERE 1<>1"
    ExecuteProcessReturnErr DB, A_Sql$, A_ErrCode
    If A_ErrCode > 0 Then Exit Function
    IsTableCollumExist = True
    

End Function

Function GetExcelColName(ByVal Col%) As String
'�N�ƭ��ഫ��Excel�������W��
Dim A_Num%

    A_Num% = Col% Mod 26
    GetExcelColName = IIf(Col% > 26, Chr((Col% - 1) \ 26 + 64), "") + _
                      IIf(A_Num% = 0, "Z", Chr(A_Num% + 64))
End Function

Function IsOSUpperXP() As Boolean
'�P�_OS��XP�H�W��Version
    IsOSUpperXP = False
    
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = Len(osvi)
    If GetVersionEx(osvi) = 0 Then
       Exit Function
    End If
    
    Const XPMajorVersion = 5
    If osvi.dwMajorVersion > XPMajorVersion Then
        IsOSUpperXP = True
    End If
End Function

Function GetSecurityPwdMinLen() As String
'���o�K�X�̤p����(�Y�ťչw�]��0���ޱ�)
Dim A_STR$

    A_STR$ = GetGUISvrIniStr("Security", "PwdMinLen")
    GetSecurityPwdMinLen = IIf(Trim(A_STR$) = "", "0", A_STR$)
End Function

Function GetSecurityPwdFixedLen() As String
'S020308013���o�K�X�T�w����(�Y�ťչw�]��0���ޱ�)
Dim A_STR$

    A_STR$ = GetGUISvrIniStr("Security", "PwdFixedLen")
    GetSecurityPwdFixedLen = IIf(Trim(A_STR$) = "", "0", A_STR$)
End Function

Function GetSecurityPwdComplexity() As String
'S020308013���o�K�X�O�_�ޱ�������,���]�w���ޱ�
Dim A_STR$

    A_STR$ = UCase(GetGUISvrIniStr("Security", "PwdComplexity"))
    GetSecurityPwdComplexity = IIf(Trim(A_STR$) = "", "N", A_STR$)
End Function

Function GetSecurityPwdFailedWaitTime() As String
'S020308013���o�K�X��J���ѶW�L�t�έ����,�ݵ��ݳ]�w���ɶ�(��),�~�i�ATRY(�Y�ťչw�]��0���ޱ�)
Dim A_STR$

    A_STR$ = UCase(GetGUISvrIniStr("Security", "PwdFailedWaitTime"))
    GetSecurityPwdFailedWaitTime = IIf(Trim(A_STR$) = "", "0", A_STR$)
End Function

Function CheckInvoiceWithDraw(ByVal DB As Database, ByVal A_Company$, ByVal A_Invoice$, ByVal A_InvoiceType$, Optional ByRef A_Msg$ = "") As Boolean
'�ˬd�@�o�o���O�_��32A/35A�榡
Dim A_Sql$, DY As Recordset
    CheckInvoiceWithDraw = True
    
    If Not (A_InvoiceType$ = "32" Or A_InvoiceType$ = "35") Then Exit Function
    If Trim(A_Company$) = "" Or Trim(A_Invoice$) = "" Then Exit Function
    
    A_Sql$ = "SELECT Top 1 B3102,B3103 FROM B31 WHERE B3101='" & A_Company$ & "'"
    A_Sql$ = A_Sql$ + " AND B3103 >'" & Left(Trim(A_Invoice$), 2) & "'"
    A_Sql$ = A_Sql$ + " AND B3103 <='" & Trim(A_Invoice$) & "'"
    A_Sql$ = A_Sql$ + " AND B3105 >='" & Right(Trim(A_Invoice$), 8) & "'"
    A_Sql$ = A_Sql$ + " AND B3104 ='" & A_InvoiceType$ & "'"
    A_Sql$ = A_Sql$ + " AND B3123 ='A'"
    A_Sql$ = A_Sql$ + " ORDER BY B3103"
    CreateDynasetODBC DB, DY, A_Sql$, "DY", True
    
    If Not (DY.BOF And DY.EOF) Then
        CheckInvoiceWithDraw = False
        'Arthur 96/04/01 �ѩ��ˮֶȾA�Ω�x�W�o���|�k�W�w, ������ܰT��, �L���H��w���A�O��
        A_Msg$ = "���o�����@�o�o��, �G���B�@�ߥH0��J."
        A_Msg$ = A_Msg$ + Chr(13) + Chr(10) + "�ЦܷJ�`�o��" + DateOut(DY.Fields("B3102") & "") + ", "
        A_Msg$ = A_Msg$ + Trim(DY.Fields("B3103") & "")
        A_Msg$ = A_Msg$ + "��, �N�o�����B�վ㬰�o���b�B!"
        MsgBox A_Msg$, vbInformation, App.Title
    End If
    
    DY.Close
End Function

'*** Add For Vista 96/6/25 By Jennifer
'�P�_�t�άO�_�w�ҥΩ�Vista���ҤU
Sub EnableVistaClient()
Dim A_Path$

    A_Path$ = GetIniStr("FilePath", "ServerPath", "GUI.INI") & "GUI.INI"
    G_IsVistaClient = (StrComp(GetIniStr("VistaClient", "Enable", A_Path$), "1", vbTextCompare) = 0)
    If G_IsVistaClient Then
       G_VistaClientTitle = Trim(GetIniStr("VistaClient", "Title", A_Path$))
       If G_VistaClientTitle = "" Then G_VistaClientTitle = "[Windows 7]"
    End If
End Sub

'*** Add For Vista 96/6/25 By Jennifer
Function GetEngine() As DAO.DBEngine
'�YG_IsVistaClient = True, �Y��ܨt�Τw��Vista�W����, �h�ϥ�DAO 3.6 Object Library. �_�h�ϥ�DAO 3.51 Object Library.
    If G_IsVistaClient Then
        If DBEngine36 Is Nothing Then
            Set DBEngine36 = CreateObject("DAO.DBEngine.36")
        End If
        Set GetEngine = DBEngine36
    Else
        Set GetEngine = DBEngine
    End If
End Function

Sub ResizeImage(ByVal Pnl As Control, ByVal img As Image, ByVal imageFileName As String, _
Optional ByVal padding As Integer = 60, Optional ByVal resizeType As Integer = 1)
'������Y��Ϥ�
'981020�̰ѼƦۦ�M�w�O�_������Y�p�Ω�j--------------Edit By Yvonne
'resizeType�Ѽ�1-->�P����Y�� 2-->�u�B�z�P����Y�p 3-->�u�B�z�P��ҩ�j
Dim A_Hf#, A_Wf#, A_Ho#, A_Wo#, A_M#
    
    A_Hf# = Pnl.Height
    A_Wf# = Pnl.Width
    
    img.Stretch = False
    img.Picture = LoadPicture(imageFileName)
    Select Case resizeType
        Case 1
        Case 2
            If Not (Pnl.Width < img.Width Or Pnl.Height < img.Height) Then
                GoTo NotResize
            End If
        Case 3
            If Not (Pnl.Width > img.Width Or Pnl.Height > img.Height) Then
                GoTo NotResize
            End If
    End Select
    
    A_Ho# = img.Height
    A_Wo# = img.Width
    
    If A_Ho# / (A_Hf# - padding * 2) > A_Wo# / (A_Wf# - padding * 2) Then
        A_M# = A_Ho# / (A_Hf# - padding * 2)
    Else
        A_M# = A_Wo# / (A_Wf# - padding * 2)
    End If
    
    img.Width = A_Wo# / A_M#
    img.Height = A_Ho# / A_M#
    
NotResize:
    img.Stretch = True
    
    img.Left = padding
    img.Top = padding
    
    If Pnl.Width > img.Width Then img.Left = padding + _
        CInt((Pnl.Width - img.Width - padding * 2) / 2)
    If Pnl.Height > img.Height Then img.Top = padding + _
        CInt((Pnl.Height - img.Height - padding * 2) / 2)
End Sub

Function IsTableExist(DB As Database, ByVal A_TableName$) As Boolean
'�P�_��ƪ�O�_�s�b
'Function�Ǧ^�� : Boolean (True:���w�s�b False:��椣�s�b)
On Error Resume Next

    IsTableExist = False
    If Trim(DB.Connect) = "" Then
       Debug.Print DB.TableDefs(A_TableName$).Name
       If Err = 0 Then IsTableExist = True
    Else
       Dim rs As Recordset
       Set rs = DB.OpenRecordset( _
            "SELECT COUNT(*) FROM sysobjects WHERE id=OBJECT_ID('" & A_TableName$ & "')", _
            dbOpenSnapshot, dbSQLPassThrough)
       If Err = 0 Then
          If Val(rs(0) & "") > 0 Then IsTableExist = True
       End If
       rs.Close
       Set rs = Nothing
    End If
End Function

Function IsIndexExist(DB As Database, ByVal A_TableName$, ByVal A_IndexName$) As Boolean
'�P�_���ެO�_�s�b
'Function�Ǧ^�� : Boolean (True:���ޤw�s�b False:���ޤ��s�b)
On Error Resume Next

    IsIndexExist = False
    If Trim(DB.Connect) = "" Then
       Debug.Print DB.TableDefs(A_TableName$).Indexes(A_IndexName$).Name
       If Err = 0 Then IsIndexExist = True
    Else
       Dim rs As Recordset
       Set rs = DB.OpenRecordset( _
            "SELECT COUNT(*) FROM sysindexes WHERE id=OBJECT_ID('" & A_TableName$ & "')" & _
            " AND name='" & A_IndexName$ & "'", _
            dbOpenSnapshot, dbSQLPassThrough)
       If Err = 0 Then
          If Val(rs(0) & "") > 0 Then IsIndexExist = True
       End If
       rs.Close
       Set rs = Nothing
    End If
End Function

Function IsFieldExist(DB As Database, ByVal A_TableName$, ByVal A_FieldName$) As Boolean
'�P�_���O�_�s�b
'Function�Ǧ^�� : Boolean (True:���w�s�b False:��줣�s�b)
On Error Resume Next

    IsFieldExist = False
    If Trim(DB.Connect) = "" Then
       Debug.Print DB.TableDefs(A_TableName$).Fields(A_FieldName$).Name
       If Err = 0 Then IsFieldExist = True
    Else
       Dim rs As Recordset
       Set rs = DB.OpenRecordset( _
            "SELECT COUNT(*) FROM syscolumns WHERE id=OBJECT_ID('" & A_TableName$ & "')" & _
            " AND name='" & A_FieldName$ & "'", _
            dbOpenSnapshot, dbSQLPassThrough)
       If Err = 0 Then
          If Val(rs(0) & "") > 0 Then IsFieldExist = True
       End If
       rs.Close
       Set rs = Nothing
    End If
End Function

Function AddTableField(DB As Database, ByVal A_TableName$, ByVal A_FieldName$, ByVal A_DataType%, _
Optional ByVal A_Size& = 1, Optional ByVal A_DoExistCheck% = True) As Boolean
'�[�J������
On Error GoTo MyError

    AddTableField = False
    If A_DataType% <> G_Data_String And A_DataType% <> G_Data_Numeric Then Exit Function
    If A_DoExistCheck% Then
       If IsFieldExist(DB, A_TableName$, A_FieldName$) Then Exit Function
    End If

    If Trim(DB.Connect) = "" Then
        Dim A_TD As TableDef, A_Fld As Field
        Set A_TD = DB.TableDefs(A_TableName$)
        If A_DataType% = G_Data_String Then
           Set A_Fld = A_TD.CreateField(A_FieldName$, dbText, A_Size&)
           A_Fld.DefaultValue = """ """
        Else
           Set A_Fld = A_TD.CreateField(A_FieldName$, dbDouble)
           A_Fld.DefaultValue = 0
        End If
        A_TD.Fields.Append A_Fld
    Else
        Dim A_Sql$
        A_Sql$ = "ALTER TABLE " & A_TableName$
        A_Sql$ = A_Sql$ & " ADD " & A_FieldName$
        If A_DataType% = G_Data_String Then
           A_Sql$ = A_Sql$ & " VARCHAR(" & CStr(A_Size&) & ") NOT NULL DEFAULT ' '"
        Else
           A_Sql$ = A_Sql$ & " NUMERIC(25,4) NOT NULL DEFAULT 0"
        End If
        ExecuteProcess DB, A_Sql$
    End If
    
    AddTableField = True
    Exit Function
    
MyError:
    AddTableField = False
End Function

Function IsIndexColumnExist(DB As Database, ByVal A_TableName$, ByVal A_IndexName$, ByVal A_FieldName$) As Boolean
'�P�_���ޤ��O�_�s�b�Y���
'Function�Ǧ^�� : Boolean (True:���ޤw���s�b False:������줣�s�b)
On Error Resume Next

    IsIndexColumnExist = False
    If Trim(DB.Connect) = "" Then
       IsIndexColumnExist = (InStr(DB.TableDefs(A_TableName$).Indexes(A_IndexName$).Fields, A_FieldName$) > 0)
    Else
       Dim A_Sql$
       A_Sql$ = "SELECT COUNT(*) FROM sysindexkeys"
       A_Sql$ = A_Sql$ & " WHERE id=OBJECT_ID('[" & A_TableName$ & "]')"
       A_Sql$ = A_Sql$ & " AND indid="
       A_Sql$ = A_Sql$ & " (SELECT indid FROM sysindexes WHERE"
       A_Sql$ = A_Sql$ & " id=OBJECT_ID('[" & A_TableName$ & "]') AND name='" & A_IndexName$ & "')"
       A_Sql$ = A_Sql$ & " AND colid="
       A_Sql$ = A_Sql$ & " (SELECT colid FROM syscolumns WHERE"
       A_Sql$ = A_Sql$ & " id=OBJECT_ID('[" & A_TableName$ & "]') AND name='" & A_FieldName$ & "')"
       
       Dim rs As Recordset
       Set rs = DB.OpenRecordset(A_Sql$, dbOpenSnapshot, dbSQLPassThrough)
       If Err = 0 Then
          If Val(rs(0) & "") > 0 Then IsIndexColumnExist = True
       End If
       rs.Close
       Set rs = Nothing
    End If
End Function

Function GetDBTextFieldLen(DB As Database, ByVal A_TableName$, ByVal A_FieldName$) As Long
'���o��Ʈw����r��쪺����
On Error GoTo MyError
Dim A_Size&

    A_Size& = 0
    If Not IsFieldExist(DB, A_TableName$, A_FieldName$) Then Exit Function
    
    If Trim(DB.Connect) = "" Then
       A_Size& = DB.TableDefs(A_TableName$).Fields(A_FieldName$).Size
    Else
       Dim rs As Recordset
       Set rs = DB.OpenRecordset( _
            "SELECT prec FROM syscolumns WHERE id=OBJECT_ID('" & A_TableName$ & "')" & _
            " AND name='" & A_FieldName$ & "'", _
            dbOpenSnapshot, dbSQLPassThrough)
       If Not (rs.BOF And rs.EOF) Then A_Size& = Val(rs(0) & "")
       rs.Close
       Set rs = Nothing
    End If
    GetDBTextFieldLen = A_Size&
    Exit Function
    
MyError:
    MsgBox Error$, vbCritical, App.Title
End Function

Sub LockControl(A_Ctl As Control, A_Lock As Boolean)
    A_Ctl.Locked = A_Lock
    A_Ctl.TabStop = IIf(A_Lock, False, True)
End Sub

'�a�^���w��Folder���|
'20091112 Add By Yvonne
Function OpenFolderDialog(Frm As Form, folder As String) As Boolean
Const BIF_RETURNONLYFSDIRS = 1
Const BIF_DONTGOBELOWDOMAIN = 2
Const MAX_PATH = 1024
Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

    OpenFolderDialog = False
    folder = ""
    
    szTitle = "Choose a folder : "
    With tBrowseInfo
       .hWndOwner = Frm.hwnd
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    OpenFolderDialog = lpIDList
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       folder = sBuffer
    End If
End Function

'�a�}�ഫ�����N�q���Ϥ�, �A�Ω�x�W
Sub AddressConvert_TW(ByVal A_Address$, A_Zipcode$, A_Country$, A_Region$, A_StateProvince$, A_County$, A_City$, A_District$, A_Street$, A_StreetSection$)
    Dim a!, A_Str1$, A_Str2$, A_Li$, A_Lin$
    
    '���P�_�l���ϸ�
    A_Str1$ = ""
    For a! = 1 To Len(A_Address$)
        If Mid(A_Address$, a!, 1) >= "0" And Mid(A_Address$, a!, 1) <= "9" Then
           A_Str1$ = A_Str1$ + Mid(A_Address$, a!, 1)
        Else
           Exit For
        End If
    Next
    If A_Str1$ <> "" Then
       A_Zipcode$ = A_Str1$
    Else
       A_Zipcode$ = ""
    End If
    A_Str2$ = Right(A_Address$, Len(A_Address$) - Len(A_Str1$))
    
    '�P�_��
    A_Str1$ = ""
    a! = InStr(1, A_Str2$, "��")
    If a! > 0 Then
       A_Str1$ = GetLenStr(A_Str2$, 1, a! + 3)
    End If
    A_County$ = A_Str1$
    A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    
    '�P�_��
    A_Str1$ = ""
    a! = InStr(1, A_Str2$, "��")
    If a! > 0 Then
       A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
    End If
    A_City$ = A_Str1$
    A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    
   '�P�_��
    If A_City$ = "" Then
        A_Str1$ = ""
        a! = InStr(1, A_Str2$, "��")
        If a! > 0 Then
           A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
        End If
        A_City$ = A_Str1$
        A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    End If
    
    '�P�_��
    If A_City$ = "" Then
        A_Str1$ = ""
        a! = InStr(1, A_Str2$, "�m")
        If a! > 0 Then
           A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
        End If
        A_City$ = A_Str1$
        A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    End If
    
    '�P�_��
    A_Str1$ = ""
    a! = InStr(1, A_Str2$, "��")
    If a! > 0 Then
       A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
    End If
    A_District$ = A_Str1$
    A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    
    If A_District$ = "" Then
        A_Str1$ = ""
        a! = InStr(1, A_Str2$, "��")
        If a! > 0 Then
           A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
        End If
        A_District$ = A_Str1$
        A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    End If
    '�P�_���í簣
    A_Str1$ = ""
    a! = InStr(1, A_Str2$, "��")
    If a! > 0 Then
       A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
    End If
    A_Li$ = A_Str1$
    A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    '�P�_�F�í簣
    A_Str1$ = ""
    a! = InStr(1, A_Str2$, "�F")
    If a! > 0 Then
       A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
    End If
    A_Lin$ = A_Str1$
    A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    
    '�P�_��
    A_Str1$ = ""
    a! = InStr(1, A_Str2$, "��")
    If a! > 0 Then
       A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
    End If
    If A_Str1$ <> "" Then
        Do Until Right(A_Str1$, 1) = "��"
           A_Str1$ = Left(A_Str1$, Len(A_Str1$) - 1)
        Loop
    End If
    A_Street$ = A_Str1$
    A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    
    If A_Street$ = "" Then
        A_Str1$ = ""
        a! = InStr(1, A_Str2$, "��")
        If a! > 0 Then
           A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
        End If
        If A_Str1$ <> "" Then
            Do Until Right(A_Str1$, 1) = "��"
               A_Str1$ = Left(A_Str1$, Len(A_Str1$) - 1)
            Loop
        End If
        A_Street$ = A_Str1$
        A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    End If
    If A_Street$ = "" Then
        A_Str1$ = ""
        a! = InStr(1, A_Str2$, "�D")
        If a! > 0 Then
           A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
        End If
        If A_Str1$ <> "" Then
            Do Until Right(A_Str1$, 1) = "�D"
               A_Str1$ = Left(A_Str1$, Len(A_Str1$) - 1)
            Loop
        End If
        A_Street$ = A_Str1$
        A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    End If
    
    If A_Street$ = "" Then
       If A_Li$ <> "" Then
          A_Street$ = A_Li$
       ElseIf A_Lin$ <> "" Then
          A_Street$ = A_Lin$
       End If
    End If
    '�P�_�q
    A_Str1$ = ""
    a! = InStr(1, A_Str2$, "�q")
    If a! > 0 Then
       A_Str1$ = GetLenStr(A_Str2$, 1, a! + a!)
       Do Until Right(A_Str1$, 1) = "�q"
          A_Str1$ = Left(A_Str1$, Len(A_Str1$) - 1)
       Loop
    End If
    A_StreetSection$ = A_Str1$
    A_Str2$ = Right(A_Str2$, Len(A_Str2$) - Len(A_Str1$))
    '�P�_�l�F
    A_Str1$ = ""
    a! = InStr(1, A_Str2$, "�l�F")
    If a! > 0 Then
       A_Str1$ = GetLenStr(A_Str2$, 1, a! + a! - 2)
    End If
    If Trim(A_City$) = "" Then
       A_City$ = A_Str1$
    End If
End Sub

Function IsWinForTaiwan() As Boolean
'�P�_OS�O�_�����媩
'1028:����(�x�W),1033:�^��(����)
    IsWinForTaiwan = (GetUserDefaultLangID() = 1028)
End Function

Function XlsFldUseChinaDate() As Boolean
'�P�_Excel�����O�_�ϥΰ�����榡(EMD),�Y���]�w�w�]���ҥ�.
    XlsFldUseChinaDate = Not (GetSvrINIStrA(DB_ARTHGUI, "Customer", "XlsFldUseChinaDate") = "N")
End Function


Function CvrString2Character(ByVal A_STR$) As String
''desc:�N����r�ꤤ�R�X�Ÿ��� ' or | �ন�r���ۥ[
''�Ҧp:���|��='��?&Chr$(124)&'��'
Dim A_Temp$, A_Temp2$
Dim I%

    CvrString2Character = "''"
    If Trim$(A_STR$) = "" Then Exit Function
    
    A_Temp$ = "'"
    For I% = 1 To Len(A_STR$)
        A_Temp2 = Mid(A_STR$, I%, 1)
        If A_Temp2$ = Chr(39) Or A_Temp2$ = Chr(124) Then
           A_Temp$ = A_Temp$ & "'&" & "Chr$(" & Trim(Asc(A_Temp2)) & ")&'"
        Else
           A_Temp$ = A_Temp$ & A_Temp2
        End If
    Next I%
    If A_Temp2$ = Chr(39) Or A_Temp2$ = Chr(124) Then   '�R�X�Ÿ��� ' or |
       A_Temp$ = Left(A_Temp$, Len(A_Temp$) - 2)
    Else
       A_Temp$ = A_Temp$ & "'"
    End If
    CvrString2Character = A_Temp$

End Function

'Add By Yvonne 20110704
Function ClearSpecialChar(ByVal A_STR$, Optional CompareMethod = vbTextCompare) As String
'�M���r�ꤤ�������Tab�Ÿ�
Dim A_NewStr$
    
    A_NewStr$ = A_STR$
    A_NewStr$ = Replace(A_NewStr$, Chr$(13) & Chr$(10), "", 1, , CompareMethod)
    A_NewStr$ = Replace(A_NewStr$, Chr$(13), "", 1, , CompareMethod)
    A_NewStr$ = Replace(A_NewStr$, Chr$(10), "", 1, , CompareMethod)
    A_NewStr$ = Replace(A_NewStr$, Chr$(9), "", 1, , CompareMethod)
    '
    ClearSpecialChar = A_NewStr$
End Function

'S010801047 �[�J�@��Panel��ܥثe��Ʈw�ƥ��ɳ̷s���ɶ�
Function AddCurrentBAKDatetimeStatusBarPanel(sb As StatusBar, ByVal PanelIndex As Integer, ByVal Path$)
Dim Pnl As Panel, A_Date$
Dim fso As Object
Dim d As Object
Dim f As Object

On Local Error GoTo MY_Error

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.getfolder(Path$)
    
    For Each f In d.Files
        If UCase(fso.getExtensionName(f.Name)) = "BAK" Then
            If Trim(A_Date$) = "" Then
                A_Date$ = Format(f.DateCreated, "yyyymmdd")
            Else
                If A_Date$ < Format(f.DateCreated, "yyyymmdd") Then
                    A_Date$ = Format(f.DateCreated, "yyyymmdd")
                End If
            End If
        End If
    Next
    
    Set Pnl = sb.Panels.Add(PanelIndex)
    With Pnl
         .text = GetCaption("DBBACKUP", "BAKDATE", "�ƥ��ɤ��") & ":" & Format(A_Date$, "####/##/##")
         .style = sbrText
         .AutoSize = sbrContents
         .MinWidth = 2500
         .Alignment = sbrCenter
         .Bevel = sbrInset
    End With
    Set fso = Nothing
    Set f = Nothing
    Set d = Nothing
    Exit Function
    
MY_Error:
        MsgBox Error(Err)
        Exit Function
End Function

Function CheckPwdComplexity(ByVal Pwd$) As Boolean
'S020308013�ˮֱK�X�������ʬO�_�ŦX(�^��j�g,�^��p�g,�Ʀr,�S��Ÿ�)�|�ػݥ]�t�T��
Dim I&, A_IsNum As Boolean, A_IsUpperCase As Boolean, A_IsLowercase As Boolean, A_IsParticular As Boolean
Dim A_HaveCnt&

    CheckPwdComplexity = False
    
    A_IsNum = False: A_IsUpperCase = False: A_IsLowercase = False: A_IsParticular = False
    For I& = 1 To Len(Pwd$)
        Select Case Asc(Mid(Pwd$, I&, 1))
            Case Asc("0") To Asc("9")
                A_IsNum = True
            Case Asc("A") To Asc("Z")
                A_IsUpperCase = True
            Case Asc("a") To Asc("z")
                A_IsLowercase = True
            Case Asc("`"), Asc("~"), Asc("!"), Asc("@"), Asc("#"), Asc("$"), Asc("%"), Asc("^"), Asc("&"), Asc("*") _
                , Asc("("), Asc(")"), Asc("-"), Asc("_"), Asc("="), Asc("+"), Asc("["), Asc("{"), Asc("]"), Asc("}") _
                , Asc("\"), Asc("|"), Asc(";"), Asc(":"), Asc("'"), Asc(""""), Asc(","), Asc("<"), Asc("."), Asc(">") _
                , Asc("/"), Asc("?")
                A_IsParticular = True
        End Select
    Next
    
    A_HaveCnt& = 0
    If A_IsNum = True Then A_HaveCnt& = A_HaveCnt& + 1
    If A_IsUpperCase = True Then A_HaveCnt& = A_HaveCnt& + 1
    If A_IsLowercase = True Then A_HaveCnt& = A_HaveCnt& + 1
    If A_IsParticular = True Then A_HaveCnt& = A_HaveCnt& + 1
    
    If A_HaveCnt& >= 3 Then CheckPwdComplexity = True
End Function

'S020527055 20130531
Sub SQLCompose(ByVal Table$)
'�զXSQL���O,�f�tInsertFields�{�Ǩϥ�
Dim A_Tmp$, A_Str1$, A_Str2$, A_Sql$

'S021114036 �]�ǲ�ñ�֮ɡA�ݲզ귥�����r��A�G�Ni%�ܼƩ��̤j(1021115 by Lidia)
Dim I As Currency
    
    A_Tmp$ = Chr(0) & Chr(128)
    I = InStr(1, G_Str, A_Tmp$)
    If I <> 0 Then
       A_Str1$ = Left(G_Str, I - 1)
       A_Str2$ = Right(G_Str, Len(G_Str) - (I + 1))
    End If
    A_Str1$ = "(" & A_Str1$ & ")"
    If Right(A_Str2$, 1) = "," Then
       A_Str2$ = Left(A_Str2$, Len(A_Str2$) - 1)
    End If
    A_Sql$ = "Insert into " & Table$ & Space(1) & A_Str1
    A_Sql$ = A_Sql$ & " values " & "(" & A_Str2$ & ")"
    G_Str = A_Sql$
End Sub

Function Check_Executable(ByVal A_System$, ByVal A_PgmName$, ByVal A_APOpt$, A_Msg$, Optional ByVal A_HaveDataChk% = False) As Integer
'102/12/09 �ˮ֬O�_���v�{�� Function----�B��(S021015032)
Dim a$, A_Sql$
    G_IllegalTerminal = GetSIniStr(A_System$, "illegal_terminal")
    G_Authority = GetSIniStr(A_System$, "authority")
    
    If UCase$(Trim$(G_DUserId)) = "GUI" Then
       Check_Executable = True
       Exit Function
    End If
    
    If Len(Trim$(A_PgmName$)) > 10 Then
       a$ = Mid$(Trim$(A_PgmName$), 1, 10)
    Else
       a$ = Trim$(A_PgmName$)
    End If
    Check_Executable = False

    A_Msg$ = ""
    If G_Terminal_Check Then
       If Not HaveTerminalLicense(A_APOpt$) Then
          A_Msg$ = G_IllegalTerminal
          Exit Function
       End If
    End If
    '
    If A_HaveDataChk% = True Then
        '�ˮ֬O�_�����@�s�ձ��v�L�A�Y�L�h�L���󳣥i�ϥ�
        A_Sql$ = "Select * From A07"
        A_Sql$ = A_Sql$ & " where A0702='" & a$ & "'"
        A_Sql$ = A_Sql$ & " and A0703='Y'"
        A_Sql$ = A_Sql$ & " order by A0701,A0704,A0702"
        CreateDynasetODBC DB_ARTHGUI, DY_A07, A_Sql$, "DY_A07", True
        If DY_A07.BOF And DY_A07.EOF Then Check_Executable = True: Exit Function
    End If
    '
    A_Sql$ = "Select * From A07"
    A_Sql$ = A_Sql$ & " where A0701='" & G_UserGroup & "'"
    A_Sql$ = A_Sql$ & " and A0702='" & a$ & "'"
    A_Sql$ = A_Sql$ & " order by A0701,A0704,A0702"
    CreateDynasetODBC DB_ARTHGUI, DY_A07, A_Sql$, "DY_A07", True
    If Not (DY_A07.BOF And DY_A07.EOF) Then
       Select Case UCase(Trim(DY_A07.Fields("A0703") & ""))
           Case "Y"
                Check_Executable = True
           Case "N"
                A_Msg$ = G_Authority
       End Select
    Else
       A_Msg$ = G_Authority
    End If
End Function
