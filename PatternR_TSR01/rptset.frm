VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form rptset 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "印表設定"
   ClientHeight    =   3030
   ClientLeft      =   300
   ClientTop       =   435
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3030
   ScaleWidth      =   6480
   Begin VsOcxLib.VideoSoftElastic Vse_background 
      Height          =   2655
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6480
      _Version        =   327680
      _ExtentX        =   11430
      _ExtentY        =   4683
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Align           =   5
      BevelOuter      =   6
      BevelOuterDir   =   1
      Begin VB.CheckBox chk_DefaultPrinter 
         Caption         =   "將印表機設為預設印表機"
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   2190
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Frame Fra_Printer 
         Caption         =   "列印設定"
         Height          =   2115
         Left            =   6510
         TabIndex        =   23
         Top             =   0
         Width           =   6345
         Begin VB.CommandButton Cmd_SCancel 
            Appearance      =   0  'Flat
            Caption         =   "取消(&L)"
            Height          =   360
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1680
            Width           =   1330
         End
         Begin VB.CommandButton Cmd_SOk 
            Appearance      =   0  'Flat
            Caption         =   "確定(&O)"
            Height          =   360
            Left            =   4920
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   330
            Width           =   1330
         End
         Begin VB.OptionButton Opt_VStyle 
            Caption         =   "直印"
            Height          =   360
            Left            =   1320
            TabIndex        =   9
            Top             =   1200
            Width           =   1800
         End
         Begin VB.ComboBox Cbo_Size 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   720
            Width           =   3525
         End
         Begin VB.ComboBox Cbo_Printer 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   330
            Width           =   3525
         End
         Begin VB.OptionButton Opt_HStyle 
            Caption         =   "橫印"
            Height          =   360
            Left            =   1320
            TabIndex        =   10
            Top             =   1680
            Width           =   1800
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   3150
            Top             =   1620
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   3120
            Top             =   1140
            Width           =   480
         End
         Begin VB.Label Lbl_Printer 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "印表機"
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   90
            TabIndex        =   26
            Top             =   360
            Width           =   1300
         End
         Begin VB.Label Lbl_Paper 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "紙張大小"
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   90
            TabIndex        =   25
            Top             =   780
            Width           =   1305
         End
         Begin VB.Label Lbl_Style 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "列印方向"
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   90
            TabIndex        =   24
            Top             =   1230
            Width           =   1305
         End
      End
      Begin VB.CommandButton cmd_printer 
         Appearance      =   0  'Flat
         Caption         =   "印表機設定(&P)"
         Height          =   420
         Left            =   4670
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   1730
      End
      Begin VB.CommandButton cmd_font 
         Appearance      =   0  'Flat
         Caption         =   "字型設定(&F)"
         Height          =   420
         Left            =   4670
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   570
         Width           =   1730
      End
      Begin VB.CommandButton cmd_ok 
         Appearance      =   0  'Flat
         Caption         =   "確定(&Enter)"
         Height          =   420
         Left            =   4670
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1020
         Width           =   1730
      End
      Begin VB.CommandButton cmd_exit 
         Appearance      =   0  'Flat
         Caption         =   "取消(Es&c)"
         Height          =   420
         Left            =   4670
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2130
         Width           =   1730
      End
      Begin VB.TextBox rptfontsize 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2220
         MaxLength       =   5
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   900
         Width           =   2370
      End
      Begin VB.TextBox rptline 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   2220
         MaxLength       =   3
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1710
         Width           =   2370
      End
      Begin MSComDlg.CommonDialog Com_Dialog 
         Left            =   6540
         Top             =   2700
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label prtwidth 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   2220
         TabIndex        =   20
         Top             =   1320
         Width           =   2370
      End
      Begin VB.Label lbl_Width 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "每行列印所需寬度"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   2130
      End
      Begin VB.Label rptneedwidth 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   2220
         TabIndex        =   22
         Top             =   120
         Width           =   2370
      End
      Begin VB.Label lbl_Col_Words 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "目前每行可印字數"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   60
         TabIndex        =   21
         Top             =   1380
         Width           =   2130
      End
      Begin VB.Label rptfont 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   2220
         TabIndex        =   19
         Top             =   510
         Width           =   2370
      End
      Begin VB.Label lbl_FontName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "字體選擇"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   60
         TabIndex        =   18
         Top             =   570
         Width           =   2130
      End
      Begin VB.Label lbl_Page_Words 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "目前每頁可印行數"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   60
         TabIndex        =   17
         Top             =   1770
         Width           =   2130
      End
      Begin VB.Label lbl_FontSize 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "字型選擇"
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   60
         TabIndex        =   16
         Top             =   960
         Width           =   2130
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   2655
      Width           =   6480
      _ExtentX        =   11430
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
Attribute VB_Name = "rptset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim retcode As Integer
Dim DB_LGUI As Database
Dim TB_CommIni As Recordset
Dim M_FontError As String
Dim M_PageError As String
Dim M_SizeError As String
Dim M_PaperSize() As String
Dim M_PageLines$
Dim M_DeviceName As String
Dim M_Paper As Integer
Dim M_Orientation As Integer
Dim M_PlatForm As Long
Dim M_FontName As String
Dim M_FontSize As String
Dim M_Position As String
Dim M_HaveDefault As Boolean

Const M_Win95 = 1
Const M_WinNT = 2

Sub SetDefaultValue()
On Local Error Resume Next
Dim A_DefaultStr$, A_ReportName$

    G_PrinterName = Printer.DeviceName

    A_ReportName$ = IIf(Me.Tag <> "", Me.Tag, App.EXEName)
    '
    A_DefaultStr$ = GetIniStr("ReportDefault", A_ReportName$, "GUI.INI")
    M_HaveDefault = (Trim(A_DefaultStr$) <> "")
    M_PageLines$ = ""
    '
    If M_HaveDefault Then
       StrCut A_DefaultStr$, "/", M_FontName, A_DefaultStr$
       StrCut A_DefaultStr$, "/", M_FontSize, A_DefaultStr$
       StrCut A_DefaultStr$, "/", M_Position, M_PageLines$
       If M_FontName <> "" Then
          rptfont.Caption = M_FontName
       Else
          rptfont.Caption = G_FontName
       End If
       If M_FontSize <> "" Then
          rptfontsize.text = Format(M_FontSize, "##0.00")
       Else
          rptfontsize.text = Format(G_FontSize, "##0.00")
       End If
       If M_PageLines$ <> "" Then
          rptline.text = Format(CvrTxt2Num(M_PageLines$), "##0")
       Else
          rptline.text = Format(G_PageSize, "##0")
       End If
       If M_Position <> "" Then
          If M_Position = 1 Then Opt_VStyle.Value = True
          If M_Position = 2 Then Opt_HStyle.Value = True
          If Printer.Orientation <> M_Position Then
             Printer.Orientation = M_Position
          End If
       Else
          If M_Orientation = 1 Then Opt_VStyle.Value = True
          If M_Orientation = 2 Then Opt_HStyle.Value = True
       End If
    Else
       rptfont.Caption = G_FontName
       rptfontsize.text = Format(G_FontSize, "##0.00")
       If M_Orientation = 1 Then Opt_VStyle.Value = True
       If M_Orientation = 2 Then Opt_HStyle.Value = True
    End If
    rptneedwidth.Caption = Format(G_RptNeedWidth, "##0")
End Sub

Sub SaveDefaultValue()
Dim A_DefaultStr$, A_ReportName$

    A_DefaultStr$ = Trim(rptfont.Caption)
    A_DefaultStr$ = A_DefaultStr$ & "/" & Trim(rptfontsize.text)
    A_DefaultStr$ = A_DefaultStr$ & "/" & M_Orientation
    A_DefaultStr$ = A_DefaultStr$ & "/" & Trim(G_PageSize)
    '
    A_ReportName$ = IIf(Me.Tag <> "", Me.Tag, App.EXEName)
    '
    retcode = OSWritePrivateProfileString%("ReportDefault", A_ReportName$, A_DefaultStr$, "GUI.INI")
End Sub


Private Sub Set_Property()
    rptset.FontBold = False
    Form_Property rptset, GetRptSet("Rptset", "formtitle"), G_Font_Name
    
    Label_Property lbl_Width, GetRptSet("Rptset", "needwidth"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property lbl_FontName, GetRptSet("Rptset", "fontname"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property lbl_FontSize, GetRptSet("Rptset", "fontsize"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property lbl_Col_Words, GetRptSet("Rptset", "col_words"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property lbl_Page_Words, GetRptSet("Rptset", "page_words"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Printer, GetRptSet("PanelDescpt", "printer"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Paper, GetRptSet("Rptset", "papersize"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Style, GetRptSet("Rptset", "orientation"), G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property rptneedwidth, "", G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property rptfont, "", G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property prtwidth, "", G_Label_Color, G_Font_Size, G_Font_Name
    
    Dim A_Caption$, A_Value$
    A_Caption$ = GetRptSet("Rptset", "DefaultPrinterCaption")  '將印表機設為預設印表機
    A_Value$ = GetRptSet("Rptset", "UpdateDefaultPrinter")
    If Trim(A_Caption$) = "" Then A_Caption$ = chk_DefaultPrinter.Caption
    Checkbox_Property chk_DefaultPrinter, A_Caption$, G_Font_Size, G_Font_Name
    chk_DefaultPrinter.Value = IIf(UCase(A_Value$) = "N", Unchecked, Checked)
    
    Frame_Property Fra_Printer, GetRptSet("Rptset", "prtset"), G_Font_Size, G_Font_Name
    Fra_Printer.Visible = False
    Fra_Printer.Move 30, 30, 6345, 2115
    
    Option_Property Opt_VStyle, GetRptSet("Rptset", "portrait"), G_Font_Size, G_Font_Name
    Option_Property Opt_HStyle, GetRptSet("Rptset", "landscape"), G_Font_Size, G_Font_Name
    
    Text_Property rptfontsize, 5, G_Font_Name
    Text_Property rptline, 3, G_Font_Name
    
    ComboBox_Property Cbo_Printer, "9.75", G_Font_Name
    ComboBox_Property Cbo_Size, "9.75", G_Font_Name

    Command_Property cmd_printer, GetRptSet("Rptset", "printer"), G_Font_Name
    Command_Property cmd_font, GetRptSet("Rptset", "font"), G_Font_Name
    Command_Property cmd_ok, GetRptSet("Rptset", "ok"), G_Font_Name
    Command_Property Cmd_exit, GetRptSet("Rptset", "cancel"), G_Font_Name
    Command_Property Cmd_SOk, GetRptSet("Rptset", "sok"), G_Font_Name
    Command_Property Cmd_SCancel, GetRptSet("Rptset", "scancel"), G_Font_Name
    
    StatusBar_ProPerty Sts_MsgLine
    VSElastic_Property Vse_Background
    
    M_FontError = GetRptSet("Rptset", "fonterror")
    M_PageError = GetRptSet("Rptset", "pageerror")
    M_SizeError = GetRptSet("Rptset", "sizeerror")
End Sub


Private Sub SetPaperSizeArray()
Erase M_PaperSize
ReDim M_PaperSize(0 To 50)

    M_PaperSize(0) = "Letter, 8 1/2 x 11 in.                                                           1"
    M_PaperSize(1) = "Letter Small, 8 1/2 x 11 in.                                                     2"
    M_PaperSize(2) = "Tabloid, 11 x 17 in.                                                             3"
    M_PaperSize(3) = "Ledger, 17 x 11 in.                                                              4"
    M_PaperSize(4) = "Legal, 8 1/2 x 14 in.                                                            5"
    M_PaperSize(5) = "Statement, 5 1/2 x 8 1/2 in.                                                     6"
    M_PaperSize(6) = "Executive, 7 1/2 x 10 1/2 in.                                                    7"
    M_PaperSize(7) = "A3, 297 x 420 mm                                                                 8"
    M_PaperSize(8) = "A4, 210 x 297 mm                                                                 9"
    M_PaperSize(9) = "A4 Small, 210 x 297 mm                                                          10"
    M_PaperSize(10) = "A5, 148 x 210 mm                                                               11"
    M_PaperSize(11) = "B4, 250 x 354 mm                                                               12"
    M_PaperSize(12) = "B5, 182 x 257 mm                                                               13"
    M_PaperSize(13) = "Folio, 8 1/2 x 13 in.                                                          14"
    M_PaperSize(14) = "Quarto, 215 x 275 mm                                                           15"
    M_PaperSize(15) = "10 x 14 in.                                                                    16"
    M_PaperSize(16) = "11 x 17 in.                                                                    17"
    M_PaperSize(17) = "Note, 8 1/2 x 11 in.                                                           18"
    M_PaperSize(18) = "Envelope #9, 3 7/8 x 8 7/8 in.                                                 19"
    M_PaperSize(19) = "Envelope #10, 4 1/8 x 9 1/2 in.                                                20"
    M_PaperSize(20) = "Envelope #11, 4 1/2 x 10 3/8 in.                                               21"
    M_PaperSize(21) = "Envelope #12, 4 1/2 x 11 in.                                                   22"
    M_PaperSize(22) = "Envelope #14, 5 x 11 1/2 in.                                                   23"
    M_PaperSize(23) = "C size sheet                                                                   24"
    M_PaperSize(24) = "D size sheet                                                                   25"
    M_PaperSize(25) = "E size sheet                                                                   26"
    M_PaperSize(26) = "Envelope DL, 110 x 220 mm                                                      27"
    M_PaperSize(27) = "Envelope C3, 324 x 458 mm                                                      29"
    M_PaperSize(28) = "Envelope C4, 229 x 324 mm                                                      30"
    M_PaperSize(29) = "Envelope C5, 162 x 229 mm                                                      28"
    M_PaperSize(30) = "Envelope C6, 114 x 162 mm                                                      31"
    M_PaperSize(31) = "Envelope C65, 114 x 229 mm                                                     32"
    M_PaperSize(32) = "Envelope B4, 250 x 353 mm                                                      33"
    M_PaperSize(33) = "Envelope B5, 176 x 250 mm                                                      34"
    M_PaperSize(34) = "Envelope B6, 176 x 125 mm                                                      35"
    M_PaperSize(35) = "Envelope, 110 x 230 mm                                                         36"
    M_PaperSize(36) = "Envelope Monarch, 3 7/8 x 7 1/2 in.                                            37"
    M_PaperSize(37) = "Envelope, 3 5/8 x 6 1/2 in.                                                    38"
    M_PaperSize(38) = "U.S. Standard Fanfold, 14 7/8 x 11 in.                                         39"
    M_PaperSize(39) = "German Standard Fanfold, 8 1/2 x 12 in.                                        40"
    M_PaperSize(40) = "German Legal Fanfold, 8 1/2 x 13 in.                                           41"
    M_PaperSize(41) = "User-defined                                                                  256"
    M_PaperSize(42) = "15 x 11 英吋 連續紙(13.2英吋)                                                 288"
    M_PaperSize(43) = "15 x 11 英吋 連續紙(13.6英吋)                                                 289"
End Sub

Private Sub SettingActivePrinter()
On Local Error Resume Next
Dim I%
    
    'Set Active Printer
    For I% = 0 To Printers.Count - 1
        Debug.Print Printers(I%).DeviceName
        If Printers(I%).DeviceName = M_DeviceName Then
           Set Printer = Printers(I%)
           Exit For
        End If
    Next I%
    'Set Active Printer PaperSize
    Printer.PaperSize = M_Paper
    'Set Active Printer Orientation
    Printer.Orientation = M_Orientation
End Sub

Private Sub Cbo_Printer_DropDown()
    DoEvents
End Sub

Private Sub Cbo_Printer_GotFocus()
    TextGotFocus
End Sub

Private Sub Cbo_Printer_LostFocus()
Dim I%, A_Index%, A_PaperSize%
Dim x As Printer
    
    TextLostFocus
    
    Me.MousePointer = HOURGLASS
    If Cbo_Printer.text <> Cbo_Printer.Tag Then
       Cbo_Printer.Tag = Cbo_Printer.text
       'Paper Size Counts
       For I% = 0 To Printers.Count - 1
           If Printers(I%).DeviceName = Cbo_Printer.text Then
              Set Printer = Printers(I%)
              Exit For
           End If
       Next I%
       'Paper Size
       Set x = Printer
       A_PaperSize% = x.PaperSize
       If Not M_HaveDefault Then M_Orientation = x.Orientation
       A_Index% = -1
       Cbo_Size.Clear
       On Error Resume Next
       For I% = 0 To UBound(M_PaperSize)
           x.PaperSize = Val(Right$(M_PaperSize(I%), 3))
           If Err Then
              Err = 0
           Else
              Cbo_Size.AddItem M_PaperSize(I%)
              If Val(Right$(M_PaperSize(I%), 3)) = A_PaperSize% Then
                 A_Index% = Cbo_Size.ListCount - 1
              End If
           End If
       Next I%
       Cbo_Size.ListIndex = IIf(A_Index% <> -1, A_Index%, 0)
       'Set Printer Orientation
       x.Orientation = M_Orientation
       If M_Orientation = 1 Then Opt_VStyle.Value = True
       If M_Orientation = 2 Then Opt_HStyle.Value = True
    End If
    Set x = Nothing
    Me.MousePointer = Default
End Sub


Private Sub Cbo_Size_DropDown()
    DoEvents
End Sub

Private Sub Cbo_Size_GotFocus()
    TextGotFocus
End Sub


Private Sub Cbo_Size_LostFocus()
    TextLostFocus
End Sub

Private Sub Cmd_Exit_Click()
    G_RptSet = False
    G_SetDefaultPrinter = -1
    G_PrinterName = ""
    
    DB_LGUI.Close
    Set DB_LGUI = Nothing
    
    Unload Me
End Sub

Private Sub cmd_font_Click()
On Local Error Resume Next

    Com_Dialog.CancelError = True
    Com_Dialog.FontName = G_FontName
    Com_Dialog.Flags = &H2&
    Com_Dialog.Action = 4
    If Err = 32755 Then
       Err = 0
    Else
       Printer.FontName = Com_Dialog.FontName
       Printer.FontSize = Com_Dialog.FontSize
       rptfont.Caption = Com_Dialog.FontName
       rptfontsize.text = Com_Dialog.FontSize
       G_FontName = rptfont.Caption
       G_FontSize = rptfontsize.text
       RptCaculate
    End If
    rptfontsize.SetFocus
End Sub

Private Sub cmd_OK_Click()
On Error Resume Next
    Dim a_DgDef, a_response
    Printer.EndDoc
    If M_PlatForm = M_Win95 Then
       SettingActivePrinter
    Else
       Printer.Orientation = M_Orientation
    End If
'Update by Li-Ming Sung
    If Val(prtwidth.Caption) < Val(rptneedwidth.Caption) Then
       a_DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2
       a_response = MsgBox(M_SizeError, a_DgDef)
       If a_response <> IDYES Then
          Exit Sub
       End If
    End If
    G_FontName = Trim(rptfont.Caption)
    G_FontSize = Val(rptfontsize.text)
    G_PageSize = Val(rptline.text)
    SaveDefaultValue
    '
    G_RptSet = True
    G_SetDefaultPrinter = Checked 'chk_DefaultPrinter.Value
    DB_LGUI.Close
    Set DB_LGUI = Nothing
    Unload Me
End Sub

Private Sub cmd_printer_Click()
On Error Resume Next
Dim I%, A_Index%
Dim x As Printer

    Me.MousePointer = HOURGLASS
    '
    If M_PlatForm = M_WinNT Then
       Fra_Printer.Visible = False
       Printer.EndDoc
       Printer.TrackDefault = True
       Com_Dialog.PrinterDefault = True
       Com_Dialog.CancelError = False
       Com_Dialog.Orientation = M_Orientation
       Com_Dialog.Action = 5
       M_Orientation = Com_Dialog.Orientation
       Printer.Orientation = M_Orientation
       RptCaculate
       rptfontsize.SetFocus
    Else
       'Printer Counts
       Cbo_Printer.Clear
       A_Index% = -1
       For I% = 0 To Printers.Count - 1
           Cbo_Printer.AddItem Printers(I%).DeviceName
           If Printers(I%).DeviceName = M_DeviceName Then
              A_Index% = I%
           End If
       Next I%
       Cbo_Printer.ListIndex = IIf(A_Index% <> -1, A_Index%, 0)
       Cbo_Printer.Tag = Cbo_Printer.text
       'Paper Size Counts
       A_Index% = -1
       Set x = Printer
       Cbo_Size.Clear
       On Error Resume Next
       For I% = 0 To UBound(M_PaperSize)
           x.PaperSize = Val(Right$(M_PaperSize(I%), 3))
           If Err Then
              Err = 0
           Else
              Cbo_Size.AddItem M_PaperSize(I%)
              If Val(Right$(M_PaperSize(I%), 3)) = M_Paper Then
                 A_Index% = Cbo_Size.ListCount - 1
              End If
           End If
       Next I%
       Cbo_Size.ListIndex = IIf(A_Index% <> -1, A_Index%, 0)
       'Print Orientation
       If M_Orientation = 1 Then Opt_VStyle.Value = True
       If M_Orientation = 2 Then Opt_HStyle.Value = True
       '
       cmd_printer.Enabled = False
       cmd_font.Enabled = False
       cmd_ok.Enabled = False
       Cmd_exit.Enabled = False
       rptfontsize.Enabled = False
       rptline.Enabled = False
       '
       Fra_Printer.Visible = True
       Cbo_Printer.SetFocus
       Set x = Nothing
    End If
    '
    Me.MousePointer = Default
End Sub

Private Sub Cmd_SCancel_Click()
    cmd_printer.Enabled = True
    cmd_font.Enabled = True
    cmd_ok.Enabled = True
    Cmd_exit.Enabled = True
    rptfontsize.Enabled = True
    rptline.Enabled = True
    Fra_Printer.Visible = False
    rptfontsize.SetFocus
End Sub

Private Sub Cmd_SOk_Click()
    cmd_printer.Enabled = True
    cmd_font.Enabled = True
    cmd_ok.Enabled = True
    Cmd_exit.Enabled = True
    rptfontsize.Enabled = True
    rptline.Enabled = True
    Fra_Printer.Visible = False
    M_DeviceName = Cbo_Printer.text
    M_Paper = Val(Right$(Cbo_Size.text, 3))
    M_Orientation = IIf(Opt_VStyle.Value, 1, 2)
    DoEvents
    RptCaculate
    rptfontsize.SetFocus
End Sub


Private Sub Form_Activate()
    DoEvents
    Sts_MsgLine.Panels(1) = GetRptSet("Rptset", "formtitle") 'S991019055
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    Sts_MsgLine.Refresh
    SetDefaultValue
    RptCaculate
    Me.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim ShiftDown, AltDown, CtrlDown, I%
  Const SHIFT_MASK = 1: Const CTRL_MASK = 2: Const ALT_MASK = 4
  ShiftDown = (Shift And SHIFT_MASK) > 0
  AltDown = (Shift And ALT_MASK) > 0
  CtrlDown = (Shift And CTRL_MASK) > 0

  Select Case KeyCode
         Case KEY_UP
              If Not TypeOf ActiveControl Is ComboBox Then
                 KeyCode = 0
                 DoEvents
                 SendKeys "+{TAB}"
              End If
         Case KEY_DOWN
              If Not TypeOf ActiveControl Is ComboBox Then
                 KeyCode = 0
                 DoEvents
                 SendKeys "{TAB}"
              End If
         Case KEY_ESCAPE
              KeyCode = 0
              If Cmd_exit.Visible And Cmd_exit.Enabled Then
                 Cmd_exit.SetFocus
                 DoEvents
                 SendKeys "{Enter}"
              End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case KEY_RETURN
            KeyAscii = 0
            SendKeys "{TAB}"
        Case KEY_ESCAPE
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Load()
    Me.MousePointer = HOURGLASS
    Me.Move (Screen.Width - Me.Width) \ 6, (Screen.Height - Me.Height) \ 6
    's980311021
    'Sts_MsgLine.Panels(1) = Me.Caption
    'Sts_MsgLine.Panels(1) = GetCaption("Rptset", "formtitle", "印表設定") 'S991019055
    OpenCommLDB
    Set_Property
    'Get PC PlatForm
    M_PlatForm = GetWinPlatform()
    If M_PlatForm = M_WinNT Then
       If Not IsWindowsNT4WithoutSP5() Then
          MsgBox "建議您先安裝NT4.0 Server Pack 5,以確保報表格式無誤 !", vbInformation, Me.Caption
       End If
    End If
    '
    If M_PlatForm = M_Win95 Then
       SetPaperSizeArray
       'Keep Old Print Setting
       M_DeviceName = Printer.DeviceName
       M_Paper = Printer.PaperSize
    End If
    '
    M_Orientation = Printer.Orientation
    '
    Me.MousePointer = Default
End Sub
Private Sub RptCaculate()
    Me.MousePointer = HOURGLASS
    Printer.ScaleMode = 1
    If M_PlatForm = M_Win95 Then
       SettingActivePrinter
    End If
    Printer.FontName = rptfont.Caption
    Printer.FontSize = rptfontsize.text
    If Val(rptfontsize.text) <= 0 Then rptfontsize.text = Format(Com_Dialog.FontSize, "##0.00")
    If Trim(Com_Dialog.FontName) <> "" Then G_FontName = Com_Dialog.FontName
    prtwidth.Caption = Format(Printer.ScaleWidth / Printer.TextWidth("A"), "##0")
    If M_PageLines$ = "" Then rptline.text = Format(Printer.ScaleHeight / Printer.TextHeight("A"), "##0")
    rptfont.Tag = Format(Printer.ScaleHeight / Printer.TextHeight("A"), "##0")
    Me.MousePointer = Default
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        G_RptSet = False
        G_SetDefaultPrinter = -1
        G_PrinterName = ""
    End If
End Sub

Private Sub Opt_HStyle_Click()
    M_Orientation = 2
    RptCaculate
End Sub

Private Sub Opt_VStyle_Click()
    M_Orientation = 1
    RptCaculate
End Sub

Private Sub rptfontsize_GotFocus()
    TextGotFocus
End Sub

Private Sub rptfontsize_LostFocus()
    TextLostFocus
    If Val(rptfontsize.text) <= 0 Then
       MsgBox M_FontError
       rptfontsize.text = Format(G_FontSize, "##0.00")
    End If
    G_FontSize = Val(rptfontsize.text)
    RptCaculate
    '字型大小變更,重算可印行數
    If Trim(rptfontsize.Tag) > "" And Val(rptfontsize.text) <> Val(rptfontsize.Tag) Then
        rptline.text = Format(Printer.ScaleHeight / Printer.TextHeight("A"), "##0")
    End If
    rptfontsize.Tag = Val(rptfontsize.text)
End Sub

Private Sub rptline_GotFocus()
    TextGotFocus
End Sub

Private Sub rptline_LostFocus()
Dim A_Line%
    
    TextLostFocus
    A_Line% = Val(rptfont.Tag)
    If Val(rptline.text) > A_Line% Or Val(rptline.text) <= 0 Then
       MsgBox M_PageError
       rptline.text = Format(A_Line%, "##0")
    End If
    G_PageSize = A_Line%
End Sub


Function GetRptSet(ByVal Section$, ByVal Topic$) As String
    GetRptSet = " "
    If Trim(DB_LGUI.Connect) <> "" Then
        Dim A_Sql$
        A_Sql$ = "SELECT TOPICVALUE FROM INI"
        A_Sql$ = A_Sql$ & " WHERE SECTION='" & Section$ & "'"
        A_Sql$ = A_Sql$ & " AND TOPIC='" & Topic$ & "'"
        Set DY_INICommon = DB_LGUI.OpenRecordset(A_Sql$, dbOpenSnapshot, dbSQLPassThrough)
        If Not (DY_INICommon.BOF And DY_INICommon.EOF) Then
            GetRptSet = Trim(DY_INICommon.Fields("TOPICVALUE") & "")
        End If
        DY_INICommon.Close
    Else
        TB_CommIni.Seek "=", Section$, Topic$
        If Not TB_CommIni.NoMatch Then
           GetRptSet = TB_CommIni.Fields("TOPICVALUE") & ""
        End If
    End If
End Function


Sub OpenCommLDB()
On Local Error Resume Next
Dim A_Path As String
Dim A_ConnectMethod As String
    
    Screen.MousePointer = HOURGLASS
   'Pick Local INI DataPath String (LGUI.mdb)
    A_Path = GetIniStr("DBPath", "Path3", "GUI.INI")
    A_ConnectMethod = GetIniStr("DBPath", "Connect3", "GUI.INI")
    Set DB_LGUI = GetEngine.OpenDatabase(A_Path, False, False, A_ConnectMethod)
    If Err Then
       If Trim$(A_ConnectMethod) = "" Then   'Access DataBase
          If Err = 3043 Then
             Err = 0
             DB_LGUI.Close
             Set DB_LGUI = GetEngine.OpenDatabase(A_Path, False, False, A_ConnectMethod)
          ElseIf Err = 3049 Then
             Err = 0
             GetEngine.RepairDatabase A_Path
             Set DB_LGUI = GetEngine.OpenDatabase(A_Path, False, False, A_ConnectMethod)
          End If
       End If
    End If
    If Err Then
       MsgBox Error(Err), MB_ICONEXCLAMATION, App.Title
       End
    End If
    If Trim$(A_ConnectMethod) <> "" Then DB_LGUI.QueryTimeout = 0
    'Open Table
    If Trim(DB_LGUI.Connect) = "" Then
        Set TB_CommIni = DB_LGUI.OpenRecordset("INI", dbOpenTable)
        TB_CommIni.index = "INI"
    End If
    Screen.MousePointer = Default
End Sub

