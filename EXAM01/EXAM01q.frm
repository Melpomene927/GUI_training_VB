VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frm_EXAM01q 
   Caption         =   "�|�p��ظ�Ƭd��"
   ClientHeight    =   2580
   ClientLeft      =   3270
   ClientTop       =   2685
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EXAM01q.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   6225
   Begin VsOcxLib.VideoSoftElastic Vse_Background 
      Height          =   2205
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   6225
      _Version        =   327680
      _ExtentX        =   10980
      _ExtentY        =   3889
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ConvInfo        =   1418783674
      Align           =   5
      BevelOuter      =   6
      Picture         =   "EXAM01q.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "EXAM01q.frx":0326
      Begin VB.TextBox Txt_A1601e 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   1
         Top             =   90
         Width           =   1395
      End
      Begin VB.TextBox Txt_A1601s 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   0
         Top             =   90
         Width           =   1395
      End
      Begin VB.TextBox Txt_A1602 
         DataField       =   "z"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1395
         TabIndex        =   5
         Top             =   1380
         Width           =   3240
      End
      Begin VB.TextBox Txt_A1609 
         DataField       =   "z"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1395
         TabIndex        =   4
         Top             =   960
         Width           =   3240
      End
      Begin VB.TextBox Txt_A1628e 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   3
         Top             =   540
         Width           =   1395
      End
      Begin VB.TextBox Txt_A1628s 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1395
         MaxLength       =   10
         TabIndex        =   2
         Top             =   540
         Width           =   1395
      End
      Begin Threed.SSCommand Cmd_Help 
         Height          =   405
         Left            =   4740
         TabIndex        =   6
         Top             =   90
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "���U F1"
         ForeColor       =   0
      End
      Begin Threed.SSCommand Cmd_Add 
         Height          =   405
         Left            =   4740
         TabIndex        =   8
         Top             =   990
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�s�WF4"
         ForeColor       =   0
      End
      Begin Threed.SSCommand Cmd_Exit 
         Height          =   405
         Left            =   4740
         TabIndex        =   9
         Top             =   1680
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "����Esc"
         ForeColor       =   0
      End
      Begin Threed.SSCommand Cmd_Ok 
         Height          =   405
         Left            =   4740
         TabIndex        =   7
         Top             =   540
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�d��F2"
         ForeColor       =   0
      End
      Begin VB.Label Lbl_A1602 
         Caption         =   "�Ȥ�W��"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   17
         Top             =   1410
         Width           =   1380
      End
      Begin VB.Label Lbl_A1609 
         Caption         =   "������/�νs"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   16
         Top             =   990
         Width           =   1380
      End
      Begin VB.Label Lbl_A1628 
         Caption         =   "�ͤ�/���ߤ�"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   14
         Top             =   570
         Width           =   1380
      End
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "��"
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   1
         Left            =   2850
         TabIndex        =   15
         Top             =   570
         Width           =   375
      End
      Begin VB.Label Lbl_A1601 
         Caption         =   "�Ȥ�s��"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   45
         TabIndex        =   12
         Top             =   135
         Width           =   1380
      End
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "System"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   13
         Top             =   150
         Width           =   300
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   2205
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_EXAM01q"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'========================================================================
'   Coding Rule
'========================================================================
'�b���B�w�q���Ҧ��ܼ�, �@�ߥHM�}�Y, �pM_AAA$, M_BBB#, M_CCC&
'�B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
' $: ��r
' #: �Ҧ��Ʀr�B��(���B�μƶq)
' &: �{���j���ܼ�
' %: ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
' �ť�: �N��VARIENT, �ʺA�ܼ�
'========================================================================
'�۩w�ܼ�
'Dim m_A4101Flag%
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

'���n�ܼ�
Dim m_FieldError%    '���ܼƦb�P�_���O�_���~, �����^�����줧�ʧ@
Dim m_ExitTrigger%   '���ܼƦb�P�_������O�_�QĲ�o, �N����ثe���b�B�z���@�~
'========================================================================
'====================================
'   User Defined Fucntions
'====================================

'========================================================================
' Procedure : CheckRoutine_A1601 (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : check data correctness of Txt_A1601s & Txt_A1601e
' Details   : check over the conditions as below:
'                   1.both A1601s, A1601e DataRange not exceed each other
'========================================================================
Private Function CheckRoutine_A1601() As Boolean
    CheckRoutine_A1601 = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim$(Txt_A1601e) = "" Then Txt_A1601e = Txt_A1601s
    
    If Not CheckDataRange(sts_msgline, Trim$(Txt_A1601s), Trim$(Txt_A1601e)) Then
        '==================
        'if from s to e
        'do not focus back (since it's correct to entering from s to e)
        '==================
        If ActiveControl.TabIndex = Txt_A1601e.TabIndex Then
            '�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
            m_FieldError% = Txt_A1601e.TabIndex
        Else
            m_FieldError% = Txt_A1601s.TabIndex
            Txt_A1601s.SetFocus
        End If
        Exit Function
    End If
       
    CheckRoutine_A1601 = True
End Function

'========================================================================
' Procedure : CheckRoutine_A1628s (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : check data correctness of Txt_A1628s
' Details   : check over the conditions as below:
'                   1.valid date format
'                   2.date range not exceed Txt_A1628e
'========================================================================
Private Function CheckRoutine_A1628s() As Boolean
    CheckRoutine_A1628s = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
'    If Trim(Txt_A1628s) = "" Then
'       Txt_A1628s = GetCurrentDay(0)
'    Else
    If Not Trim(Txt_A1628s) = "" Then
       If Not IsDateValidate(Txt_A1628s) Then
          sts_msgline.Panels(1) = G_Pnl_A1628$ & G_DateError
          m_FieldError% = Txt_A1628s.TabIndex
          Txt_A1628s.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(sts_msgline, Trim$(Txt_A1628s), Trim$(Txt_A1628e)) Then
       If ActiveControl.TabIndex = Txt_A1628e.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A1628s.TabIndex
       Else
          m_FieldError% = Txt_A1628s.TabIndex
          Txt_A1628s.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A1628s = True
End Function

'========================================================================
' Module    : frm_EXAM01q
' Procedure : CheckRoutine_A1628e
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : check data correctness of Txt_A1628e
' Details   : check over the conditions as below:
'                   1.valid date format
'                   2.date range not exceed Txt_A1628s
'========================================================================
Private Function CheckRoutine_A1628e() As Boolean
    CheckRoutine_A1628e = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
'    If Trim(Txt_A1628e) = "" Then
'       Txt_A1628e = GetCurrentDay(0)
'    Else
    If Not Trim(Txt_A1628e) = "" Then
       If Not IsDateValidate(Txt_A1628e) Then
          sts_msgline.Panels(1) = G_Pnl_A1628$ & G_DateError
          m_FieldError% = Txt_A1628e.TabIndex
          Txt_A1628e.SetFocus
          Exit Function
       End If
    End If
    
    If Not CheckDateRange(sts_msgline, Trim$(Txt_A1628s), Trim$(Txt_A1628e)) Then
       If ActiveControl.TabIndex = Txt_A1628s.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A1628s.TabIndex
       Else
          m_FieldError% = Txt_A1628e.TabIndex
          Txt_A1628e.SetFocus
       End If
       Exit Function
    End If
    
    CheckRoutine_A1628e = True
End Function

'========================================================================
' Procedure : DataPrepare_A16 (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Prepare data for Help Spread
' Details   :
'       @modified: Do Nothing, Since Help procedure has been done
'                  with frm_GD
'========================================================================
Private Sub DataPrepare_A16(Txt As TextBox)
Dim A_Sql$      'SQL Message
Dim A_A1601$    'PKey of A16 (�Ȥ�/�t�ӽs��)
    
    
    Me.MousePointer = HOURGLASS
    
'    A_A1601$ = Trim(Txt)    'parameter is

    '�}�_�ɮ�
    'concate SQL Message
'    A_Sql$ = "Select A1601, A0202 A From A02"

    'generate wildcard compare SQL Statement
'    A_Sql$ = A_Sql$ & " Where A1601 Like '" & A_A1601 & GetLikeStr(DB_ARTHGUI, True) & "'"
'    A_Sql$ = A_Sql$ & " Order by A1601"
'
    'Old statements that belongs to EXAM01(A15)
'    If Len(A_A1502$) > 4 Then
'       A_Sql$ = A_Sql$ & " and A1502='" & Mid$(A_A1502$, 1, 4) & "'"
'       A_Sql$ = A_Sql$ & " and A1503 Like '" & Mid$(A_A1502$, 5) & GetLikeStr(DB_ARTHGUI, True) & "'"
'    Else
'       A_Sql$ = A_Sql$ & " and A1502 Like '" & A_A1502$ & GetLikeStr(DB_ARTHGUI, True) & "'"
'    End If
    
    'open dynaset of A02
'    CreateDynasetODBC DB_ARTHGUI, DY_A16, A_Sql$, "DY_A16", True
'    If DY_A16.BOF And DY_A16.EOF Then
'       Me.MousePointer = Default
'       Sts_MsgLine.Panels(1) = G_NoReference
'       Exit Sub
'    End If
    
    
'    With Spd_Help
'
'    '�]�w���U����(Spd_Help)������ݩ�
'        .UnitType = 2          '<---- @!!! Fix property, DO NOT CHANGE IT. !!!
'
'        Spread_Property Spd_Help, 0, 2, WHITE, G_Font_Size, G_Font_Name    'row: 0, col: 2
'        Spread_Col_Property Spd_Help, 1, TextWidth("X") * 7, G_Pnl_A1601$  'col1 header: A1601
'        Spread_Col_Property Spd_Help, 2, TextWidth("X") * 16, G_Pnl_A1601$ 'col2 header: A0202
'        Spread_DataType_Property Spd_Help, 1, SS_CELL_TYPE_EDIT, "", "", 6
'        Spread_DataType_Property Spd_Help, 2, SS_CELL_TYPE_EDIT, "", "", 12
'
'        .Row = -1
'        .Col = -1: .Lock = True
'        .Col = 1: .TypeHAlign = 2
'
'        '�N����\�JSpread��
'        Do Until DY_A16.EOF
'           .MaxRows = .MaxRows + 1
'           .Row = Spd_Help.MaxRows
'           .Col = 1
'           .text = Trim(DY_A16.Fields("A1601") & "")
'           .Col = 2
'           .text = Trim(DY_A16.Fields("A0202") & "")
'           DY_A16.MoveNext
'        Loop
'
'        '�]�w���U��������ܦ�m
'        SetHelpWindowPos Fra_Help, Spd_Help, 330, 90, 4305, 2025
'        .Tag = Txt.TabIndex
'        .SetFocus
'    End With
    
    Me.MousePointer = Default
End Sub

'========================================================================
' Procedure : IsAllFieldsCheck
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Do Full Check over current form's components
' Details   :
'========================================================================
Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    
'����d�ߩΦs�ɫe���N�Ҧ��ˮ����A���@��

    If Not CheckRoutine_A1601 Then Exit Function
    If Not CheckRoutine_A1628s() Then Exit Function
    If Not CheckRoutine_A1628e() Then Exit Function
    
    DoEvents
    
    IsAllFieldsCheck = True
End Function

'========================================================================
' Module    : frm_EXAM01q
' Procedure : OpenMainFile
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Get The Information from Textboxes and push to V-pattern
' Details   : Get 1.A1602 2.A1609 3. A1601 4.A1628 as Where Clause
'             Concate the SQL Statement and Open Dynaset As Global var.
'========================================================================
Private Sub OpenMainFile()
On Local Error GoTo MyError
Dim A_Sql$
Dim A_A1601s$
Dim A_A1601e$
Dim A_A1602$
Dim A_A1609$
Dim A_A1628s$
Dim A_A1628e$
    
'Keep TextBox ��Ʀ��ܼ�
    A_A1601s$ = Trim(Txt_A1601s)
    A_A1601e$ = Trim(Txt_A1601e)
    A_A1628s$ = Trim(DateIn(Txt_A1628s))
    A_A1628e$ = Trim(DateIn(Txt_A1628e))
    A_A1609$ = Trim(Txt_A1609)
    A_A1602$ = Trim(Txt_A1602)
    
    
'�}�Ҹ��
    'get the required Columns as SPEC
    A_Sql$ = "Select A1601,A1602,A1605,A1606,A1609,A1628,"
    A_Sql$ = A_Sql$ & " A16121,A16122,A16123 From A16 Where A1613='1'"
    
    'where clause: A1601
    If A_A1601s$ <> "" Then
       A_Sql$ = A_Sql$ & " And A1601>='" & A_A1601s$ & "' "
    End If
    If A_A1601e$ <> "" Then
       A_Sql$ = A_Sql$ & " And A1601<='" & A_A1601e$ & "' "
    End If
    
    'where clause A1628
    If A_A1628s$ <> "" Then
       A_Sql$ = A_Sql$ & " And A1628>='" & A_A1628s$ & "' "
    End If
    If A_A1628e$ <> "" Then
       A_Sql$ = A_Sql$ & " And A1628<='" & A_A1628e$ & "' "
    End If
    
    'where clause A1609
    If A_A1609$ <> "" Then
       A_Sql$ = A_Sql$ & " And A1609 Like'" & A_A1609$ & GetLikeStr(DB_ARTHGUI, True) & "'"
    End If
        
    'where clause A1602
    If A_A1602$ <> "" Then
       A_Sql$ = A_Sql$ & " And A1602 Like'" & A_A1602$ & GetLikeStr(DB_ARTHGUI, True) & "'"
    End If
    
    A_Sql$ = A_Sql$ & "Order by A1601"
    
    CreateDynasetODBC DB_ARTHGUI, DY_A16, A_Sql$, "DY_A16", True
    Exit Sub
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Procedure : Set_Property  (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Initializing
' Details   : init: 1.form          (caption, font, color)
'                   2.Panel & Label (caption, font, color)
'                   3.Help Frame    (caption, font, color)
'                   4.TextBox       (font, MaxLength)
'                   5.Command button(caption, font)
'========================================================================
Private Sub Set_Property()
    '�]�w��Form�����D,�r�ΤΦ�t
    Form_Property Me, G_Form_EXAM01q, G_Font_Name
    
    '�]Form���Ҧ�Panel, Label�����D, �r�ΤΦ�t
    Label_Property Lbl_A1601, G_Pnl_A1601$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1602, G_Pnl_A1602$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1628, G_Pnl_A1628$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A1609, G_Pnl_A1609$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Sign(0), G_Pnl_Dash$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Sign(1), G_Pnl_Dash$, G_Label_Color, G_Font_Size, G_Font_Name
    
    '�]Form��Help Frame�����D, �r�ΤΦ�t
'    Label_Property Fra_Help, "", COLOR_SKY, G_Font_Size, G_Font_Name
'    Fra_Help.Visible = False
'
    '�]Form���Ҧ�Text Box ���r�ΤΥi��H����
    Text_Property Txt_A1601s, 10, G_Font_Name
    Text_Property Txt_A1601e, 10, G_Font_Name
    Text_Property Txt_A1628s, 8, G_Font_Name
    Text_Property Txt_A1628e, 8, G_Font_Name
    Text_Property Txt_A1609, 15, G_Font_Name
    Text_Property Txt_A1602, 12, G_Font_Name
    
    '�]Form���Ҧ�Combo Box ���r��
    '    ComboBox_Property Cbo_A1501, G_Font_Size, G_Font_Name
        
    '�]Form���Ҧ�Command�����D�Φr��
    Command_Property cmd_Help, G_CmdHelp, G_Font_Name
    Command_Property cmd_ok, G_CmdSearch, G_Font_Name
    Command_Property cmd_add, G_CmdAdd, G_Font_Name
    Command_Property cmd_Exit, G_CmdExit, G_Font_Name
    
    '�H�U���зǫ��O, ���o�ק�
    VSElastic_Property vse_background
    StatusBar_ProPerty sts_msgline
End Sub

'====================================
'   Command Buttom Events
'====================================

'========================================================================
' Module    : frm_EXAM01q
' Procedure : cmd_add_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Doing Add Operation, Goto D-form
' Details   :
'========================================================================
Private Sub cmd_add_Click()
'�N�@�~���A�]�w���s�W���A
    G_AP_STATE = G_AP_STATE_ADD
    
'����Q�e��, Show�XDetail�e��
    DoEvents
    Me.Hide
    frm_EXAM01.Show
End Sub

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Cmd_Exit_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Exit Program
' Details   :
'========================================================================
Private Sub Cmd_Exit_Click()
'�����ثe����,���X��L�B�z�{��
    m_ExitTrigger% = True
    CloseFileDB
    End
End Sub

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Cmd_Ok_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Doing Update & Delete, Goto V-form with DY_A16 opened
' Details   : Calling "OpenMainFile" to open the dynaset by the clauses
'             that input in the textboxes A1602, A1609, A1601, A1628.
'========================================================================
Private Sub Cmd_Ok_Click()
    Me.MousePointer = HOURGLASS
    
    sts_msgline.Panels(1) = G_Process
    sts_msgline.Refresh
    
'�w�惡�e���������ˮ���찵PageCheck
    If Not IsAllFieldsCheck() Then
       Me.MousePointer = Default
       Exit Sub
    End If

'�}�Ҭd�߸��
    OpenMainFile
    
'�N�����ܨ�V�e��
    If Not (DY_A16.BOF And DY_A16.EOF) Then
       DoEvents
       Me.Hide
       frm_EXAM01v.Show
    Else
       sts_msgline.Panels(1) = G_NoQueryData
    End If
    
    Me.MousePointer = Default
End Sub

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Cmd_Help_Click
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Open HLP file
' Details   :
'========================================================================
Private Sub Cmd_Help_Click()
Dim a$

    a$ = "notepad " + G_Help_Path + "EXAM01q.HLP"
    retcode = Shell(a$, 4)
End Sub

'====================================
'   Form Events
'====================================

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Form_Activate
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Initial & Prepare Data
' Details   :
'========================================================================
Private Sub Form_Activate()
Dim A_A1601$
    sts_msgline.Panels(2) = GetCurrentDay(1)
    Me.Refresh
    m_FieldError% = -1
    m_ExitTrigger% = False
    
'�P�_�O�_�Ѩ�L���U�e���^��, �ӫD��������
    If Trim(G_FormFrom$) <> "" Then
        G_FormFrom$ = ""
        
        'Take out return value and push to correct Textbox
        StrCut frm_GD.Tag, Chr$(KEY_TAB), A_A1601$, ""
        Select Case G_Hlp_Return
            Case Txt_A1601s.TabIndex
                Txt_A1601s.text = A_A1601$
            Case Txt_A1601e.TabIndex
                Txt_A1601e.text = A_A1601$
        End Select

        Exit Sub
    Else
        '.....                '�Ĥ@������ɤ��ǳưʧ@
        'Do Something Here��
        
    End If
    G_AP_STATE = G_AP_STATE_QUERY  '�]�w�@�~���A
    sts_msgline.Panels(1) = SetMessage(G_AP_STATE)
    
    '�NForm��m��ù������h
    frm_EXAM01q.ZOrder 0
    If frm_EXAM01q.Visible Then Txt_A1601s.SetFocus
End Sub

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Form_KeyDown
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : Handle Key Events
' Details   : Handling: F1���U, F2�d��, F4�s�W, ESC���}
'========================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case KEY_F1
                If ActiveControl.TabIndex = Txt_A1601s.TabIndex Then Exit Sub
                If ActiveControl.TabIndex = Txt_A1601e.TabIndex Then Exit Sub
                KeyCode = 0
                If cmd_Help.Visible = True And cmd_Help.Enabled = True Then
                   cmd_Help.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F2
                KeyCode = 0
                If cmd_ok.Visible = True And cmd_ok.Enabled = True Then
                   cmd_ok.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F4
                KeyCode = 0
                If cmd_add.Visible = True And cmd_add.Enabled = True Then
                   cmd_add.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_ESCAPE
                KeyCode = 0
                If cmd_Exit.Visible = True And cmd_Exit.Enabled = True Then
                   cmd_Exit.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    End Select
End Sub

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Form_KeyPress
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   :
' Details   :
'========================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
    sts_msgline.Panels(1) = SetMessage(G_AP_STATE)
    
'�D�ʱN��ƿ�J�Ѥp�g�ର�j�g
'  �Y���Y����줣�ݭn�ഫ��, �����H���L
   'If ActiveControl.TabIndex = txt_xxx.TabIndex Then GoTo Form_KeyPress_A
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Form_KeyPress_A:
'��J���N�r��(ENTER���~), �N��Ʋ����ܼƳ]��TRUE
    'If ActiveControl.TabIndex <> Spd_EXAM01.TabIndex Then
       KeyPress KeyAscii           'Enter�ɦ۰ʸ���U�@���, spread���~
    'End If
End Sub

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Form_Load
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   : First Entering this Form, Preparing
' Details   :
'========================================================================
Private Sub Form_Load()
    FormCenter Me
    Set_Property
End Sub

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Form_QueryUnload
' @ Author  : Mike_chang
' @ Date    : 2015/8/28
' Purpose   :
' Details   :
'========================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim MSG ' Declare variable.

    If UnloadMode > 0 Then
       ' If exiting the application.
       MSG = GetSIniStr("PgmMsg", "g_gui_run")
    Else
       ' If just closing the form.
       Cmd_Exit_Click
    End If
    ' If user clicks the 'No' button, stop QueryUnload.
    If MsgBox(MSG, 36, Me.Caption) = 7 Then
       Cancel = True
    Else
       Cmd_Exit_Click
    End If
    
End Sub

'====================================
'   Textbox Events
'====================================


'========================================================================
' Procedure : Txt_A1601e_DblClick (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'       @modified: Call frm_GD to help user input data
'========================================================================
Private Sub Txt_A1601e_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    G_FormFrom$ = frm_GD.Name
    frm_GD.Tag = "1"
    frm_GD.Show vbModal
    G_Hlp_Return = Txt_A1601e.TabIndex
End Sub


'========================================================================
' Procedure : Txt_A1601e_KeyDown (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A1601e_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If KeyCode = KEY_F1 Then Txt_A1601e_DblClick
End Sub

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Txt_A1601e_GotFocus
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A1601e_GotFocus()
    TextHelpGotFocus
End Sub

'========================================================================
' Module    : frm_EXAM01q
' Procedure : Txt_A1601e_LostFocus
' @ Author  : Mike_chang
' @ Date    : 2015/8/27
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A1601e_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A1601e.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A1601()
End Sub

'========================================================================
' Procedure : Txt_A1601s_DblClick (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'       @modified: Call frm_GD to help user input data
'========================================================================
Private Sub Txt_A1601s_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
'    Txt_A1601s_KeyDown KEY_F1, 0
    G_FormFrom$ = frm_GD.Name
    frm_GD.Tag = "1"
    frm_GD.Show vbModal
    G_Hlp_Return = Txt_A1601s.TabIndex
    
End Sub


'========================================================================
' Procedure : Txt_A1601s_KeyDown (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A1601s_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If KeyCode = KEY_F1 Then Txt_A1601s_DblClick
End Sub

Private Sub Txt_A1601s_GotFocus()
    TextHelpGotFocus
End Sub

'========================================================================
' Procedure : Txt_A1601s_LostFocus (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Do checking
' Details   :
'========================================================================
Private Sub Txt_A1601s_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A1601s.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A1601()
End Sub

Private Sub Txt_A1602_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A1602_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A1609_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A1609_LostFocus()
    TextLostFocus
End Sub


'========================================================================
' Procedure : Txt_A1628e_GotFocus (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A1628e_GotFocus()
    TextGotFocus
End Sub

'========================================================================
' Procedure : Txt_A1628e_LostFocus (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A1628e_LostFocus()
     TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A1628e.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A1628e()
End Sub

'========================================================================
' Procedure : Txt_A1628s_GotFocus (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   :
' Details   :
'========================================================================
Private Sub Txt_A1628s_GotFocus()
    TextGotFocus
End Sub

'========================================================================
' Procedure : Txt_A1628s_LostFocus (frm_EXAM01q)
' @ Author  : Mike_chang
' @ Date    : 2015/9/4
' Purpose   : Do checking
' Details   :
'========================================================================
Private Sub Txt_A1628s_LostFocus()
     TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A1628s.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A1628s()
End Sub

