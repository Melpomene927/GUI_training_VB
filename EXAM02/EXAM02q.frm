VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frm_EXAM02q 
   Caption         =   "���u�򥻸�Ƭd��"
   ClientHeight    =   2445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2445
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VsOcxLib.VideoSoftElastic Vse_Background 
      Height          =   2070
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6270
      _Version        =   327680
      _ExtentX        =   11060
      _ExtentY        =   3651
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
      Picture         =   "EXAM02q.frx":0000
      BevelOuterDir   =   1
      MouseIcon       =   "EXAM02q.frx":001C
      Begin VB.Frame Fra_Help 
         BackColor       =   &H00FFFF80&
         Height          =   825
         Left            =   3330
         TabIndex        =   12
         Top             =   855
         Visible         =   0   'False
         Width           =   855
         Begin FPSpread.vaSpread Spd_Help 
            Height          =   495
            Left            =   90
            OleObjectBlob   =   "EXAM02q.frx":0038
            TabIndex        =   21
            Top             =   210
            Width           =   615
         End
      End
      Begin VB.TextBox Txt_A0804s 
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
         TabIndex        =   3
         Top             =   450
         Width           =   1395
      End
      Begin VB.TextBox Txt_A0804e 
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
         TabIndex        =   4
         Top             =   450
         Width           =   1395
      End
      Begin VB.TextBox Txt_A0809 
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
         TabIndex        =   6
         Top             =   1230
         Width           =   3240
      End
      Begin VB.TextBox Txt_A0802 
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
         TabIndex        =   7
         Top             =   1605
         Width           =   3240
      End
      Begin VB.TextBox Txt_A0801s 
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
         TabIndex        =   1
         Top             =   90
         Width           =   1395
      End
      Begin VB.TextBox Txt_A0801e 
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
         TabIndex        =   2
         Top             =   90
         Width           =   1395
      End
      Begin VB.ComboBox Cbo_A0824 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "EXAM02q.frx":0268
         Left            =   1395
         List            =   "EXAM02q.frx":026A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   810
         Width           =   3270
      End
      Begin Threed.SSCommand Cmd_Help 
         Height          =   405
         Left            =   4740
         TabIndex        =   8
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
         TabIndex        =   10
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
         TabIndex        =   11
         Top             =   1560
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
         TabIndex        =   9
         Top             =   540
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�d��F2"
         ForeColor       =   0
      End
      Begin VB.Label Lbl_A0804 
         Caption         =   "�����s��"
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
         Left            =   90
         TabIndex        =   15
         Top             =   525
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
         TabIndex        =   14
         Top             =   150
         Width           =   300
      End
      Begin VB.Label Lbl_A0801 
         Caption         =   "���u�s��"
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
         Left            =   90
         TabIndex        =   13
         Top             =   165
         Width           =   1380
      End
      Begin VB.Label Lbl_Sign 
         Alignment       =   2  'Center
         Caption         =   "��"
         ForeColor       =   &H00404040&
         Height          =   300
         Index           =   1
         Left            =   2850
         TabIndex        =   16
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Lbl_A0809 
         Caption         =   "�����Ҧr��"
         DataField       =   "z"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   90
         TabIndex        =   18
         Top             =   1290
         Width           =   1380
      End
      Begin VB.Label Lbl_A0802 
         Caption         =   "���u�m�W"
         DataField       =   "z"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   90
         TabIndex        =   19
         Top             =   1665
         Width           =   1380
      End
      Begin VB.Label Lbl_A0824 
         Caption         =   "���q�O"
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
         Left            =   90
         TabIndex        =   17
         Top             =   885
         Width           =   1380
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   2070
      Width           =   6270
      _ExtentX        =   11060
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
Attribute VB_Name = "Frm_EXAM02q"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'�b���B�w�q���Ҧ��ܼ�, �@�ߥHM�}�Y, �pM_AAA$, M_BBB#, M_CCC&
'�B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
' $: ��r
' #: �Ҧ��Ʀr�B��(���B�μƶq)
' &: �{���j���ܼ�
' %: ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
' �ť�: �N��VARIENT, �ʺA�ܼ�

'�۩w�ܼ�
'Dim m_A4101Flag%
'Dim m_aa$
'Dim m_bb#
'Dim m_cc&

'���n�ܼ�
Dim m_FieldError%    '���ܼƦb�P�_���O�_���~, �����^�����줧�ʧ@
Dim m_ExitTrigger%   '���ܼƦb�P�_������O�_�QĲ�o, �N����ثe���b�B�z���@�~

Private Sub CBO_A0824_Prepare()
On Local Error GoTo MyError
Dim A_Sql$
Dim DY_Tmp As Recordset

    '���M��Combo Box���e
    Cbo_A0824.Clear
    
    '�[�J�ťտﶵ
    Cbo_A0824.AddItem ""
    
    '�}�_�ɮ�
    A_Sql$ = "Select A0101,A0102 From A01 ORDER BY A0101"
    CreateDynasetODBC DB_ARTHGUI, DY_Tmp, A_Sql$, "DY_TMP", True

    '�N����\�JCombo Box��
    Do While Not DY_Tmp.EOF
       Cbo_A0824.AddItem Format(Trim$(DY_Tmp.Fields("A0101") & ""), "!@@@") & Trim$(DY_Tmp.Fields("A0102") & "")
       DY_Tmp.MoveNext
    Loop
    DY_Tmp.Close

    '�YCombo Box�������, ���b�Ĥ@��
    If Cbo_A0824.ListCount > 0 Then Cbo_A0824.ListIndex = 0
    Exit Sub
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Function CheckRoutine_A0801() As Boolean
    CheckRoutine_A0801 = False

    '�]�w�ܼƪ�l��
    m_FieldError% = -1
    
    '�W�[�Q�n�����ˬd
    If Trim$(Txt_A0801e) = "" Then Txt_A0801e = Txt_A0801s
    
    If Not CheckDataRange(Sts_MsgLine, Trim$(Txt_A0801s), Trim$(Txt_A0801e)) Then
       If ActiveControl.TabIndex = Txt_A0801e.TabIndex Then
          '�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A0801e.TabIndex
       Else
          m_FieldError% = Txt_A0801s.TabIndex
          Txt_A0801s.SetFocus
       End If
       Exit Function
    End If
       
    CheckRoutine_A0801 = True
End Function

Private Function CheckRoutine_A0804() As Boolean
    CheckRoutine_A0804 = False

'�]�w�ܼƪ�l��
    m_FieldError% = -1
    
'�W�[�Q�n�����ˬd
    If Trim$(Txt_A0804e) = "" Then Txt_A0804e = Txt_A0804s
    
    If Not CheckDataRange(Sts_MsgLine, Trim$(Txt_A0804s), Trim$(Txt_A0804e)) Then
       '==================
       'if from s to e
       'do not focus back (since it's correct to entering from s to e)
       '==================
       If ActiveControl.TabIndex = Txt_A0804e.TabIndex Then
'�Y�����~, �N�ܼƭȳ]�w����Control��TabIndex
          m_FieldError% = Txt_A0804e.TabIndex
       Else
          m_FieldError% = Txt_A0804s.TabIndex
          Txt_A0804s.SetFocus
       End If
       Exit Function
    End If
       
    CheckRoutine_A0804 = True
End Function

Private Sub DataPrepare_A02(Txt As TextBox)
'PrepareData for Txt_A0804s, Txt_A0804e
Dim A_Sql$                  'SQL Message
Dim DY_Tmp As Recordset     'Temporary Dynaset
    Me.MousePointer = HOURGLASS
    
    '�}�_�ɮ�
    'concate SQL Message
    A_Sql$ = "Select A0201 ,A0202 From A02"
    'generate wildcard compare SQL Statement
    If Txt.text <> "" Then
        A_Sql$ = A_Sql$ & " Where A0201 Like '" & Txt.text & GetLikeStr(DB_ARTHGUI, True) & "'"
    End If
    A_Sql$ = A_Sql$ & " Order by A0201"
    
    'open dynaset of A02
    CreateDynasetODBC DB_ARTHGUI, DY_Tmp, A_Sql$, "DY_TMP", True
    If DY_Tmp.BOF And DY_Tmp.EOF Then
       Me.MousePointer = Default
       Sts_MsgLine.Panels(1) = G_NoReference
       Exit Sub
    End If
    
    With Spd_Help
         '�]�w���U����(Spd_Help)������ݩ�
         .UnitType = 2
         Spread_Property Spd_Help, 0, 2, WHITE, G_Font_Size, G_Font_Name
         Spread_Col_Property Spd_Help, 1, TextWidth("X") * 6, G_Pnl_A0201$
         Spread_Col_Property Spd_Help, 2, TextWidth("X") * 12, G_Pnl_A0201$
         Spread_DataType_Property Spd_Help, 1, SS_CELL_TYPE_EDIT, "", "", 6
         Spread_DataType_Property Spd_Help, 2, SS_CELL_TYPE_EDIT, "", "", 12
         
         .Row = -1
         .Col = -1: .Lock = True
         .Col = 1: .TypeHAlign = 2
    
         '�N����\�JSpread��
         Do Until DY_Tmp.EOF
            .MaxRows = .MaxRows + 1
            .Row = Spd_Help.MaxRows
            .Col = 1
            .text = Trim(DY_Tmp.Fields("A0201") & "")
            .Col = 2
            .text = Trim(DY_Tmp.Fields("A0202") & "")
            DY_Tmp.MoveNext
         Loop
         DY_Tmp.Close
         
         '�]�w���U��������ܦ�m
         SetHelpWindowPos Fra_Help, Spd_Help, 330, 90, 4305, 2025
         .Tag = Txt.TabIndex    'set return control tab index
         .SetFocus
    End With
    Me.MousePointer = Default
End Sub

Private Sub DataPrepare_A08(Txt As TextBox)
'PrepareData for Txt_A0801s, Txt_A0801e
Dim A_Sql$                  'SQL Message
Dim DY_Tmp As Recordset     'Temporary Dynaset
    Me.MousePointer = HOURGLASS
    
    '�}�_�ɮ�
    'concate SQL Message
    A_Sql$ = "Select A0801 ,A0802 From A08"
    
    'generate wildcard compare SQL Statement
    If Txt.text <> "" Then
        A_Sql$ = A_Sql$ & " Where A0801 Like '" & Txt.text & GetLikeStr(DB_ARTHGUI, True) & "'"
    End If
    A_Sql$ = A_Sql$ & " Order by A0801"
    
    'open dynaset of A08
    CreateDynasetODBC DB_ARTHGUI, DY_Tmp, A_Sql$, "DY_TMP", True
    If DY_Tmp.BOF And DY_Tmp.EOF Then
       Me.MousePointer = Default
       Sts_MsgLine.Panels(1) = G_NoReference
       Exit Sub
    End If
    
    With Spd_Help
         '�]�w���U����(Spd_Help)������ݩ�
         .UnitType = 2
         Spread_Property Spd_Help, 0, 2, WHITE, G_Font_Size, G_Font_Name
         Spread_Col_Property Spd_Help, 1, TextWidth("X") * 10, G_Pnl_A0801$
         Spread_Col_Property Spd_Help, 2, TextWidth("X") * 12, G_Pnl_A0802$
         Spread_DataType_Property Spd_Help, 1, SS_CELL_TYPE_EDIT, "", "", 10
         Spread_DataType_Property Spd_Help, 2, SS_CELL_TYPE_EDIT, "", "", 12
         
         .Row = -1
         .Col = -1: .Lock = True
         .Col = 1: .TypeHAlign = 2
    
         '�N����\�JSpread��
         Do Until DY_Tmp.EOF
            .MaxRows = .MaxRows + 1
            .Row = Spd_Help.MaxRows
            .Col = 1
            .text = Trim(DY_Tmp.Fields("A0801") & "")
            .Col = 2
            .text = Trim(DY_Tmp.Fields("A0802") & "")
            DY_Tmp.MoveNext
         Loop
         DY_Tmp.Close
    
         '�]�w���U��������ܦ�m
         SetHelpWindowPos Fra_Help, Spd_Help, 330, 90, 4305, 2025
         .Tag = Txt.TabIndex    'set return control tab index
         .SetFocus
    End With
    
    Me.MousePointer = Default
End Sub

Private Function IsAllFieldsCheck() As Boolean
    IsAllFieldsCheck = False
    
'����d�ߩΦs�ɫe���N�Ҧ��ˮ����A���@��

    If Not CheckRoutine_A0801 Then Exit Function
    If Not CheckRoutine_A0804 Then Exit Function
    
    DoEvents
    
    IsAllFieldsCheck = True
End Function

Private Sub OpenMainFile()
On Local Error GoTo MyError
Dim A_Sql$
Dim A_A0801s$, A_A0801e$    'txt
Dim A_A0804s$, A_A0804e$    'txt
Dim A_A0824$                'cbo
Dim A_A0809$                'txt
Dim A_A0802$                'txt
    
    'Keep TextBox ��Ʀ��ܼ�
    A_A0801s$ = Trim(Txt_A0801s)
    A_A0801e$ = Trim(Txt_A0801e)
    A_A0804s$ = Trim(Txt_A0804s)
    A_A0804e$ = Trim(Txt_A0804e)
    StrCut Cbo_A0824.text, Space(1), A_A0824$, ""
    A_A0809$ = Trim(Txt_A0809)
    A_A0802$ = Trim(Txt_A0802)
    
    '�}�Ҹ��
    'get the required Columns as SPEC
    'Associated column A0102 relating A0824 display'�q��' if null
    A_Sql$ = "Select A08.*, ISNULL(A01.A0102,'�q��') As A0102 From A08"
    A_Sql$ = A_Sql$ & " LEFT JOIN A01 On A01.A0101 = A08.A0824"
    A_Sql$ = A_Sql$ & " Where 1=1"
    
    'where clause: A0824 (allow empty)
    If A_A0801s$ <> "" Then
        A_Sql$ = A_Sql$ & " And A0824 = '" & A_A0824$ & "'"
    End If
    
    'where clause: A0801
    If A_A0801s$ <> "" Then
       A_Sql$ = A_Sql$ & " And A0801>='" & A_A0801s$ & "'"
    End If
    If A_A0801e$ <> "" Then
       A_Sql$ = A_Sql$ & " And A0801<='" & A_A0801e$ & "'"
    End If
    
    'where clause A0804
    If A_A0804s$ <> "" Then
       A_Sql$ = A_Sql$ & " And A0804>='" & A_A0804s$ & "'"
    End If
    If A_A0804e$ <> "" Then
       A_Sql$ = A_Sql$ & " And A0804<='" & A_A0804e$ & "'"
    End If
    
    'where clause A0809
    If A_A0809$ <> "" Then
       A_Sql$ = A_Sql$ & " And A0809 Like'" & A_A0809$ _
                & GetLikeStr(DB_ARTHGUI, True) & "'"
    End If
    
    'where clause A0802
    If A_A0802$ <> "" Then
       A_Sql$ = A_Sql$ & " And A0802 Like'" & GetLikeStr(DB_ARTHGUI, True) _
                & A_A0802$ & GetLikeStr(DB_ARTHGUI, True) & "'"
    End If
    
    A_Sql$ = A_Sql$ & "Order by A0801"
    
    CreateDynasetODBC DB_ARTHGUI, DY_A08, A_Sql$, "DY_A08", True
    Exit Sub
    
MyError:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

Private Sub Set_Property()
'�]�w��Form�����D,�r�ΤΦ�t
    Form_Property Me, G_Form_EXAM02q, G_Font_Name
    
'�]Form���Ҧ�Panel, Label�����D, �r�ΤΦ�t
    Label_Property Lbl_A0801, G_Pnl_A0801$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0804, G_Pnl_A0804$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0809, G_Pnl_A0809$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0802, G_Pnl_A0802$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_A0824, G_Pnl_A0824$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Sign(0), G_Pnl_Dash$, G_Label_Color, G_Font_Size, G_Font_Name
    Label_Property Lbl_Sign(1), G_Pnl_Dash$, G_Label_Color, G_Font_Size, G_Font_Name
    
'�]Form��Help Frame�����D, �r�ΤΦ�t
    Label_Property Fra_Help, "", COLOR_SKY, G_Font_Size, G_Font_Name
    Fra_Help.Visible = False
   
'�]Form���Ҧ�Text Box ���r�ΤΥi��H����
    Text_Property Txt_A0801s, 6, G_Font_Name
    Text_Property Txt_A0801e, 6, G_Font_Name
    Text_Property Txt_A0804s, 8, G_Font_Name
    Text_Property Txt_A0804e, 8, G_Font_Name
    Text_Property Txt_A0809, 8, G_Font_Name
    Text_Property Txt_A0802, 8, G_Font_Name
    
'�]Form���Ҧ�Combo Box ���r��
    ComboBox_Property Cbo_A0824, G_Font_Size, G_Font_Name
    
'�]Form���Ҧ�Command�����D�Φr��
    Command_Property cmd_help, G_CmdHelp, G_Font_Name
    Command_Property cmd_ok, G_CmdSearch, G_Font_Name
    Command_Property cmd_add, G_CmdAdd, G_Font_Name
    Command_Property cmd_exit, G_CmdExit, G_Font_Name
    
'�H�U���зǫ��O, ���o�ק�
    VSElastic_Property Vse_background
    StatusBar_ProPerty Sts_MsgLine
End Sub

Private Sub Cbo_A0824_DropDown()
Dim A_A0824$
    DoEvents
    
    '�N�ثeCombo Box�W���N�XKeep�U��
    StrCut Cbo_A0824.text, Space(1), A_A0824$, ""
    
    '���s�ǳƦ�Combo Box�����e
    CBO_A0824_Prepare
    
    '�NCombo Box�W��ListIndex���VKeep�U�Ӫ����
    CboStrCut Cbo_A0824, A_A0824$, Space(1)
End Sub

Private Sub Cbo_A0824_GotFocus()
    TextGotFocus
End Sub

Private Sub Cbo_A0824_LostFocus()
    TextLostFocus
End Sub

Private Sub cmd_add_Click()
'�N�@�~���A�]�w���s�W���A
    G_AP_STATE = G_AP_STATE_ADD
    
'����Q�e��, Show�XDetail�e��
    DoEvents
    Me.Hide
    frm_EXAM02.Show
End Sub

Private Sub Cmd_Ok_Click()
    Me.MousePointer = HOURGLASS
    
    Sts_MsgLine.Panels(1) = G_Process
    Sts_MsgLine.Refresh
    
'�w�惡�e���������ˮ���찵PageCheck
    If Not IsAllFieldsCheck() Then
       Me.MousePointer = Default
       Exit Sub
    End If

'�}�Ҭd�߸��
    OpenMainFile
    
'�N�����ܨ�V�e��
    If Not (DY_A08.BOF And DY_A08.EOF) Then
       DoEvents
       Me.Hide
       Frm_EXAM02v.Show
    Else
       Sts_MsgLine.Panels(1) = G_NoQueryData
    End If
    
    Me.MousePointer = Default
End Sub

Private Sub Cmd_Exit_Click()
'�����ثe����,���X��L�B�z�{��
    m_ExitTrigger% = True
    CloseFileDB
    End
End Sub

Private Sub Cmd_Help_Click()
Dim a$

    a$ = "notepad " + G_Help_Path + "EXAM02q.HLP"
    retcode = Shell(a$, 4)
End Sub

Private Sub Form_Activate()
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    Me.Refresh
    m_FieldError% = -1
    m_ExitTrigger% = False
    
'�P�_�O�_�Ѩ�L���U�e���^��, �ӫD��������
    If Trim(G_FormFrom$) <> "" Then
       G_FormFrom$ = ""
       '.....                '�[�J�ҭn�]�w���ʧ@
       '.....
       Exit Sub
    Else
       '.....                '�Ĥ@������ɤ��ǳưʧ@
       'Do Something Here��
       
    End If
    G_AP_STATE = G_AP_STATE_QUERY  '�]�w�@�~���A
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    
    '�NForm��m��ù������h
    Frm_EXAM02q.ZOrder 0
    If Frm_EXAM02q.Visible Then Txt_A0801s.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case KEY_F1
                If ActiveControl.TabIndex = Txt_A0801s.TabIndex Then Exit Sub
                If ActiveControl.TabIndex = Txt_A0801e.TabIndex Then Exit Sub
                If ActiveControl.TabIndex = Txt_A0804s.TabIndex Then Exit Sub
                If ActiveControl.TabIndex = Txt_A0804e.TabIndex Then Exit Sub
                KeyCode = 0
                If cmd_help.Visible = True And cmd_help.Enabled = True Then
                   cmd_help.SetFocus
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
                If cmd_exit.Visible = True And cmd_exit.Enabled = True Then
                   cmd_exit.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    
'�D�ʱN��ƿ�J�Ѥp�g�ର�j�g
'  �Y���Y����줣�ݭn�ഫ��, �����H���L
   'If ActiveControl.TabIndex = txt_xxx.TabIndex Then GoTo Form_KeyPress_A
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Form_KeyPress_A:
'��J���N�r��(ENTER���~), �N��Ʋ����ܼƳ]��TRUE
    'If ActiveControl.TabIndex <> Spd_PATTERNM.TabIndex Then
       KeyPress KeyAscii           'Enter�ɦ۰ʸ���U�@���, spread���~
    'End If
End Sub

Private Sub Form_Load()
    FormCenter Me
    Set_Property
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim MSG ' Declare variable.

    If UnloadMode > 0 Then
       ' If exiting the application.
       MSG = GetCaption("PgmMsg", "g_gui_run", "���t�Υثe���b����,�n������?")
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

Private Sub Spd_Help_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim A_Code$, A_Value$

    Me.MousePointer = HOURGLASS
    
    'KEEP�ۻ��U�����I�諸���
    With Spd_Help
         'redirect to Pkey
         .Row = .ActiveRow
         .Col = 1
         A_Code$ = Trim(.text)
         .Col = 2
         A_Value$ = Trim(.text)
    
         '�NKEEP����Ʊa�J�e��
         Select Case Val(.Tag)
           Case Txt_A0801s.TabIndex
                Txt_A0801s = A_Code$
           Case Txt_A0801e.TabIndex
                Txt_A0801e = A_Code$
           Case Txt_A0804s.TabIndex
                Txt_A0804s = A_Code$
           Case Txt_A0804e.TabIndex
                Txt_A0804e = A_Code$
         End Select
    End With
    
    '���û��U����
    Fra_Help.Visible = False
    
    Me.MousePointer = Default
End Sub

Private Sub Spd_Help_GotFocus()
    SpreadGotFocus Spd_Help.ActiveCol, Spd_Help.ActiveRow
End Sub

Private Sub Spd_Help_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_RETURN Then
       Spd_Help_DblClick Spd_Help.ActiveCol, Spd_Help.ActiveRow
    End If
End Sub

Private Sub Spd_Help_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'�зǫ��O,���o�ק�
    SpreadLostFocus Col, Row
    If NewCol > 0 Then SpreadGotFocus NewCol, NewRow
End Sub

Private Sub Spd_Help_LostFocus()
    Fra_Help.Visible = False
    Select Case Val(Spd_Help.Tag)
      Case Txt_A0801s.TabIndex
           Txt_A0801s.SetFocus
      Case Txt_A0801e.TabIndex
           Txt_A0801e.SetFocus
      Case Txt_A0804s.TabIndex
           Txt_A0804s.SetFocus
      Case Txt_A0804e.TabIndex
           Txt_A0804e.SetFocus
    End Select
End Sub

Private Sub Txt_A0801e_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A0801e_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A0801e_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If m_FieldError% <> -1 Then Exit Sub
    If KeyCode = KEY_F1 Then DataPrepare_A08 Txt_A0801e
End Sub

Private Sub Txt_A0801e_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A0801e_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If Fra_Help.Visible = True Then Exit Sub
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0801e.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0801()
End Sub

Private Sub Txt_A0801s_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A0801s_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A0801s_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If m_FieldError% <> -1 Then Exit Sub
    If KeyCode = KEY_F1 Then DataPrepare_A08 Txt_A0801s
End Sub

Private Sub Txt_A0801s_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A0801s_LostFocus()
    TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If Fra_Help.Visible = True Then Exit Sub
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0801s.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0801()
End Sub

Private Sub Txt_A0802_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0802_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0809_GotFocus()
    TextGotFocus
End Sub

Private Sub Txt_A0809_LostFocus()
    TextLostFocus
End Sub

Private Sub Txt_A0804e_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A0804e_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A0804e_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If m_FieldError% <> -1 Then Exit Sub
    If KeyCode = KEY_F1 Then DataPrepare_A02 Txt_A0804e
End Sub

Private Sub Txt_A0804e_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A0804e_LostFocus()
     TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0804e.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0804()
End Sub

Private Sub Txt_A0804s_DblClick()
'�Y��즳���ѻ��U���,���U�ƹ�, �Ҷ��B�z���ƶ�
    Txt_A0804s_KeyDown KEY_F1, 0
End Sub

Private Sub Txt_A0804s_KeyDown(KeyCode As Integer, Shift As Integer)
'�Y��즳���ѻ��U���,���UF1, �Ҷ��B�z���ƶ�
    If m_FieldError% <> -1 Then Exit Sub
    If KeyCode = KEY_F1 Then DataPrepare_A02 Txt_A0804s
End Sub

Private Sub Txt_A0804s_GotFocus()
    TextHelpGotFocus
End Sub

Private Sub Txt_A0804s_LostFocus()
     TextLostFocus
    
'�P�_�H�U���p�o�ͮ�, ����������B�z
    If (TypeOf ActiveControl Is SSCommand) Then Exit Sub
    If m_FieldError% <> -1 And m_FieldError% <> Txt_A0804s.TabIndex Then Exit Sub
    ' ....

'�ۧ��ˬd
    retcode = CheckRoutine_A0804()
End Sub



