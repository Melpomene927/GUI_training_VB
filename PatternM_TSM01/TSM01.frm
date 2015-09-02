VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{2037E3AD-18D6-101C-8158-221E4B551F8E}#5.0#0"; "Vsocx32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Frm_TSM01 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "���O �򥻸�ƺ޲z(M Pattern)"
   ClientHeight    =   4650
   ClientLeft      =   795
   ClientTop       =   705
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "TSM01.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   8250
   Begin VsOcxLib.VideoSoftElastic Vse_background 
      Height          =   4275
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   8250
      _Version        =   327680
      _ExtentX        =   14552
      _ExtentY        =   7541
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
      Picture         =   "TSM01.frx":030A
      BevelOuterDir   =   1
      MouseIcon       =   "TSM01.frx":0326
      Begin FPSpread.vaSpread Spd_TSM01 
         Height          =   3645
         Left            =   60
         OleObjectBlob   =   "TSM01.frx":0342
         TabIndex        =   0
         Top             =   540
         Width           =   6765
      End
      Begin Threed.SSCommand cmd_delete 
         Height          =   405
         Left            =   6900
         TabIndex        =   2
         Top             =   540
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�R��F3"
      End
      Begin Threed.SSCommand cmd_help 
         Height          =   405
         Left            =   6900
         TabIndex        =   1
         Top             =   90
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "���UF1"
      End
      Begin Threed.SSCommand cmd_exit 
         Height          =   405
         Left            =   6900
         TabIndex        =   3
         Top             =   3780
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2293
         _ExtentY        =   714
         _StockProps     =   78
         Caption         =   "�� �� ESC"
      End
   End
   Begin ComctlLib.StatusBar Sts_MsgLine 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4275
      Width           =   8250
      _ExtentX        =   14552
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
Attribute VB_Name = "Frm_TSM01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

'========================================================================
' Coding Rules
'========================================================================
'�b���B�w�q���Ҧ��ܼ�, �@�ߥHM�}�Y,
'       Ex: M_AAA$, M_BBB#, M_CCC&
'
'�B�ܼƤ��κA, �@�ߦb�̫�@�X�ϧO, �d�Ҧp�U:
' $: String ��r
' #: Double �Ҧ��Ʀr�B��(���B�μƶq)
' &: Long �{���j���ܼ�
' %: Integer ���@�ǨϥΩ�O�Χ_�γ~���ܼ� (TRUE / FALSE )
' �ť�: VARIENT, �ʺA�ܼ�
'========================================================================

'-- Fixed Variables (���n�ܼ�) :
Dim m_FieldError%    '���ܼƦb�P�_���O�_���~, �����^�����줧�ʧ@
Dim m_ExitTrigger%   '���ܼƦb�P�_������O�_�QĲ�o, �N����ثe���b�B�z���@�~

'-- Additional Variables (�۩w�ܼ�) :
'Dim m_A0101Flag%
Dim m_aa$
Dim m_bb#
Dim m_cc&

'================================
'    User Define Function, Sub
'================================

'========================================================================
' Procedure : Set_Property
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Called while Initializing form
'========================================================================
Private Sub Set_Property()

'�]�w��Form�����D,�r�ΤΦ�t
    Form_Property Me, G_Form_PATTERN$, G_Font_Name

'�]Form���Ҧ�Panel, Label�����D, �r�ΤΦ�t

    '===========
    ' @Modify:
    '   No Label
    '===========
'    Label_Property Lbl_A0101, G_Pnl_A0101$, G_Label_Color, G_Font_Size, G_Font_Name
    
'�]Form���Ҧ�Command�����D�Φr��
    Command_Property cmd_help, G_CmdHelp, G_Font_Name
    Command_Property cmd_delete, G_CmdDel, G_Font_Name
    Command_Property cmd_exit, G_CmdExit, G_Font_Name

'�]Form���Ҧ�Combo Box ���r��

    '===========
    ' @Modify:
    '   No CBO
    '===========
'    ComboBox_Property Cbo_A0101, G_Font_Size, G_Font_Name
    
'�]Form��Spread���ݩ�
    '===========
    ' @Modify:
    '   New property
    '===========
    Set_Spread_Property

'�H�U���зǫ��O, ���o�ק�
    VSElastic_Property Vse_background
    StatusBar_ProPerty Sts_MsgLine
End Sub

'========================================================================
' Procedure : Set_Spread_Property
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Initializing Spread
'   1. Set Cols, Rows
'   2. Set Col Headers, Col default width
'   3. Set Celltype of each one
'   4. Set Frozen Column
'   5. Set Column alignment
'   6. Set Hidden Columns
'========================================================================
Private Sub Set_Spread_Property()
    Spd_TSM01.UnitType = 2

'1. �]�w��Spread�����Ƥ�����
    Spread_Property Spd_TSM01, 0, 6, WHITE, G_Font_Size, G_Font_Name

'2. �]�w��Spread���U����D����ܼe��, 0�N�����줣���
    Spread_Col_Property Spd_TSM01, 1, TextWidth("A") * 8, G_Pnl_A0101$
    Spread_Col_Property Spd_TSM01, 2, TextWidth("A") * 10, G_Pnl_A0102$
    Spread_Col_Property Spd_TSM01, 3, TextWidth("A") * 18, G_Pnl_A0104$
    Spread_Col_Property Spd_TSM01, 4, TextWidth("A") * 12, G_Pnl_A0111$
    Spread_Col_Property Spd_TSM01, 5, TextWidth("A") * 0, "A0101o" 'p-key
    Spread_Col_Property Spd_TSM01, 6, TextWidth("A") * 0, "Change/Add/No Change" 'p-key

'3. �]�w��Spread���U���ݩʤ���ܦr��
    'SS_CELL_TYPE_EDIT        = ��r�i��J
    'SS_CELL_TYPE_FLOAT       = �Ʀr�i��J
    'SS_CELL_TYPE_STATIC_TEXT = �����
    'SS_CELL_TYPE_CHECKBOX    = �I�ﶵ��
    Spread_DataType_Property Spd_TSM01, 1, SS_CELL_TYPE_EDIT, "", "", 2
    Spread_DataType_Property Spd_TSM01, 2, SS_CELL_TYPE_EDIT, "", "", 12
    Spread_DataType_Property Spd_TSM01, 3, SS_CELL_TYPE_EDIT, "", "", 40
    Spread_DataType_Property Spd_TSM01, 4, SS_CELL_TYPE_EDIT, "", "", 15
    Spread_DataType_Property Spd_TSM01, 5, SS_CELL_TYPE_EDIT, "", "", 2
    Spread_DataType_Property Spd_TSM01, 6, SS_CELL_TYPE_EDIT, "", "", 1
    
    Spd_TSM01.EditEnterAction = SS_CELL_EDITMODE_EXIT_NEXT

'4. �T�w�V�k���ʮ�, �ҭ�����
    Spd_TSM01.ColsFrozen = 1

'5. �w�q�Y����m����m���]�w 0:���a  1:�k�a  2:�m��
    Spd_TSM01.Row = -1
    Spd_TSM01.Col = 1: Spd_TSM01.TypeHAlign = 2

'6. �w�q�Y����m�Q�O�@�L�k���
    Spd_TSM01.Col = 5:  Spd_TSM01.ColHidden = True
    Spd_TSM01.Col = 6:  Spd_TSM01.ColHidden = True
End Sub


'========================================================================
' Procedure : Cbo_A0101_Check
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : check while ComboBox change
'========================================================================
'Function Cbo_A0101_Check() As Boolean
'
'    Cbo_A0101_Check = False
'
'    ' Initializing
'    m_FieldError% = -1
'
'    'Check combobox item empty string
'    If Trim(Cbo_A0101) = "" Then
'       Sts_MsgLine.Panels(1) = Lbl_A0101 & G_MustInput  'ErrorMsg @ stsBar
'       m_FieldError% = Cbo_A0101.TabIndex               'Record Err Component
'       Cbo_A0101.SetFocus
'       Exit Function
'    End If
'
'    Cbo_A0101_Check = True
'End Function


'========================================================================
' Procedure : SpreadLineCheck
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Check the Primary Key of each record: not empty
'========================================================================
Function SpreadLineCheck(ByVal Row As Long, Col As Long) As Boolean
    With Spd_TSM01
        
        'Initialize
         .Row = Row
         SpreadLineCheck = False
        
        'Check Primary Key columns:
        '   Add more while there's more then one Pkey
         If SpreadCheck_1(Row) = False Then
            Col = 1
            Exit Function
         End If
        
        'If SpreadCheck_2(row) = False Then
        '   Col = 2
        '   Exit Function
        'End If
        

        SpreadLineCheck = True
    End With
End Function

'========================================================================
' Procedure : SpreadCheck_1
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Check individual column which is Pkey
'========================================================================

Function SpreadCheck_1(ByVal Row As Long) As Boolean
Dim A_A0101$    'Pkey of the table (Before modified)
Dim A_A0101o$   'Pre Pkey
Dim A_Action$   'Action record of modification
    
    SpreadCheck_1 = False
    With Spd_TSM01
         .Row = Row

        '���oAction Code(�s�W�έק�)
         .Col = 1
         A_A0101$ = Trim(.text)     'fetch Pkey
         .Col = 5
         A_A0101o$ = Trim(.text)    'fetch Pre Pkey
         .Col = 6
         A_Action$ = Trim(.text)    'fetch Action record


        'Check Pkey empty
         .Col = 1                   'allocate to column1:A0101 (Pkey)
         If Trim(.text) = "" Then
            'raise err msg @ stsBar
            Sts_MsgLine.Panels(1) = G_Pnl_A0101$ & G_MustInput
            Exit Function
         End If

        'Check Pkey duplicate while modify data
        'Allocate to column1: A0101 (Pkey)
         .Col = 1
         If A_Action$ = "A" Then    'Action: Add data
         
            If IsKeyExist(A_A0101$) = True Then
               'raise err msg @ stsBar
               Sts_MsgLine.Panels(1) = G_Pnl_A0101$ & G_RecordExist
               Exit Function
            End If
            
         ElseIf A_Action$ = "U" Then 'Action: Update data
         
            'check only if Pkey is modified
            If IsKeyChanged(.text, A_A0101o$) = True Then
                If IsKeyExist(A_A0101$) = True Then
                    'raise err msg @ stsBar
                    Sts_MsgLine.Panels(1) = G_Pnl_A0101$ & G_RecordExist
                    Exit Function
                End If
            End If
            
         End If

        'Pass Pkey Check: return True
         SpreadCheck_1 = True
    End With
End Function

'========================================================================
' Procedure : OpenMainFile
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Load Records to Spread
'========================================================================

Private Sub OpenMainFile()
On Local Error GoTo My_Error
Dim A_Sql$      'SQL Message

    
    'Concate the SQL Message String
    A_Sql$ = _
        "Select A0101,A0102,A0104,A0111 From A01 " & _
        "Order by A0101;"
    
    'Open RecordSet by [GUI_common_component]
    CreateDynasetODBC DB_ARTHGUI, DY_A01, A_Sql$, "DY_A01", True
    Exit Sub
    
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Procedure : Delete_Process_A01
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Do deletion to the highlight row of Spread
'========================================================================

Private Sub Delete_Process_A01(ByVal A_A0101$)
On Local Error GoTo My_Error
Dim A_Sql$      'SQL Message

    '�U�R����ƫ��O
    A_Sql$ = "DELETE From A01 " & _
             "Where A0101='" & Trim(A_A0101$) & "'"
             
    'Execute SQL Message by [GUI_common_component]
    ExecuteProcess DB_ARTHGUI, A_Sql$
    Exit Sub
    
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Procedure : IsKeyChanged
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Boolean Function to determind whether Pkey has changed
'                   True if A_4102 <> A_4102o
'========================================================================

Private Function IsKeyChanged(ByVal A_A0101$, ByVal A_A0101o$) As Boolean

   IsKeyChanged = False
   If UCase$(A_A0101$) <> UCase$(A_A0101o$) Then
      IsKeyChanged = True
   End If
   
End Function

'========================================================================
' Procedure : IsKeyExist
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Boolean Function to determind whether
'========================================================================

Private Function IsKeyExist(ByVal A_A0101$) As Boolean
On Local Error GoTo My_Error
Dim A_Sql$      'SQL Message
    
    'Initialize
    IsKeyExist = False
    
    'Concate SQL Message
    A_Sql$ = "Select A0101 From A01 " & _
             "Where A0101 = '" & Trim(A_A0101$) & "' " & _
             "Order by A0101"
             
    'Open Recordset By [GUI_common_componet]
    CreateDynasetODBC DB_ARTHGUI, DY_A01, A_Sql$, "DY_A01", True
    
    'Check if Pkey already exists
    If Not (DY_A01.BOF And DY_A01.EOF) Then
        IsKeyExist = True
    End If
    Exit Function
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Function

'========================================================================
' Procedure : MoveDB2Field
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Fetch records form Database to Spread
'========================================================================

Private Sub MoveDB2Field()
On Local Error GoTo My_Error
    
    With Spd_TSM01
         'Initialize
         .MaxRows = 0   'Clear Spread
         
         'Fetch data from Recordset that has already opened
         Do While Not DY_A01.EOF And Not m_ExitTrigger%
            .MaxRows = .MaxRows + 1 'New Row
            
            'Allocate to last row & write data
            .Row = .MaxRows
            .Col = 1
            .text = Trim(DY_A01.Fields("A0101") & "")
            .Col = 2
            .text = Trim(DY_A01.Fields("A0102") & "")
            .Col = 3
            .text = Trim(DY_A01.Fields("A0104") & "")
            .Col = 4
            .text = Trim(DY_A01.Fields("A0111") & "")
            .Col = 5
            .text = Trim(DY_A01.Fields("A0101") & "")
            .Col = 6
            .text = ""
            DY_A01.MoveNext
         Loop
         
        .MaxRows = .MaxRows + 1     'One more row for new record
'         Cbo_A0101.Tag = Cbo_A0101.text 'Keep Pre 1st Pkey for check
    End With
    Exit Sub
    
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'========================================================================
' Procedure : MoveField2DB
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Save record to Database after modified
'========================================================================

Private Sub MoveField2DB(ByVal Row As Long)
On Local Error GoTo My_Error
Dim A_A0101$    'Col1
Dim A_A0102$    'Col2
Dim A_A0104$    'Col3
Dim A_A0111$    'Col4
Dim A_A0101o$   'Pre Pkey for Update
Dim A_Action$   'Action record: A/U
    
    'Set mouse curser to loading
    Me.MousePointer = HOURGLASS
    
    
    With Spd_TSM01
        'Fetch data from Spread
        .Row = Row
        .Col = 1: A_A0101$ = Trim(.text)
        .Col = 2: A_A0102$ = Trim(.text)
        .Col = 3: A_A0104$ = Trim(.text)
        .Col = 4: A_A0111$ = Trim(.text)
        .Col = 5: A_A0101o$ = Trim(.text)
        .Col = 6: A_Action$ = Trim(.text)
        
        'Write to Global String which pass data
        G_Str = ""
        If UCase$(A_Action$) = UCase$("U") Then
           'Updating
           UpdateString "A01005", GetCurrentDate(), G_Data_String
           UpdateString "A01006", GetCurrentTime(), G_Data_String
           UpdateString "A01007", GetWorkStation(), G_Data_String
           UpdateString "A01008", GetUserId(), G_Data_String
           
           UpdateString "A0101", A_A0101$, G_Data_String
           UpdateString "A0102", A_A0102$, G_Data_String
           UpdateString "A0104", A_A0104$, G_Data_String
           UpdateString "A0111", A_A0111$, G_Data_String
           
           G_Str = G_Str & " where A0101='" & Trim(A_A0101$) & "'"
           
           SQLUpdate DB_ARTHGUI, "A01"
        Else
           'Inserting
           InsertFields "A01001", GetCurrentDate(), G_Data_String
           InsertFields "A01002", GetCurrentTime(), G_Data_String
           InsertFields "A01003", GetWorkStation(), G_Data_String
           InsertFields "A01004", GetUserId(), G_Data_String
           InsertFields "A01005", " ", G_Data_String
           InsertFields "A01006", " ", G_Data_String
           InsertFields "A01007", " ", G_Data_String
           InsertFields "A01008", " ", G_Data_String
           
           InsertFields "A0101", A_A0101$, G_Data_String
           InsertFields "A0102", A_A0102$, G_Data_String
           InsertFields "A0104", A_A0104$, G_Data_String
           InsertFields "A0111", A_A0111$, G_Data_String
           SQLInsert DB_ARTHGUI, "A01"
        End If
        '
        .Col = 5: .text = A_A0101$  'Record the Previous 2nd Pkey
        .Col = 6: .text = ""        'Clear Action record
    End With
    
    'Resume mouse curser to default
    Me.MousePointer = Default
    Exit Sub
    
My_Error:
    retcode = AccessDBErrorMessage()
    If retcode = IDOK Then Resume
    If retcode = IDCANCEL Then CloseFileDB: End
End Sub

'================================
'    Form Events Handler
'================================

'================================
'    ComboBox Events (Eliminated)
'================================
'========================================================================
'Private Sub Cbo_A0101_Click()
'
'    If m_A0101Flag% Then Exit Sub
'
''�Y����Ƥ��e���I��B�ܰʮ�, �Ҷ��B�z���ƶ�
'    If Trim(Cbo_A0101.text) <> Trim(Cbo_A0101.Tag) Then
'       Me.MousePointer = HOURGLASS
'       OpenMainFile            '���d�Ҭ����s�}��, ��ܸ��
'       MoveDB2Field
'       '....
'       '....
'       Me.MousePointer = Default
'    End If
'End Sub
'========================================================================

'========================================================================
'Private Sub Cbo_A0101_DropDown()
'    DoEvents
'
'    m_A0101Flag% = True
'
''�N�ثeCombo Box�W���N�XKeep�U��
'    Dim A_A0101$
'    StrCut Cbo_A0101.text, Space(1), A_A0101$, ""
'
''���s�ǳƦ�Combo Box�����e
'    CBO_A0101_Prepare
'
''�NCombo Box�W��ListIndex���VKeep�U�Ӫ����
'    CboStrCut Cbo_A0101, A_A0101$, Space(1)
'
'    m_A0101Flag% = False
'End Sub
'========================================================================

'========================================================================
'Sub CBO_A0101_Prepare()
'On Local Error GoTo My_Error
'Dim A_Sql$
'
''���M��Combo Box���e
'    Cbo_A0101.Clear
'
''�}�_�ɮ�
'    A_Sql$ = "SELECT TOPICVALUE FROM Sini " & _
'             "Where section='BASIC' " & _
'             "ORDER BY section,topic"
'
'    CreateDynasetODBC DB_ARTHGUI, DY_SINI, A_Sql$, "DY_SINI", True
'
''�N����\�JCombo Box��
'    Do While Not DY_SINI.EOF
'       Cbo_A0101.AddItem DY_SINI.Fields("TOPICVALUE") & ""
'       DY_SINI.MoveNext
'    Loop
'
''�YCombo Box�������, ���b�Ĥ@��
'    If Cbo_A0101.ListCount > 0 Then Cbo_A0101.ListIndex = 0
'    Exit Sub
'
'My_Error:
'    retcode = AccessDBErrorMessage()
'    If retcode = IDOK Then Resume
'    If retcode = IDCANCEL Then CloseFileDB: End
'End Sub
'========================================================================

'========================================================================
'Private Sub Cbo_A0101_GotFocus()
'    TextGotFocus
'End Sub
'
'Private Sub Cbo_A0101_LostFocus()
'    TextLostFocus
'
''�P�_�H�U���p�o�ͮ�, ����������B�z
'    If ActiveControl.TabIndex = cmd_exit.TabIndex Then Exit Sub
'    If m_FieldError% <> -1 And m_FieldError% <> Cbo_A0101.TabIndex Then Exit Sub
'    ' ....
'
''�ۧ��ˬd
'    retcode = Cbo_A0101_Check()
'End Sub
'========================================================================


'================================
'    Command Buttom Events
'================================


'========================================================================
' Procedure : cmd_delete_Click
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Do Deletion while click "Delete F3"
'========================================================================
Private Sub cmd_delete_Click()
Dim A_Def$
Dim A_Msg$
Dim A_A0101o$

    With Spd_TSM01
        '�L��Ʈɤ��B�z
        If .DataRowCnt <= 0 Then Exit Sub

        '����ܦ�, �[�jĵ�ܧ@��
        SpreadWarnLine Spd_TSM01, .ActiveRow
    
        'Concate Msgbox messgage
        A_Msg$ = G_Delete_Check
        A_Def$ = MB_OKCANCEL + MB_ICONSTOP + MB_DEFBUTTON2
        
        retcode = MsgBox(A_Msg$, A_Def$, Me.Caption)  ' Get user retcode.
        If retcode = IDCANCEL Then                    ' Evaluate retcode
           SpreadWarnLineCancel Spd_TSM01, .ActiveRow '����C���٭�
           SpreadGotFocus .ActiveCol, .ActiveRow      '���s�]�w�C��
           .SetFocus
           Exit Sub
        End If

        '�NP-Key�ǤJ�R���
        .Row = .ActiveRow
        .Col = 5: A_A0101o$ = Trim(.text)
        If Trim(A_A0101o$) <> "" Then
           Delete_Process_A01 A_A0101o$
        End If

        '�H�U���зǫ��O, ���o�ק�
        '=================================
        .Action = SS_ACTION_DELETE_ROW
        If .ActiveRow > .DataRowCnt Then
           If .ActiveRow = 1 Then
              .Row = .ActiveRow
           Else
              .Row = .ActiveRow - 1
           End If
           .Col = 1
           .Action = SS_ACTION_ACTIVE_CELL
        End If
        .MaxRows = .DataRowCnt + 1
        Sts_MsgLine.Panels(1) = G_Delete_Ok
        .SetFocus
        '=================================
    End With
End Sub


'========================================================================
' Procedure : Cmd_Exit_Click
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Exit while click "Exit ESC"
'========================================================================
Private Sub Cmd_Exit_Click()
'�����ثe����,���X��L�B�z�{��
    m_ExitTrigger% = True
    CloseFileDB
    End
End Sub

'========================================================================
' Procedure : Cmd_Help_Click
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Call HLP document while click "HELP F1"
'========================================================================
Private Sub Cmd_Help_Click()
Dim a$

'�бNPATTERNM�אּ��Form�W�r�Y�i, ��l���зǫ��O, ���o�ק�
    a$ = "notepad " + G_Help_Path + "TSM01.HLP"
    retcode = Shell(a$, 1)
End Sub


'================================
'    Form Events
'================================

'========================================================================
' Procedure : Form_Activate
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   :
'   1. Setup status bar
'   2. Load data from DB
'========================================================================
Private Sub Form_Activate()
    'Setup status bar: Date
    Sts_MsgLine.Panels(2) = GetCurrentDay(1)
    Me.Refresh
    
    'Initialize
    m_FieldError% = -1
    m_ExitTrigger% = False

    'Determind whether call by others
    If Trim(G_FormFrom$) <> "" Then
       G_FormFrom$ = ""
       '.....                '�[�J�ҭn�]�w���ʧ@
       '.....
       Exit Sub
    Else
       '.....                '�Ĥ@������ɤ��ǳưʧ@
       '.....
'       CBO_A0101_Prepare
'       OpenMainFile         '�p�Ĥ@���}�ɷǳƸ�����
'       MoveDB2Field

        'Load Data from database
        Me.MousePointer = HOURGLASS
        OpenMainFile            '���d�Ҭ����s�}��, ��ܸ��
        MoveDB2Field
        '....
        '....
        Me.MousePointer = Default
        
        'Setup status bar: Operations Msg
        G_AP_STATE = G_AP_STATE_NORMAL  '�]�w�@�~���A
        Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
    End If
End Sub


'========================================================================
' Procedure : Form_KeyDown
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Handle the Key Event
'========================================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
           Case KEY_F1
                KeyCode = 0
                If cmd_help.Visible = True And cmd_help.Enabled = True Then
                   cmd_help.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_F3
                KeyCode = 0
                If cmd_delete.Visible = True And cmd_delete.Enabled = True Then
                   cmd_delete.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
           Case KEY_PAUSE, KEY_ESCAPE
                KeyCode = 0
                If cmd_exit.Visible = True And cmd_exit.Enabled = True Then
                   cmd_exit.SetFocus
                   DoEvents
                   SendKeys "{Enter}"
                End If
    End Select
End Sub


'========================================================================
' Procedure : Form_Load
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   :
'   1. Center the form
'   2. Initialize the form
'========================================================================
Private Sub Form_Load()
    FormCenter Me                     '�e���m���B�z
    Set_Property                      '�]�w���e��������ݩ�
End Sub

'========================================================================
' Procedure : Form_KeyPress
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   :
'   1. Translate Lower Case charactor to Upper Case
'   2. Do nothing if Focus on some (specified) components
'========================================================================
Private Sub Form_KeyPress(KeyAscii As Integer)
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)

   
    '  �Y���Y����줣�ݭn�ഫ��, �����H���L
    If ActiveControl.TabIndex = Spd_TSM01.TabIndex And _
        Spd_TSM01.ActiveCol <> 1 Then GoTo Form_KeyPress_A
    'If ActiveControl.TabIndex = txt_xxx.TabIndex Then GoTo Form_KeyPress_A
    'If ActiveControl.TabIndex = txt_yyy.TabIndex Then GoTo Form_KeyPress_A
    'If ActiveControl.TabIndex = txt_zzz.TabIndex Then GoTo Form_KeyPress_A

    '�D�ʱN��ƿ�J�Ѥp�g�ର�j�g
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If

Form_KeyPress_A:
    If ActiveControl.TabIndex <> Spd_TSM01.TabIndex Then
       KeyPress KeyAscii           'Enter�ɦ۰ʸ���U�@���, spread���~
    End If
End Sub

'========================================================================
' Procedure : Form_QueryUnload
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   :
'   �YUser��������Windows, �ӥ��{�����b����, ���{���|���߰ݬO�_�n�������ۤv?
'   �H�U���зǫ��O, ���o�ק�
'========================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim MSG

    If UnloadMode > 0 Then
       MSG = GetSIniStr("PgmMsg", "g_gui_run")   ' If exiting the application.
    Else
       CloseFileDB
       End
    End If
    
' If user clicks the 'No' button, stop QueryUnload.
    If MsgBox(MSG, 36, Me.Caption) = 7 Then
       Cancel = True
    Else
       CloseFileDB
       End
    End If
End Sub

'========================================================================
' Procedure : Form_Unload
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   :
'========================================================================
Private Sub Form_Unload(Cancel As Integer)
    Cmd_Exit_Click
End Sub


'================================
'    Spread Events
'================================

'========================================================================
' Procedure : Spd_TSM01_Change
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Update Action record while value of Spread is changed
'========================================================================
Private Sub Spd_TSM01_Change(ByVal Col As Long, ByVal Row As Long)
Dim A_A0101$
Dim A_A0102$
Dim A_A0104$
Dim A_A0111$
Dim A_A0101o$

'�p����@��즳���ܧ��, �bP-key�O�ťձ��p�U, ���P�s�W,
'  �_�h���ק窱�A

    With Spd_TSM01
        .Row = Row
        .Col = 1: A_A0101$ = Trim(.text)
        .Col = 2: A_A0102$ = Trim(.text)
        .Col = 3: A_A0104$ = Trim(.text)
        .Col = 4: A_A0111$ = Trim(.text)
        .Col = 5: A_A0101o$ = Trim(.text)
        
        'Update Action Record (Column6)
        .Col = 6
        If A_A0101o$ <> "" Then 'if col5 not empty: Exist Row
           .text = "U"
        Else
            'if something in the row: New Row
            If A_A0101$ + A_A0102$ + A_A0104$ <> "" Then
                .text = "A"
            Else
                .text = ""  'no data writen
            End If
        End If
    End With
End Sub


'========================================================================
' Procedure : Spd_TSM01_GotFocus
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Change Color while got Focus
'========================================================================
Private Sub Spd_TSM01_GotFocus()
    
    SpreadGotFocus Spd_TSM01.ActiveCol, Spd_TSM01.ActiveRow
    
End Sub


'========================================================================
' Procedure : Spd_TSM01_KeyUp
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Hotfix while Typing Chinese
'========================================================================
Private Sub Spd_TSM01_KeyUp(KeyCode As Integer, Shift As Integer)

    '�зǫ��O, �קK����r�Ĥ@�Ӧr�W���h, ���o�ק�
    SpreadKeyPress Spd_TSM01, KeyCode
    
End Sub


'========================================================================
' Procedure : Spd_TSM01_LeaveCell
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Do as below while edit cell and exit
'   1. Handle conditions that don't need to take care
'   2. Add one new row if the last row was edited
'   3. Check constraint and Write to DB while leave original row
'   4. If still in the same row, check Pkey valid
'   5. Change color: Reset original and Change new one
'========================================================================
Private Sub Spd_TSM01_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'�зǫ��O, ���o�ק�
On Local Error GoTo My_Error

    SpreadLostFocus Col, Row

'1. �P�_�H�U���p�o�ͮ�, ����������B�z
    If ActiveControl.TabIndex = cmd_exit.TabIndex Then Exit Sub
    If ActiveControl.TabIndex = cmd_delete.TabIndex Then Exit Sub
    
    With Spd_TSM01

'2. �P�_�b�̫�@������J��, �۰ʼW�[�@��
'   �зǫ��O, �קK�ק�
         .Row = Row: .Col = Col
         If Trim(.text) <> "" And Row = .MaxRows Then
            .MaxRows = .MaxRows + 1
         End If

'3. �Y�����ӵ���, ���ˬd�Ҧ����O�_���T, �A�s��
'   ���P�_�ӵ��O�_������
         .Row = Row
         .Col = 6   'allocate to Action record
         
         'Check if Leaving original Row: Do Action
         If Row <> NewRow And Trim(.text) <> "" Then
         
            '�зǫ��O, �קK�ק�
            '===========================================
            Dim A_Col&
            If SpreadLineCheck(Row, A_Col&) = False Then
               Cancel = True
               .Row = Row: .Col = A_Col&
               .Action = SS_ACTION_ACTIVE_CELL
               .SetFocus
               SpreadGotFocus A_Col&, Row
               Exit Sub
            End If
            '===========================================
            
            'Write Data to DB
            MoveField2DB Row
            
            'Skip "SpreadCheck_1()"
            GoTo New_Cell
         End If

'4. �C���O�_�n�ˬd
         .Row = Row
         .Col = 6   'allocate to Action record
         
         'Still in the same row: Check Pkey valid
         If Trim(.text) <> "" Then
            Select Case Col
              Case 1
                   retcode = SpreadCheck_1(Row)
            ' Case 2
            '      retcode = SpreadCheck_2(Row)
            ' Case 3
            '      retcode = SpreadCheck_3(Row)
            End Select
         End If
    End With
    
'5. �s����C��B�z, �зǫ��O, ���o�ק�
New_Cell:
    If NewCol > 0 Then SpreadGotFocus NewCol, NewRow
    Exit Sub
    
My_Error:
    Sts_MsgLine.Panels(1) = Error(Err)
End Sub

'========================================================================
' Procedure : Spd_TSM01_MouseDown
' Author    : Mike_chang
' Date      : 2015/8/26
' Purpose   : Set status bar msg
'========================================================================
Private Sub Spd_TSM01_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Sts_MsgLine.Panels(1) = SetMessage(G_AP_STATE)
End Sub

