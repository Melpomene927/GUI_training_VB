VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "ss32x25.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmSpread 
   Caption         =   "Spread Practice"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   7245
      Left            =   60
      OleObjectBlob   =   "Spread.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   4035
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   495
      Left            =   4425
      TabIndex        =   7
      Top             =   4530
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Write Data"
      Enabled         =   0   'False
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   495
      Left            =   4425
      TabIndex        =   6
      Top             =   3750
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Load Data"
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   495
      Left            =   4425
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "結束"
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   495
      Left            =   4425
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "清值"
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1275
      Left            =   4260
      TabIndex        =   3
      Top             =   300
      Width           =   1785
      _Version        =   65536
      _ExtentX        =   3149
      _ExtentY        =   2249
      _StockProps     =   14
      Caption         =   "排列順序"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSOption SSOption_Desc 
         Height          =   285
         Left            =   330
         TabIndex        =   4
         Top             =   780
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "由大到小"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption_Asc 
         Height          =   285
         Left            =   330
         TabIndex        =   5
         Top             =   360
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "由小到大"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
   End
End
Attribute VB_Name = "frmSpread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SortAsc As Boolean
Dim num_rows As Long
Dim num_cols As Long
Dim the_array() As String

Private Sub Form_Activate()
    loadData
End Sub

Private Sub Form_Load()


    If SSOption_Asc.Value = True Then
        SortAsc = True
    Else
        SortAsc = False
    End If
    
    
End Sub

Private Sub SSCommand1_Click()
    Dim i%, j%
    
    For i% = 1 To vaSpread1.MaxRows - 1
        For j% = 1 To vaSpread1.MaxCols

            vaSpread1.col = j%
            vaSpread1.Col2 = j%
            vaSpread1.Row = i%
            vaSpread1.row2 = i%
            vaSpread1.Text = ""
            
                
        Next
    Next
End Sub

Private Sub SSCommand2_Click()
    Unload Me
End Sub

Private Sub SSCommand3_Click()
    loadData
End Sub

Private Sub SSOption_Asc_Click(Value As Integer)
    SortAsc = True
End Sub

Private Sub SSOption_Desc_Click(Value As Integer)
    SortAsc = False
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim strAsc As String
    If SortAsc Then
        strAsc = " Asc"
    Else
        strAsc = " Desc"
    End If
    
    If (BlockRow = -1 And BlockRow2 = -1) Then
        MsgBox "hit col: " + Str(BlockCol) + strAsc + _
            "（°o°；）", vbInformation + vbOKOnly, "(’⊙ω⊙`)！"
        
        bubbleSort BlockCol
        
    End If
    
End Sub

Private Sub vaSpread1_Click(ByVal col As Long, ByVal Row As Long)
    
'    Debug.Print "vaSpread1_Click"
'    Debug.Print "Col: " + Str(Col)
'    Debug.Print "Row: " + Str(Row)
'    Debug.Print "-----------"

End Sub

Public Sub loadData()
Dim file_name As String
Dim fnum As Integer
Dim whole_file As String
Dim lines As Variant
Dim one_line As Variant

Dim R As Long
Dim C As Long
    
    'build file path
    file_name = App.Path
    If Right$(file_name, 1) <> "\" Then file_name = _
        file_name & "\"
    file_name = file_name & "data.csv"

    ' Load the file.
    fnum = FreeFile
    Open file_name For Input As fnum
    whole_file = Input$(LOF(fnum), #fnum)
    Close fnum

    ' Break the file into lines.
    lines = Split(whole_file, vbCrLf)

    ' Dimension the array.
    num_rows = UBound(lines)
    one_line = Split(lines(0), ",")
    num_cols = UBound(one_line)
    ReDim the_array(num_rows, num_cols)

    ' Copy the data into the array.
    For R = 0 To num_rows
        If Len(lines(R)) > 0 Then
            one_line = Split(lines(R), ",")
            For C = 0 To num_cols
                the_array(R, C) = one_line(C)
            Next C
        End If
    Next R

    ' Prove we have the data loaded.
    For R = 0 To num_rows
        For C = 0 To num_cols
            setCol C + 1
            setRow R + 1
            vaSpread1.Clip = the_array(R, C)
        Next C
    Next R
    
End Sub

Public Sub bubbleSort(col As Long)
Dim i As Long
Dim j As Long
Dim amount As Long
Dim a As Variant, b As Variant

With vaSpread1
    amount = num_rows
    
    For i = 1 To amount
        For j = 1 To amount - i
            setCol col
            setRow j
            a = .Clip
            setRow j + 1
            b = .Clip
            
            If SortAsc Then
                If compare(a, b) Then
                    swapRow j, j + 1
                End If
            Else
                If compare(b, a) Then
                    swapRow j, j + 1
                End If
            End If
        Next
    Next
End With
End Sub


Public Sub swapRow(row1 As Long, row2 As Long)
Dim valueR1 As Variant
Dim valueR2 As Variant
Dim i As Long
    With vaSpread1
        For i = 1 To .MaxCols
            '===== set column number ======
            setCol i
            
            '===== keep r1 =====
            setRow row1
            valueR1 = .Clip
            
            '===== keep r2, edit r2 =====
            setRow row2
            valueR2 = .Clip
            .Clip = valueR1
            
            '===== edit r1 =====
            setRow row1
            .Clip = valueR2
        Next
    End With
End Sub


Public Sub setRow(i As Long)
    With vaSpread1
        .Row = i
        .row2 = i
    End With
End Sub

Public Sub setCol(i As Long)
    With vaSpread1
        .col = i
        .Col2 = i
    End With
End Sub


Public Function compare(a As Variant, b As Variant) As Boolean
    Dim ans As Boolean
    ans = False
    If Asc(a) > 57 Then
'        MsgBox "◢▆▅▄▃崩╰(〒皿〒)╯潰▃▄▅▇◣" _
'            , vbCritical + vbOKOnly, "(’⊙ω⊙`)！"
        If StrComp(a, b) > 0 Then
            ans = True
        End If
    Else
        If CDbl(a) > CDbl(b) Then
            ans = True
        End If
    End If
    
    
    
    compare = ans
End Function
