VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "myStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private stack As Collection
Const stackSize As Double = 20


Public Function reset()
    Dim i As Double
    If Me.isEmpty Then
        Return
    End If
    
    For i = stack.Count To 1 Step -1
        stack.Remove i
        
    Next i
    
'    Debug.Print stack.Count
    
End Function


Public Function push(ByVal Cin As String)
    stack.Add Cin
    Debug.Print "push: " & stack.Item(stack.Count)
    
End Function


Public Function check(ByVal Cin As Double) As String
    If isEmpty Then
        check = ""
    Else
        check = stack.Item(Cin)
    End If
End Function

Public Function getSize() As Double
    getSize = stack.Count
End Function

Public Function isEmpty() As Boolean
    If stack.Count = 0 Then
        isEmpty = True
    Else
        isEmpty = False
    End If
End Function

Public Function pop() As String
    Dim i As Double
    Dim ret As String
    
    ret = ""
    i = stack.Count
    If i > 0 Then
        ret = stack.Item(i)
        stack.Remove (i)
    End If
    
    Debug.Print "pop: " & ret
    
    pop = ret
End Function

Private Sub Class_Initialize()
    Dim str As String
'    MsgBox "Class Init"
    Set stack = New Collection
End Sub

