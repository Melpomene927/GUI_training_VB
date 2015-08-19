VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   2385
   ClientTop       =   3330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DB As Database
Dim RS As Recordset

Private Sub Form_Load()

'    Set DB = OpenDatabase("E:\HT2000\ARTHGUI\DATA\GUI.mdb", False, False, "")
'
'    Set RS = DB.OpenRecordset("SELECT * FROM A01 WHERE A0101='T'", dbOpenDynaset)
'    RS.AddNew
'    RS.Fields("A0101") = "T"
'    RS.Fields("A0102") = "testcompany"
'    RS.Update
'    DoEvents

'    RS.LockEdits = False
'    RS.Edit
'    RS.Fields("A0102") = RS.Fields("A0102") & "T"
'    RS.Update
'    DoEvents
    
'    RS.Delete
'    DoEvents

'    Set RS = DB.OpenRecordset("SELECT * FROM A01", dbOpenSnapshot)
'    If Not (RS.BOF And RS.EOF) Then
'        RS.MoveLast
'        Debug.Print RS.RecordCount
'        RS.MoveFirst
'    Else
'        Debug.Print "0"
'    End If
    
'    Do While Not RS.EOF
'        Debug.Print RS.Fields("A0101") & ":" & RS.Fields("A0102")
'        RS.MoveNext
'    Loop
'    DoEvent

'Problems
    
    Set DB = OpenDatabase("", False, False, "ODBC;DSN=FamilyGroup;UID=SA;PWD=7669588")
    Set RS = DB.OpenRecordset("SELECT * FROM E_Personal_information WHERE PID=2", dbOpenSnapshot)

'    RS.AddNew
'    RS.Fields("A0101") = "S"
'    RS.Fields("A0102") = "testcompany"
'    RS.Update
    Debug.Print RS.Fields("name")
    DoEvents
    
'    RS.LockEdits = False
'    RS.Edit
'    RS.Fields("A0102") = RS.Fields("A0102") & "T"
'    RS.Update
'    DoEvents

'    RS.Delete
'    DoEvents


'    Set DB = OpenDatabase("", False, False, "ODBC;DSN=ARTHGUI;UID=SA;PWD=7669588")
'    Set RS = DB.OpenRecordset("SELECT * FROM A01", dbOpenSnapshot, dbSQLPassThrough)
'    If Not (RS.BOF And RS.EOF) Then
'        RS.MoveLast
'        Debug.Print RS.RecordCount
'        RS.MoveFirst
'    Else
'        Debug.Print "0"
'    End If
'
'    Do While Not RS.EOF
'        Debug.Print RS.Fields("A0101") & ":" & RS.Fields("A0102")
'        RS.MoveNext
'    Loop
'    DoEvents
    
'    DB.Close
    
'    Set DB = OpenDatabase("", False, False, "ODBC;DSN=ARTHGUI;UID=SA;PWD=7669588")
'
'    DB.Execute "INSERT INTO A01 (A0101,A0102) VALUES ('0','TEST')", dbSQLPassThrough
'
'    DB.Execute "UPDATE A01 SET A0102 = A0102 + 'QQ' WHERE A0101='0'", dbSQLPassThrough
'
'    DB.Execute "DELETE FROM A01 WHERE A0101='0'", dbSQLPassThrough
'
'
'    DoEvents
    RS.Close
    DB.Close

'
    
End Sub
