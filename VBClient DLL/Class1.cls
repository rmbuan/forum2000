VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public cnn As ADODB.Connection
Public rsTest As ADODB.Recordset
Public cnnString As String
    
Public Function Connect()

'    cnnString = "FileDSN=" + App.Path + "\forums.dsn"
        
    'MsgBox cnnString
    ' Open a connection using a DSN and ODBC tags.
    Set cnn = New ADODB.Connection
    cnn.ConnectionString = "DSN=Discussion;"
    cnn.Open
        
End Function

Public Function SQLSelect(SQLString As String, ForumName As String)
    
    'rsTest.d
    Set rsTest = cnn.Execute(SQLString & "WHERE ForumID = " & ForumName)

End Function

Public Function SQLSelect2(SQLString As String)
    
    'rsTest.d
    Set rsTest = cnn.Execute(SQLString)

End Function

