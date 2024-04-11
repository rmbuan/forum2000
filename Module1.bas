Attribute VB_Name = "VBClient"
' For the forum form
Public fMainForm As frmForums
' Global string for the forum name
' This is later used in other forms
Public ForumName As String
' Global string for message name
Public MessageName As Variant
' This is for COM Objects!
'Public cnn1 As Project1.Class1
' Our ADODB Connection
Public cnn1 As ADODB.Connection
' Our constant string for our ADODB ConenctionString
Public strCnn As String

' This method is used for unified connection to the database. Since we only need to connect
' to the database once, we will do it here in the module.
Sub connect()
    
    ' Changed the ConnectionString so that we can run this on any machine.
    strCnn = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\Discussions.mdb;"
    Set cnn1 = New ADODB.Connection
    cnn1.ConnectionString = strCnn '"DSN=Discussion;"
    cnn1.Open

End Sub

' Our Main Sub. This Sub is called by our program when it is first loaded to the memory.
' It then calls the corresponding commands that follows below.
Sub Main()
    
    ' Show our splash screen
    frmSplash.Show
    ' Refresh our splash screen
    frmSplash.Refresh
    
    ' Screwed up timer! :)
    For i = 1 To 10000000
    Next i
    
    ' Create a new instance of frmForums
    Set fMainForm = New frmForums
    ' Load the new instance
    Load fMainForm
    ' Unload our splash screen
    Unload frmSplash
    ' Show our main form
    fMainForm.Show

End Sub

