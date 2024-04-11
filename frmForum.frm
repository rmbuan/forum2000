VERSION 5.00
Begin VB.Form frmForum 
   BackColor       =   &H00FFFFFF&
   Caption         =   "View/Edit Existing Forum"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveForum 
      Caption         =   "Save Forum"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdDeleteForum 
      Caption         =   "Delete Forum"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame fEditForum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Forum"
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.TextBox txtDescription 
         Height          =   3255
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtForumID 
         Height          =   615
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Forum Description:"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblForumID 
         BackStyle       =   0  'Transparent
         Caption         =   "Forum ID:"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmForum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myRS As ADODB.Recordset
Dim myRS2 As ADODB.Recordset

Private Sub cmdDeleteForum_Click()

    Set myRS2 = New ADODB.Recordset
    myRS2.CursorType = adOpenKeyset
    myRS2.LockType = adLockOptimistic
    myRS2.ActiveConnection = cnn1
    myRS2.Open ("select * from Messages where ForumID = '" & ForumName & "'")
    
    While Not myRS2.EOF
        'MsgBox myRS2.Fields
        myRS2.Delete
        myRS2.Update
        myRS2.MoveNext
    Wend
            
    myRS.Delete
    myRS.Update
    Unload Me
    
End Sub

Private Sub cmdOk_Click()

    cmdSaveForum_Click
    Unload Me
    
End Sub


Private Sub cmdSaveForum_Click()
    If txtForumID <> "" Then
        If txtDescription <> "" Then
            myRS.Fields("ForumID") = ForumName
            myRS.Fields("ForumDescription") = txtDescription.Text
            myRS.Update
        Else
        End If
    Else
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Set myRS = New ADODB.Recordset
    myRS.CursorType = adOpenKeyset
    myRS.LockType = adLockOptimistic
    myRS.ActiveConnection = cnn1
    myRS.Open ("select * from Forums where ForumID = '" & ForumName & "'")
    
    If Not myRS.EOF Then
        txtForumID.Text = myRS.Fields("ForumID")
        txtDescription.Text = myRS.Fields("ForumDescription")
    End If
    
End Sub
