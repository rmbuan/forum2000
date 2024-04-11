VERSION 5.00
Begin VB.Form frmNewMessage 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Post a new message!"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveMessage 
      Caption         =   "Save Message"
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame fNewMessage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create New Message"
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtPostedBy 
         Height          =   615
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtMessage 
         Height          =   3255
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1920
         Width           =   4455
      End
      Begin VB.TextBox txtSubject 
         Height          =   615
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblBy 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Posted by:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblMessage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Message:"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblSubject 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Subject:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmNewMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myRS As ADODB.Recordset

Private Sub cmdOk_Click()
    
    cmdSaveMessage_Click
    Unload Me

End Sub

Private Sub cmdSaveMessage_Click()

    If txtSubject <> "" Then
        If txtPostedBy <> "" Then
            If txtSubject <> "" Then
                myRS.AddNew
                myRS.Fields("ForumID") = ForumName
                myRS.Fields("DatePosted") = Now()
                myRS.Fields("Subject") = txtSubject.Text
                myRS.Fields("Message") = txtMessage.Text
                myRS.Fields("PostedBy") = txtPostedBy.Text
                myRS.Update
            Else
            End If
        Else
            If txtSubject <> "" Then
                myRS.AddNew
                myRS.Fields("ForumID") = ForumName
                myRS.Fields("DatePosted") = Now()
                myRS.Fields("Subject") = txtSubject.Text
                myRS.Fields("Message") = txtMessage.Text
                myRS.Fields("PostedBy") = "Anonymous Coward"
                myRS.Update
            Else
            End If
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
    myRS.Open ("select * from Messages where ForumID = '" & ForumName & "'")

End Sub

