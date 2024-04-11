VERSION 5.00
Begin VB.Form frmMessage 
   BackColor       =   &H00FFFFFF&
   Caption         =   "View Message"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDeleteMessage 
      Caption         =   "Delete Message"
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   720
      Width           =   1455
   End
   Begin VB.Frame fNewMessage 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Message"
      Height          =   6015
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtDatePosted 
         Height          =   615
         Left            =   1920
         TabIndex        =   10
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtSubject 
         Height          =   615
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtMessage 
         Height          =   3255
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   2520
         Width           =   4455
      End
      Begin VB.TextBox txtPostedBy 
         Height          =   615
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblDatePosted 
         BackStyle       =   0  'Transparent
         Caption         =   "Date posted:"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblSubject 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblMessage 
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lblBy 
         BackStyle       =   0  'Transparent
         Caption         =   "Posted by:"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSaveMessage 
      Caption         =   "Save Message"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myRS As ADODB.Recordset

Private Sub cmdDeleteMessage_Click()

    myRS.Delete
    Unload Me
    
End Sub

Private Sub cmdOk_Click()

    Unload Me

End Sub

Private Sub cmdSaveMessage_Click()
    If txtSubject <> "" Then
        If txtPostedBy <> "" Then
            If txtSubject <> "" Then
'                myRS.EditMode
                'myRS.Fields("ForumID") = ForumName
                'myRS.Fields("DatePosted") = Now()
                myRS.Fields("Subject") = txtSubject.Text
                myRS.Fields("Message") = txtMessage.Text
                myRS.Fields("PostedBy") = txtPostedBy.Text
                myRS.Update
            Else
            End If
        Else
        End If
    Else
    End If
End Sub

Private Sub Form_Activate()
    
    Set myRS = New ADODB.Recordset
    myRS.CursorType = adOpenKeyset
    myRS.LockType = adLockOptimistic
    myRS.ActiveConnection = cnn1
    myRS.Open ("select * from Messages where MessageID = " & MessageName)
    
    If Not myRS.EOF Then
        txtSubject.Text = myRS.Fields("Subject")
        txtPostedBy.Text = myRS.Fields("PostedBy")
        txtDatePosted.Text = myRS.Fields("DatePosted")
        txtMessage.Text = myRS.Fields("Message")
'        myRS.MoveNext
    End If

End Sub

