VERSION 5.00
Begin VB.Form frmNewForum 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Create a New Forum"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fNewForum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create New Forum"
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtForumID 
         Height          =   615
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtDescription 
         Height          =   3255
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label lblForumID 
         BackStyle       =   0  'Transparent
         Caption         =   "Forum ID:"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Forum Description:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSaveForum 
      Caption         =   "Save Forum"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      DownPicture     =   "frmNewForum.frx":0000
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmNewForum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myRS As ADODB.Recordset

Private Sub cmdOk_Click()

    cmdSaveForum_Click
    Unload Me

End Sub


Private Sub cmdSaveForum_Click()

    If txtForumID <> "" Then
        If txtDescription <> "" Then
            myRS.AddNew
            myRS.Fields("ForumID") = txtForumID.Text
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
    myRS.Open ("select * from Forums")

End Sub
