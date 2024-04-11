VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmForumMessages 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5085
   ClientLeft      =   3570
   ClientTop       =   2940
   ClientWidth     =   10830
   Icon            =   "frmForumMessages.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   10830
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   9240
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewMessage 
      Caption         =   "View Message"
      Height          =   375
      Left            =   9240
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdPostMessage 
      Caption         =   "Post Message"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblForumMessages 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmForumMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myRS As ADODB.Recordset

Private Sub cmdOk_Click()
    
    Unload Me

End Sub

Private Sub cmdPostMessage_Click()

    frmNewMessage.Show 1

End Sub

Private Sub cmdViewMessage_Click()

    MessageName = ListView1.SelectedItem
    frmMessage.Show 1

End Sub

Private Sub Form_Activate()
    
    Set myRS = New ADODB.Recordset
    
    myRS.ActiveConnection = cnn1
      
    myRS.Open ("select * from Messages where ForumID = '" & ForumName & "'")
    
        
  Dim itmX As ListItem
  ListView1.ColumnHeaders. _
   Add , , "Message ID", ListView1.Width / 6
'  ListView1.ColumnHeaders. _
'   Add , , "Forum ID", ListView1.Width / 5
  ListView1.ColumnHeaders. _
   Add , , "Date Created", ListView1.Width / 4
  ListView1.ColumnHeaders. _
   Add , , "Message Subject", ListView1.Width / 3
  ListView1.ColumnHeaders. _
   Add , , "Posted by", ListView1.Width / 1

 ListView1.View = lvwReport
 ListView1.SmallIcons = frmForums.ImageList1
 
    If Not myRS.EOF Then
    Else
        Dim result As Variant
        result = MsgBox("There are no messages for Forum " & ForumName & vbCrLf & "Do you want to post a new message now?", vbYesNo, "Forum 2000")
        If result = vbYes Then
            frmNewMessage.Show 1
        Else
            Unload Me
        End If
    End If
    
    ListView1.ListItems.Clear
    If Not myRS.EOF Then
        myRS.Requery
        myRS.MoveFirst
    End If
    
    While Not myRS.EOF
        Set itmX = ListView1.ListItems. _
          Add(, , CStr(myRS.Fields("MessageID")), , 4)
'        itmX.SubItems(1) = myRS.Fields("ForumID")
        itmX.SubItems(1) = myRS.Fields("DatePosted")
        itmX.SubItems(2) = myRS.Fields("Subject")
'        itmX.SubItems(4) = cnn2.rsTest.Fields("ForumURL")
        itmX.SubItems(3) = myRS.Fields("PostedBy")
        myRS.MoveNext
    Wend
    
    lblForumMessages.Caption = "Forum " & ForumName & " Messages"
    Me.Caption = lblForumMessages.Caption
    
End Sub

Private Sub ListView1_DblClick()
 '   MsgBox ListView1.SelectedItem
    
    MessageName = ListView1.SelectedItem
    
    frmMessage.Show 1
    
End Sub
