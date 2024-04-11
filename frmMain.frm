VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmForums 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MB Software Forum 2000 v1.0.1"
   ClientHeight    =   5250
   ClientLeft      =   3000
   ClientTop       =   2655
   ClientWidth     =   8325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1006
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOpenForum 
      Caption         =   "Open Forum"
      Height          =   375
      Left            =   6480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Forum"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New Forum"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   5530
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4980
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9499
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "25/05/2000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:25 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please select a forum, or create a new one"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   4635
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmForums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnn As ADODB.Connection

Dim myRS1 As ADODB.Recordset
Dim itmX As ListItem
   

Private Sub cmdEdit_Click()
    
    ForumName = ListView1.SelectedItem
    frmForum.Show vbModal, Me
   
End Sub

Private Sub cmdNew_Click()

    frmNewForum.Show 1
'    If Not cnn1.rsTest.EOF Then
'        MsgBox "Requery!"
'        cnn1.rsTest.Requery
'        cnn1.rsTest.Requery
'        cnn1.rsTest.MoveFirst
'    End If
 
End Sub

Private Sub cmdOpenForum_Click()
'    MsgBox ListView1.SelectedItem
    
    ForumName = ListView1.SelectedItem
    
    frmForumMessages.Show 1

End Sub

Private Sub Form_Activate()
 
    display
    
End Sub

Private Sub Form_Load()
'cnn1.connect
    connect

'"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Discussions.mdb;Mode=Read|Write|Share Deny None;Persist Security Info=False"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End

End Sub
Sub display()

    Set myRS1 = New ADODB.Recordset
    myRS1.CursorType = adOpenKeyset
    myRS1.LockType = adLockOptimistic
    myRS1.ActiveConnection = cnn1
    myRS1.Open ("select * from Forums")
       
    ListView1.ListItems.Clear
'    MsgBox "Nothing in the listitem"
    ListView1.ColumnHeaders. _
        Add , , "Forum ID", ListView1.Width / 5
    ListView1.ColumnHeaders. _
        Add , , "Forum Description", ListView1.Width / 1
 
    ListView1.View = lvwReport
'    ListView1.SmallIcons = ImageList1
'    MsgBox "MoveFirst"
    If Not myRS1.EOF Then
    Else
        Dim result As Variant
        result = MsgBox("There are no forums, Do you want to create a new message now?", vbYesNo, "Forum 2000")
        If result = vbYes Then
            frmNewForum.Show 1
        Else
            Unload Me
        End If
    End If
    
    While Not myRS1.EOF
 '   MsgBox "im in the EOF"
        Set itmX = ListView1.ListItems. _
            Add(, , CStr(myRS1.Fields("ForumID")), , 3)
        itmX.SubItems(1) = myRS1.Fields("ForumDescription")
        myRS1.MoveNext
    Wend
       
    'myRS1.Close
    'Set myRS1 = Nothing
End Sub

Private Sub ListView1_DblClick()

 '   MsgBox ListView1.SelectedItem
    
    ForumName = ListView1.SelectedItem
    
    frmForumMessages.Show 1
    
End Sub

Private Sub mnuHelpAbout_Click()
    
    frmAbout.Show vbModal, Me

End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    End

End Sub

