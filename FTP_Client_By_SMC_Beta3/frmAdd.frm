VERSION 5.00
Begin VB.Form frmAdd 
   Caption         =   "Add Entry To Address Book"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   4725
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "File Transfer Mode"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   3135
      Begin VB.OptionButton optAscii 
         Caption         =   "ASCII File Transfer"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton optBin 
         Caption         =   "Binary File Transfer"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Connection Mode"
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   3135
      Begin VB.OptionButton optActive 
         Caption         =   "Active Connection"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optPassive 
         Caption         =   "Passive Connection"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TxtSiteName 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TxtServer 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox TxtPort 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "21"
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox TxtUser 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox TxtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "ftp://"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Site Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Private Sub Command1_Click()

rs.AddNew
DoEvents
    rs.Fields("Site Name") = TxtSiteName.Text
    rs.Fields("FTP") = TxtServer.Text
    rs.Fields("Port") = TxtPort.Text
    rs.Fields("UserName") = TxtUser.Text
    rs.Fields("PassWord") = TxtPass.Text
If optBin.Value = True Then
rs.Fields("Transfer") = 1
Else
rs.Fields("Transfer") = 2
End If

If optActive.Value = True Then
rs.Fields("Connection") = 1
Else
rs.Fields("Connection") = 2
End If
rs.Update
DoEvents
frmconnect.List1.Clear
Call frmconnect.LoadDatabaseList
DoEvents
DoEvents
    TxtSiteName.Text = ""
    TxtServer.Text = ""
    TxtPort.Text = "21"
    TxtUser.Text = ""
    TxtPass.Text = ""
    optBin.Value = True
    optActive.Value = True
    MsgBox "The information has been added to your address book."
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\Data.mdb")
    Set rs = db.OpenRecordset("SELECT * FROM Master " & "ORDER BY [Site Name]")
End Sub
