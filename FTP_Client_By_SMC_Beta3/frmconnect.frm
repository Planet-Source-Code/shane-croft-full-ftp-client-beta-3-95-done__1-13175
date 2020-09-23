VERSION 5.00
Begin VB.Form frmconnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect to..."
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Connection Mode"
      Height          =   855
      Left            =   3923
      TabIndex        =   30
      Top             =   2520
      Width           =   3375
      Begin VB.OptionButton optPassive 
         Caption         =   "Passive Connection"
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton optActive 
         Caption         =   "Active Connection"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Transfer Mode"
      Height          =   855
      Left            =   323
      TabIndex        =   27
      Top             =   2520
      Width           =   3375
      Begin VB.OptionButton optBin 
         Caption         =   "Binary File Transfer"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optAscii 
         Caption         =   "ASCII File Transfer"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quick Connect"
      Height          =   1215
      Left            =   0
      TabIndex        =   16
      Top             =   3600
      Width           =   7575
      Begin VB.CommandButton Command6 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6240
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Connect"
         Height          =   375
         Left            =   6240
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtPass2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3600
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtUser2 
         Height          =   285
         Left            =   3600
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox TxtPort2 
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Text            =   "21"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtServer2 
         Height          =   285
         Left            =   600
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Password:"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Username:"
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "ftp://"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   375
      Left            =   6240
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox TxtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox TxtUser 
      Height          =   285
      Left            =   3600
      TabIndex        =   10
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox TxtPort 
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Text            =   "21"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox TxtServer 
      Height          =   285
      Left            =   3600
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox TxtSiteName 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Password:"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Username:"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Site Name:"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ftp://"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ftp Site Address Book"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmconnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Private Sub Command1_Click()
DoEvents
DoEvents
frmconnect.Hide
DoEvents
DoEvents
frmmain.TxtConnectedTo.Text = frmconnect.TxtServer.Text
frmmain.Caption = "FTP Client By SMC - Connected To - " & frmconnect.TxtSiteName.Text
    If frmmain.mFTP.OpenConnection(TxtServer.Text, TxtPort.Text, TxtUser.Text, TxtPass.Text) Then
        
        If frmconnect.optActive = True Then
        frmmain.mFTP.SetModeActive
        Else
        frmmain.mFTP.SetModePassive
        End If
        
        If frmconnect.optBin = True Then
        frmmain.mFTP.SetTransferBinary
        Else
        frmmain.mFTP.SetTransferASCII
        End If
        
        frmmain.mFTP.SetFTPDirectory "/"
        frmmain.RefreshDirectoryListing
    End If
frmconnect.Hide

End Sub

Private Sub Command2_Click()
If TxtSiteName.Text = "" Then
MsgBox "You must choose a Address Book entry to edit."
Exit Sub
End If


frmEdit.TxtSiteName.Text = frmconnect.TxtSiteName.Text
frmEdit.TxtServer.Text = frmconnect.TxtServer.Text
frmEdit.TxtPort.Text = frmconnect.TxtPort.Text
frmEdit.TxtUser.Text = frmconnect.TxtUser.Text
frmEdit.TxtPass.Text = frmconnect.TxtPass.Text
DoEvents
If frmconnect.optBin.Value = True Then
frmEdit.optBin.Value = True
Else
frmEdit.optAscii.Value = True
End If

If frmconnect.optActive.Value = True Then
frmEdit.optActive.Value = True
Else
frmEdit.optPassive.Value = True
End If
frmEdit.Show vbModal, Me
End Sub

Private Sub Command3_Click()
frmAdd.Show vbModal, Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
rs.FindFirst "[ID] = " & List1.ItemData(List1.ListIndex)
rs.Delete
List1.RemoveItem List1.ListIndex
List1.Clear
Call LoadDatabaseList
DoEvents
DoEvents
    TxtSiteName.Text = ""
    TxtServer.Text = ""
    TxtPort.Text = "21"
    TxtUser.Text = ""
    TxtPass.Text = ""
    optBin.Value = True
    optActive.Value = True

End Sub

Private Sub Command5_Click()
DoEvents
DoEvents
frmconnect.Hide
DoEvents
DoEvents
frmmain.TxtConnectedTo.Text = frmconnect.TxtServer2.Text
frmmain.Caption = "FTP Client By SMC - Connected To - " & frmconnect.TxtServer2.Text
    If frmmain.mFTP.OpenConnection(TxtServer2.Text, TxtPort2.Text, TxtUser2.Text, TxtPass2.Text) Then
        frmmain.mFTP.SetFTPDirectory "/"
        frmmain.RefreshDirectoryListing
    End If
frmconnect.Hide
End Sub

Private Sub Command6_Click()
frmconnect.Hide
End Sub

Private Sub Form_Load()
    Set db = OpenDatabase(App.Path & "\Data.mdb")
    Set rs = db.OpenRecordset("SELECT * FROM Master " & "ORDER BY [Site Name]")
    Call LoadDatabaseList

End Sub
Public Sub LoadDatabaseList()
    Set db = OpenDatabase(App.Path & "\Data.mdb")
    Set rs = db.OpenRecordset("SELECT * FROM Master " & "ORDER BY [Site Name]")
    
    ' Populate the list box
    Do Until rs.EOF
        frmconnect.List1.AddItem rs.Fields("Site Name")
        frmconnect.List1.ItemData(frmconnect.List1.NewIndex) = rs.Fields("ID")
        
        rs.MoveNext
    Loop
End Sub
Private Sub List1_Click()
    rs.FindFirst "[ID] = " & List1.ItemData(List1.ListIndex)
    
    TxtSiteName.Text = rs.Fields("Site Name") & ""
    TxtServer.Text = rs.Fields("FTP") & ""
    TxtPort.Text = rs.Fields("Port") & ""
    TxtUser.Text = rs.Fields("UserName") & ""
    TxtPass.Text = rs.Fields("PassWord") & ""
If rs.Fields("Transfer") = 1 Then
optBin.Value = True
Else
optAscii.Value = True
End If

If rs.Fields("Connection") = 1 Then
optActive.Value = True
Else
optPassive.Value = True
End If
End Sub
