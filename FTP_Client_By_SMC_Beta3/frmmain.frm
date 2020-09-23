VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "FTP Client by SMC"
   ClientHeight    =   7845
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1440
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8760
      TabIndex        =   21
      Text            =   "0"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   5640
   End
   Begin VB.TextBox TxtTotalBytesQueued 
      Height          =   285
      Left            =   8160
      TabIndex        =   15
      Text            =   "0"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   5640
   End
   Begin VB.TextBox TxtConnectedTo 
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtRemotePath 
      Height          =   285
      Left            =   5280
      TabIndex        =   8
      Top             =   1800
      Width           =   4935
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   5760
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3625
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Local Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Remote Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Size (Bytes)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Command"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3960
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Height          =   330
      Left            =   10320
      Picture         =   "frmmain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   375
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3255
      Left            =   5280
      TabIndex        =   1
      Top             =   2160
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   5741
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2471
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTBHeader 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   2143
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmmain.frx":0102
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6360
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":01D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":077D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0D21
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":12C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1869
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1E0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":23B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2955
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2EF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":349D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3A41
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3FE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4589
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4B2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4C41
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4D61
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   688
      ButtonWidth     =   767
      ButtonHeight    =   688
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      HotImageList    =   "imlToolbarHot"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Connect"
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Connect"
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Disconnect"
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop operation"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh contents of current directory"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Transfer"
            Object.ToolTipText     =   "Transfer Queue"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Download"
            Object.ToolTipText     =   "Add files to queue for download"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Upload"
            Object.ToolTipText     =   "Add files to queue for upload"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CreateDirectory"
            Object.ToolTipText     =   "Create Directory..."
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete File"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rename"
            Object.ToolTipText     =   "Rename File..."
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Large Icons"
            Object.ToolTipText     =   "View Large Icons"
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Small Icons"
            Object.ToolTipText     =   "View Small Icons"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View List"
            Object.ToolTipText     =   "View List"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "View Details"
            Object.ToolTipText     =   "View Details"
            ImageIndex      =   13
            Style           =   2
            Value           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarHot 
      Left            =   6960
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   20
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4E75
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5419
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":59BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5F61
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6505
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":6AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":704D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":75F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7B95
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8139
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":86DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8C81
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9225
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":97C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":98E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":99FD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9B11
            Key             =   "dir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9C6B
            Key             =   "file"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A0BD
            Key             =   "dll"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":A50F
            Key             =   "web"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":CCC1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Elapsed Time: 00:00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Estimated Time: 00:00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Total Size - 0.00 KB"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Total Files Queued - 0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Total Files Queued Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   4320
      Width           =   5175
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "0 KB of  0 KB Transfered"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Speed: 0 Kbps"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3120
      Width           =   5175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Current File:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   2520
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Time Left: 00:00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Files Queued"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   5520
      Width           =   10695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Current File Progress"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Menu menufile 
      Caption         =   "File"
      Begin VB.Menu menuconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu menuDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu menuline2 
         Caption         =   "-"
      End
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menucommands 
      Caption         =   "Commands"
      Begin VB.Menu menuremove 
         Caption         =   "Remove Item From Queue"
      End
      Begin VB.Menu menuremoveall 
         Caption         =   "Remove All From Queue"
      End
      Begin VB.Menu menuline1 
         Caption         =   "-"
      End
      Begin VB.Menu menutransfer 
         Caption         =   "Transfer Queue"
      End
      Begin VB.Menu menudownload 
         Caption         =   "Download Files"
      End
      Begin VB.Menu menuupload 
         Caption         =   "Upload Files"
      End
      Begin VB.Menu menustop 
         Caption         =   "Stop"
      End
      Begin VB.Menu menuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu menuMD 
         Caption         =   "Make New Dir."
      End
      Begin VB.Menu menuFtpdelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu menuRename 
         Caption         =   "Rename Files"
      End
   End
   Begin VB.Menu menutools 
      Caption         =   "Tools"
      Begin VB.Menu menusettings 
         Caption         =   "Settings"
         Enabled         =   0   'False
      End
      Begin VB.Menu menulinez 
         Caption         =   "-"
      End
      Begin VB.Menu menuUpdate 
         Caption         =   "Automatic Update"
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "Help"
      Begin VB.Menu menuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents mFTP As cFTP
Attribute mFTP.VB_VarHelpID = -1
Private Header As Variant
Private BeginTransfer                   As Single
Private TransferRate                    As Single
Private Declare Function ClipCursor Lib "user32" _
    (lpRect As Any) As Long

Private FilePathName As String
Private Filename As String
Private FormName As String

Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Private NewVersion As String
Private OldVersion As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub RefreshDirectoryListing()
    Dim Item As cDirItem
    Dim lstX As ListItem
    Dim sAttr As String
    mFTP_StateChanged (FTP_RETRIEVING_DIRECTORY_INFO)
    DoEvents
    mFTP.GetDirectoryListing "*.*"
    DoEvents
    ListView2.ListItems.Clear
    DoEvents
    Set lstX = ListView2.ListItems.Add(, , "..")
    lstX.SmallIcon = 1
    TxtRemotePath.Text = mFTP.GetFTPDirectory
    DoEvents
    DoEvents
    For Each Item In mFTP.Directory
         sAttr = ""
         With Item
              If .Archive Then sAttr = sAttr & " A " Else sAttr = sAttr & " - "
              If .Compressed Then sAttr = sAttr & " C " Else sAttr = sAttr & " - "
              If .Directory Then sAttr = sAttr & " D " Else sAttr = sAttr & " - "
              If .Hidden Then sAttr = sAttr & " H " Else sAttr = sAttr & " - "
              If .Normal Then sAttr = sAttr & " N " Else sAttr = sAttr & " - "
              If .Offline Then sAttr = sAttr & " O " Else sAttr = sAttr & " - "
              If .ReadOnly Then sAttr = sAttr & " R " Else sAttr = sAttr & " - "
              If .System Then sAttr = sAttr & " S " Else sAttr = sAttr & " - "
              If .Temporary Then sAttr = sAttr & " T " Else sAttr = sAttr & " - "
         End With
         
         Set lstX = ListView2.ListItems.Add(, , Item.Filename)
         DoEvents
         With lstX
            If Item.Directory Then
               .SmallIcon = 1
               .SubItems(1) = "< Directory >"
               DoEvents
            Else
               .SmallIcon = 2
               DoEvents
               .SubItems(1) = Item.FileSize
               DoEvents
            End If
            DoEvents
         End With
         DoEvents
    Next
    DoEvents
    TxtRemotePath.Text = mFTP.GetFTPDirectory
    mFTP_StateChanged (FTP_DIRECTORY_INFO_COMPLETED)
End Sub

Private Sub Command2_Click()
            mFTP.SetFTPDirectory ".."
            TxtRemotePath.Text = mFTP.GetFTPDirectory
            DoEvents
            DoEvents
            RefreshDirectoryListing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub menuconnect_Click()
            frmconnect.Show vbModal, Me

End Sub

Private Sub menuDisconnect_Click()
    mFTP.CloseConnection
    ListView2.ListItems.Clear
    Timer2.Enabled = False

End Sub

Private Sub menudownload_Click()
Call ListView2_DblClick
End Sub

Private Sub menuFtpdelete_Click()
   If Not (ListView2.SelectedItem Is Nothing) Then
      If MsgBox("Are you sure you want to delete " & ListView2.SelectedItem.Text & "?", vbQuestion + vbYesNo, "Delete?") = vbYes Then
         If mFTP.Directory(ListView2.SelectedItem.Text).Directory Then
            If mFTP.RemoveFTPDirectory(ListView2.SelectedItem.Text) Then
               MsgBox "Directory " & sName & " was successfully removed.", vbInformation
               RefreshDirectoryListing
            Else
               MsgBox mFTP.GetLastErrorMessage
            End If
         Else
            If mFTP.DeleteFTPFile(ListView2.SelectedItem.Text) Then
               MsgBox "The file " & sName & " was successfully deleted.", vbInformation
               RefreshDirectoryListing
            Else
               MsgBox mFTP.GetLastErrorMessage
            End If
         End If
      End If
   End If
End Sub

Private Sub menuMD_Click()
   Dim sName As String
   sName = Trim(InputBox("Please enter a name for this directory:"))
   If sName <> "" Then
      If mFTP.CreateFTPDirectory(sName) Then
         MsgBox "Directory " & sName & " was successfully created.", vbInformation
         RefreshDirectoryListing
      Else
         MsgBox mFTP.GetLastErrorMessage
      End If
   End If
End Sub

Private Sub menuRefresh_Click()
RefreshDirectoryListing
End Sub

Private Sub menuremove_Click()
        TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text - ListView3.SelectedItem.SubItems(4)
        DoEvents
        DoEvents

ListView3.ListItems.Remove ListView3.SelectedItem.Index
End Sub

Private Sub menuremoveall_Click()
ListView3.ListItems.Clear
TxtTotalBytesQueued.Text = "0"
End Sub

Private Sub menuRename_Click()
    If Not (ListView2.SelectedItem Is Nothing) Then
       If ListView2.SelectedItem.Text <> ".." Then
            If Not mFTP.Directory(ListView2.SelectedItem.Text).Directory Then
               Dim sNewName As String
               sNewName = Trim(InputBox("Please enter a new name for this file:"))
               If sNewName <> "" Then
                  If mFTP.RenameFTPFile(ListView2.SelectedItem.Text, sNewName) Then
                     MsgBox "File was successfuly renamed"
                     RefreshDirectoryListing
                  Else
                     MsgBox mFTP.GetLastErrorMessage
                  End If
               End If
            End If
         End If
      End If
End Sub

Private Sub menustop_Click()
On Error Resume Next
        Dim pstrmessage As String
        pstrmessage = MsgBox("This will stop your current transfer and disconnect you from the site, are you sure you want to continue?", vbYesNo)
        If pstrmessage = vbYes Then
            mFTP.CloseConnection
            ListView2.ListItems.Clear
            Timer2.Enabled = False
            MsgBox "Proccess Stopped"
        End If


End Sub

Private Sub menutransfer_Click()
    Dim lTimer As Long
    Dim strRemote As String
    Dim strLocal As String
    Dim y As String
    If ListView3.ListItems.Count = 0 Then
    MsgBox "The are no files queued."
    Exit Sub
    End If
    
    Do Until ListView3.ListItems.Count = 0
        If frmconnect.optActive = True Then
        frmmain.mFTP.SetModeActive
        DoEvents
        Else
        frmmain.mFTP.SetModePassive
        DoEvents
        End If
        DoEvents
        If frmconnect.optBin = True Then
        frmmain.mFTP.SetTransferBinary
        DoEvents
        Else
        frmmain.mFTP.SetTransferASCII
        DoEvents
        End If
               DoEvents
               DoEvents
   BeginTransfer = Timer
   Timer2.Enabled = True
   DoEvents
        Label7.Caption = "Current File: " & ListView3.SelectedItem.Text
        DoEvents
        y = TxtRemotePath.Text
        DoEvents
        DoEvents
        If y = ListView3.SelectedItem.SubItems(2) Then
        
        Else
        mFTP.SetFTPDirectory ListView3.SelectedItem.SubItems(2)
        DoEvents
        DoEvents
        RefreshDirectoryListing
        DoEvents
        DoEvents
        End If
         DoEvents
         DoEvents
         DoEvents
         DoEvents
         
         If ListView3.SelectedItem.SubItems(5) = "Download" Then
          strRemote = ListView3.SelectedItem.Text
          strLocal = ListView3.SelectedItem.SubItems(1) & "\" & ListView3.SelectedItem.Text

               If mFTP.FTPDownloadFile(strLocal, strRemote) Then
                
                Else
                MsgBox mFTP.GetLastErrorMessage & "Unable To Complete Request."
                Exit Sub
                End If
                DoEvents
                DoEvents
        End If
        
        If ListView3.SelectedItem.SubItems(5) = "Upload" Then
          strRemote = ListView3.SelectedItem.Text
          strLocal = ListView3.SelectedItem.SubItems(1)

               If mFTP.FTPUploadFile(strLocal, strRemote) Then
                
                Else
                MsgBox mFTP.GetLastErrorMessage & "Unable To Complete Request."
                Exit Sub
                End If
                DoEvents
                DoEvents
        End If
        TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text - ListView3.SelectedItem.SubItems(4)
        DoEvents
        DoEvents
        ListView3.ListItems.Remove 1
        DoEvents
  Loop
RefreshDirectoryListing
Timer2.Enabled = False
Label10.Caption = "Elapsed Time: 00:00:00"
Text1.Text = "0"
DoEvents
End Sub

Private Sub menuupload_Click()
    Dim vFiles As Variant
    Dim lFile As Long
    Dim y As Long
    With CD1
        .Filename = "" 'Clear the filename
        .CancelError = False 'Gives an error if cancel is pressed
        .DialogTitle = "Select File(s)...  (Multi Select)"
        .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly 'Falgs, allows Multi select, Explorer style and hide the Read only tag
        .Filter = "All files (*.*)|*.*"
        .MaxFileSize = 9999
        .ShowOpen
        vFiles = Split(.Filename, Chr(0)) 'Splits the filename up in segments
    If UBound(vFiles) = 0 Then ' If there is only 1 file then do this
    Open .Filename For Binary Access Read As #1
    Size = LOF(1)
    Close #1
    DoEvents
    DoEvents
    Set Item2 = ListView3.ListItems.Add(, , .FileTitle)
    Item2.SubItems(1) = .Filename
    DoEvents
    Item2.SubItems(2) = mFTP.GetFTPDirectory
    DoEvents
    Item2.SubItems(3) = frmmain.TxtConnectedTo.Text
    DoEvents
    Item2.SubItems(4) = Size
    y = Item2.SubItems(4)
    DoEvents
    DoEvents
    TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text + y
    Text1.Text = Text1.Text + y
    DoEvents
    DoEvents
    DoEvents
    Item2.SubItems(5) = "Upload"
    DoEvents
    Else
    For lFile = 1 To UBound(vFiles) ' More than 1 file then do this until there are no more files
    Open vFiles(0) + "\" & vFiles(lFile) For Binary Access Read As #1
    Size = LOF(1)
    Close #1
    DoEvents
    DoEvents
    Set Item2 = ListView3.ListItems.Add(, , vFiles(lFile))
    DoEvents
    Item2.SubItems(1) = vFiles(0) + "\" & vFiles(lFile)
    DoEvents
    Item2.SubItems(2) = mFTP.GetFTPDirectory
    DoEvents
    Item2.SubItems(3) = frmmain.TxtConnectedTo.Text
    DoEvents
    Item2.SubItems(4) = Size
    y = Item2.SubItems(4)
    DoEvents
    DoEvents
    TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text + y
    Text1.Text = Text1.Text + y
    DoEvents
    DoEvents
    DoEvents
    Item2.SubItems(5) = "Upload"
    DoEvents
    Next
    End If
    End With
End Sub

Public Sub mFTP_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
On Error Resume Next
Dim j As Long
Dim j3 As Long
TransferRate = Format(Int(lCurrentBytes / (Timer - BeginTransfer)) / 1024, "####.00")
    PB.Max = lTotalBytes
    PB.Min = 0
  j = PB.Value

  DoEvents
        PB.Value = lCurrentBytes
         DoEvents
        PB.ToolTipText = PB.Value & " Bytes of " & PB.Max & " Bytes Transfered"
        DoEvents
        Label11.Caption = PB.Value \ 1024 & " KB of  " & PB.Max \ 1024 & " KB Transfered"
        DoEvents
        Label4.Caption = Format$(CLng((j / PB.Max) * 100)) + "%"
        DoEvents
        Label8.Caption = "Speed: " & Format(TransferRate, "##.#0#") & " Kbps"
        DoEvents
        Label3.Caption = "Time Left: " & ConvertTime(Int(((PB.Max - PB.Value) / 1024) / TransferRate))
        DoEvents
        Label9.Caption = "Estimated Time: " & ConvertTime(Int(((Text1.Text) / 1024) / TransferRate))
        DoEvents
        If PB.Value = PB.Max Then
        Label4.Caption = "100%"
        End If

End Sub

Private Sub Form_Load()
   Set mFTP = New cFTP
End Sub
Public Function ConvertTime(ByVal TheTime As Single) As String
    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function
Private Sub ListView2_DblClick()
On Error Resume Next
Dim y As Long
y = ListView2.SelectedItem.SubItems(1)

DoEvents
    If Not (ListView2.SelectedItem Is Nothing) Then
         If ListView2.SelectedItem.Text = ".." Then
            mFTP.SetFTPDirectory ListView2.SelectedItem.Text
            TxtRemotePath.Text = mFTP.GetFTPDirectory
DoEvents
         Else
            If mFTP.Directory(ListView2.SelectedItem.Text).Directory Then
               mFTP.SetFTPDirectory ListView2.SelectedItem.Text
               TxtRemotePath.Text = mFTP.GetFTPDirectory
               DoEvents
            End If
DoEvents
            
            If Not mFTP.Directory(ListView2.SelectedItem.Text).Directory Then
            DoEvents
            Set Item2 = ListView3.ListItems.Add(, , ListView2.SelectedItem.Text)
            Item2.SubItems(1) = App.Path & "\Downloads"
            DoEvents
            Item2.SubItems(2) = mFTP.GetFTPDirectory
            DoEvents
            Item2.SubItems(3) = frmmain.TxtConnectedTo.Text
            DoEvents
            Item2.SubItems(4) = y
            TxtTotalBytesQueued.Text = TxtTotalBytesQueued.Text + y
            Text1.Text = Text1.Text + y
            DoEvents
            Item2.SubItems(5) = "Download"
            Exit Sub
            End If
            End If
            End If
            DoEvents
            RefreshDirectoryListing
DoEvents
End Sub
Public Function DepleteChr(Chars As String, Optional ReplaceChr As String) As String
    Dim ChrCnt As Long
    ReplaceChr = Left(ReplaceChr, 1)
    If ReplaceChr = "" Then ReplaceChr = " "


    Do
        ChrCnt = InStr(1, Chars, ReplaceChr)
        If ChrCnt = 0 Then Exit Do
        Chars = Left(Chars, ChrCnt - 1) & Right(Chars, Len(Chars) - ChrCnt)
    Loop
    DepleteChr = Chars
End Function

Private Sub menuUpdate_Click()
FrmUpdate.Show vbModal, Me
End Sub

Private Sub mFTP_StateChanged(State As FTP_CONNECTION_STATES)
    Select Case State
        Case FTP_CONNECTION_RESOLVING_HOST
            RTBHeader.SelText = "FTP_CONNECTION_RESOLVING_HOST" & vbNewLine
        Case FTP_CONNECTION_HOST_RESOLVED
            RTBHeader.SelText = "FTP_CONNECTION_HOST_RESOLVED" & vbNewLine
        Case FTP_CONNECTION_CONNECTED
            RTBHeader.SelText = "FTP_CONNECTION_CONNECTED" & vbNewLine
        Case FTP_CONNECTION_AUTHENTICATION
            RTBHeader.SelText = "FTP_CONNECTION_AUTHENTICATION" & vbNewLine
        Case FTP_USER_LOGGED
            RTBHeader.SelText = "FTP_USER_LOGGED" & vbNewLine
        Case FTP_ESTABLISHING_DATA_CONNECTION
            RTBHeader.SelText = "FTP_ESTABLISHING_DATA_CONNECTION" & vbNewLine
        Case FTP_DATA_CONNECTION_ESTABLISHED
            RTBHeader.SelText = "FTP_DATA_CONNECTION_ESTABLISHED" & vbNewLine
        Case FTP_RETRIEVING_DIRECTORY_INFO
            RTBHeader.SelText = "FTP_RETRIEVING_DIRECTORY_INFO" & vbNewLine
        Case FTP_DIRECTORY_INFO_COMPLETED
            RTBHeader.SelText = "FTP_DIRECTORY_INFO_COMPLETED" & vbNewLine
        Case FTP_TRANSFER_STARTING
            RTBHeader.SelText = "FTP_TRANSFER_STARTING" & vbNewLine
        Case FTP_TRANSFER_COMLETED
            RTBHeader.SelText = "FTP_TRANSFER_COMLETED" & vbNewLine
    End Select
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "Connect"
        Call menuconnect_Click
        Case "Disconnect"
        Call menuDisconnect_Click
        Case "Stop"
        Call menustop_Click
        Case "Refresh"
            RefreshDirectoryListing
        Case "Transfer"
        Call menutransfer_Click
        Case "Download"
        Call menudownload_Click
        Case "Upload"
        Call menuupload_Click
        Case "CreateDirectory"
          Call menuMD_Click
        Case "Delete"
            menuFtpdelete_Click
        Case "Rename"
           Call menuRename_Click
        Case "View Large Icons"
            ListView2.View = lvwIcon
        Case "View Small Icons"
            ListView2.View = lvwSmallIcon
        Case "View List"
            ListView2.View = lvwList
        Case "View Details"
            ListView2.View = lvwReport
    End Select
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "Total Files Queued - " & ListView3.ListItems.Count
Label5.Caption = "Total Size - " & Format(TxtTotalBytesQueued.Text / 1024, "0.00") & " KB"
End Sub

Private Sub Timer2_Timer()
        If PB.Value = PB.Max Then
            Label10.Caption = "Elapsed Time: " & Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
        Else
            Sec = Sec + 1
            If Sec >= 60 Then
                Sec = 0
                Min = Min + 1
            ElseIf Min >= 60 Then
                Min = 0
                Hr = Hr + 1
            End If
           Label10.Caption = "Elapsed Time: " & Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")

        End If

End Sub
