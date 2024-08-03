VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmMain 
   Caption         =   "CCAT"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSDBGrid.DBGrid grdData 
      Bindings        =   "frmMain.frx":030A
      Height          =   1845
      Left            =   7440
      OleObjectBlob   =   "frmMain.frx":031E
      TabIndex        =   9
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6075
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4485
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar barLoad 
      Height          =   300
      Left            =   15
      TabIndex        =   8
      ToolTipText     =   "% Operation Complete"
      Top             =   5325
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   6360
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CF1
            Key             =   "Session"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1245
            Key             =   "DB_CLOSED"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1359
            Key             =   "DB_OPEN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":146D
            Key             =   "HD_CLOSED"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19C1
            Key             =   "HD_OPEN"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F15
            Key             =   "CD_CLOSED"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2149
            Key             =   "CD_OPEN"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":237D
            Key             =   "TAPE_CLOSED"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2491
            Key             =   "TAPE_OPEN"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25A5
            Key             =   "MSG_CLOSED"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26B9
            Key             =   "MSG_OPEN"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27CD
            Key             =   "SIG"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C21
            Key             =   "LOB"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3075
            Key             =   "GEO"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34C9
            Key             =   "EVT"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":391D
            Key             =   "Query"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E61
            Key             =   "Data"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":417B
            Key             =   "ClosedBook"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45CD
            Key             =   "OpenBook"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A1F
            Key             =   "DB3_CLOSED"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B31
            Key             =   "DB3_OPEN"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C43
            Key             =   "DB4_CLOSED"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D55
            Key             =   "DB4_OPEN"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLargeIcons 
      Left            =   5640
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E67
            Key             =   "Session"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52BB
            Key             =   "DB"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":570F
            Key             =   "HD"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B63
            Key             =   "CD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5FB7
            Key             =   "MSG"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":640B
            Key             =   "TAPE"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":669F
            Key             =   "QUERY"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   7
      Top             =   690
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10560
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   10560
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ListView:"
         Height          =   270
         Index           =   1
         Left            =   2078
         TabIndex        =   4
         Tag             =   " ListView:"
         Top             =   12
         Width           =   3216
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TreeView:"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Tag             =   " TreeView:"
         Top             =   12
         Width           =   2016
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   1111
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "Create a new database"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            Object.ToolTipText     =   "Open an existing database"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            Key             =   "Remove"
            Object.ToolTipText     =   "Remove a database from the session"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete a database file"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prop"
            Key             =   "Properties"
            Object.ToolTipText     =   "Edit/View properties"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Arch"
            Key             =   "Archive"
            Object.ToolTipText     =   "Add an archive"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Filter"
            Key             =   "Filter"
            Object.ToolTipText     =   "Display the filter options"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save data to file"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Large"
            Key             =   "View Large Icons"
            Object.ToolTipText     =   "View Large Icons"
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Small"
            Key             =   "View Small Icons"
            Object.ToolTipText     =   "View Small Icons"
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List"
            Key             =   "View List"
            Object.ToolTipText     =   "View List"
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Detail"
            Key             =   "View Details"
            Object.ToolTipText     =   "View Details"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tutor"
            Key             =   "Tutorial"
            Object.ToolTipText     =   "Run tutorials"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5295
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13344
            MinWidth        =   882
            Text            =   "Status"
            TextSave        =   "Status"
            Object.ToolTipText     =   "Current CCAT Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            Key             =   "SECURITY"
            Object.ToolTipText     =   "Current Classification Level"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2170
            MinWidth        =   882
            TextSave        =   "11/14/2006"
            Object.ToolTipText     =   "Today's Date"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1640
            MinWidth        =   882
            TextSave        =   "1:57 PM"
            Object.ToolTipText     =   "Current Time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5520
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6120
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D1B
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E2D
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F3F
            Key             =   "Remove"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7053
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7165
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7277
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7389
            Key             =   "Archive"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":749D
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":78F1
            Key             =   "View Large Icons"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A03
            Key             =   "View Small Icons"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B15
            Key             =   "View List"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C27
            Key             =   "View Details"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D39
            Key             =   "Tutorial"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7E4D
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2052
      TabIndex        =   5
      Top             =   705
      Width           =   3216
      _ExtentX        =   5662
      _ExtentY        =   8467
      Arrange         =   2
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4800
      HelpContextID   =   243
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   8467
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      PathSeparator   =   "->"
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin VB.Image imgSplitter 
      Height          =   4788
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileRemove 
         Caption         =   "&Remove"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "&Add"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditProperties 
         Caption         =   "&Properties"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEditFilter 
         Caption         =   "&Filter"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Lar&ge Icons"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "S&mall Icons"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&Details"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Arrange &Icons"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsTime 
         Caption         =   "Convert &Time Values"
      End
      Begin VB.Menu mnuToolsDeg 
         Caption         =   "Convert &Degree Values"
      End
      Begin VB.Menu mnuToolsSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsViewINI 
         Caption         =   "&View CCAT.INI"
      End
      Begin VB.Menu mnuToolsRemapINI 
         Caption         =   "&Re-map INI file"
      End
      Begin VB.Menu mnuToolsSaveQuery 
         Caption         =   "Save Custom &Query"
      End
      Begin VB.Menu mnuToolsSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsExecuteSQL 
         Caption         =   "Execute &SQL Command"
      End
      Begin VB.Menu mnuToolsUpdateDB 
         Caption         =   "&Upgrade Database"
      End
      Begin VB.Menu mnuToolsEOB 
         Caption         =   "Enter &EOB Table"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuToolsSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsCreate 
         Caption         =   "Create Default VarStruct"
      End
      Begin VB.Menu mnuToolsImport 
         Caption         =   "Import Default VarStruct"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpTutorials 
         Caption         =   "&Tutorials"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Popup"
      Begin VB.Menu mnuPopNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuPopOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuPopCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuPopDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuPopProperties 
         Caption         =   "&Properties"
      End
      Begin VB.Menu mnuPopAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuPopSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuPopBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopCopyCan 
         Caption         =   "Canned Reports"
      End
      Begin VB.Menu mnuPopCopyVS 
         Caption         =   "Copy VarStruct"
      End
      Begin VB.Menu mnuPopPasteVS 
         Caption         =   "Paste VarStruct"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopMsg 
         Caption         =   "Process Msg"
      End
      Begin VB.Menu mnuPopTemplate 
         Caption         =   "Msg Template"
      End
      Begin VB.Menu mnuPopupHelp 
         Caption         =   "&What's This?"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuData 
      Caption         =   "Data"
      Begin VB.Menu mnuDataFilter 
         Caption         =   "Filter"
      End
      Begin VB.Menu mnuDataLastQuery 
         Caption         =   "Execute last query"
      End
      Begin VB.Menu mnuDataHideCol 
         Caption         =   "Hide Column"
      End
      Begin VB.Menu mnuDataShowAllCol 
         Caption         =   "Show All Columns"
      End
      Begin VB.Menu mnuDataSortCol 
         Caption         =   "Sort by Column"
      End
      Begin VB.Menu mnuDataShowValue 
         Caption         =   "Show related to value"
         Begin VB.Menu mnuDataEQ 
            Caption         =   "field = value"
         End
         Begin VB.Menu mnuDataNE 
            Caption         =   "field <> value"
         End
         Begin VB.Menu mnuDataLT 
            Caption         =   "field < value"
         End
         Begin VB.Menu mnuDataLE 
            Caption         =   "field <= value"
         End
         Begin VB.Menu mnuDataGT 
            Caption         =   "field > value"
         End
         Begin VB.Menu mnuDataGE 
            Caption         =   "field >= value"
         End
      End
      Begin VB.Menu mnuDataShowAllRows 
         Caption         =   "Show All Rows"
      End
      Begin VB.Menu mnuDataExport 
         Caption         =   "Export Data"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' FORM:     frmMain
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Main user interface and routines for manipulating the interface controls
'           or responding to user interaction
' REVISIONS:
'   v1.3.0  TAE Replaced Token controls with INI function calls
'   v1.4.0  TAE Added images for the Stored Query tree node
'           TAE Fixed bug - now right-click menu uses real field names, not aliases
'           TAE Added menu items for time/angle conversions to the "Help" menu
'   v1.5.0  TAE Fixed bug - right-click would sometimes cause error 9 and crash.
'               Problem was field parsing if all columns were shown.
'           TAE Changed time value menu to handle date/time values
'           TAE Added code to handle the case where the first column was hidden
'           TAE Added mouse pointer change to indicate busy state when selecting message node
'           TAE Added version info to window title
'           TAE Added context-sensitive help information to each menu and control
'           TAE Added code to trap the F1 key in the Grid, Tree View, and List View to bring
'               up context-sensitive help.
'           TAE Added code to force full screen window
'           TAE Added Tools menu and several tool menu items
'           TAE Moved the status panel text update to a separate method that can be called externally
'           TAE Moved the progress bar setup to a separate method that can be called externally
'           TAE Added code to save the current database's version in the structure
'           TAE Added code to enable some right-click menu items for query nodes
'           TAE Added code to display a query node's SQL statement (Properties menu)
'           TAE Added code to launch the Save Query function (Add menu)
'           TAE Added code to clear the current query and launch the Filter form (New menu)
'           TAE Added code to launch Keith's ConfigEditor to view the INI file
'           TAE Modified data popup menu to hide and re-configure menu items based on query type and content
'               Modified data popup menu to requery the database if the recordset is invalid
'               Modified data popup menu to create additional value-related menu items
'           TAE Modified Hide COlumn menu item to handle aliased columns
'           TAE Added menu items and code to handle multiple value-related filters
'           TAE Disabled the mnuDataShowValue menu handler because of the new sub-menu
'           TAE Modified the "Sort by Column" menu handler to re-order the sort list to put the
'               newly selected field first, and remove it from the list if it appears more than
'               once.
'           TAE Added code to requery the database if the export type is a DAS file
'   v1.6.0  TAE Changed the tree view minimum sizing limit to 1/4 of the window width
'           TAE Changed old Archive options references to the new Archive Wizard
'           TAE Added refresh display to delete actions to prevent invalid data on the screen
'           TAE Added code to prevent the Grid view from being connected to data that was being deleted or modified
'           TAE Changed archive form properties to use a new Archive Properties form
'           TAE Modified the code that handles clicking on tree view nodes
'           TAE Added storage of the currently selected archive node to the global Current structure
'           TAE Added code to automatically size the columns of the list view to the contents
'           TAE Updated the event file field list
'           TAE Modified resizing to occur only when the window is visible
'   v1.6.1  TAE Added verbose logging calls
'
Option Explicit     ' Forces variables to be declared before they can be used
'
' Interface to the Help System
'Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
'
' Variables and constants
Dim dbPrevious As String
Dim mblnIsMoving As Boolean     ' True while the user is moving the splitter bar
Dim cpyBufFull As Boolean       ' True if the varStruct table has been copied
Const mintSPLIT_LIMIT = 1500    ' Minimum width for the tree view or list view windows
'
' EVENT:    Form_Load
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Configure user interface controls and initialize the program
' TRIGGER:  When the program is executed, frmMain is activated.  This is the first
'           routine that is run.
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    '+v1.5
    ' Set the form caption to show the version
    frmMain.Caption = "CCAT - " & basCCAT.sGet_Version
    '-v1.5
    '
    ' Open the log file
    giLog_File = FreeFile
    Open App.Path & "\" & App.EXEName & ".log" For Output As giLog_File
    '
    '+v1.6.1TE
    Close giLog_File
    '-v1.6.1
    '
    ' Write the log file header
    basCCAT.WriteLogEntry "EVENT    : frmMain Load (Start)"
    '
    ' Display the splash screen
    frmSplash.Show
    frmSplash.Refresh
    '
    ' Initialize the security routines
    frmSecurity.InitializeSecurity
    '
    ' Initialize the translator routines
    basCCAT.Initialize_Translator
    '
    ' Assign the help file
    App.HelpFile = App.Path & DAS_HELP_PATH & CCAT_HELP_FILE
    '
    ' Configure the grid control
    frmMain.grdData.Visible = False
    frmMain.grdData.RecordSelectors = False
    '
    ' Using current screen geometry and resolution settings, compute the correct height
    ' for the status bar to minimize distortion of the security banners
    frmMain.sbStatusBar.Height = (Screen.TwipsPerPixelY * frmSecurity.imlBanners.ImageHeight) + 30
    '
    ' Position and size the form based on the settings saved from the last session.
    ' Settings are saved in the system registry.  The registry key is the application
    ' name (App.Title) and the category is "Settings".  The default values follow the
    ' parameter names.
    '+v1.5
    'frmMain.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    'frmMain.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    'frmMain.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    'frmMain.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    frmMain.imgSplitter.Left = GetSetting(App.Title, "Settings", "Splitter", frmMain.Width / 3)
    '
    ' Force window to full screen
    frmMain.WindowState = vbMaximized
    '-v1.5
    '
    ' Load the appropriate security banner into the status bar panel
    ' The appropriate security banner is determined by looking up a token based on the
    ' current classification level.  The returned alias is the image tag from the
    ' security image list.
    frmMain.sbStatusBar.Panels("SECURITY").Picture = frmSecurity.imlBanners.ListImages(frmSecurity.strGetAlias("Images", "SECURITYPIC" & frmSecurity.Tag, "UNCLASSIFIED")).Picture
    '
    ' Assign image lists to interface controls.
    ' By assigning image lists at run time, the developer can add/remove images from the
    ' image lists at design time.
    frmMain.tvTreeView.ImageList = frmMain.imlSmallIcons
    frmMain.lvListView.Icons = frmMain.imlLargeIcons
    frmMain.lvListView.SmallIcons = frmMain.imlSmallIcons
    frmMain.tbToolBar.ImageList = frmMain.imlToolbarIcons
    '
    ' Hide the progress bar and data grid
    frmMain.barLoad.Visible = False
    frmMain.grdData.Visible = False
    '
    ' Hide the popup menus
    frmMain.mnuPop.Visible = False
    frmMain.mnuData.Visible = False
    frmMain.mnuHelpTutorials.Visible = False
    '
    ' Remove the text captions from the toolbar buttons and assign icons
    frmMain.tbToolBar.Buttons("New").Caption = ""
    frmMain.tbToolBar.Buttons("New").Image = "New"
    frmMain.tbToolBar.Buttons("Open").Caption = ""
    frmMain.tbToolBar.Buttons("Open").Image = "Open"
    frmMain.tbToolBar.Buttons("Remove").Caption = ""
    frmMain.tbToolBar.Buttons("Remove").Image = "Remove"
    frmMain.tbToolBar.Buttons("Delete").Caption = ""
    frmMain.tbToolBar.Buttons("Delete").Image = "Delete"
    frmMain.tbToolBar.Buttons("Properties").Caption = ""
    frmMain.tbToolBar.Buttons("Properties").Image = "Properties"
    frmMain.tbToolBar.Buttons("Archive").Caption = ""
    frmMain.tbToolBar.Buttons("Archive").Image = "Archive"
    frmMain.tbToolBar.Buttons("Save").Caption = ""
    frmMain.tbToolBar.Buttons("Save").Image = "Save"
    frmMain.tbToolBar.Buttons("Filter").Caption = ""
    frmMain.tbToolBar.Buttons("Filter").Image = "Filter"
    frmMain.tbToolBar.Buttons("View Large Icons").Caption = ""
    frmMain.tbToolBar.Buttons("View Large Icons").Image = "View Large Icons"
    frmMain.tbToolBar.Buttons("View Small Icons").Caption = ""
    frmMain.tbToolBar.Buttons("View Small Icons").Image = "View Small Icons"
    frmMain.tbToolBar.Buttons("View List").Caption = ""
    frmMain.tbToolBar.Buttons("View List").Image = "View List"
    frmMain.tbToolBar.Buttons("View Details").Caption = ""
    frmMain.tbToolBar.Buttons("View Details").Image = "View Details"
    frmMain.tbToolBar.Buttons("Tutorial").Caption = ""
    frmMain.tbToolBar.Buttons("Tutorial").Image = "Tutorial"
    frmMain.tbToolBar.Buttons("Tutorial").Visible = False
    frmMain.tbToolBar.Buttons("Help").Caption = ""
    frmMain.tbToolBar.Buttons("Help").Image = "Help"
    '
    ' Create the root session node
    basCCAT.Create_New_Session
    '
    ' Read the session file
    basCCAT.Read_Session_File
    '
    ' Configure the interface for session mode by triggering the Node Click event
    frmMain.ChangeMode gsSESSION
    '
    '+v1.5
    ' Set help contexts
    frmMain.HelpContextID = basCCAT.IDH_GUI_MAIN
    frmMain.grdData.HelpContextID = basCCAT.IDH_GUI_GRID
    frmMain.lvListView.HelpContextID = basCCAT.IDH_GUI_LISTVIEW
    frmMain.mnuData.HelpContextID = basCCAT.IDH_GUI_DATA_MENU
    frmMain.mnuDataExport.HelpContextID = basCCAT.IDH_GUI_FILE_SAVE
    frmMain.mnuDataFilter.HelpContextID = basCCAT.IDH_GUI_FILTER
    frmMain.mnuDataHideCol.HelpContextID = basCCAT.IDH_GUI_DATA_HIDE
    frmMain.mnuDataShowAllCol.HelpContextID = basCCAT.IDH_GUI_DATA_COL
    frmMain.mnuDataShowAllRows.HelpContextID = basCCAT.IDH_GUI_DATA_ROW
    frmMain.mnuDataShowValue.HelpContextID = basCCAT.IDH_GUI_DATA_VALUE
    frmMain.mnuDataSortCol.HelpContextID = basCCAT.IDH_GUI_DATA_SORT
    frmMain.mnuEdit.HelpContextID = basCCAT.IDH_GUI_EDIT
    frmMain.mnuEditAdd.HelpContextID = basCCAT.IDH_GUI_EDIT_ADD
    frmMain.mnuEditDelete.HelpContextID = basCCAT.IDH_GUI_FILE_DELETE
    frmMain.mnuEditFilter.HelpContextID = basCCAT.IDH_GUI_EDIT_FILTER
    frmMain.mnuEditProperties.HelpContextID = basCCAT.IDH_GUI_EDIT_PROPERTIES
    frmMain.mnuFile.HelpContextID = basCCAT.IDH_GUI_FILE
    frmMain.mnuFileClose.HelpContextID = basCCAT.IDH_GUI_CLOSE
    frmMain.mnuFileDelete.HelpContextID = basCCAT.IDH_GUI_FILE_DELETE
    frmMain.mnuFileNew.HelpContextID = basCCAT.IDH_GUI_FILE_NEW
    frmMain.mnuFileOpen.HelpContextID = basCCAT.IDH_GUI_FILE_OPEN
    frmMain.mnuFileRemove.HelpContextID = basCCAT.IDH_GUI_FILE_REMOVE
    frmMain.mnuFileSave.HelpContextID = basCCAT.IDH_GUI_FILE_SAVE
    frmMain.mnuHelp.HelpContextID = basCCAT.IDH_GUI_HELP
    frmMain.mnuHelpAbout.HelpContextID = basCCAT.IDH_GUI_ABOUT
    frmMain.mnuHelpContents.HelpContextID = basCCAT.IDH_GUI_HELP_CONTENTS
    frmMain.mnuTools.HelpContextID = basCCAT.IDH_GUI_TOOLS
    frmMain.mnuToolsDeg.HelpContextID = basCCAT.IDH_GUI_TOOLS_DEGREE
    frmMain.mnuHelpSearchForHelpOn.HelpContextID = basCCAT.IDH_GUI_MAIN
    frmMain.mnuToolsTime.HelpContextID = basCCAT.IDH_GUI_TOOLS_TIME
    frmMain.mnuListViewMode(0).HelpContextID = basCCAT.IDH_GUI_VIEW_LARGE
    frmMain.mnuListViewMode(1).HelpContextID = basCCAT.IDH_GUI_VIEW_SMALL
    frmMain.mnuListViewMode(2).HelpContextID = basCCAT.IDH_GUI_VIEW_LIST
    frmMain.mnuListViewMode(3).HelpContextID = basCCAT.IDH_GUI_VIEW_DETAILS
    frmMain.mnuPopAdd.HelpContextID = basCCAT.IDH_GUI_EDIT_ADD
    frmMain.mnuView.HelpContextID = basCCAT.IDH_GUI_VIEW
    frmMain.mnuPopCut.HelpContextID = basCCAT.IDH_GUI_FILE_REMOVE
    frmMain.mnuPopDelete.HelpContextID = basCCAT.IDH_GUI_FILE_DELETE
    frmMain.mnuPopNew.HelpContextID = basCCAT.IDH_GUI_FILE_NEW
    frmMain.mnuPopOpen.HelpContextID = basCCAT.IDH_GUI_FILE_OPEN
    frmMain.mnuPopProperties.HelpContextID = basCCAT.IDH_GUI_EDIT_PROPERTIES
    frmMain.mnuPopSave.HelpContextID = basCCAT.IDH_GUI_FILE_SAVE
    frmMain.mnuViewArrangeIcons.HelpContextID = basCCAT.IDH_GUI_VIEW_ARRANGE
    frmMain.mnuViewRefresh.HelpContextID = basCCAT.IDH_GUI_VIEW_REFRESH
    frmMain.mnuViewStatusBar.HelpContextID = basCCAT.IDH_GUI_VIEW_STATUS
    frmMain.mnuViewToolbar.HelpContextID = basCCAT.IDH_GUI_VIEW_TOOLBAR
    frmMain.tvTreeView.HelpContextID = basCCAT.IDH_GUI_TREE
    frmMain.mnuToolsExecuteSQL.HelpContextID = basCCAT.IDH_GUI_TOOLS_SQL
    frmMain.mnuToolsRemapINI.HelpContextID = basCCAT.IDH_GUI_TOOLS_REMAP
    frmMain.mnuToolsSaveQuery.HelpContextID = basCCAT.IDH_GUI_TOOLS_SAVE
    frmMain.mnuToolsUpdateDB.HelpContextID = basCCAT.IDH_GUI_TOOLS_UPDATE
    frmMain.mnuToolsViewINI.HelpContextID = basCCAT.IDH_GUI_TOOLS_INI
    '-v1.5
    '
    ' Remove the splash screen
    Unload frmSplash
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain Load (End)"
    '-v1.6.1
    '
    '+v1.7SV
    cpyBufFull = False
    '-v1.7SV
End Sub
'
' EVENT:    Form_Paint
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Updates the user interface
' TRIGGER:  When the main form is exposed after being hidden behind another form
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Paint()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain Paint (Start)"
    '-v1.6.1
    '
    ' Get the listview mode from the registry.
    frmMain.lvListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    '
    ' Configure the toolbar and menu based on the mode
    frmMain.mnuListViewMode(frmMain.lvListView.View).Checked = True
    Select Case lvListView.View
        '
        ' View large icons
        Case lvwIcon
            frmMain.tbToolBar.Buttons("View Large Icons").Value = tbrPressed
        '
        ' View small icons
        Case lvwSmallIcon
            frmMain.tbToolBar.Buttons("View Small Icons").Value = tbrPressed
        '
        ' View List
        Case lvwList
            frmMain.tbToolBar.Buttons("View List").Value = tbrPressed
        '
        ' View Details
        Case lvwReport
            frmMain.tbToolBar.Buttons("View Details").Value = tbrPressed
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain Paint (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Unload
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Housekeeping chores
' TRIGGER:  User exits the program
' INPUT:    None
' OUTPUT:   If "intCancel" is 0, then the form should be unloaded, if "intCancel" is
'           any other value, the form is NOT unloaded.
' NOTES:
Private Sub Form_Unload(intCancel As Integer)
    Dim pintForm As Integer ' Current form
    Dim pintFile As Integer ' Session file
    Dim pnodDB As Node      ' Database node
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmMain Unload (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intCancel
    End If
    '-v1.6.1
    '
    ' Close all sub forms
    On Error Resume Next
    For pintForm = Forms.Count - 1 To 1 Step -1
        Unload Forms(pintForm)
    Next
    On Error GoTo 0
    '
    ' Save the current position and size settings to the system registry
    If frmMain.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", frmMain.Left
        SaveSetting App.Title, "Settings", "MainTop", frmMain.Top
        SaveSetting App.Title, "Settings", "MainWidth", frmMain.Width
        SaveSetting App.Title, "Settings", "MainHeight", frmMain.Height
        SaveSetting App.Title, "Settings", "Splitter", frmMain.imgSplitter.Left
    End If
    '
    ' Save the current list view display mode
    SaveSetting App.Title, "Settings", "ViewMode", frmMain.lvListView.View
    '
    ' Note log file
    basCCAT.WriteLogEntry "          MAIN: Window settings saved in system registry"
    '
    ' Save session
    basCCAT.WriteLogEntry "          MAIN: Saving session information"
    '
    ' Find an available file ID
    pintFile = FreeFile
    '
    ' Suppress error reporting
    On Error Resume Next
    '
    ' Open the session file
    Open App.Path & "\" & App.EXEName & ".ses" For Output As pintFile
    '
    '
    If Err.Number = NO_ERROR Then
        '
        ' Restore error reporting
        On Error GoTo 0
        '
        ' See if there is a session node
        If frmMain.blnNodeExists(gsSESSION) Then Set pnodDB = frmMain.tvTreeView.Nodes(gsSESSION).Child
        '
        ' Go through all database nodes until the end
        While Not pnodDB Is Nothing
            '
            ' Write the database file name to the session file and the log file
            Print #pintFile, pnodDB.Key
            basCCAT.WriteLogEntry "              Session file: " & pnodDB.Key
            '
            ' Move to the next database node
            Set pnodDB = pnodDB.Next
        Wend
        '
        ' Close the session file
        Close pintFile
        '
        ' Close out the log file
        basCCAT.WriteLogEntry "MAIN: Log file ended"
        Close giLog_File
    Else
        '
        ' Close out the log file
        basCCAT.WriteLogEntry "MAIN: Log file ended"
        Close
        '
        ' Delete the session file
        Kill App.Path & "\" & App.EXEName & ".ses"
    End If
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    End
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain Unload (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Resize
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Adjust control location and size
' TRIGGER:  User resizes the form
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Resize()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain Resize (Start)"
    '-v1.6.1
    '
    ' Repress error reporting
    On Error Resume Next
    '
    ' Check for a minimum width
    If frmMain.Width < 3000 Then frmMain.Width = 3000
    '
    ' Adjust the controls to the new size
    frmMain.SizeControls imgSplitter.Left
    '
    ' Re-enable errors
    On Error GoTo 0
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain Resize (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    grdData_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Changes to data mode
' TRIGGER:  User clicks anywhere in the data grid
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub grdData_GotFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.grdData GotFocus (Start)"
    '-v1.6.1
    '
    frmMain.ChangeMode gsDATA
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.grdData GotFocus (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' EVENT:    grdData_KeyUp
' AUTHOR:   Tom Elkins
' PURPOSE:  Monitors keystrokes when the grid has the focus
' TRIGGER:  User presses a key anywhere in the data grid
' INPUT:    "intKey_Code" is the system code for the key that was pressed
'           "intShift" is a code to indicate if the intShift/Ctrl/Alt key(s) were pressed as well
' OUTPUT:   None
' NOTES:
Private Sub grdData_KeyUp(intKey_Code As Integer, intShift As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.grdData KeyUp (Start)"
    '-v1.6.1
    '
    Select Case intKey_Code
        '
        ' Trap the F1 key
        Case vbKeyF1:
            '
            ' Launch the help file at the Grid section
            basCCAT.HtmlHelp frmMain.hwnd, App.HelpFile, basCCAT.HH_HELP_CONTEXT, basCCAT.IDH_GUI_GRID
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.grdData KeyUp (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
' EVENT:    grdData_MouseUp
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure and display the popup menu
' TRIGGER:  User released a mouse button
' INPUT:    "intButton" indicates which mouse button was clicked
'           "intShift" indicates if the Shift key was pressed
'           "sngMouse_X" and "sngMouse_Y" are the mouse location
' OUTPUT:   None
' NOTES:
Private Sub grdData_MouseUp(intButton As Integer, intShift As Integer, sngMouse_X As Single, sngMouse_Y As Single)
    Dim plngCol As Long         ' current column
    Dim plngRow As Long         ' current row
    Dim pstrField As String     ' Actual field name
    Dim pastrFields() As String 'v1.5
    Dim pstrValue As String     'v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.grdData MouseUp (Start)"
    '-v1.6.1
    '
    '+v1.5
    ' See if the grid is valid.  If not, requery the database
    If frmMain.Data1.Recordset Is Nothing Then basDatabase.QueryData
    '-v1.5
    '
    ' Save the mouse position
    guGUI.fMouse_X = sngMouse_X
    guGUI.fMouse_Y = sngMouse_Y
    '
    ' Check for the right mouse button
    If intButton = vbRightButton Then
        '
        '+v1.6.1TE
        If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmMain.grdData MouseUp (Right mouse button clicked)"
        '-v1.6.1
        '
        '+v1.5
        ' If there is a last query, show the menu option
        frmMain.mnuDataLastQuery.Visible = (basDatabase.LastQuery <> "")
        '-v1.5
        '
        ' Get the current row and column in the Grid
        plngCol = frmMain.grdData.ColContaining(sngMouse_X)
        plngRow = frmMain.grdData.RowContaining(sngMouse_Y)
        '
        '+v1.6.1TE
        If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmMain.grdData MouseUp (Cell selected @ " & plngCol & ", " & plngRow & ")"
        '-v1.6.1
        '
        '
        ' If there is a filter applied, enable the "Show All Rows" menu option
        If Len(guCurrent.uSQL.sFilter) > 0 Then
            frmMain.mnuDataShowAllRows.Visible = True
        Else
            frmMain.mnuDataShowAllRows.Visible = False
        End If
        '
        '+v1.5
        ' If there is an aggregate function, hide the "Show All Columns" menu option
        frmMain.mnuDataShowAllCol.Visible = (InStr(1, UCase(guCurrent.uSQL.sQuery), "GROUP BY") = 0)
        frmMain.mnuDataShowAllRows.Visible = frmMain.mnuDataShowAllCol.Visible
        '-v1.5
        '
        ' Check that the user was in a column
        If plngCol >= 0 Then
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmMain.grdData MouseUp (Valid column)"
            '-v1.6.1
            '
            ' Extract the actual field name, ignoring the alias
            '+v1.5
            ' Old code (in the Else clause) was causing an error 9 - subscript out of range
            ' The problem was if the sFields member contained a "*" -- the Split command
            ' returned a 0-element array.  The new code checks for the "*" and uses the
            ' grid column for the field name -- if the "*" is used, all fields are displayed
            ' with no aliases; therefore, the column caption contains the actual field name
            If guCurrent.uSQL.sFields = "*" Then
                pstrField = frmMain.grdData.Columns(plngCol).Caption
            Else
                '+v1.5
                'pstrField = Split(Trim(Split(guCurrent.uSQL.sFields, ",")(plngCol)), " ")(0)
                pstrField = frmMain.Data1.Recordset.Fields(plngCol).Name
                '-v1.5
            End If
            '-v1.5
            '
            ' Enable the menu option to sort by values in the current column
            frmMain.mnuDataSortCol.Caption = "Sort by " & pstrField
            frmMain.mnuDataSortCol.Visible = True
            '
            '+v1.5
            ' Enable the menu option to hide the current column
            frmMain.mnuDataHideCol.Caption = "Hide Column '" & pstrField & "'"
            frmMain.mnuDataHideCol.Visible = True
            '-v1.5
        Else
            '
            ' Hide the sort menu option
            frmMain.mnuDataSortCol.Visible = False
            '
            '+v1.5
            ' Hide the hide column menu option
            frmMain.mnuDataHideCol.Visible = False
            '-v1.5
        End If
        '
        ' Check that the user is in a valid row and column
        If plngRow >= 0 And plngCol >= 0 Then
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmMain.grdData MouseUp (Valid row)"
            '-v1.6.1
            '
            '+v1.5
            ' Configure the menu item that allows the user to grab records with a
            ' particular value
            Select Case (frmMain.Data1.Recordset.Fields(frmMain.grdData.Columns(plngCol).DataField).Type)
                '
                ' Format for a text value
                Case dbText:
                    pstrValue = "'" & frmMain.grdData.Columns(plngCol).CellValue(frmMain.grdData.RowBookmark(plngRow)) & "'"
                '
                ' Time value
                Case dbDate:
                    '+v1.5
                    'frmMain.mnuDataShowValue.Caption = pstrField & " = TIMEVALUE(""" & frmMain.grdData.Columns(plngCol).CellValue(frmMain.grdData.RowBookmark(plngRow)) & """)"
                    ' Change syntax to handle date/time values
                    pstrValue = "#" & frmMain.grdData.Columns(plngCol).CellValue(frmMain.grdData.RowBookmark(plngRow)) & "#"
                    '-v1.5
                Case Else:
                    '
                    ' Format for a numeric value
                    pstrValue = frmMain.grdData.Columns(plngCol).CellValue(frmMain.grdData.RowBookmark(plngRow))
            End Select
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmMain.grdData MouseUp (Configure Menus)"
            '-v1.6.1
            '
            '
            ' Construct the relational menus
            frmMain.mnuDataEQ.Caption = pstrField & " = " & pstrValue
            frmMain.mnuDataGE.Caption = pstrField & " >= " & pstrValue
            frmMain.mnuDataGT.Caption = pstrField & " > " & pstrValue
            frmMain.mnuDataLE.Caption = pstrField & " <= " & pstrValue
            frmMain.mnuDataLT.Caption = pstrField & " < " & pstrValue
            frmMain.mnuDataNE.Caption = pstrField & " <> " & pstrValue
            '
            ' Display the menu option
            'frmMain.mnuDataShowValue.Visible = True
            frmMain.mnuDataShowValue.Visible = frmMain.mnuDataShowAllCol.Visible
            '-v1.5
        Else
            '
            ' Hide the menu option
            frmMain.mnuDataShowValue.Visible = False
        End If
        '
        ' If there is only one column, hide the "Hide Column" menu
        frmMain.mnuDataHideCol.Visible = (frmMain.grdData.Columns.Count > 1)
        '
        ' Display the popup menu
        frmMain.PopupMenu frmMain.mnuData
        '
        '+v1.5
        ' Show the query
        frmMain.lblTitle(1).Caption = basDatabase.CurrentQuery
        '-v1.5
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.grdData MouseUp (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    grdData_RowResize
' AUTHOR:   Tom Elkins
' PURPOSE:  Prevent the rows in the grid from changing size
' TRIGGER:  User clicked between rows in the data grid
' INPUT:    None
' OUTPUT:   "intCancel" -- if set to true, cancels the resize request
' NOTES:    If the user clicks between rows, the rows resize themselves.
'           This event traps that and prevents the rows from changing size.
Private Sub grdData_RowResize(intCancel As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.grdData RowResize (Start)"
    '-v1.6.1
    '
    intCancel = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.grdData RowResize (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    imgSplitter_MouseDown
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Aids in resizing the treeview and listview windows
' TRIGGER:  User holds the mouse button down on the splitter bar
' INPUT:    "intButton" is which mouse button is being pressed
'           "intShift" determines if the intShift key is being pressed
'           "sngMouse_X" and "sngMouse_Y" are the current mouse position
' OUTPUT:   None
' NOTES:
Private Sub imgSplitter_MouseDown(intButton As Integer, intShift As Integer, sngMouse_X As Single, sngMouse_Y As Single)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.imgSplitter MouseDown (Start)"
    '-v1.6.1
    '
    ' Use control-level addressing
    With frmMain.imgSplitter
        '
        ' Move the splitter bar shadow image to the current mouse position
        frmMain.picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    '
    ' Show the splitter bar shadow
    frmMain.picSplitter.Visible = True
    '
    ' Set the moving indicator to true
    mblnIsMoving = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.imgSplitter MouseDown (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    imgSplitter_MouseMove
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Move the splitter shadow and check for limits
' TRIGGER:  The user moved the mouse over the splitter bar
' INPUT:    "intButton" indicates which mouse button is being used
'           "intShift" indicates whether the intShift key is being pressed
'           "sngMouse_X" and "sngMouse_Y" are the current mouse position
' OUTPUT:   None
' NOTES:
Private Sub imgSplitter_MouseMove(intButton As Integer, intShift As Integer, sngMouse_X As Single, sngMouse_Y As Single)
    Dim psngPosition As Single
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.imgSplitter MouseMove (Start)"
    '-v1.6.1
    '
    ' Only process if the user is moving the splitter bar
    If mblnIsMoving Then
        '
        ' Compute the new position
        psngPosition = sngMouse_X + frmMain.imgSplitter.Left
        '
        ' Check for limits
        If psngPosition < mintSPLIT_LIMIT Then
            '
            ' Set to the minimum
            frmMain.picSplitter.Left = mintSPLIT_LIMIT
        ElseIf psngPosition > frmMain.Width - mintSPLIT_LIMIT Then
            '
            ' Set to the maximum
            frmMain.picSplitter.Left = frmMain.Width - mintSPLIT_LIMIT
        Else
            '
            ' Set to the current position
            frmMain.picSplitter.Left = psngPosition
        End If
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.imgSplitter MouseMove (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    imgSplitter_MouseUp
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Set the new splitter position
' TRIGGER:  The user releases the mouse button after dragging the splitter bar
' INPUT:    "intButton" indicates which mouse button was released
'           "intShift" indicates whether the intShift key is pressed
'           "sngMouse_X" and "sngMouse_Y" are the current mouse position
' OUTPUT:   None
' NOTES:
Private Sub imgSplitter_MouseUp(intButton As Integer, intShift As Integer, sngMouse_X As Single, sngMouse_Y As Single)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.imgSplitter MouseUp (Start)"
    '-v1.6.1
    '
    ' Adjust the controls on the form to the new splitter position
    frmMain.SizeControls frmMain.picSplitter.Left
    '
    ' Hide the splitter shadow image
    frmMain.picSplitter.Visible = False
    '
    ' Set the moving indicator to false
    mblnIsMoving = False
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.imgSplitter Mouseup (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  SizeControls
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Position controls on the main form based on the position of the splitter bar
' INPUT:    "sngBar_X" is the position of the splitter bar
' OUTPUT:   None
' NOTES:
Sub SizeControls(sngBar_X As Single)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMain.SizeControls (Start)"
    '-v1.6.1
    '
    '+v1.6TE
    ' Only update the form controls if the form is NOT minimized
    If frmMain.WindowState <> vbMinimized Then
    '-v1.6
        '
        ' Set the min and max width
        '+v1.6TE
        'If sngBar_X < 1500 Then sngBar_X = 1500
        If sngBar_X < mintSPLIT_LIMIT Then sngBar_X = frmMain.Width / 4
        '-v1.6
        If sngBar_X > (frmMain.Width - mintSPLIT_LIMIT) Then sngBar_X = frmMain.Width - mintSPLIT_LIMIT
        '
        ' Size the tree view
        frmMain.tvTreeView.Width = sngBar_X
        '
        ' Position the splitter bar and listview
        frmMain.imgSplitter.Left = sngBar_X
        frmMain.lvListView.Left = sngBar_X + 40
        frmMain.grdData.Left = frmMain.lvListView.Left
        '
        ' Size the list view
        frmMain.lvListView.Width = frmMain.Width - (frmMain.tvTreeView.Width + 140)
        frmMain.grdData.Width = frmMain.lvListView.Width
        '
        ' Position and size the captions
        frmMain.lblTitle(0).Width = frmMain.tvTreeView.Width
        frmMain.lblTitle(1).Left = frmMain.lvListView.Left + 20
        frmMain.lblTitle(1).Width = frmMain.lvListView.Width - 40
        '
        ' Set the top based on toolbar visibility
        If frmMain.tbToolBar.Visible Then
            frmMain.tvTreeView.Top = frmMain.tbToolBar.Height + frmMain.picTitles.Height
        Else
            frmMain.tvTreeView.Top = frmMain.picTitles.Height
        End If
        '
        ' Position the list view and splitter
        frmMain.lvListView.Top = frmMain.tvTreeView.Top
        frmMain.imgSplitter.Top = frmMain.tvTreeView.Top
        frmMain.grdData.Top = frmMain.lvListView.Top
        '
        ' Set the height based on status bar visibility
        If frmMain.sbStatusBar.Visible Then
            frmMain.tvTreeView.Height = frmMain.ScaleHeight - (frmMain.picTitles.Top + frmMain.picTitles.Height + frmMain.sbStatusBar.Height)
        Else
            frmMain.tvTreeView.Height = frmMain.ScaleHeight - (frmMain.picTitles.Top + frmMain.picTitles.Height)
        End If
        '
        ' Size the list view and splitter
        frmMain.lvListView.Height = frmMain.tvTreeView.Height
        frmMain.imgSplitter.Height = frmMain.tvTreeView.Height
        frmMain.grdData.Height = frmMain.lvListView.Height
        '
        ' Restore error reporting
        On Error GoTo 0
    '
    '+v1.6TE
    End If
    '-v1.6
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMain.SizeControls (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    lvListView_DblClick
' AUTHOR:   Tom Elkins
' PURPOSE:  Process a list view item as if it were a tree view node
' TRIGGER:  User double-clicks on an item in the list view
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub lvListView_DblClick()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView DblClick (Start)"
    '-v1.6.1
    '
    ' See if there are any items in the list view window
    If frmMain.lvListView.ListItems.Count > 0 Then
        '
        ' Find the equivalent node in the tree view and force it to be selected
        frmMain.tvTreeView.Nodes(frmMain.lvListView.SelectedItem.Key).Selected = True
        '
        ' Trigger the Node Click event for that node
        tvTreeView_NodeClick frmMain.tvTreeView.Nodes(frmMain.lvListView.SelectedItem.Key)
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView DblClick (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    lvListView_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Change mode to the currently selected item
' TRIGGER:  User entered the list view
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub lvListView_GotFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView GotFocus (Start)"
    '-v1.6.1
    '
    ' See if there are any items in the list view
    If frmMain.lvListView.ListItems.Count > 0 Then
        '
        ' Change mode to the selected item
        frmMain.ChangeMode frmMain.lvListView.SelectedItem.Tag
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView GotFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    lvListView_ItemClick
' AUTHOR:   Tom Elkins
' PURPOSE:  Perform an operation on a list view liChosen
' TRIGGER:  User clicked on an liChosen in the list view
' INPUT:    "liChosen" is the liChosen the user clicked on
' OUTPUT:   None
' NOTES:
Private Sub lvListView_ItemClick(ByVal liChosen As MSComctlLib.ListItem)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView ItemClick (Start)"
    '-v1.6.1
    '
    ' Continue processing only if the selected item is different than the previous item
    If liChosen.Key <> guGUI.sItem Then
        '
        ' Save the key value of the selected item
        guGUI.sItem = liChosen.Key
        '
        ' Action is based on item type
        Select Case liChosen.Tag
            '
            ' Database
            Case gsDATABASE:
                '
                ' Only change database if item is different than the current database
                If liChosen.Key <> guCurrent.sName Then
                    '
                    ' Check for an existing database
                    If Not guCurrent.DB Is Nothing Then
                        '
                        ' Close the current database
                        basCCAT.WriteLogEntry "MAIN: lvListView_ItemClick: Closing database " & guCurrent.sName
                        guCurrent.DB.Close
                    End If
                    '
                    ' Open specified database
                    basCCAT.WriteLogEntry "MAIN: lvListView_ItemClick: Opening database " & liChosen.Key
                    Set guCurrent.DB = OpenDatabase(liChosen.Key)
                    '
                    ' Save database info
                    guCurrent.sName = guCurrent.DB.Name
                    guCurrent.iArchive = 0
                    guCurrent.sMessage = ""
                    guCurrent.fVersion = guCurrent.DB.Version 'v1.5 - database version
                    '
                    ' Set the caption
                    frmMain.sbStatusBar.Panels(1).Text = guCurrent.sName
                End If
            '
            ' Message
            Case gsMESSAGE:
                '
                ' Extract and save the database info
                guCurrent.iArchive = basCCAT.iExtract_ArchiveID(liChosen.Key)
                guCurrent.iMessage = basCCAT.iExtract_MessageID(liChosen.Key)
                guCurrent.sMessage = basCCAT.GetAlias("Message Names", "CC_MSGID" & guCurrent.iMessage, "UNKNOWN_ID" & guCurrent.iMessage)
            '
            ' Data
            Case gsDATA:
                '
                ' Extract and save the database info
                guCurrent.iArchive = basCCAT.iExtract_ArchiveID(frmMain.tvTreeView.SelectedItem.Key)
                guCurrent.iMessage = basCCAT.iExtract_MessageID(frmMain.tvTreeView.SelectedItem.Key)
                guCurrent.sMessage = basCCAT.GetAlias("Message Names", "CC_MSGID" & guCurrent.iMessage, "UNKNOWN_ID" & guCurrent.iMessage)
            '
            ' Others
            Case Else
                '
                ' Save the other database info
                guCurrent.iArchive = basCCAT.iExtract_ArchiveID(liChosen.Key)
                guCurrent.sMessage = basCCAT.GetAlias("Message Names", "CC_MSGID" & basCCAT.iExtract_MessageID(liChosen.Key), "UNKNOWN")
        End Select
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView ItemClick (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    lvListView_KeyUp
' AUTHOR:   Tom Elkins
' PURPOSE:  Trap keyboard key presses and act accordingly
' TRIGGER:  User presses a key while in the List view
' INPUT:    "intKey_Code" is the identifier of the key that was pressed
'           "intShift" indicates whether the Shift key was pressed
' OUTPUT:   None
' NOTES:
Private Sub lvListView_KeyUp(intKey_Code As Integer, intShift As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView KeyUp (Start)"
    '-v1.6.1
    '
    ' Action depends on the key pressed
    Select Case intKey_Code
        '
        '+v1.5
        ' Trap the F1 key
        Case vbKeyF1:
            '
            ' Display context-sensitive help for the List View
            basCCAT.HtmlHelp frmMain.hwnd, App.HelpFile, basCCAT.HH_HELP_CONTEXT, basCCAT.IDH_GUI_LISTVIEW
        '-v1.5
        '
        ' Make the Enter key act like a double-click
        Case vbKeyReturn:
            '
            ' Trigger the double-click event
            lvListView_DblClick
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView KeyUp (End)"
    '-v1.6.1
    '
End Sub

'
' EVENT:    lvListView_MouseUp
' AUTHOR:   Tom Elkins
' PURPOSE:  Display a popup menu of options
' TRIGGER:  User clicked the right mouse button
' INPUT:    "intButton" indicates which mouse button was pressed
'           "intKeys" indicates the state of the SHIFT, CTRL, and ALT keys
'           "sngMouse_X" and "sngMouse_Y" are the current mouse position
' OUTPUT:   None
' NOTES:    The MouseUp event is triggered after the ItemClick event, so the menu is
'           configured for the clicked item before this event is triggered.
Private Sub lvListView_MouseUp(intButton As Integer, intKeys As Integer, sngMouse_X As Single, sngMouse_Y As Single)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView MouseUp (Start)"
    '-v1.6.1
    '
    ' See if there is anything in the list
    If frmMain.lvListView.ListItems.Count > 0 Then
        '
        ' Configure the interface for the selected item
        ' The Tag property contains the type of item (Session, Database, Archive, or Message)
        ' The SmallIcon property contains the key of the icon used in the display
        If frmMain.lvListView.SelectedItem.Tag <> guGUI.sMode And frmMain.lvListView.SelectedItem.Key <> guGUI.sItem Then
            frmMain.ChangeMode frmMain.lvListView.SelectedItem.Tag
        End If
        '
        ' Check for the right mouse button
        If intButton = vbRightButton Then
            '
            ' Display the menu
            frmMain.PopupMenu frmMain.mnuPop
        End If
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.lvListView MouseUp (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' EVENT:    mnuDataEQ_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds a equality filter to the query
' TRIGGER:  User selected the "<field> = <value>" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataEQ_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataEQ Click (Start)"
    '-v1.6.1
    '
    ' Add the filter and re-query the database
    basDatabase.AddValueFilter frmMain.mnuDataEQ.Caption
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataEQ Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
' EVENT:    mnuDataExport_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Export data from the grid to a file
' TRIGGER:  User selected the "Export Data" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataExport_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataExport Click (Start)"
    '-v1.6.1
    '
    ' Trigger the "File Save" menu click event
    mnuFileSave_Click
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataExport Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuDataFilter_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Display the filter options form
' TRIGGER:  User selected the "Filter" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataFilter_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataFilter Click (Start)"
    '-v1.6.1
    '
    ' Show the filter form
    frmFilter.Show vbModal
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataFilter Click (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' EVENT:    mnuDataGE_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds an inequality filter to the query
' TRIGGER:  User selected the "<field> >= <value>" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataGE_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataGE Click (Start)"
    '-v1.6.1
    '
    ' Append the filter and re-execute the query
    basDatabase.AddValueFilter frmMain.mnuDataGE.Caption
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataGE Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' EVENT:    mnuDataGT_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds an inequality filter to the query
' TRIGGER:  User selected the "<field> > <value>" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataGT_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataGT Click (Start)"
    '-v1.6.1
    '
    ' Append the filter and re-execute the query
    basDatabase.AddValueFilter frmMain.mnuDataGT.Caption
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataGT Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
' EVENT:    mnuDataHideCol_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Remove a column from the grid display
' TRIGGER:  User selected the "Hide Column" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataHideCol_Click()
    Dim vntColumn As Variant      ' Current column in the grid
    Dim astrColumns() As String   'v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataHideCol Click (Start)"
    '-v1.6.1
    '
    ' If the shortcut for all fields is used, enter the name of all of the fields
    If guCurrent.uSQL.sFields = "*" Then
        '
        ' Reset the field list
        guCurrent.uSQL.sFields = ""
        '
        ' Loop through each column in the grid
        For Each vntColumn In frmMain.grdData.Columns
            '
            ' Add the name of the column to the field list
            guCurrent.uSQL.sFields = guCurrent.uSQL.sFields & vntColumn.DataField & ", "
        Next vntColumn
        '
        ' Remove the extraneous characters at the end
        guCurrent.uSQL.sFields = Mid(guCurrent.uSQL.sFields, 1, Len(guCurrent.uSQL.sFields) - 2)
    End If
    '
    '+v1.5
    ' Separate the field list into components
    astrColumns = Split(guCurrent.uSQL.sFields, ",")
    '
    ' See if there is only one column
    If UBound(astrColumns) = 0 Then
        '
        ' Warn the user he cannot remove the only column
        MsgBox "You cannot remove the only field in the query!", vbOKOnly Or vbExclamation, "Cannot Hide Column"
    Else
        '
        ' Blank out the column the user selected
        astrColumns(frmMain.grdData.Columns(frmMain.grdData.ColContaining(guGUI.fMouse_X)).ColIndex) = ""
        '
        ' Re-form the modified field list
        guCurrent.uSQL.sFields = Join(astrColumns, ",")
        ''
        '' Look for the column name in the list and replace it with a blank
        ''guCurrent.uSQL.sFields = Replace(guCurrent.uSQL.sFields, frmMain.grdData.Columns(frmMain.grdData.ColContaining(guGUI.fMouse_X)).DataField, "", 1, 1)
        '
        ' Replace the blank field with a single comma
        'guCurrent.uSQL.sFields = Replace(guCurrent.uSQL.sFields, ", ,", ",")
        guCurrent.uSQL.sFields = Replace(guCurrent.uSQL.sFields, ",,", ",")
    '-v1.5
        '
        ' If the last field was removed, also remove the comma and space at the end of the list
        While Right(guCurrent.uSQL.sFields, 1) = "," Or Right(guCurrent.uSQL.sFields, 1) = " "
            guCurrent.uSQL.sFields = Mid(guCurrent.uSQL.sFields, 1, Len(guCurrent.uSQL.sFields) - 1)
        Wend
        '+v1.5
        ' If the first field was removed, also remove the comma and space at the beginning of the list
        While Left(guCurrent.uSQL.sFields, 1) = "," Or Left(guCurrent.uSQL.sFields, 1) = " "
            guCurrent.uSQL.sFields = Mid(guCurrent.uSQL.sFields, 2)
        Wend
        '-v1.5
        '
        ' Re-execute the query
        '+v1.5
        'basDatabase.Requery_Data
        basDatabase.QueryData basDatabase.sCreate_SQL
        '-v1.5
    '
    '+v1.5
    End If
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataHideCol Click (Start)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' EVENT:    mnuDataLastQuery_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Executes the last saved query
' TRIGGER:  User selected the "execute last query" menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataLastQuery_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataLastQuery Click (Start)"
    '-v1.6.1
    '
    ' Re-query the database
    basDatabase.QueryData
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataLastQuery Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' EVENT:    mnuDataLE_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds an inequality filter to the query
' TRIGGER:  User selected the "<field> <= <value>" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataLE_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataLE Click (Start)"
    '-v1.6.1
    '
    ' Append the filter and re-execute the query
    basDatabase.AddValueFilter frmMain.mnuDataLE.Caption
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataLE Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' EVENT:    mnuDataLT_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds an inequality filter to the query
' TRIGGER:  User selected the "<field> < <value>" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataLT_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataLT Click (Start)"
    '-v1.6.1
    '
    ' Append the filter and re-execute the query
    basDatabase.AddValueFilter frmMain.mnuDataLT.Caption
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataLT Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' EVENT:    mnuDataNE_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds an negation filter to the query
' TRIGGER:  User selected the "<field> <> <value>" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataNE_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataNE Click (Start)"
    '-v1.6.1
    '
    ' Append the filter and re-execute the query
    basDatabase.AddValueFilter frmMain.mnuDataNE.Caption
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataNE Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
' EVENT:    mnuDataShowAllCol_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Show all columns from the database
' TRIGGER:  User selected the "Show All Columns" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataShowAllCol_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataShowAllCol Click (Start)"
    '-v1.6.1
    '
    ' Set the field list to all fields
    guCurrent.uSQL.sFields = "*"
    '
    ' Re-execute the query
    '+v1.5
    'basDatabase.Requery_Data
    basDatabase.QueryData basDatabase.sCreate_SQL
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataShowAllCol Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuDataShowAllRows_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Remove the filter from the query
' TRIGGER:  User selected the "Show All Rows" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataShowAllRows_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataShowAllRows Click (Start)"
    '-v1.6.1
    '
    ' Remove the filter
    guCurrent.uSQL.sFilter = ""
    '
    ' Re-execute the query
    '+v1.5
    'basDatabase.Requery_Data
    basDatabase.QueryData basDatabase.sCreate_SQL
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataShowAllRows Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuDataShowValue_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Modify the filter to show records of a specific value
' TRIGGER:  User selected the "<Column> = <Value>" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataShowValue_Click()
'+v1.5
'    Dim bProcess As Boolean
'    '
'    ' Default value
'    bProcess = False
'    '
'    ' See if there is already a filter
'    If Len(guCurrent.uSQL.sFilter) > 0 Then
'        '
'        ' See if this filter is already applied
'        If InStr(1, guCurrent.uSQL.sFilter, frmMain.mnuDataShowValue.Caption) = 0 Then
'            '
'            ' Add this clause to the filter
'            guCurrent.uSQL.sFilter = guCurrent.uSQL.sFilter & " AND " & frmMain.mnuDataShowValue.Caption
'            '
'            ' Set flag to process
'            bProcess = True
'        End If
'    Else
'        '
'        ' Set the filter to the clause
'        guCurrent.uSQL.sFilter = frmMain.mnuDataShowValue.Caption
'        '
'        ' Set flag to process
'        bProcess = True
'    End If
'    '
'    ' Re-execute the query
'    '+v1.5
'    'If bProcess Then basDatabase.Requery_Data
'    If bProcess Then basDatabase.QueryData basDatabase.sCreate_SQL
'    '-v1.5
'-v1.5
End Sub
'
' EVENT:    mnuDataSortCol_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Sort the data based on a column
' TRIGGER:  User selected the "Sort by <Column>" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuDataSortCol_Click()
    Dim pstrField As String
    Dim blnProcess As Boolean
    Dim astrOrder() As String     'v1.5
    Dim strTmp As String          'v1.5
    Dim intIndex As Integer       'v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataSortCol Click (Start)"
    '-v1.6.1
    '
    ' Set default
    blnProcess = False
    '
    ' Save the field name
    ' The field name could be aliased (<field> AS <alias>), so the original field name must be
    ' extracted from the field list.  This is done by first breaking the field list into a
    ' string array (using the SPLIT command), then using the index of the column containing
    ' the mouse, then extracting the first word in the field entry.
    ' For example, if the 19th entry in the field clause is "Status AS Active", the 18th column
    ' in the grid display (0 is the first column) will be named "Active".  This is not a valid
    ' name in the database, so the query will fail if "Active" is used to sort.  The ColContaining
    ' property returns the column where the mouse is (18 in this example).  The field list is
    ' broken up into an array, using the "," as the delimiter, and the current column is used
    ' to select the specific field entry ("Status AS Active").  The string is stripped of any
    ' leading and trailing spaces using the TRIM command, then is re-split at each space character
    ' so "Status AS Active" becomes "Status", "AS", and "Active".  The first (0) entry is the
    ' actual field name ("Status") and is used to sort on.
    pstrField = Split(Trim(Split(guCurrent.uSQL.sFields, ",")(frmMain.grdData.ColContaining(guGUI.fMouse_X))), " ")(0)
    '
    ' See if there is a sort list
    If Len(guCurrent.uSQL.sOrder) > 0 Then
        '
        ' See if the field is NOT already on the list
        If InStr(1, guCurrent.uSQL.sOrder, pstrField) = 0 Then
            '
            ' Add the column name to the sort list
            guCurrent.uSQL.sOrder = pstrField & ", " & guCurrent.uSQL.sOrder
            '
            ' Set flag to process
            blnProcess = True
        '
        '+v1.5
        Else
            '
            ' The field already exists in the list, so parse the list to get the individual fields
            astrOrder = Split(guCurrent.uSQL.sOrder, ",")
            '
            ' Start at the beginning of the list
            intIndex = LBound(astrOrder)
            '
            ' Look through the list until the selected field is found
            While Trim(astrOrder(intIndex)) <> pstrField
                intIndex = intIndex + 1
            Wend
            '
            ' Move the previous items in the list to the right
            While intIndex > 0
                astrOrder(intIndex) = astrOrder(intIndex - 1)
                intIndex = intIndex - 1
            Wend
            '
            ' Put the selected field in the first position
            astrOrder(LBound(astrOrder)) = pstrField
            '
            ' Re-form the sort list
            guCurrent.uSQL.sOrder = Join(astrOrder, ",")
            '
            ' Continue with processing
            blnProcess = True
        '-v1.5
        End If
    Else
        '
        ' Set the sort list
        guCurrent.uSQL.sOrder = pstrField
        '
        ' Set the processing flag
        blnProcess = True
    End If
    '
    ' Re-execute the query
    '+v1.5
    'If blnProcess Then basDatabase.Requery_Data
    If blnProcess Then basDatabase.QueryData basDatabase.sCreate_SQL
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuDataSortCol Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuEditAdd_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds an archive to the current database
' TRIGGER:  User clicked on the Edit-->Add menu item
'           User clicked on the Popup-->Add menu itme
'           User clicked on the "Archive" toolbar button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuEditAdd_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuEditAdd Click (Start)"
    '-v1.6.1
    '
    ' Call the Add Archive routine
    '+v1.6TE
    'frmArchive.Add_Archive
    frmWizard.Show vbModal, frmMain
    '-v1.6
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuEditAdd Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuEditDelete_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Deletes an archive from the current database
' TRIGGER:  User clicked on the Edit-->Delete menu item
'           User clicked on the Popup-->Delete menu item
'           User clicked on the "Delete" toolbar button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuEditDelete_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuEditDelete Click (Start)"
    '-v1.6.1
    '
    ' Action is based on the selected object type
    Select Case frmMain.ActiveControl.SelectedItem.Tag
        '
        ' Archives
        Case gsARCHIVE:
            '
            ' Confirm with the user
            If MsgBox("This will permanently delete the archive information from the database!" & vbCr & "Are you sure?", vbYesNo, "Confirm Delete") = vbYes Then
                '
                ' Delete the archive
                basDatabase.Delete_Archive frmMain.ActiveControl.SelectedItem.Key
                '
                '+v1.6TE
                ' Refresh the display to make sure invalid data is not showing
                Me.RefreshDisplay
                '-v1.6
            End If
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuEditDelete Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuEditFilter_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Display the filter options form
' TRIGGER:  User selected the "Filter" menu option
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuEditFilter_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuEditFilter Click (Start)"
    '-v1.6.1
    '
    ' Call the Data-->Filter menu event
    mnuDataFilter_Click
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuEditFilter Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuEditProperties_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  View/Edit the properties of the selected item/node
' TRIGGER:  User clicked on the Edit-->Properites menu item
'           User clicked on the Popup-->Properties menu item
'           User clicked on the "Properties" toolbar button
' INPUT:    None
' OUTPUT:   None
' NOTES:    Since the selected item could be a Node from the TreeView or an Item
'           from the ListView, the ActiveControl reference is used to handle either
'           case.
Private Sub mnuEditProperties_Click()
    Dim pstrTmp As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuEditProperties Click (Start)"
    '-v1.6.1
    '
    ' Make sure only TreeView Nodes or ListView Items are processed
    If TypeOf frmMain.ActiveControl Is TreeView Or _
       TypeOf frmMain.ActiveControl Is ListView Then
        '
        ' Action is based on the type of the selected node/item
        ' The type is kept in the Tag property and can be Session, Database, Archive,
        ' or Message
        '
        '+v1.6.1TE
        If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmMain.mnuEditProperties Click (Node is " & frmMain.ActiveControl.SelectedItem.Tag & ")"
        '-v1.6.1
        '
        Select Case frmMain.ActiveControl.SelectedItem.Tag
            '
            ' Database
            Case gsDATABASE:
                '
                ' Edit the Info table properties
                ' The database is stored in the guCurrent.DB variable
                If frmDBInfo.blnEditInfoTable(guCurrent.DB) Then
                End If
            '
            ' Archive
            Case gsARCHIVE:
                '
                ' Edit the archive properties
                '+v1.6TE
                '
                ' Close out the data grid in case it is using one of the tables
                frmMain.Data1.RecordSource = basDatabase.TBL_ARCHIVES
                frmMain.Data1.Refresh
                '
                'If frmArchive.bEdit_Archive_Properties(frmMain.ActiveControl.SelectedItem.Key) Then
                pstrTmp = frmArchiveProp.EditArchiveInfo(frmMain.ActiveControl.SelectedItem.Text)
                If pstrTmp <> "" Then frmMain.ActiveControl.SelectedItem.Text = pstrTmp
                '
                ' Update the text for the selected item
                frmMain.lblTitle(1).Caption = frmMain.tvTreeView.SelectedItem.FullPath
                '
                ' Display the contents
                basDatabase.Display_Archive_Messages pstrTmp
                'End If
                '-v1.6
            '
            ' Message
            Case gsMESSAGE:
                '
                ' Edit the Message properties
                frmMessage.DisplayMessageProperties frmMain.ActiveControl.SelectedItem.Key
            '
            '+v1.5
            ' Query
            Case gsQUERY:
                Dim pstrQuery As String     ' The entire query
                Dim pstrClause As String    ' The where clause
                Dim pintQuery As Integer    ' Query index
                '
                ' Construct and display the query
                pintQuery = CInt(Val(Mid(frmMain.tvTreeView.SelectedItem.Key, InStr(1, frmMain.tvTreeView.SelectedItem.Key, basDatabase.SEP_QUERY) + 1)))
                pstrQuery = "SELECT " & basCCAT.GetAlias("Queries", "QUERY_FIELDS" & pintQuery, "*") & " FROM " & frmMain.tvTreeView.SelectedItem.Parent.Parent.Text & basDatabase.TBL_DATA
                If basCCAT.GetAlias("Queries", "QUERY" & pintQuery, "") <> "" Then pstrQuery = pstrQuery & " WHERE " & basCCAT.GetAlias("Queries", "QUERY" & pintQuery, "")
                If basCCAT.GetAlias("Queries", "QUERY_SORT" & pintQuery, "") <> "" Then pstrQuery = pstrQuery & " ORDER BY " & basCCAT.GetAlias("Queries", "QUERY_SORT" & pintQuery, "")
                MsgBox "Query #" & pintQuery & " : " & basCCAT.GetAlias("Queries", "QUERY_TITLE" & pintQuery, "(UNTITLED)") & vbCr & pstrQuery, vbOKOnly Or vbInformation Or vbMsgBoxHelpButton, "Query Property", App.HelpFile, basCCAT.IDH_TOKEN_SQL
            '-v1.5
        End Select
    End If
    '
    ' Update the text for the selected item
    frmMain.lblTitle(1).Caption = frmMain.tvTreeView.SelectedItem.FullPath
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuEditProperties Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuFileRemove_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Remove a database entry from the tree view and list view
'           or remove a data record from the list view
' TRIGGER:  User clicked on the File-->Remove menu item
'           User clicked on the Popup-->Cut menu item
'           User clicked on the "Remove" toolbar button
' INPUT:    None
' OUTPUT:   None
' NOTES:    Since the user could have clicked on a database entry in either the
'           tree view or list view, the ActiveControl object is used to handle
'           either case.
Private Sub mnuFileRemove_Click()
    Dim pintRecord As Integer  ' Record counter
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileRemove Click (Start)"
    '-v1.6.1
    '
    ' Action depends on type
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmMain.mnuFileRemove Click (Type is " & frmMain.ActiveControl.SelectedItem.Tag & ")"
    '-v1.6.1
    '
    Select Case frmMain.ActiveControl.SelectedItem.Tag
        Case gsDATABASE:
            '
            ' Remove the database
            basCCAT.Remove_Database frmMain.ActiveControl.SelectedItem.Key
            '
            ' Trigger a nodeclick event for the new selected item
            tvTreeView_NodeClick frmMain.tvTreeView.SelectedItem
            frmMain.ChangeMode frmMain.tvTreeView.SelectedItem.Tag
        '
        ' Data record
        Case gsDATA:
            '
            ' Loop through all of the items in the list view
            For pintRecord = frmMain.lvListView.ListItems.Count To 1 Step -1
                '
                ' Check if the current record is selected
                If frmMain.lvListView.ListItems(pintRecord).Selected Then
                    '
                    ' Note it
                    basCCAT.WriteLogEntry "   Record #" & pintRecord & " (" & frmMain.lvListView.ListItems(pintRecord).Text & ") removed from the list"
                    '
                    ' Remove the record
                    frmMain.lvListView.ListItems.Remove pintRecord
                End If
            Next pintRecord
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileRemove Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuFileSave_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Saves data to an external file
' TRIGGER:  User clicked on the File-->Save menu item
'           User clicked on the Popup-->Save menu item
'           User clicked on the "Save" toolbar button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuFileSave_Click()
    Dim pintFile As Integer         ' File number
    Dim plngNum_Exported As Long    ' Number of records exported
    Dim pintPos As Integer
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileSave Click (Start)"
    '-v1.6.1
    '
    ' Use control-level addressing
    With frmMain.dlgCommonDialog
        '
        ' Change the title on the dialog
        .DialogTitle = "Save data as..."
        '
        ' Blank out the file name
        .FileName = ""
        '
        ' Set the filters to the various export types
        .Filter = "DAS Signal Activity file (*.sig)|*.sig|DAS Track file (*.mtf)|*.mtf|DAS Line-of-Bearing file (*.mtf)|*.mtf|DAS Geolocation file (*.mtf)|*.mtf|DAS Stationary Target file (*.stf)|*.stf|DAS Event file (*.evt)|*.evt|Comma-delimited Text (*.csv)|*.csv"
        '
        ' Set the flags
        .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
        '
        ' Set the default file type
        .FilterIndex = guExport.iFile_Type
        '
        ' Give it to the user
        .ShowSave
        '
        ' Save the filename
        guExport.sFile = .FileName
    End With
    '
    ' Check for a filename
    If Len(guExport.sFile) > 0 Then
        '
        ' Log the event
        basCCAT.WriteLogEntry "MAIN: mnuFileSave_Click: Output File = " & guExport.sFile
        '
        ' Save the file type
        basCCAT.guExport.iFile_Type = frmMain.dlgCommonDialog.FilterIndex
        '
        ' Select fields based on file type
        Select Case basCCAT.guExport.iFile_Type
            '
            ' Signal file
            Case 1:
                guCurrent.uSQL.sFields = "ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Other_Data"
            '
            ' Moving target track file
            Case 2:
                guCurrent.uSQL.sFields = "ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Other_Data"
            '
            ' Moving target LOB file
            Case 3:
                guCurrent.uSQL.sFields = "ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Range, Bearing, Elevation, Other_Data"
            '
            ' Moving target geolocation file
            Case 4:
                guCurrent.uSQL.sFields = "ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, XX, XY, YY, Other_Data"
            '
            ' Stationary target file
            Case 5:
                guCurrent.uSQL.sFields = "ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Parent, Parent_ID, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Other_Data"
            '
            ' Event file
            Case 6:
                '
                '+v1.6TE
                'guCurrent.uSQL.sFields = "ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Other_Data"
                guCurrent.uSQL.sFields = "ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Common_ID, Range, Other_Data"
                '-v1.6
        End Select
        '
        ' Remove sorting or grouping clauses
        If basCCAT.guExport.iFile_Type <= 6 Then
            guCurrent.uSQL.sOrder = "ReportTime"
            pintPos = InStr(1, UCase(guCurrent.uSQL.sFilter), "GROUP BY")
            If pintPos > 0 Then guCurrent.uSQL.sFilter = Left(guCurrent.uSQL.sFilter, pintPos - 1)
            '
            '+v1.5
            basDatabase.QueryData basDatabase.sCreate_SQL
            '-v1.5
        End If
        'Debug.Print guCurrent.uSQL.sFilter
        '
        ' Find an available file ID
        pintFile = FreeFile
        '
        ' Open/create the specified file
        Open basCCAT.guExport.sFile For Output As #pintFile
        '
        ' Write the DAS Header
        basCCAT.Write_DAS_Header pintFile
        '
        ' Export the data table
        '+v1.5
        'plngNum_Exported = basDatabase.lExport_Table(pintFile)
        plngNum_Exported = basDatabase.lExportGrid(pintFile)
        '-v1.5
        '
        '
        Close pintFile
        '
        ' Check for exported records
        If plngNum_Exported > 0 Then
            '
            ' Report the completion of the operation
            If MsgBox(plngNum_Exported & " records written to file " & basCCAT.guExport.sFile & vbCr & "Do you want to view the file?", vbYesNo Or vbQuestion, "Export Complete") = vbYes Then
                '
                ' Show the text file
                Shell "Notepad " & basCCAT.guExport.sFile, vbNormalFocus
            End If
        Else
            '
            ' Inform the user
            MsgBox "No records exported with current settings" & vbCr & "Please check your export settings and try again.", vbInformation Or vbOKOnly, "No Records Exported"
        End If
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileSave Click (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' EVENT:    mnuPop_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Enables/disables menu items depending on the current mode
' TRIGGER:  The user right-clicked in the GUI
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuPop_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPop Click (Start)"
    '-v1.6.1
    '
    ' See what type of item was active
    Select Case frmMain.ActiveControl.SelectedItem.Tag
        '
        ' Query node
        Case basCCAT.gsQUERY:
            frmMain.mnuPopAdd.Enabled = True
            frmMain.mnuPopCut.Enabled = False
            frmMain.mnuPopDelete.Enabled = False
            frmMain.mnuPopNew.Enabled = True
            frmMain.mnuPopOpen.Enabled = False
            frmMain.mnuPopProperties.Enabled = True
            frmMain.mnuPopSave.Enabled = False
            frmMain.mnuPopMsg.Enabled = False       'v1.7BB
            frmMain.mnuPopTemplate.Enabled = False  'v1.7BB
            frmMain.mnuPopPasteVS.Enabled = False   'v1.7SV
            frmMain.mnuPopCopyVS.Enabled = False    'v1.7SV

        End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPop Click (End)"
    '-v1.6.1
    '
End Sub


Private Sub mnuPopCopyCan_Click()
           frmCanRpt.Show , frmMain
End Sub

Private Sub mnuPopMsg_Click()
Dim rsMessage As Recordset
Dim strMsg As String

    '' Open a recordset to the Archive Message table
    Set rsMessage = guCurrent.DB.OpenRecordset("SELECT * FROM [" & guCurrent.sArchive & "_Message] WHERE Msg_Name = " & "'" & guCurrent.sMessage & "'")
    
    If (rsMessage.NoMatch = False) Then
        'If (rsMessage!Proc_Msg = False) Then
            'Process Message
            'frmtreeproc.Show vbModal, frmMain
            
            frmtreeproc.Show , frmMain
'            If (basTOC.Add_ProcMsg_Record(guCurrent.iMessage) = True) Then
'                rsMessage.Edit
'                rsMessage!Proc_Msg = True
'                rsMessage.Update
'            End If
            'frmtreeproc.Show vbModal, frmMain
        'End If
        'Display Results
    End If
    rsMessage.Close


End Sub
'
'+v1.7
' EVENT:    mnuPopCopyVS_Click
' AUTHOR:   Shaun Vogel
' PURPOSE:  Copy varStruct table from current database to paste into another db.
' TRIGGER:  The user right-clicked in the GUI
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuPopCopyVS_Click()
    Dim i As Integer
    
    If guGUI.sMode = gsARCHIVE Then
        ' Data record
        guPrevious = guCurrent
        dbPrevious = guCurrent.DB.Name
        cpyBufFull = True
        frmMain.mnuPopPasteVS.Enabled = True
    End If
   
End Sub

'
'+v1.7
' EVENT:    mnuPopPasteVS_Click
' AUTHOR:   Shaun Vogel
' PURPOSE:  Paste the varStruct table from previous database to current database.
' TRIGGER:  The user right-clicked in the GUI
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuPopPasteVS_Click()
    Dim i As Integer
    Dim dbVarStrTblPst As Recordset
    Dim dbVarStrTblCpy As Recordset
    'Dim guPrevious As Database
    
    'If MsgBox("WARNING!  All records in the varStruct table will be overwritten." & vbCr & "Do you wish to continue?", vbYesNo) = vbYes Then
    If MsgBox("WARNING!  You are about to OVERWRITE the varStruct table in the currently selected Archive," & _
               vbCr & "with the varStruct table from " & guPrevious.sArchive & vbCr & vbCr & "Do you wish to continue?", vbYesNo) = vbYes Then
        MousePointer = vbHourglass
        Set guPrevious.DB = OpenDatabase(dbPrevious)
        If guCurrent.DB Is Nothing Then
            MsgBox ("Unable to OPEN Paste buffer Database.  Try repeating Copy varStruct command")
            Return
        End If
        Set dbVarStrTblPst = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_VarStruct", dbOpenDynaset)
        Set dbVarStrTblCpy = guPrevious.DB.OpenRecordset(guPrevious.sArchive & "_VarStruct", dbOpenDynaset)
        
        If guGUI.sMode = gsARCHIVE Then
            dbVarStrTblCpy.MoveFirst
            dbVarStrTblPst.MoveFirst
            
            While Not dbVarStrTblPst.EOF
                dbVarStrTblPst.Delete
                dbVarStrTblPst.MoveNext
            Wend
            dbVarStrTblPst.MoveFirst
            
            While Not dbVarStrTblCpy.EOF
                dbVarStrTblPst.AddNew
                dbVarStrTblPst!varStructID = dbVarStrTblCpy!varStructID
                dbVarStrTblPst!msgid = dbVarStrTblCpy!msgid
                dbVarStrTblPst!fieldname = dbVarStrTblCpy!fieldname
                dbVarStrTblPst!FieldSize = dbVarStrTblCpy!FieldSize
                dbVarStrTblPst!DataType = dbVarStrTblCpy!DataType
                dbVarStrTblPst!ConvType = dbVarStrTblCpy!ConvType
                dbVarStrTblPst!fieldlabel = dbVarStrTblCpy!fieldlabel
                dbVarStrTblPst!DasField = dbVarStrTblCpy!DasField
                dbVarStrTblPst!MultiEntry = dbVarStrTblCpy!MultiEntry
                dbVarStrTblPst!MultiRecPtr = dbVarStrTblCpy!MultiRecPtr
                dbVarStrTblPst!StructLevel = dbVarStrTblCpy!StructLevel
                       
                dbVarStrTblPst.Update
                dbVarStrTblCpy.MoveNext
            Wend
            dbVarStrTblCpy.Close
            dbVarStrTblPst.Close
            MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub mnuPopTemplate_Click()

    frmTree.Show vbModal, frmMain

End Sub

'-v1.5
'
'+v1.5
' EVENT:    mnuTools_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Allows configuration of the tools menu items before displaying them
' TRIGGER:  User selected the "Tools" menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    VB does not allow altering the "Visible" property of subordinate menu items at this point
Private Sub mnuTools_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuTools Click (Start)"
    '-v1.6.1
    '
    ' Enable the remap item if the selected item is an Archive
    frmMain.mnuToolsRemapINI.Enabled = (frmMain.sbStatusBar.Panels(1).Tag = basCCAT.gsARCHIVE)
    '
    ' Set the caption for the remap item based on state
    If frmMain.mnuToolsRemapINI.Enabled Then
        frmMain.mnuToolsRemapINI.Caption = "Re-map '" & frmMain.tvTreeView.SelectedItem.Text & "'"
    Else
        frmMain.mnuToolsRemapINI.Caption = "Cannot re-map INI file"
    End If
    '
    ' Enable the SQL item if there is a database and it, or its children, are selected
    frmMain.mnuToolsExecuteSQL.Enabled = (Not basDatabase.guCurrent.DB Is Nothing) And (frmMain.sbStatusBar.Panels(1).Tag <> basCCAT.gsSESSION)
    '
    ' Enable the Update item if the SQL item is enabled and the database version is below the current version
    frmMain.mnuToolsUpdateDB.Enabled = frmMain.mnuToolsExecuteSQL And (basDatabase.guCurrent.fVersion < basDatabase.CURRENT_DB_VERSION)
    '
    ' Add the name of the database to the Update item (if enabled)
    If frmMain.mnuToolsUpdateDB.Enabled Then frmMain.mnuToolsUpdateDB.Caption = "Update Database '" & frmMain.tvTreeView.Nodes(basDatabase.guCurrent.sName).Text & "'"
    '
    ' Enable the query item if there is a query
    frmMain.mnuToolsSaveQuery.Enabled = (basDatabase.guCurrent.uSQL.sQuery <> "")
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuTools Click (End)"
    '-v1.6.1
    '
End Sub

Private Sub mnuToolsCanRun_Click()

    
       
            frmCanRpt.Show , frmMain
        
        

    
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    'If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basTOC.Get_Message_Name (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
   ' Err.Raise Err.Number, "CCAT:Get_Message_Name", Err.Description
End Sub

Private Sub mnuToolsCan_Click()

End Sub

'
'+v1.7BB
' ROUTINE:  mnuToolsCreate_Click
' AUTHOR:   Shaun Vogel
' PURPOSE:  Creates a Default VarStruct table in the current open DB and saves
'           the current open VarStruct into it.
' INPUT:
' OUTPUT:   defaultVarStruct
'
' NOTES:

Private Sub mnuToolsCreate_Click()
   If (guCurrent.iArchive = 0) Then
      MsgBox ("You must select a Database Archive first")
      Exit Sub
   End If
      
   VSOps.CreateDefaultVS
End Sub

'
'+v1.7BB
' ROUTINE:  mnuToolsImport_Click
' AUTHOR:   Shaun Vogel
' PURPOSE:  Imports the Default VarStruct table in the current open DB and saves
'           the current open VarStruct.
' INPUT:    defaultVarStruct
' OUTPUT:
'
' NOTES:

Private Sub mnuToolsImport_Click()
   If (guCurrent.iArchive = 0) Then
      MsgBox ("You must select a Database Archive first")
      Exit Sub
   End If

   VSOps.ImportDefaultVS
End Sub



'-v1.5
'
'+v1.5
'' EVENT:    mnuHelpDeg_Click
' EVENT:    mnuToolsDeg_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Calls the Convert form for degree conversion
'' TRIGGER:  User selected "Convert Degree Values" from the "Help" menu
' TRIGGER:  User selected "Convert Degree Values" from the "Tools" menu
' INPUT:    None
' OUTPUT:   None
' NOTES:
'Private Sub mnuHelpDeg_Click()
Private Sub mnuToolsDeg_Click()
'-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsDeg Click (Start)"
    '-v1.6.1
    '
    ' Set time mode to false
    frmConvert.InTimeMode = False
    '
    ' Display the form
    frmConvert.Show
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsDeg Click (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' EVENT:    mnuToolsExecuteSQL_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Allows user-entered SQL commands to be sent directly to the database
' TRIGGER:  User selected "Execute SQL Commands" from the "Tools" menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    This can be VERY dangerous in the hands of people with no/little SQL knowledge
Private Sub mnuToolsExecuteSQL_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsExecuteSQL Click (Start)"
    '-v1.6.1
    '
    basDatabase.ExecuteSQLAction
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsExecuteSQL Click (End)"
    '-v1.6.1
    '
End Sub

'-v1.5
'
'+v1.5
' EVENT:    mnuToolsRemapINI_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Re-maps some numeric values in the database to new values in the INI file
' TRIGGER:  User selected "Re-map" from the "Tools" menu
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuToolsRemapINI_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsRemapINI Click (Start)"
    '-v1.6.1
    '
    basDatabase.RemapINI
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsRemapINI Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' EVENT:    mnuToolsSaveQuery_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Allows the user to enter and save a query as a Stored Query in the INI file
' TRIGGER:  User selected "Save Custom Query" from the "Tools" menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    The query is parsed and saved in the INI file.
'           The saved query appears in the "Stored Queries" branch of the Tree View
Private Sub mnuToolsSaveQuery_Click()
    Dim pstrNew_Query As String ' New query string
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsSaveQuery Click (Start)"
    '-v1.6.1
    '
    ' Prompt the user to enter/confirm the query
    If basDatabase.guCurrent.uSQL.sQuery <> "" Then
        pstrNew_Query = InputBox("Are you sure you want to save this query?", "Confirm Query Save", basDatabase.guCurrent.uSQL.sQuery, , , App.HelpFile, basCCAT.IDH_TOKEN_SQL)
    Else
        pstrNew_Query = InputBox("Enter the SQL statement you want to save", "Enter Custom Query", "", , , App.HelpFile, basCCAT.IDH_TOKEN_SQL)
    End If
    '
    ' See if the query is the same as the current one
    If pstrNew_Query <> basDatabase.guCurrent.uSQL.sQuery Then
        '
        ' Parse the new query into its components
        basDatabase.Parse_SQL pstrNew_Query
        '
        ' Save the query
        frmFilter.SaveCustomFilter
        '
        ' Refresh the display
        mnuViewRefresh_Click
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsSaveQuery Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
'' EVENT:    mnuHelpTime_Click
' EVENT:    mnuToolsTime_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Calls the Convert form for time conversion
'' TRIGGER:  User selected "Convert Time Values" from the "Help" menu
' TRIGGER:  User selected "Convert Time Values" from the "Tools" menu
' INPUT:    None
' OUTPUT:   None
' NOTES:
'Private Sub mnuHelpTime_Click()
Private Sub mnuToolsTime_Click()
'-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsTime Click (Start)"
    '-v1.6.1
    '
    ' Set time mode to true
    frmConvert.InTimeMode = True
    '
    ' Display the form
    frmConvert.Show
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsTime Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuListViewMode_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the interface for the specified display mode
' TRIGGER:  User changed view mode from the View menu
'           User pressed view mode button in the toolbar
' INPUT:    "intNew_Mode" is the value of the new display mode
' OUTPUT:   None
' NOTES:    The menu items corresponding to the list view modes are in a control array.
'           The indices for each item matches the list view mode value the control
'           represents.
'           Value Mode          Description  Menu Index  Toolbar Button
'             0   lvwIcon       Large Icons  0           "View Large Icons"
'             1   lvwSmallIcon  Small Icons  1           "View Small Icons"
'             2   lvwList       List         2           "View List"
'             3   lvwReport     Details      3           "View Details"
Private Sub mnuListViewMode_Click(intNew_Mode As Integer)
    Dim pintOld_Mode As Integer        ' Save the old mode
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmMain.mnuListViewMode Click (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intNew_Mode
    End If
    '-v1.6.1
    '
    ' Save the current listview display mode
    pintOld_Mode = frmMain.lvListView.View
    '
    ' Set the list view to the selected mode
    frmMain.lvListView.View = intNew_Mode
    '
    ' Uncheck the old mode, and check the new one
    frmMain.mnuListViewMode(pintOld_Mode).Checked = False
    frmMain.mnuListViewMode(intNew_Mode).Checked = True
    '
    ' Raise the old button, and press the new one
    frmMain.tbToolBar.Buttons(frmMain.tbToolBar.Buttons("View Large Icons").Index + pintOld_Mode).Value = tbrUnpressed
    frmMain.tbToolBar.Buttons(frmMain.tbToolBar.Buttons("View Large Icons").Index + intNew_Mode).Value = tbrPressed
    '
    ' Save the current mode in the registry
    SaveSetting App.Title, "Settings", "ViewMode", frmMain.lvListView.View
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuListViewMode Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuPopAdd_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds an archive to the database
' TRIGGER:  User right-clicked and selected the Add item from the popup menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    Calls the Edit-->Add menu event
Private Sub mnuPopAdd_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopAdd Click (Start)"
    '-v1.6.1
    '
    '+v1.5
    ' Action based on the currently selected item
    Select Case frmMain.ActiveControl.SelectedItem.Tag
        '
        ' Query node
        Case basCCAT.gsQUERY:
            '
            ' Save a custom query
            mnuToolsSaveQuery_Click

        Case Else
    '-v1.5
            '
            ' Trigger the Edit-->Add menu click event
            mnuEditAdd_Click
    '+v1.5
    End Select
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopAdd Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuPopCut_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Removes a database from the tree
' TRIGGER:  User right-clicked and selected the Cut item from the popup menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    Calls the File-->Remove menu event
Private Sub mnuPopCut_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopCut Click (Start)"
    '-v1.6.1
    '
    ' Trigger the File-->Remove menu click event
    mnuFileRemove_Click
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopCut Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuPopDelete_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Deletes a database file or a table record
' TRIGGER:  User right-clicked and selected the Delete item from the popup menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    Calls the File-->Delete menu event or the Edit-->Delete menu event
Private Sub mnuPopDelete_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopDelete Click (Start)"
    '-v1.6.1
    '
    ' Action is based on mode
    If guGUI.sMode = gsDATABASE Then
        '
        ' Trigger the File-->Delete menu click event
        mnuFileDelete_Click
    Else
        '
        ' Trigger the Edit-->Delete menu click event
        mnuEditDelete_Click
    End If
    '
    '+v1.6TE
    ' Refresh the display to make sure invalid data is not showing
    Me.RefreshDisplay
    '-v1.6
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopDelete Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuPopNew_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Create a new database
' TRIGGER:  User right-clicked and selected the New item from the popup menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    Calls the File-->New menu event
Private Sub mnuPopNew_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopNew Click (Start)"
    '-v1.6.1
    '
    '+v1.5
    ' Action based on selected item
    Select Case frmMain.ActiveControl.SelectedItem.Tag
        '
        ' Query node
        Case basCCAT.gsQUERY:
            '
            ' Clear the query
            frmFilter.ClearQuery
            '
            ' Show the filter construction form
            frmFilter.Show vbModal, frmMain
            
        Case Else
    '-v1.5
            '
            ' Trigger the File-->New menu click event
            mnuFileNew_Click
    '+v1.5
    End Select
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopNew Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuPopOpen_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Opens an existing database
' TRIGGER:  User right-clicked and selected the Open item from the popup menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    Calls the File-->Open menu event
Private Sub mnuPopOpen_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopOpen Click (Start)"
    '-v1.6.1
    '
    ' Trigger the File-->Open menu click event
    mnuFileOpen_Click
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopOpen Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuPopProperties_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Edits the properties of the selected object
' TRIGGER:  User right-clicked and selected the Properties item from the popup menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    Calls the Edit-->Properties menu event
Private Sub mnuPopProperties_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopProperties Click (Start)"
    '-v1.6.1
    '
    ' Trigger the Edit-->Properties menu click event
    mnuEditProperties_Click
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopProperties Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuPopSave_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Saves the selected message data to a file
' TRIGGER:  User right-clicked and selected the Save item from the popup menu
' INPUT:    None
' OUTPUT:   None
' NOTES:    Calls the File-->Save menu event
Private Sub mnuPopSave_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopSave Click (Start)"
    '-v1.6.1
    '
    ' Trigger the File-->Save menu click event
    mnuFileSave_Click
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuPopSave Click (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' EVENT:    mnuToolsUpdateDB_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Upgrades the selected database to the current version
' TRIGGER:  User clicked on the Tools-->Upgrade... menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:    This can take a LONG time if the database is large
Private Sub mnuToolsUpdateDB_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsUpdateDB Click (Start)"
    '-v1.6.1
    '
    ' Warn the user
    If MsgBox("This operation may take a LONG time depending on the size of the database." & vbCr & "Also, this operation cannot be stopped once started." & vbCr & "Are you sure you want to proceed with the upgrade?", vbYesNo Or vbQuestion Or vbMsgBoxHelpButton, "Confirm Database Upgrade", App.HelpFile, basCCAT.IDH_GUI_TOOLS_UPDATE) = vbYes Then
        '
        ' Upgrade the database
        basDatabase.UpgradeDatabase
        '
        ' Make sure the tree view has focus
        frmMain.tvTreeView.SetFocus
        '
        ' Refresh the display
        mnuViewRefresh_Click
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsUpdateDB Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' EVENT:    mnuToolsViewINI_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Display the CCAT.INI file in an editor
' TRIGGER:  User clicked on the Tools-->View CCAT.INI menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:    Uses Keith's ConfigEditor, which currently does not accept command-line arguments
Private Sub mnuToolsViewINI_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsViewINI Click (Start)"
    '-v1.6.1
    '
    ' Launch the ConfigEditor
    Shell App.Path & "\ConfigEditor.exe " & basCCAT.gsCCAT_INI_Path, vbNormalFocus
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuToolsViewINI Click (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
' EVENT:    mnuViewArrangeIcons_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Arranges the icons in the List View
' TRIGGER:  User clicked on the View-->Arrange Icons menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:    Sometimes icons are scattered about the List View area, including being
'           placed on top of another icon.  The little routine forces order to the
'           icons by placing them in rows as wide as the list view area.
Private Sub mnuViewArrangeIcons_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuViewArrangeIcons Click (Start)"
    '-v1.6.1
    '
    ' Line icons along the top
    frmMain.lvListView.Arrange = lvwAutoTop
    '
    ' Redraw the list view
    frmMain.lvListView.Refresh
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuViewArrangeIcon Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    sbStatusBar_PanelClick
' AUTHOR:   Tom Elkins
' PURPOSE:  Process status bar panel clicks
' TRIGGER:  User clicked on a status bar panel
' INPUT:    "pnlClicked" is the panel clicked
' OUTPUT:   None
' NOTES:    Use the Key or Index property of the Panel object to determine which one
'           was clicked.
Private Sub sbStatusBar_PanelClick(ByVal pnlClicked As MSComctlLib.Panel)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmMain.sbStatusBar PanelClick (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & pnlClicked.Key
    End If
    '-v1.6.1
    '
    ' Action is based on which panel was clicked
    Select Case pnlClicked.Key
        '
        ' Security
        Case "SECURITY":
            '
            ' Display the classification banner
            frmSecurity.Show vbModal, frmMain
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.sbStatusBar PanelClick (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    tbToolBar_ButtonClick
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Process toolbar button clicks
' TRIGGER:  User clicked on one of the toolbar buttons
' INPUT:    "btnClicked" is the button object that the user pressed
' OUTPUT:   None
' NOTES:
Private Sub tbToolBar_ButtonClick(ByVal btnClicked As MSComctlLib.Button)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmMain.tbToolBar ButtonClick (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & btnClicked.Key
    End If
    '-v1.6.1
    '
    '
    ' Choose action based on the Key value of the button
    ' Most of the actions will be triggers for the equivalent menu event
    Select Case btnClicked.Key
        '
        ' New Database
        Case "New"
            '
            ' Trigger the File-->New menu click event
            mnuFileNew_Click
        '
        ' Open Database
        Case "Open"
            '
            ' Trigger the File-->Open menu click event
            mnuFileOpen_Click
        '
        ' Remove Database from Tree
        Case "Remove"
            '
            ' Trigger the File-->Remove menu click event
            mnuFileRemove_Click
        '
        ' Delete
        Case "Delete"
            '
            ' Action is based on mode.  Current mode is stored in the guGUI.sMode
            ' variable
            If guGUI.sMode = gsDATABASE Then
                '
                ' Trigger the File-->Delete menu click event
                mnuFileDelete_Click
            Else
                '
                ' Trigger the Edit-->Delete menu click event
                mnuEditDelete_Click
            End If
        '
        ' View Properties
        Case "Properties"
            '
            ' Trigger the Edit-->Properties menu click event
            mnuEditProperties_Click
        '
        ' Process Archive
        Case "Archive"
            '
            ' Trigger the Edit-->Add menu click event
            mnuEditAdd_Click
        '
        ' Filter Data
        Case "Filter"
            '
            ' Trigger the Data-->Filter menu click event
            mnuDataFilter_Click
        '
        ' Save Data
        Case "Save"
            '
            ' Trigger the File-->Save menu click event
            mnuFileSave_Click
        '
        ' View Large Icon mode
        Case "View Large Icons"
            '
            ' Trigger the "View-->View Large Icon" menu click event
            mnuListViewMode_Click lvwIcon
        '
        ' View Small Icon mode
        Case "View Small Icons"
            '
            ' Trigger the "View-->View Small Icon" menu click event
            mnuListViewMode_Click lvwSmallIcon
        '
        ' View List mode
        Case "View List"
            '
            ' Trigger the "View-->View List" menu click event
            mnuListViewMode_Click lvwList
        '
        ' View Detail mode
        Case "View Details"
            '
            ' Trigger the "View-->View Details" menu click event
            mnuListViewMode_Click lvwReport
        '
        ' Show Help
        Case "Help"
            '
            ' Trigger the Help-->Contents menu click event
            mnuHelpContents_Click
    End Select
    '
    ' Resume error reporting
    On Error GoTo 0
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tbToolbar ButtonClick (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuHelpAbout_Click
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Display the brag box
' TRIGGER:  User clicked on the Help-->About menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuHelpAbout_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuHelpAbout Click (Start)"
    '-v1.6.1
    '
    ' Display the form modally, so the user must respond to it.
    frmAbout.Show vbModal, frmMain
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuHelpAbout Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuHelpSearchForHelpOn_Click
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Configures the Help system to search for a word
' TRIGGER:  User clicked on the Help-->Search menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuHelpSearchForHelpOn_Click()
    '
    '+v1.5
    ' Change the return value data type
    'Dim plngReturn As Integer
    Dim plngReturn As Long
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuHelpSearchForHelpOn Click (Start)"
    '-v1.6.1
    '
    ' If there is no helpfile for this project display a message to the user.
    ' Set the Help File for the application in the Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        '
        ' Inform the user there is no help file
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, frmMain.Caption
    Else
        '
        ' Suppress error reporting
        On Error Resume Next
        '
        ' Launch the help system
        '+v1.5
        'plngReturn = OSWinHelp(frmMain.hwnd, App.HelpFile, 261, 0)
        plngReturn = basCCAT.HtmlHelp(frmMain.hwnd, App.HelpFile, basCCAT.HH_HELP_TOPIC, 0)
        '-v1.5
        '
        ' Check for errors
        If Err Then
            '
            ' Display the error
            MsgBox Err.Description
        End If
        '
        ' Resume error reporting
        On Error GoTo 0
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuHelpSearchForHelpOn Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuHelpContents_Click
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Configures the Help system to display the table of contents
' TRIGGER:  User clicked on the Help-->Contents menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuHelpContents_Click()
    '+v1.5
    ' Change the return value type
    'Dim plngReturn As integer
    Dim plngReturn As Long
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuHelpContents Click (Start)"
    '-v1.6.1
    '
    ' See if a help file was assigned
    ' Set the Help File for this application in the Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        '
        ' Inform the user that there is no help file available
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, frmMain.Caption
    Else
        '
        ' Suppress Error reporting
        On Error Resume Next
        '
        ' Launch the Help system
        '+v1.5
        'plngReturn = OSWinHelp(frmMain.hwnd, App.HelpFile, 3, 0)
        plngReturn = basCCAT.HtmlHelp(frmMain.hwnd, App.HelpFile, basCCAT.HH_HELP_TOPIC, 0)
        '-v1.5
        '
        ' Check for errors
        If Err Then
            '
            ' Display the error
            MsgBox Err.Description
        End If
        '
        ' Resume error reporting
        On Error GoTo 0
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuHelpContents Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuViewRefresh_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Redisplays the tree view or list view areas
' TRIGGER:  User clicked on the View-->Refresh menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:    The method of refreshing depends on whether the Tree View or the List View
'           was the active control at the time the user requested this operation.
'           For the Tree View, we step through the database nodes and re-add the nodes.
'           The Add Node routines will skip any existing nodes, but will add new ones.
'           For the List View, we re-execute the list view display routine for the
'           currently selected node in the Tree View.
Private Sub mnuViewRefresh_Click()
    Dim pnodCurrent As Node
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuViewRefresh Click (Start)"
    '-v1.6.1
    '
    ' See if the active control was the Tree View
    If TypeOf frmMain.ActiveControl Is TreeView Then
        '
        '+v1.6.1TE
        If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmMain.mnuViewRefresh Click (Treeview)"
        '-v1.6.1
        '
        ' Start at the session node
        Set pnodCurrent = frmMain.tvTreeView.Nodes(gsSESSION)
        '
        ' Proceed only if there are database nodes
        If pnodCurrent.Children > 0 Then
            '
            ' Move to the first database node
            Set pnodCurrent = pnodCurrent.Child
            '
            ' Continue until there are no more database nodes
            While Not pnodCurrent Is Nothing
                '
                ' Close the current database
                If Not guCurrent.DB Is Nothing Then guCurrent.DB.Close
                '
                ' Open the selected database
                Set guCurrent.DB = OpenDatabase(pnodCurrent.Key)
                '
                ' Save the database name
                guCurrent.sName = guCurrent.DB.Name
                guCurrent.fVersion = guCurrent.DB.Version 'v1.5 database version
                '
                ' Start the add-node process
                basDatabase.Add_Database_Node
                '
                ' Move to the next database node
                Set pnodCurrent = pnodCurrent.Next
            Wend
        End If
        '
        ' Force the session node to be the selected node
        frmMain.tvTreeView.Nodes(gsSESSION).Selected = True
        '
        ' Trigger the NodeClick event
        tvTreeView_NodeClick frmMain.tvTreeView.Nodes(gsSESSION)
    End If
    '
    ' Check if the active control is the List View
    If TypeOf frmMain.ActiveControl Is ListView Then
        '
        '+v1.6.1TE
        If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuViewRefresh Click (ListView)"
        '-v1.6.1
        '
        ' Action depends on type of selected node
        Select Case frmMain.tvTreeView.SelectedItem.Tag
            '
            ' Session
            Case gsSESSION:
                '
                ' Re-display the database information
                basDatabase.Display_Session_Databases frmMain.tvTreeView.SelectedItem
            '
            ' Database
            Case gsDATABASE:
                '
                ' Re-display the database archives
                basDatabase.Display_Database_Archives
            '
            ' Archive
            Case gsARCHIVE:
                '
                ' Re-display the archive summary table
                basDatabase.Display_Archive_Messages frmMain.tvTreeView.SelectedItem.Key
            '
            ' Message
            Case gsMESSAGE:
                '
                ' Re-display the message data
                basDatabase.Display_Message_Details frmMain.tvTreeView.SelectedItem
        End Select
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuViewRefresh Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuViewStatusBar_Click
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Toggle the display of the status bar on/off
' TRIGGER:  User clicked on the View-->Status Bar menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuViewStatusBar_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuViewStatusBar Click (Start)"
    '-v1.6.1
    '
    ' Update the check mark on the menu item
    frmMain.mnuViewStatusBar.Checked = Not frmMain.mnuViewStatusBar.Checked
    '
    ' Toggle the status bar visibility
    frmMain.sbStatusBar.Visible = frmMain.mnuViewStatusBar.Checked
    '
    ' Adjust the controls on the main form for the presence/absence of the status bar
    frmMain.SizeControls frmMain.imgSplitter.Left
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuViewStatusBar Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuViewToolbar_Click
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Toggle the display of the toolbar on/off
' TRIGGER:  User clicked on the View-->Tool Bar menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuViewToolbar_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuViewToolbar Click (Start)"
    '-v1.6.1
    '
    ' Update the check mark on the menu item
    frmMain.mnuViewToolbar.Checked = Not frmMain.mnuViewToolbar.Checked
    '
    ' Toggle the toolbar visibility
    frmMain.tbToolBar.Visible = frmMain.mnuViewToolbar.Checked
    '
    ' Adjust the controls on the main form for the presence absence of the tool bar
    frmMain.SizeControls frmMain.imgSplitter.Left
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuViewToolbar Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuFileClose_Click
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Exit the program
' TRIGGER:  User clicked on the File-->Exit menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuFileClose_Click()
    '
    ' Diagnostic Log
    '+v1.6.1TE
    basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileClose Click (Start)"
    'basCCAT.WriteLogEntry Format(Now, "hh:nn:ss") & ": MAIN: mnuFileClose_Click: Exiting Application"
    '-v1.6.1
    '
    ' Unload the form
    Unload frmMain
End Sub
'
' EVENT:    mnuFileDelete_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Delete the selected file
' TRIGGER:  User clicked on the File-->Delete menu item
'           User clicked on the Popup-->Delete menu item
'           User clicked on the "Delete" toolbar button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuFileDelete_Click()
    Dim pstrKenny As String     ' File to be deleted
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileDelete Click (Start)"
    '-v1.6.1
    '
    ' Ensure that only TreeView Nodes or ListView Items trigger this event
    If TypeOf frmMain.ActiveControl Is TreeView Or _
       TypeOf frmMain.ActiveControl Is ListView Then
        '
        ' Get the filename from the tag property of the selected item. The selected
        ' item could be a Node from the TreeView or an Item from the ListView.  Using
        ' the ActiveControl handles either case.  There should probably be some more
        ' logic inserted to make sure the ActiveControl is really the ListView or the
        ' TreeView
        pstrKenny = frmMain.ActiveControl.SelectedItem.Key
        '
        ' Confirm delete with the user
        If MsgBox("This action will permanently delete file" & _
            vbCr & pstrKenny & vbCr & "Are you sure?", _
            vbYesNo, "Confirm Delete") = vbYes Then
            '
            ' Log the event
            '+v1.6.1TE
            basCCAT.WriteLogEntry "INFO     : frmMain.mnuFileDelete Click (User deleting file " & pstrKenny & ")"
            '-v1.6.1
            '
            ' Close the database
            guCurrent.DB.Close
            '
            ' User said "Yes", so delete the file
            Kill pstrKenny
            '
            ' Remove the entry from the tree view and/or list view
            mnuFileRemove_Click
        End If
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileDelete Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuFileNew_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Create a new database
' TRIGGER:  User clicked on the File-->New menu item
'           User clicked on the Popup-->New menu item
'           User clicked on the "New" toolbar button
Private Sub mnuFileNew_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileNew Click (Start)"
    '-v1.6.1
    '
    ' Use control-level addressing
    With frmMain.dlgCommonDialog
        '
        ' Set the title of the dialog
        .DialogTitle = "Choose Location and Name For New Database"
        '
        ' Disable the error if the user clicks on "Cancel"
        .CancelError = False
        '
        ' Set the flags and attributes dialog
        .DefaultExt = "mdb"
        .Filter = "Databases (*.mdb)|*.mdb"
        .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
        '
        ' Show the dialog
        .ShowSave
        '
        ' Check for a filename
        If Len(.FileName) > 0 Then
            '
            ' Create the new database
            If basDatabase.bCreate_New_Database(.FileName) Then
                '
                ' Open the new database
                basDatabase.Open_Existing_Database .FileName
            End If
        End If
    End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileNew Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    mnuFileOpen_Click
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Find and open a database file
' TRIGGER:  User clicked on the File-->Open menu item
'           User clicked on the Open toolbar button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub mnuFileOpen_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileOpen Click (Start)"
    '-v1.6.1
    '
    ' Use control-level addressing
    With frmMain.dlgCommonDialog
        '
        ' Set the title of the dialog
        .DialogTitle = "Open Database File"
        '
        ' Disable the error if the user clicks on "Cancel"
        .CancelError = False
        '
        ' Set the flags and attributes dialog
        .Filter = "Databases (*.mdb)|*.mdb"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        '
        ' Show the dialog
        .ShowOpen
        '
        ' Check for a filename
        If Len(.FileName) > 0 Then
            '
            ' Open the datbase
            basDatabase.Open_Existing_Database .FileName
        End If
    End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.mnuFileOpen Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    tvTreeView_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Forces a mode change to the selected node
' TRIGGER:  The TreeView receives focus
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub tvTreeView_GotFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView GotFocus (Start)"
    '-v1.6.1
    '
    ' Check node type to current mode
    If frmMain.tvTreeView.SelectedItem.Tag <> guGUI.sMode Then
        '
        ' Force a mode change to the selected node
        frmMain.ChangeMode frmMain.tvTreeView.SelectedItem.Tag
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView GotFocus (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' EVENT:    tvTreeView_KeyUp
' AUTHOR:   Tom Elkins
' PURPOSE:  Traps keystrokes
' TRIGGER:  The user presses a key while the tree view has focus
' INPUT:    "intKey_Code" is the internal number representing the key that was pressed
'           "intShift" is a numeric code indicating the state of the intShift/Ctrl/Alt keys
' OUTPUT:   None
' NOTES:
Private Sub tvTreeView_KeyUp(intKey_Code As Integer, intShift As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView KeyUp (Start)"
    '-v1.6.1
    '
    Select Case intKey_Code
        '
        ' Trap the F1 key
        Case vbKeyF1:
            '
            ' Display help about the tree view
            basCCAT.HtmlHelp frmMain.hwnd, App.HelpFile, basCCAT.HH_HELP_CONTEXT, basCCAT.IDH_GUI_TREE
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView KeyUp (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
' EVENT:    tvTreeView_MouseDown
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the user interface for the selected node
' TRIGGER:  User clicked a mouse Button in the Tree View
' INPUT:    "intButton" indicates which mouse Button was pressed
'           "intShift" indicates whether the Shift key was pressed
'           "sngMouse_X" and "sngMouse_Y" are the current mouse coordinates
' OUTPUT:   None
' NOTES:
Private Sub tvTreeView_MouseDown(intButton As Integer, intShift As Integer, sngMouse_X As Single, sngMouse_Y As Single)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView MouseDown (Start)"
    '-v1.6.1
    '
    ' Check for the right mouse Button
    guGUI.bRight_Button = (intButton = vbRightButton)
    '
    ' Check node type to current mode
    If frmMain.tvTreeView.SelectedItem.Tag <> guGUI.sMode Then
        '
        ' Force a mode change to the selected node
        frmMain.ChangeMode frmMain.tvTreeView.SelectedItem.Tag
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView MouseDown (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    tvTreeView_MouseUp
' AUTHOR:   Tom Elkins
' PURPOSE:  Display a popup menu
' TRIGGER:  User clicked the tree view with the right mouse button
' INPUT:    "intButton" indicates which mouse button was pressed
'           "intKeys" indicates the state of the SHIFT, CTRL, and ALT keys
'           "sngMouse_X" and "sngMouse_Y" are the current mouse position
' OUTPUT:   None
' NOTES:    The MouseUp event is triggered after the NodeClick event; therefore,
'           the popup menu will have been configured by the time the mouse-up event
'           is triggered.
Private Sub tvTreeView_MouseUp(intButton As Integer, intKeys As Integer, sngMouse_X As Single, sngMouse_Y As Single)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView MouseUp (Start)"
    '-v1.6.1
    '
    ' Check for mode change
    If frmMain.tvTreeView.SelectedItem.Tag <> guGUI.sMode Then
        '
        ' Configure the interface for the selected node
        ' The mode is equivalent to the Node type, which is stored in the Tag property.
        ' nodChosen.Image is the key of the icon for the selected nodChosen.
        frmMain.ChangeMode frmMain.tvTreeView.SelectedItem.Tag
    End If
    '
    ' Check for the right mouse button
    If intButton = vbRightButton Then
        '
        ' Display the popup menu
        frmMain.PopupMenu frmMain.mnuPop
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView MouseUp (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    tvTreeView_NodeClick
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the user interface when a node is selected
' TRIGGER:  User clicked on a node in the tree view
' INPUT:    "nodChosen" is the node that was selected
' OUTPUT:   None
' NOTES:
Private Sub tvTreeView_NodeClick(ByVal nodChosen As MSComctlLib.Node)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView NodeClick (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & nodChosen.Text & " [" & nodChosen.Key & "]"
    End If
    '-v1.6.1
    '
    '
    ' Trap errors
    On Error GoTo ERR_HANDLER
    '
    ' Check for the left button
    If Not guGUI.bRight_Button Then
        '
        ' Update status
        frmMain.sbStatusBar.Panels(1).Text = "Selected " & nodChosen.Tag & " " & nodChosen.Text
        '
        ' Change mode
        frmMain.ChangeMode nodChosen.Tag
    '
    '+v1.6TE
    End If
    '-v1.6
        '
        ' Save the selected node
        guGUI.sNode = nodChosen.Key
        '
        ' Action is based on the type of node, stored in the Tag field
        Select Case nodChosen.Tag
            '
            ' Session node
            Case gsSESSION:
                '
                '+v1.6TE
                If Not guGUI.bRight_Button Then
                    '
                    ' Display the databases in the sesion
                    basDatabase.Display_Session_Databases nodChosen
                End If
                '-v1.6
            '
            ' Database node
            Case gsDATABASE:
                '
                ' See if selected database is different than current database
                If guCurrent.DB Is Nothing Then
                    Set guCurrent.DB = OpenDatabase(nodChosen.Key)
                ElseIf nodChosen.Key <> guCurrent.sName Then
                    '
                    ' Close the old database
                    guCurrent.DB.Close
                    '
                    ' Open the new database
                    Set guCurrent.DB = OpenDatabase(nodChosen.Key)
                End If
                '
                ' Save the database info
                guCurrent.sName = guCurrent.DB.Name
                guCurrent.iArchive = 0
                '+v1.6TE
                guCurrent.sArchive = ""
                '-v1.6TE
                guCurrent.iMessage = 0
                guCurrent.sMessage = ""
                guCurrent.fVersion = guCurrent.DB.Version 'v1.5 database version
                '
                '+v1.6TE
                If Not guGUI.bRight_Button Then
                    '
                    ' Display the archives stored in the database
                    basDatabase.Display_Database_Archives
                    '
                    ' Force the focus back to the tree view
                    frmMain.tvTreeView.SetFocus
                End If
                '-v1.6
            '
            ' Archive node
            Case gsARCHIVE:
                '
                ' See if selected database is different than current database
                If nodChosen.Parent.Key <> guCurrent.sName Then
                    '
                    ' Close the old database
                    guCurrent.DB.Close
                    '
                    ' Open the new database
                    Set guCurrent.DB = OpenDatabase(nodChosen.Parent.Key)
                    '
                    ' Save the db info
                    guCurrent.sName = guCurrent.DB.Name
                End If
                '
                ' Save the archive info
                guCurrent.iArchive = basCCAT.iExtract_ArchiveID(nodChosen.Key)
                '+v1.6TE
                guCurrent.sArchive = nodChosen.Text
                '-v1.6
                guCurrent.iMessage = 0
                guCurrent.sMessage = ""
                guCurrent.fVersion = guCurrent.DB.Version 'v1.5 database version
                '
                '+1.6TE
                If Not guGUI.bRight_Button Then
                    '
                    ' Display the messages stored in the archive
                    basDatabase.Display_Archive_Messages nodChosen.Key
                End If
                '-v1.6
            '
            ' Message node
            Case gsMESSAGE:
                '
                ' See if selected database is different than current database
                If nodChosen.Parent.Parent.Key <> guCurrent.sName Then
                    '
                    ' Close the old database
                    guCurrent.DB.Close
                    '
                    ' Open the new database
                    Set guCurrent.DB = OpenDatabase(nodChosen.Parent.Parent.Key)
                    '
                    ' Save the database info
                    guCurrent.sName = guCurrent.DB.Name
                End If
                '
                ' Save the other database info
                guCurrent.iArchive = basCCAT.iExtract_ArchiveID(nodChosen.Key)
                '+v1.6TE
                guCurrent.sArchive = nodChosen.Parent.Text
                '-v1.6
                guCurrent.iMessage = basCCAT.iExtract_MessageID(nodChosen.Key)
                guCurrent.sMessage = basCCAT.GetAlias("Message Names", "CC_MSGID" & guCurrent.iMessage, "UNKNOWN_ID" & guCurrent.iMessage)
                guCurrent.fVersion = guCurrent.DB.Version 'v1.5 database version
                '
                '+v1.6TE
                If Not guGUI.bRight_Button Then
                    '
                    ' Display the data for the selected message
                    basDatabase.Display_Message_Details nodChosen
                End If
                '-v1.6
            
            Case gsTOCMSG:
               '
                ' See if selected database is different than current database
                If nodChosen.Parent.Parent.Parent.Key <> guCurrent.sName Then
                    '
                    ' Close the old database
                    guCurrent.DB.Close
                    '
                    ' Open the new database
                    Set guCurrent.DB = OpenDatabase(nodChosen.Parent.Parent.Parent.Key)
                    '
                    ' Save the database info
                    guCurrent.sName = guCurrent.DB.Name
                End If
                '
                ' Save the other database info
                guCurrent.iArchive = basCCAT.iExtract_ArchiveID(nodChosen.Key)
                '+v1.6TE
                guCurrent.sArchive = nodChosen.Parent.Parent.Text
                '-v1.6
                guCurrent.iMessage = basCCAT.iExtract_MessageID(nodChosen.Key)
                guCurrent.sMessage = basTOC.Get_Message_Name(guCurrent.iMessage)
                'guCurrent.sMessage = basCCAT.GetAlias("Message Names", "CC_MSGID" & guCurrent.iMessage, "UNKNOWN_ID" & guCurrent.iMessage)
                guCurrent.fVersion = guCurrent.DB.Version 'v1.5 database version
                '
                '+v1.6TE
                If Not guGUI.bRight_Button Then
                    '
                    ' Display the data for the selected message
                    basTOC.Display_TOCMsg_Details
                End If
                '-v1.6
             
            
            '
            ' Query node
            Case gsQUERY:
                '
                ' See if selected database is different than current database
                If nodChosen.Parent.Parent.Parent.Key <> guCurrent.sName Then
                    '
                    ' Close the old database
                    guCurrent.DB.Close
                    '
                    ' Open the new database
                    Set guCurrent.DB = OpenDatabase(nodChosen.Parent.Parent.Parent.Key)
                    '
                    ' Save the database info
                    guCurrent.sName = guCurrent.DB.Name
                End If
                '
                ' Save the other info
                guCurrent.iArchive = basCCAT.iExtract_ArchiveID(nodChosen.Key)
                '+v1.6TE
                guCurrent.sArchive = nodChosen.Parent.Parent.Text
                '-v1.6
                guCurrent.iMessage = 0
                guCurrent.sMessage = "Query " & nodChosen.Text
                guCurrent.fVersion = guCurrent.DB.Version 'v1.5 database version
                '
                '+v1.6TE
                If Not guGUI.bRight_Button Then
                    '
                    ' Execute the query
                    basDatabase.Display_Query_Results CInt(Val(Mid(nodChosen.Key, InStr(1, nodChosen.Key, basDatabase.SEP_QUERY) + 1)))
                End If
                '-v1.6
        End Select
        '
        '+v1.6TE
        '
        ' Resize the columns
        frmMain.SizeColumns
        '
        ' Update the header
        If Not guGUI.bRight_Button Then
            '
            ' Update the List View caption
            frmMain.lblTitle(1).Caption = nodChosen.FullPath
        End If
        '-v1.6
    '
    '+v1.6TE
    'Else
    '    '
    '    ' Reset the button trap
    '    guGUI.bRight_Button = False
    'End If
    '-v1.6
    '
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.tvTreeView NodeClick (End)"
    '-v1.6.1
    '
    Exit Sub
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmMain.tvTreeView NodeClick (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "while processing database " & guCurrent.DB.Name, vbOKOnly, "Error Processing Database"
    Set guCurrent.DB = Nothing
    On Error GoTo 0
End Sub
'
' ROUTINE:  ChangeMode
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the interface for the specified mode
' INPUT:    "strMode" is the name of the new mode.  Use the constants gsSESSION,
'           gsDATABASE, gsARCHIVE, or gsMESSAGE as defined in basCCAT.  These
'           constants are applied to the Tag property of nodes and items.
' OUTPUT:   None
' NOTES:
Public Sub ChangeMode(strMode As String)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmMain.ChangeMode (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & strMode
    End If
    '-v1.6.1
    '
    ' See if mode changed
    If strMode <> guGUI.sMode Then
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : frmMain.ChangeMode (From " & guGUI.sMode & " To " & strMode & ")"
        '
        ' Store new mode
        guGUI.sMode = strMode
        '
        ' Configure the toolbar buttons
        frmMain.tbToolBar.Buttons("New").Enabled = (strMode = gsSESSION)
        frmMain.tbToolBar.Buttons("Open").Enabled = (strMode = gsSESSION)
        frmMain.tbToolBar.Buttons("Remove").Enabled = (strMode = gsDATABASE) Or (strMode = gsDATA)
        If strMode = gsDATA Then
            frmMain.tbToolBar.Buttons("Remove").ToolTipText = "Remove the selected data record from the list"
        Else
            frmMain.tbToolBar.Buttons("Remove").ToolTipText = "Remove the selected database from the session"
        End If
        frmMain.tbToolBar.Buttons("Delete").Enabled = (strMode = gsDATABASE) Or (strMode = gsARCHIVE)
        If strMode = gsDATABASE Then
            frmMain.tbToolBar.Buttons("Delete").ToolTipText = "Delete the selected database file"
        Else
            frmMain.tbToolBar.Buttons("Delete").ToolTipText = "Delete the selected archive from the database"
        End If
        frmMain.tbToolBar.Buttons("Properties").Enabled = (strMode = gsDATABASE) Or (strMode = gsARCHIVE) Or (strMode = gsMESSAGE)
        frmMain.tbToolBar.Buttons("Archive").Enabled = (strMode = gsDATABASE)
        'frmMain.tbToolBar.Buttons("Save").Enabled = (strMode = gsMESSAGE) Or (strMode = gsDATA) Or (strMode = gsQUERY)
        frmMain.tbToolBar.Buttons("Save").Enabled = (strMode = gsDATA) Or (strMode = gsQUERY)
        'frmMain.tbToolBar.Buttons("Filter").Enabled = (strMode = gsMESSAGE) Or (strMode = gsDATA) Or (strMode = gsQUERY)
        frmMain.tbToolBar.Buttons("Filter").Enabled = (strMode = gsDATA) Or (strMode = gsQUERY)
        '
        ' Configure the menu options
        frmMain.mnuFileOpen.Enabled = (strMode = gsSESSION)
        frmMain.mnuFileNew.Enabled = (strMode = gsSESSION)
        frmMain.mnuFileDelete.Enabled = (strMode = gsDATABASE)
        frmMain.mnuFileSave.Enabled = (strMode = gsMESSAGE) Or (strMode = gsDATA)
        frmMain.mnuFileRemove.Enabled = (strMode = gsDATABASE) Or (strMode = gsDATA)
        frmMain.mnuEditAdd.Enabled = (strMode = gsDATABASE)
        frmMain.mnuEditDelete.Enabled = (strMode = gsDATABASE) Or (strMode = gsARCHIVE)
        frmMain.mnuEditProperties.Enabled = (strMode = gsDATABASE) Or (strMode = gsARCHIVE) Or (strMode = gsMESSAGE)
        frmMain.mnuEditFilter.Enabled = (strMode = gsQUERY) Or (strMode = gsDATA)
        frmMain.mnuViewArrangeIcons.Enabled = (strMode = gsSESSION) Or (strMode = gsDATABASE) Or (strMode = gsARCHIVE)
        frmMain.mnuViewRefresh.Enabled = (strMode = gsSESSION) Or (strMode = gsDATABASE) Or (strMode = gsARCHIVE)
        frmMain.mnuListViewMode(0).Enabled = (strMode = gsSESSION) Or (strMode = gsDATABASE) Or (strMode = gsARCHIVE)
        frmMain.mnuListViewMode(1).Enabled = (strMode = gsSESSION) Or (strMode = gsDATABASE) Or (strMode = gsARCHIVE)
        frmMain.mnuListViewMode(2).Enabled = (strMode = gsSESSION) Or (strMode = gsDATABASE) Or (strMode = gsARCHIVE)
        frmMain.mnuListViewMode(3).Enabled = (strMode = gsSESSION) Or (strMode = gsDATABASE) Or (strMode = gsARCHIVE)
        '
        ' Configure the popup menu options
        frmMain.mnuPopAdd.Enabled = (strMode = gsDATABASE)
        frmMain.mnuPopCut.Enabled = (strMode = gsDATABASE) Or (strMode = gsDATA)
        frmMain.mnuPopDelete.Enabled = (strMode = gsDATABASE) Or (strMode = gsARCHIVE)
        frmMain.mnuPopNew.Enabled = (strMode = gsSESSION)
        frmMain.mnuPopOpen.Enabled = (strMode = gsSESSION)
        frmMain.mnuPopProperties.Enabled = (strMode = gsDATABASE) Or (strMode = gsARCHIVE) Or (strMode = gsMESSAGE)
        'frmMain.mnuPopSave.Enabled = (strMode = gsMESSAGE)
        frmMain.mnuPopSave.Enabled = False
        frmMain.mnuPopMsg.Enabled = (strMode = gsTOCMSG)        '
        frmMain.mnuPopTemplate.Enabled = (strMode = gsTOCMSG)
        frmMain.mnuPopCopyCan.Enabled = (strMode = gsARCHIVE)
        frmMain.mnuPopCopyVS.Enabled = (strMode = gsARCHIVE)    'v1.7SV
        frmMain.mnuPopPasteVS.Enabled = (strMode = gsARCHIVE) And cpyBufFull    'v1.7SV
        ' Save the current mode in the status bar panel tag field
        frmMain.sbStatusBar.Panels(1).Tag = strMode
        '
        '+v1.5
        ' Check the INI file to enable/disable the advanced SQL options
        ' You cannot modify the Visible property of a menu item from the parent menu's Click event,
        ' so this is a way to periodically check and update the status (in case the user changes
        ' the INI file).
        frmMain.mnuToolsExecuteSQL.Visible = (basCCAT.GetNumber("Miscellaneous Operations", "ADVSQL", 0) = 1)
        '-v1.5
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMain.ChangeMode (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: blnNodeExists
' AUTHOR:   Tom Elkins
' PURPOSE:  Check for the existence of a node in the Tree View
' INPUT:    "pstrKey" is the key value for the node in question
' OUTPUT:   TRUE if the specified node exists
'           FALSE if the node does not exist
' NOTES:
Public Function blnNodeExists(pstrKey As String) As Boolean
    Dim pnodCurrent As Node  ' The specified node
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : frmMain.blnNodeExists (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & pstrKey
    End If
    '-v1.6.1
    '
    '
    ' Suppress error reporting
    On Error Resume Next
    '
    ' Attempt to access the node in question
    Set pnodCurrent = frmMain.tvTreeView.Nodes.Item(pstrKey)
    '
    ' Check for errors
    Select Case Err.Number
        '
        ' If there was no error, the node exists
        Case NO_ERROR:
            blnNodeExists = True
        '
        ' Error 35601 means the specified node does not exist in the collection
        Case 35601:
            blnNodeExists = False
        '
        ' If we get a different error, report it.
        Case Else:
            MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "while in frmMain.blnNodeExists", , "Unexpected Error"
            blnNodeExists = False
    End Select
    '
    ' Resume error reporting
    On Error GoTo 0
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmMain.blnNodeExists (End)"
    '-v1.6.1
    '
End Function
'
'+v1.5
' ROUTINE:  UpdateStatusText
' AUTHOR:   Tom Elkins
' PURPOSE:  Displays the supplied text in the status bar panel
' INPUT:    "strMsg" is the text to be displayed
'           "strIconKey" is the key for an icon to be displayed
' OUTPUT:   None
' NOTES:
Public Sub UpdateStatusText(strMsg As String, Optional strIconKey As String)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmMain.UpdateStatusText (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & strMsg & "," & strIconKey
    End If
    '-v1.6.1
    '
    '
    ' Display the text
    frmMain.sbStatusBar.Panels(1).Text = strMsg
    '
    ' Find the specified icon
    If strIconKey <> "" Then frmMain.sbStatusBar.Panels(1).Picture = frmMain.imlSmallIcons.ListImages(strIconKey).Picture
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMain.UpdateStatusText (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' ROUTINE:  ShowProgressBar
' AUTHOR:   Tom Elkins
' PURPOSE:  Locates, sizes, scales, and displays the progress bar in first panel of the status bar
' INPUT:    "lngMin" is the minimum value
'           "lngMax" is the maximum value
'           "lngStart" is the starting value
' OUTPUT:   None
' NOTES:    Uses the current physical dimensions of the status bar panel to locate and size the
'           progress bar.
Public Sub ShowProgressBar(lngMin As Long, lngMax As Long, lngStart As Long)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmMain.ShowProgressBar (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & lngMin & ", " & lngMax & ", " & lngStart
    End If
    '-v1.6.1
    '
    '
    ' Position the progress bar in the second status bar panel
    frmMain.barLoad.Move frmMain.sbStatusBar.Panels(1).Left, frmMain.sbStatusBar.Top, frmMain.sbStatusBar.Panels(1).Width, frmMain.sbStatusBar.Height
    '
    ' Set the limits and the starting point
    frmMain.barLoad.Max = lngMax
    frmMain.barLoad.Min = lngMin
    frmMain.barLoad.Value = lngStart
    '
    ' Show the bar
    frmMain.barLoad.Visible = True
    DoEvents
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMain.ShowProgressBar (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' ROUTINE:  RefreshDisplay
' AUTHOR:   Tom Elkins
' PURPOSE:  Forces the tree view to redraw
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Sub RefreshDisplay()
    Dim pnodCurrent As Node
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMain.RefreshDisplay (Start)"
    '-v1.6.1
    '
    ' Start at the session node
    Set pnodCurrent = frmMain.tvTreeView.Nodes(gsSESSION)
    '
    ' Proceed only if there are database nodes
    If pnodCurrent.Children > 0 Then
        '
        ' Move to the first database node
        Set pnodCurrent = pnodCurrent.Child
        '
        ' Continue until there are no more database nodes
        While Not pnodCurrent Is Nothing
            '
            ' Close the current database
            If Not guCurrent.DB Is Nothing Then guCurrent.DB.Close
            '
            ' Open the selected database
            Set guCurrent.DB = OpenDatabase(pnodCurrent.Key)
            '
            ' Save the database name
            guCurrent.sName = guCurrent.DB.Name
            guCurrent.fVersion = guCurrent.DB.Version 'v1.5 database version
            '
            ' Start the add-node process
            basDatabase.Add_Database_Node
            '
            ' Move to the next database node
            Set pnodCurrent = pnodCurrent.Next
        Wend
    End If
    '
    ' Force the session node to be the selected node
    frmMain.tvTreeView.Nodes(gsSESSION).Selected = True
    '
    ' Trigger the NodeClick event
    tvTreeView_NodeClick frmMain.tvTreeView.Nodes(gsSESSION)
    '
    ' Action depends on type of selected node
    Select Case frmMain.tvTreeView.SelectedItem.Tag
        '
        ' Session
        Case gsSESSION:
            '
            ' Re-display the database information
            basDatabase.Display_Session_Databases frmMain.tvTreeView.SelectedItem
        '
        ' Database
        Case gsDATABASE:
            '
            ' Re-display the database archives
            basDatabase.Display_Database_Archives
        '
        ' Archive
        Case gsARCHIVE:
            '
            ' Re-display the archive summary table
            basDatabase.Display_Archive_Messages frmMain.tvTreeView.SelectedItem.Key
        '
        ' Message
        Case gsMESSAGE:
            '
            ' Re-display the message data
            basDatabase.Display_Message_Details frmMain.tvTreeView.SelectedItem
            
        Case gsTOCMSG:
        
        
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMain.RefreshDisplay (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.6TE
' ROUTINE:  SizeColumns
' AUTHOR:   Tom Elkins
' PURPOSE:  Sizes the columns of the list view to the contents
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Sub SizeColumns()
    Dim pintCol As Integer      ' Current column
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMain.SizeColumns (Start)"
    '-v1.6.1
    '
    ' Loop through the columns
    For pintCol = 0 To frmMain.lvListView.ColumnHeaders.Count - 1
        '
        ' Send the Windows resize message to each column
        basCCAT.SendMessage frmMain.lvListView.hwnd, &H1000 + 30, pintCol, -1
    Next pintCol
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMain.SizeColumns (End)"
    '-v1.6.1
    '
End Sub
'-v1.6
'
