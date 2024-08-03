VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmArchive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archive Options"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6105
   ControlBox      =   0   'False
   Icon            =   "frmArchive.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   4890
      TabIndex        =   36
      ToolTipText     =   "Get detailed help"
      Top             =   4470
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgHelp 
      Left            =   105
      Top             =   4410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnRestore 
      Caption         =   "Restore"
      Height          =   375
      Left            =   1410
      TabIndex        =   14
      Top             =   4470
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3735
      TabIndex        =   13
      ToolTipText     =   "Ignore changes to info"
      Top             =   4470
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2580
      TabIndex        =   12
      ToolTipText     =   "Save information about this archive"
      Top             =   4470
      Width           =   1095
   End
   Begin VB.Frame fraPage 
      Caption         =   "fraPage3"
      Height          =   3660
      Index           =   3
      Left            =   -20000
      TabIndex        =   15
      Top             =   555
      Width           =   5775
      Begin VB.CommandButton btnStopTranslating 
         Height          =   600
         Left            =   2520
         Picture         =   "frmArchive.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Translate filtered file to database"
         Top             =   1200
         Width           =   600
      End
      Begin VB.CommandButton btnProcess 
         Caption         =   "Translate"
         Height          =   360
         Left            =   255
         TabIndex        =   37
         ToolTipText     =   "Translate filtered file to database"
         Top             =   1245
         Width           =   1080
      End
      Begin VB.TextBox txtFilterFile 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1065
         TabIndex        =   35
         Text            =   "txtFilterFile"
         ToolTipText     =   "Name of filtered file"
         Top             =   705
         Width           =   4650
      End
      Begin MSComctlLib.ProgressBar barProgress 
         Height          =   390
         Left            =   45
         TabIndex        =   16
         ToolTipText     =   "Progress"
         Top             =   2145
         Width           =   5700
         _ExtentX        =   10054
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblProcessInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblProcessInfo"
         Height          =   240
         Left            =   75
         TabIndex        =   42
         Top             =   1905
         Width           =   5550
      End
      Begin VB.Line linShadow 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   15
         X2              =   5745
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Line linDivide 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   15
         X2              =   5745
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label lblArcFilter 
         AutoSize        =   -1  'True
         Caption         =   "Filtered File:"
         Height          =   195
         Left            =   195
         TabIndex        =   34
         Top             =   750
         Width           =   840
      End
      Begin VB.Label lblArcProcessed 
         AutoSize        =   -1  'True
         Caption         =   "Last Processed:"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label lblProcessed 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblProcessed"
         Height          =   315
         Left            =   1350
         TabIndex        =   25
         ToolTipText     =   "Last time this file was translated"
         Top             =   300
         Width           =   2775
      End
      Begin VB.Line linDivide 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   15
         X2              =   5760
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line linShadow 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   5745
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label lblArcStart 
         AutoSize        =   -1  'True
         Caption         =   "First Time"
         Height          =   195
         Left            =   105
         TabIndex        =   24
         Top             =   2730
         Width           =   675
      End
      Begin VB.Label lblArcEnd 
         AutoSize        =   -1  'True
         Caption         =   "Last Time"
         Height          =   195
         Left            =   3015
         TabIndex        =   23
         Top             =   2730
         Width           =   690
      End
      Begin VB.Label lblStart 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblStart"
         Height          =   285
         Left            =   870
         TabIndex        =   22
         ToolTipText     =   "Earliest time found in archive"
         Top             =   2700
         Width           =   2070
      End
      Begin VB.Label lblEnd 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblEnd"
         Height          =   285
         Left            =   3825
         TabIndex        =   21
         ToolTipText     =   "Latest time found in archive"
         Top             =   2700
         Width           =   1905
      End
      Begin VB.Label lblArcMsg 
         AutoSize        =   -1  'True
         Caption         =   "Messages Processed"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   3090
         Width           =   1515
      End
      Begin VB.Label lblNumMsg 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumMsg"
         Height          =   285
         Left            =   1665
         TabIndex        =   19
         ToolTipText     =   "Number of messages processed"
         Top             =   3075
         Width           =   1275
      End
      Begin VB.Label lblNumBytes 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumBytes"
         Height          =   285
         Left            =   3825
         TabIndex        =   18
         ToolTipText     =   "Number of bytes processed"
         Top             =   3075
         Width           =   1410
      End
      Begin VB.Label lblBytes 
         AutoSize        =   -1  'True
         Caption         =   "Bytes"
         Height          =   195
         Left            =   3330
         TabIndex        =   17
         Top             =   3120
         Width           =   390
      End
   End
   Begin VB.Frame fraPage 
      Caption         =   "fraPage0"
      Height          =   3645
      Index           =   0
      Left            =   -20000
      TabIndex        =   27
      Top             =   570
      Width           =   5595
      Begin VB.CheckBox chkRewind 
         Caption         =   "Rewind Tape when done"
         Height          =   225
         Left            =   390
         TabIndex        =   54
         Top             =   2070
         Width           =   2145
      End
      Begin VB.TextBox txtArchiveName 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text1"
         ToolTipText     =   "Extracted archive file name"
         Top             =   1305
         Width           =   3975
      End
      Begin VB.CommandButton btnExtract 
         Caption         =   "Extract Files"
         Height          =   375
         Left            =   2850
         TabIndex        =   2
         Top             =   2010
         Width           =   1365
      End
      Begin VB.CommandButton btnTapeInit 
         Caption         =   "Initialize Tape"
         Height          =   330
         Left            =   1845
         TabIndex        =   0
         Top             =   540
         Width           =   1410
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save To..."
         Height          =   330
         Left            =   4365
         TabIndex        =   1
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Shape shpStatus3 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   90
         Shape           =   3  'Circle
         Top             =   1755
         Width           =   195
      End
      Begin VB.Shape shpStatus2 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   90
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   195
      End
      Begin VB.Shape shpStatus1 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   90
         Shape           =   3  'Circle
         Top             =   225
         Width           =   195
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         Caption         =   "Step 3 - Extract the files from the tape"
         Height          =   195
         Left            =   360
         TabIndex        =   51
         Top             =   1755
         Width           =   2655
      End
      Begin VB.Label lblStep1 
         AutoSize        =   -1  'True
         Caption         =   "Step 1 - Initialize the Tape"
         Height          =   195
         Left            =   315
         TabIndex        =   50
         Top             =   225
         Width           =   1845
      End
      Begin VB.Label lblTapeControls 
         Caption         =   "Rew  End   Prev Next  Rec Eject"
         Height          =   285
         Left            =   1305
         TabIndex        =   49
         Top             =   2025
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label lblTapeData 
         Caption         =   "Archive Data file: N/A"
         Height          =   195
         Left            =   405
         TabIndex        =   48
         Top             =   2970
         Width           =   5010
      End
      Begin VB.Label lblTapeJunk 
         Caption         =   "Archive Header file: N/A"
         Height          =   195
         Left            =   405
         TabIndex        =   47
         Top             =   2700
         Width           =   5055
      End
      Begin VB.Label lblTapeTime 
         Caption         =   "Archive Time stamp: N/A"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   405
         TabIndex        =   46
         Top             =   2430
         Width           =   5010
      End
      Begin VB.Label lblInstruction 
         AutoSize        =   -1  'True
         Caption         =   "Step 2 - Select destination and name for extracted archive file."
         Height          =   195
         Left            =   315
         TabIndex        =   45
         Top             =   1080
         Width           =   4395
      End
      Begin VB.Label lblArchiveName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   330
         Left            =   360
         TabIndex        =   44
         Top             =   1305
         Width           =   3975
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraPage 
      Caption         =   "fraPage2"
      Height          =   3645
      Index           =   2
      Left            =   210
      TabIndex        =   38
      Top             =   570
      Width           =   5595
      Begin MSComctlLib.ListView lvMessages 
         Height          =   3210
         Left            =   60
         TabIndex        =   39
         ToolTipText     =   "Select/Deselect messages to translate"
         Top             =   390
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   5662
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Message"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
      End
   End
   Begin VB.Frame fraPage 
      Caption         =   "fraPage1"
      Height          =   3645
      Index           =   1
      Left            =   -20000
      TabIndex        =   28
      Top             =   555
      Width           =   5595
      Begin VB.TextBox txtFile 
         BackColor       =   &H8000000F&
         Height          =   330
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   1575
         Width           =   3255
      End
      Begin VB.ComboBox cmbClass 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Text            =   "UNCLASSIFIED"
         ToolTipText     =   "Select the archive classification"
         Top             =   2520
         Width           =   2415
      End
      Begin VB.CommandButton btnFilter 
         Caption         =   "Filter and Translate"
         Height          =   330
         Left            =   1320
         TabIndex        =   10
         ToolTipText     =   "Filter and translate the archive"
         Top             =   3015
         Width           =   1740
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Text            =   "txtName"
         ToolTipText     =   "Enter a name for the archive"
         Top             =   750
         Width           =   2205
      End
      Begin VB.CommandButton btnBrowse 
         Caption         =   "Browse"
         Height          =   315
         Left            =   4620
         TabIndex        =   7
         ToolTipText     =   "Select an archive file"
         Top             =   1590
         Width           =   795
      End
      Begin VB.OptionButton optMedia 
         Caption         =   "8mm Tape"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   1170
         Width           =   1215
      End
      Begin VB.OptionButton optMedia 
         Caption         =   "Disk"
         Height          =   375
         Index           =   1
         Left            =   2580
         TabIndex        =   5
         ToolTipText     =   "Choose the media for the archive"
         Top             =   1170
         Width           =   735
      End
      Begin VB.OptionButton optMedia 
         Caption         =   "CD/DVD"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   6
         Top             =   1170
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   345
         Left            =   1320
         TabIndex        =   8
         ToolTipText     =   "Select the date offset for mission time"
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   60358659
         CurrentDate     =   36272
      End
      Begin VB.Label lblFile 
         BorderStyle     =   1  'Fixed Single
         Caption         =   ".lblFile"
         Height          =   315
         Left            =   1320
         TabIndex        =   43
         ToolTipText     =   "Use BROWSE to select an archive file"
         Top             =   1590
         Width           =   3255
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         Caption         =   "Classification"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   2550
         Width           =   915
      End
      Begin VB.Label lblArcDate 
         AutoSize        =   -1  'True
         Caption         =   "Archive Date"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   2100
         Width           =   930
      End
      Begin VB.Label lblArcID 
         AutoSize        =   -1  'True
         Caption         =   "Archive ID"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   420
         Width           =   750
      End
      Begin VB.Label lblID 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblID"
         Height          =   255
         Left            =   1320
         TabIndex        =   32
         ToolTipText     =   "Database identifier"
         Top             =   390
         Width           =   375
      End
      Begin VB.Label lblArcName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Archive Name"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label lblArcFile 
         AutoSize        =   -1  'True
         Caption         =   "Archive File"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   1650
         Width           =   825
      End
      Begin VB.Label lblArcMedia 
         Caption         =   "Archive Media"
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   1230
         Width           =   1035
      End
   End
   Begin MSComctlLib.TabStrip tabOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tape"
            Key             =   "Tape"
            Object.ToolTipText     =   "Tape operations"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Raw Archive"
            Key             =   "Raw"
            Object.ToolTipText     =   "Raw archive operations"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Messages"
            Key             =   "Messages"
            Object.ToolTipText     =   "Message operations"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Filtered Archive"
            Key             =   "Filtered"
            Object.ToolTipText     =   "Filtered archive operations"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' FORM:     frmArchive
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Interface for processing archives
' REVISIONS:
'   v1.3.0  TAE Added Keith's tape controls
'   v1.3.1  TAE Changed some label controls to read-only text controls
'   v1.3.2  TAE Added tape extract button
'           TAE Added ability to save filtered file to a user-specified location
'           TAE Added ability to select destination for archive files
'           TAE Added button to initialize tape before other functions
'           TAE Tied some message boxes to help files and added help button
'   v1.5.0  TAE Use the archive date as the base for time calculations
'           TAE Use date/time values directly for calculation rather than convert to/from TSecs
'           TAE Format date/time text to fit existing controls and to be compatible with the
'               old database schema
'           TAE Added a Stop button that is enabled during translation.  When clicked, the
'               software will stop translation gracefully.
'           TAE Repaired the help button to use the HTML help file.  When the help button is
'               pressed, the help for the current tab is displayed.
'           TAE Added F1 help to the form.  The help page for the current tab is displayed.
'
Option Explicit
'v1.3
' Keith's tape interface
Private Declare Function InitializeTape Lib "ReadTape.dll" () As Long
Private Declare Function ParsePathStr Lib "ReadTape.dll" (ByVal sMyString As String) As String
Private Declare Function Rewind Lib "ReadTape.dll" () As Long
Private Declare Function SpaceForward Lib "ReadTape.dll" () As Long
Private Declare Function SpaceBackward Lib "ReadTape.dll" () As Long
Private Declare Function FindEndOfData Lib "ReadTape.dll" () As Long
Private Declare Function ReadTapeFile Lib "ReadTape.dll" () As Long
Private Declare Function ScanTape Lib "ReadTape.dll" () As Long
Private Declare Function EjectTape Lib "ReadTape.dll" () As Long
'
' Form-level constants
Const TAB_TAPE = 0
Const TAB_RAW = 1
Const TAB_MESSAGE = 2
Const TAB_FILTERED = 3
Const OPT_TAPE = 0
Const OPT_HD = 1
Const OPT_CD = 2
Const FILE_RAW = 1
Const FILE_FILTERED = 2
'
' EVENT:    btnBrowse_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Allows the user to select an archive file
' TRIGGER:  User clicked on the "Browse" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnBrowse_Click()
    '
    ' Use control-level addressing
    With frmMain.dlgCommonDialog
        '
        ' Set the file filter.
        ' Use the extensions specified in Brad's Raw2FiltRaw module.  If it changes
        ' there, the filter will also change to accomodate.
        .Filter = "COMPASS CALL Archives (*" & Raw2filtraw.sArchiveExt & ")|*" & Raw2filtraw.sArchiveExt & "|Filtered Archive Files (*" & Raw2filtraw.sFilteredExt & ")|*" & Raw2filtraw.sFilteredExt
        '
        ' Set flags
        .Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
        '
        ' Set the dialog title
        .DialogTitle = "Select archive file to process"
        '
        ' Set the default file to the one specified in the archive record
        .FileName = frmArchive.txtFile.Text
        '
        ' Save the old name
        .Tag = frmArchive.txtFile.Text
        '
        ' Allow the user to select a file
        .ShowOpen
        '
        ' Save the filename if the user selected one
        If .FileName <> .Tag Then
            '
            ' Save the new file name
            frmArchive.txtFile.Text = .FileName
            guArchive.sFile = .FileName
            '
            ' Store the file type
            guArchive.iType = .FilterIndex
            '
            ' Reset the classification
            frmArchive.cmbClass.Enabled = True
            guArchive.sClass = ""
            frmArchive.cmbClass.ListIndex = 0
            '
            ' Reset the date
            If Dir(.FileName) <> "" Then
                frmArchive.dtDate.Enabled = True
                frmArchive.dtDate.Value = FileDateTime(.FileName)
                guArchive.sDate = Format(frmArchive.dtDate.Value, "mm/dd/yyyy")
            Else
                frmArchive.dtDate.Value = Date
                guArchive.sDate = Date
            End If
            '
            '+v1.5
            ' Update the archive date
            guCurrent.uArchive.dtArchiveDate = guArchive.sDate
            '-v1.5
            '
            ' Log the event
            basCCAT.WriteLogEntry "ARCHIVE: BROWSE: User selected archive file " & .FileName
        End If
    End With
    '
    ' Validate the current file information
    frmArchive.Check_For_Valid_Info
End Sub
'
' EVENT:    btnExtract_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Extracts files from the tape device
' TRIGGER:  User clicked on the "Extract" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnExtract_Click()
    Dim sFile() As String
    '
    ' Indicate operation is starting
    Me.shpStatus3.FillColor = vbYellow
    '
    ' Extract the time header file
    ' Add ".hdr1" extension to the base file name
    sFile = Split(ParsePathStr(Me.txtArchiveName.Text & ".hdr1"), "/")
    '
    ' Display the filename without the path
    Me.lblTapeTime.Caption = "Archive Time stamp: " & sFile(UBound(sFile))
    '
    ' Color the label yellow to indicate that this is the one that is being processed
    Me.lblTapeTime.ForeColor = vbYellow
    '
    ' Update the form
    DoEvents
    '
    ' Extract the file with READTAPEFILE.  If the result is 0 then there was a problem
    If ReadTapeFile > 0 Then
'    If ReadTapeFile = 0 Then ' for testing
        '
        ' update the label with the file size
        Me.lblTapeTime.Caption = Me.lblTapeTime.Caption & " - " & FileLen(Me.txtArchiveName.Text & ".hdr1") & " bytes read"
        '
        ' Change the label color to green
        Me.lblTapeTime.ForeColor = &H8000&
        '
        ' Extract the message header file
        ' Add ".hdr2" extension to the base file name
        sFile = Split(ParsePathStr(Me.txtArchiveName.Text & ".hdr2"), "/")
        '
        ' Display the filename without the path
        Me.lblTapeJunk.Caption = "Archive Header file: " & sFile(UBound(sFile))
        '
        ' Color the label yellow
        Me.lblTapeJunk.ForeColor = vbYellow
        '
        ' Update the form
        DoEvents
        '
        ' Extract the file and check for errors
        If ReadTapeFile > 0 Then
'        If ReadTapeFile = 0 Then ' for testing
            '
            ' Add the file size to the label
            Me.lblTapeJunk.Caption = Me.lblTapeJunk.Caption & " - " & FileLen(Me.txtArchiveName.Text & ".hdr2") & " bytes read"
            '
            ' Color the label green
            Me.lblTapeJunk.ForeColor = &H8000&
            '
            ' Extract the data file
            ' Add a null string "" to the file to properly terminate the C string;
            ' otherwise PARSEPATHSTR will not read the correct string
            sFile = Split(ParsePathStr(Me.txtArchiveName.Text & ""), "/")
            '
            ' Display the filename without the path
            Me.lblTapeData.Caption = "Archive Data file: " & sFile(UBound(sFile))
            '
            ' Color the label yellow
            Me.lblTapeData.ForeColor = vbYellow
            '
            ' Update the form
            DoEvents
            '
            ' Read the archive file and check for errors
            If ReadTapeFile > 0 Then
'            If ReadTapeFile = 0 Then ' for testing
                '
                ' Add the file size to the label
                Me.lblTapeData.Caption = Me.lblTapeData.Caption & " - " & FileLen(Me.txtArchiveName.Text) & " bytes read"
                '
                ' Color the label green
                Me.lblTapeData.ForeColor = &H8000&
                '
                ' Pass the file name to the archive file name box
                Me.txtFile.Text = Me.txtArchiveName.Text
                '
                ' Color the progress light green
                Me.shpStatus3.FillColor = &H8000&
                '
                ' Update the form
                DoEvents
                '
                ' See if the rewind box is checked
                If Me.chkRewind.Value = vbChecked Then
                    '
                    ' Rewind and check for errors
                    If Rewind = 0 Then
                        MsgBox "Cannot rewind tape"
                    '
                    ' Eject the tape and check for errors
                    ElseIf EjectTape = 0 Then
                        MsgBox "Cannot eject tape"
                    Else
                        '
                        ' Tape is rewound and ejected, disable the buttons
                        Me.btnSave.Enabled = False
                    End If
                End If
                '
                ' Operation complete, disable the buttons
                Me.btnExtract.Enabled = False
                Me.chkRewind.Enabled = False
            Else
                '
                ' Indicate failure on the form
                Me.lblTapeData.Caption = Me.lblTapeData.Caption & " - FAILED"
                Me.lblTapeData.ForeColor = vbRed
                Me.shpStatus3.FillColor = vbRed
            End If
        Else
            '
            ' Indicate failure on the form
            Me.lblTapeJunk.Caption = Me.lblTapeJunk.Caption & " - FAILED"
            Me.lblTapeJunk.ForeColor = vbRed
            Me.shpStatus3.FillColor = vbRed
        End If
    Else
        '
        ' Indicate failure on the form
        Me.lblTapeTime.Caption = Me.lblTapeTime.Caption & " - FAILED"
        Me.lblTapeTime.ForeColor = vbRed
        Me.shpStatus3.FillColor = vbRed
    End If
End Sub
'
' EVENT:    btnFilter_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Filters and translates the specified archive file
' TRIGGER:  User clicked on the "Filter and Translate" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnFilter_Click()
    '
    ' Log the event
    basCCAT.WriteLogEntry ": ARCHIVE: btnFilter_Click : About to filter file " & frmArchive.txtFile.Text
    '
    ' Check for a file name
    If Len(frmArchive.txtFile.Text) > 0 Then
        '
        ' Check for the existence of the archive file, and prompt for the filtered file
        ' destination
        If (Dir(frmArchive.txtFile.Text) <> "") And Me.bSave_Filtered_File Then
            '
            ' Reset the data table
            If basDatabase.bReset_Archive_Processing Then
                '
                ' Make the Processing tab visible
                frmArchive.tabOptions.Tabs(TAB_FILTERED + 1).Selected = True
                '
                ' Set the limits of the progress bar
                frmArchive.barProgress.Min = 0
                frmArchive.barProgress.Value = 0
                '
                ' Set the max value to the number of simulated messages or the size of the file
                If frmMain.GetNumber("Miscellaneous Operations", "SIMULATE_PROCESSING", 0) > 0 Then
                    frmArchive.barProgress.Max = frmMain.GetNumber("Miscellaneous Operations", "SIMULATE_PROCESSING", 0)
                Else
                    frmArchive.barProgress.Max = FileLen(frmArchive.txtFile.Text)
                End If
                '
                ' Reset the form values for processing
                frmArchive.lblNumBytes.Caption = 0
                frmArchive.lblNumMsg.Caption = 0
                frmArchive.lblStart.Caption = ""
                frmArchive.lblEnd.Caption = ""
                frmArchive.lblProcessed.Caption = Now
                frmArchive.lblProcessInfo.Caption = "Filtering Archive..."
                '
                ' Set the initial start time to an impossible value, and the rest to 0
                '+v1.5
                'guCurrent.uArchive.dStart_Time = 400 * 86400#
                'guCurrent.uArchive.dEnd_Time = 0#
                '
                ' Use date/time values directly
                ' Adding 1 adds 1 day
                guCurrent.uArchive.dtStart_Time = guCurrent.uArchive.dtArchiveDate + 1
                guCurrent.uArchive.dtEnd_Time = 0#
                '-v1.5
                '
                guCurrent.uArchive.lNum_Messages = 0
                '
                ' Disable the buttons on the form
                frmArchive.btnBrowse.Enabled = False
                frmArchive.btnCancel.Enabled = False
                frmArchive.btnFilter.Enabled = False
                frmArchive.btnHelp.Enabled = False
                frmArchive.btnOK.Enabled = False
                frmArchive.btnProcess.Enabled = False
                frmArchive.btnRestore.Enabled = False
                frmArchive.tabOptions.Enabled = False
                '
                ' Log the start time for the filtering operation
                basCCAT.WriteLogEntry String(70, "*")
                basCCAT.WriteLogEntry "ARCHIVE: btnFilter_Click: Started filtering archive"
                basCCAT.WriteLogEntry String(70, "*")
                DoEvents
                '
                ' Call the appropriate routine
                If frmMain.GetNumber("Miscellaneous Operations", "SIMULATE_PROCESSING", 0) > 0 Then
                    '
                    ' Simulated data
                    frmArchive.Simulate_Processing
                Else
                    '
                    ' Actual archive file
                    ' Currently, this routine will call the translation routine upon
                    ' completion
                    Raw2filtraw.ProcRawMain frmArchive.txtFile.Text, frmArchive.txtFilterFile.Text
                End If
                '
                ' Save the filter file name (simply replacing the archive file extension with the filter file extension
                frmArchive.lblNumMsg.Caption = guCurrent.uArchive.lNum_Messages
                frmArchive.lblNumBytes.Caption = CLng(frmArchive.barProgress.Value)
                '
                ' Log the completion of the filtering and translation processes
                basCCAT.WriteLogEntry String(70, "*")
                basCCAT.WriteLogEntry "ARCHIVE: btnFilter_Click: Filtering and translation completed"
                basCCAT.WriteLogEntry "        Original File: " & FileLen(frmArchive.txtFile.Text)
                basCCAT.WriteLogEntry "        Filtered File: " & FileLen(frmArchive.txtFilterFile.Text)
                basCCAT.WriteLogEntry "        Messages Processed: " & frmArchive.lblNumMsg.Caption
                basCCAT.WriteLogEntry "        Bytes Processed: " & frmArchive.lblNumBytes.Caption
                basCCAT.WriteLogEntry String(70, "*")
                '
                ' Update the start/stop times on the form
                '+v1.5
                'frmArchive.lblStart.Caption = basCCAT.sTSecs_To_Human_Time(guCurrent.uArchive.dStart_Time)
                'frmArchive.lblEnd.Caption = basCCAT.sTSecs_To_Human_Time(guCurrent.uArchive.dEnd_Time)
                '
                ' Use date/time values directly
                ' Shorten the text to fit existing control sizes and to be compatible with old
                ' database schema.
                frmArchive.lblStart.Caption = Format(guCurrent.uArchive.dtStart_Time, "mm/dd/yyyy hh:nn:ss")
                frmArchive.lblEnd.Caption = Format(guCurrent.uArchive.dtEnd_Time, "mm/dd/yyyy hh:nn:ss")
                '-v1.5
                '
                ' Close the summary and data tables
                guCurrent.uArchive.rsData.Close
                guCurrent.uArchive.rsSummary.Close
                '
                ' See if the entire file was processed by comparing the number of bytes
                ' processed with the size of the filtered file
                If frmArchive.barProgress.Value = frmArchive.barProgress.Max Then
                    MsgBox "Filtering and Translation Complete" & vbCr & "Original File: " & FileLen(frmArchive.txtFile.Text) & " bytes" & vbCr & "Filtered file: " & FileLen(frmArchive.txtFilterFile.Text) & " bytes", vbOKOnly Or vbInformation, "Operation Complete"
                Else
                    On Error Resume Next
                    MsgBox "An Error Occurred During Processing." & vbCr & Format(CDbl(Val(frmArchive.lblNumBytes.Caption)) * 100# / CDbl(FileLen(frmArchive.txtFilterFile.Text)), 0#) & "% Complete", vbOKOnly Or vbExclamation, "Operation Incomplete"
                    If Err.Number > 0 Then
                        MsgBox "Error #" & Err.Number & " - " & Err.Description, vbOKOnly Or vbExclamation, Err.Source & " reported an error while processing"
                    End If
                    On Error GoTo 0
                End If
                '
                ' Re-enable form buttons
                frmArchive.tabOptions.Enabled = True
                frmArchive.btnBrowse.Enabled = True
                frmArchive.btnCancel.Enabled = True
                frmArchive.btnFilter.Enabled = True
                frmArchive.btnHelp.Enabled = True
                frmArchive.btnOK.Enabled = True
                'frmArchive.btnRestore.Enabled = True
                frmArchive.tabOptions.Enabled = True
                '
                ' Check for the filter file name
                If Len(frmArchive.txtFilterFile.Text) > 0 Then
                    '
                    ' Check for the existence of the specified file
                    If Dir(frmArchive.txtFilterFile.Text) <> "" Then
                        '
                        ' Enable the translation button
                        frmArchive.btnProcess.Enabled = True
                    Else
                        '
                        ' Try getting the filename
                        If Not Me.bSave_Filtered_File Then
                            ' Remove the bad file name and disable the translation button
                            frmArchive.txtFilterFile.Text = ""
                            frmArchive.btnProcess.Enabled = False
                        End If
                    End If
                End If
            Else
                '
                ' Could not reset the tables
                basCCAT.WriteLogEntry " ERROR!   Could not reset the database tables"
                MsgBox "Could not reset the database tables." & vbCr & "Archive could not be reprocessed", vbOKOnly Or vbInformation, "Filter Process Failed"
            End If ' Reset archive process
        Else
            If (Dir(frmArchive.txtFile.Text) = "") Then
                '
                ' Archive file did not exist
                basCCAT.WriteLogEntry " ERROR!   Could not find file " & frmArchive.txtFile.Text
                MsgBox "Could not find the file " & frmArchive.txtFile.Text, vbOKOnly Or vbInformation, "Filter Process Failed"
            End If
        End If ' Existence of specified archive file
    Else
        '
        ' No file name specified
        basCCAT.WriteLogEntry " ERROR!   No file to process!"
        MsgBox "No file specified!", vbOKOnly Or vbInformation, "Filter Process Failed"
    End If ' Presence of file name
End Sub
'
' FUNCTION: bSave_Filtered_File
' AUTHOR:   Tom Elkins
' PURPOSE:  Filters and translates the specified archive file
' TRIGGER:  User clicked on the "Filter and Translate" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Function bSave_Filtered_File() As Boolean
    '
    ' Use control-level addressing
    With frmMain.dlgCommonDialog
        '
        ' Set the file filter.
        ' Use the extensions specified in Brad's Raw2FiltRaw module.  If it changes
        ' there, the filter will also change to accomodate.
        .Filter = "Filtered Archive Files (*" & Raw2filtraw.sFilteredExt & ")|*" & Raw2filtraw.sFilteredExt
        '
        ' Set flags
        .Flags = cdlOFNPathMustExist Or cdlOFNCreatePrompt Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        '
        ' Set the dialog title
        .DialogTitle = "Select filtered file location"
        '
        ' Set the default file to the one specified in the archive record
        If Len(frmArchive.txtFilterFile.Text) > 0 Then
            .FileName = frmArchive.txtFilterFile.Text
        Else
            .FileName = Replace(frmArchive.txtFile.Text, Raw2filtraw.sArchiveExt, Raw2filtraw.sFilteredExt)
        End If
        '
        ' Save the old name
        .Tag = frmArchive.txtFilterFile.Text
        '
        '
        .CancelError = True
        '
        ' Allow the user to select a file
        On Error Resume Next
        .ShowSave
        '
        '
        If Err.Number = cdlCancel Then
            .FileName = .Tag
            .CancelError = False
        End If
        On Error GoTo 0
        bSave_Filtered_File = False
        '
        ' Save the filename if the user selected one
        If .FileName <> .Tag Then
            '
            ' Save the new file name
            frmArchive.txtFilterFile.Text = .FileName
            '
            ' Log the event
            basCCAT.WriteLogEntry "ARCHIVE: FILTER: User specified filter file " & .FileName
            '
            ' Return success
            bSave_Filtered_File = True
        End If
    End With
End Function
'
' EVENT:    btnHelp_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Displays help for the Archive form
' TRIGGER:  User clicked on the "Help" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnHelp_Click()
    '+v1.5
    ''
    '' Use control-level addressing
    'With frmArchive.dlgHelp
        ''
        '' Assign the help file
        '.HelpFile = App.Path & DAS_HELP_PATH & CCAT_HELP_FILE
        ''
        '' Configure to show the help topic
        '.HelpCommand = cdlHelpContext
        ''
        '' Assign the topic
        '.HelpContext = IDH_ArchiveRaw
        ''
        '' Show the help file
        '.ShowHelp
    'End With
    basCCAT.HtmlHelp Me.hWnd, App.HelpFile, basCCAT.HH_HELP_CONTEXT, basCCAT.lGetHelpID("Archive_Tab" & Me.tabOptions.SelectedItem.Index)
    '-v1.5
End Sub
'
' EVENT:    btnProcess_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets up and directs the processing of a filtered archive file
' TRIGGER:  User clicked on the "Translate" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnProcess_Click()
    Dim iProcess As Integer     ' Flag to continue processing
    '
    ' Check for a valid file
    If Len(frmArchive.txtFilterFile.Text) > 0 Then
        '
        ' Check for the existence of the file
        If Dir(frmArchive.txtFilterFile.Text) <> "" Then
            '
            ' See if the archive has been processed
            If frmArchive.lblProcessed.Caption <> "Never" Then
                '
                ' Warn the user that this is destructive
                iProcess = MsgBox("WARNING: This operation will delete all previous data for this archive." & vbCr & "Do you wish to continue?", vbYesNo Or vbExclamation, "Warning")
            Else
                iProcess = vbYes
            End If
            '
            ' See if the processing should continue
            If iProcess = vbYes And basDatabase.bReset_Archive_Processing Then
                '
                ' Set up the progress bar
                frmArchive.barProgress.Min = 0
                frmArchive.barProgress.Value = 0
                '
                ' Set the maximum to the number of simulated records or the size of the file
                If frmMain.GetNumber("Miscellaneous Operations", "SIMULATE_PROCESSING", 0) > 0 Then
                    frmArchive.barProgress.Max = frmMain.GetNumber("Miscellaneous Operations", "SIMULATE_PROCESSING", 0)
                Else
                    frmArchive.barProgress.Max = FileLen(frmArchive.txtFilterFile.Text)
                    guCurrent.uArchive.lFile_Size = FileLen(frmArchive.txtFilterFile.Text)
                End If
                '
                ' Reset the form values for processing
                frmArchive.lblNumBytes.Caption = 0
                frmArchive.lblNumMsg.Caption = 0
                frmArchive.lblStart.Caption = ""
                frmArchive.lblEnd.Caption = ""
                frmArchive.lblProcessed.Caption = Now
                frmArchive.lblProcessInfo.Caption = "Translating Filtered Archive - 0% Complete"
                '
                ' Set the initial start time to an impossible value, and the rest to 0
                '+v1.5
                'guCurrent.uArchive.dStart_Time = 400 * 86400#
                'guCurrent.uArchive.dEnd_Time = 0#
                '
                ' Use date/time values directly
                ' Adding 1 adds 1 day
                ' 0 = 12/30/1899
                guCurrent.uArchive.dtStart_Time = guCurrent.uArchive.dtArchiveDate + 1
                guCurrent.uArchive.dtEnd_Time = 0#
                '-v1.5
                '
                guCurrent.uArchive.lNum_Messages = 0
                '
                ' Disable the buttons on the form
                frmArchive.btnBrowse.Enabled = False
                frmArchive.btnCancel.Enabled = False
                frmArchive.btnFilter.Enabled = False
                frmArchive.btnHelp.Enabled = False
                frmArchive.btnOK.Enabled = False
                frmArchive.btnProcess.Enabled = False
                frmArchive.btnRestore.Enabled = False
                frmArchive.tabOptions.Enabled = False
                '+v1.5
                ' Enable the Stop button and set the processing flag
                frmArchive.btnStopTranslating.Enabled = True
                gbProcessing = True
                '-v1.5
                '
                ' Log the start time for the filtering operation
                basCCAT.WriteLogEntry String(70, "*")
                basCCAT.WriteLogEntry "ARCHIVE: btnProcess_Click: Started translating archive"
                basCCAT.WriteLogEntry String(70, "*")
                DoEvents
                '
                ' Call the appropriate routine
                If frmMain.GetNumber("Miscellaneous Operations", "SIMULATE_PROCESSING", 0) > 0 Then
                    '
                    ' Simulated data
                    frmArchive.Simulate_Processing
                    'frmArchive.Simulate_Processing2
                Else
                    '
                    ' Actual archive file
                    Filtraw2Das.ProcFiltMain frmArchive.txtFilterFile.Text
                End If
                '
                ' Suppress error reporting
                On Error Resume Next
                '
                ' Update the final numbers on the form
                frmArchive.lblNumMsg.Caption = guCurrent.uArchive.lNum_Messages
                frmArchive.lblNumBytes.Caption = frmArchive.barProgress.Value
                '
                '+v1.5
                'frmArchive.lblStart.Caption = basCCAT.sTSecs_To_Human_Time(guCurrent.uArchive.dStart_Time)
                'frmArchive.lblEnd.Caption = basCCAT.sTSecs_To_Human_Time(guCurrent.uArchive.dEnd_Time)
                ' Use date/time values directly
                ' Shorten the text to fit existing control sizes and to be compatible with
                ' old database schema
                frmArchive.lblStart.Caption = Format(guCurrent.uArchive.dtStart_Time, "mm/dd/yyyy hh:nn:ss")
                frmArchive.lblEnd.Caption = Format(guCurrent.uArchive.dtEnd_Time, "mm/dd/yyyy hh:nn:ss")
                '-v1.5
                '
                ' Log the completion of the filtering and translation processes
                basCCAT.WriteLogEntry String(70, "*")
                basCCAT.WriteLogEntry "ARCHIVE: btnProcess_Click: Translation completed"
                basCCAT.WriteLogEntry "        Messages Processed: " & frmArchive.lblNumMsg.Caption
                basCCAT.WriteLogEntry "        Bytes Processed: " & frmArchive.lblNumBytes.Caption
                basCCAT.WriteLogEntry String(70, "*")
                '
                ' Close the summary and data tables
                guCurrent.uArchive.rsData.Close
                guCurrent.uArchive.rsSummary.Close
                '
                ' See if the entire file was processed by comparing the number of bytes
                ' processed with the size of the filtered file
                If frmArchive.barProgress.Value = frmArchive.barProgress.Max Then
                    MsgBox "Translation Complete" & vbCr & frmArchive.lblNumMsg.Caption & " messages processed", vbOKOnly Or vbInformation, "Operation Complete"
                    frmArchive.lblProcessInfo.Caption = "Translation Complete"
                    frmArchive.btnOK.Enabled = True
                Else
                    MsgBox "An Error Occurred During Processing." & vbCr & FileLen(frmArchive.txtFilterFile.Text) - frmArchive.barProgress.Value & " bytes remaining in file", vbOKOnly Or vbExclamation, "Operation Incomplete"
                    frmArchive.lblProcessInfo.Caption = "Translation Incomplete"
                    frmArchive.btnOK.Enabled = False
                End If
                '
                ' Restore the form controls
                frmArchive.tabOptions.Enabled = True
                'frmArchive.btnRestore.Enabled = True
                frmArchive.btnCancel.Enabled = True
                frmArchive.btnBrowse.Enabled = True
                frmArchive.btnFilter.Enabled = True
                frmArchive.btnHelp.Enabled = True
                frmArchive.btnProcess.Enabled = True
                frmArchive.tabOptions.Enabled = True
                '+v1.5
                ' Disable the Stop button and reset the processing flag
                frmArchive.btnStopTranslating.Enabled = False
                gbProcessing = False
                '-v1.5
                '
                ' Check for errors
                If Err.Number <> NO_ERROR Then
                    basCCAT.WriteLogEntry "ARCHIVE: PROCESS: ERROR #" & Err.Number & vbCr & Err.Description & vbCr & Err.Source
                End If
            Else
                If iProcess = vbYes Then
                    '
                    ' The user elected to cancel the operation
                    basCCAT.WriteLogEntry "        User chose to cancel the operation"
                Else
                    '
                    ' Database could not be reset
                    basCCAT.WriteLogEntry "        Archive data could not be reset"
                    MsgBox "Archive data could not be reset in the database", vbOKOnly Or vbExclamation, "Translation Operation Failed"
                End If
            End If
        Else
            '
            ' The specified file could not be used
            basCCAT.WriteLogEntry "        ERROR -- could not find filtered file " & frmArchive.txtFilterFile.Text
            MsgBox "Error -- could not find the specified filtered file!" & vbCr & frmArchive.txtFilterFile.Text, , "File Not Found"
        End If
    Else
        '
        ' Inform the user there was no file specified
        MsgBox "Error -- No archive file to process", vbOKOnly, "No File Specified"
        basCCAT.WriteLogEntry "        ERROR -- No archive file to process!"
    End If
End Sub
'
' ROUTINE:  Simulate_Processing
' AUTHOR:   Tom Elkins
' PURPOSE:  Routine for testing data processing without using an archive
' INPUT:    None
' OUTPUT:   None
' NOTES:    Set the "SIMULATE_PROCESSING" token to the desired number of messages
Public Sub Simulate_Processing()
    Dim lMsg As Long            ' Message loop counter
    Dim iMsg_Type As Integer    ' Message type
    Dim lMax_Msg As Long        ' maximum number of messages
    Dim uDAS As DAS_MASTER_RECORD
    '
    ' Suppress error reporting
    On Error Resume Next
    '
    ' Get the maximum number of messages
    lMax_Msg = frmMain.GetNumber("Miscellaneous Operations", "SIMULATE_PROCESSING", 0)
    '
    ' Generate random simulated messages
    For lMsg = 1 To lMax_Msg
        '
        ' Pick a message at random
        Randomize
        iMsg_Type = Int((Rnd * frmMain.GetNumber("Message List", "CC_MESSAGES", 1)) + 1)
        '
        ' Populate fake data structure
        uDAS.dAltitude = Rnd * 80000#
        uDAS.dBearing = Rnd * 360#
        uDAS.dElevation = Rnd * 180# - 90#
        uDAS.dFrequency = Rnd * 2000000#
        uDAS.dHeading = Rnd * 360#
        uDAS.dLatitude = Rnd * 180# - 90#
        uDAS.dLongitude = Rnd * 360# - 180#
        uDAS.dPRI = Rnd * 100000#
        uDAS.dRange = Rnd * 10000#
        uDAS.dReport_Time = CDbl(Hour(Now) * 3600#) + CDbl(Minute(Now) * 60#) + CDbl(Second(Now))
        uDAS.dSpeed = Rnd * 1000#
        uDAS.dXX = Rnd * 100000#
        uDAS.dXY = Rnd * 100000#
        uDAS.dYY = Rnd * 100000#
        uDAS.lCommon_ID = Int(Rnd * 25)
        uDAS.lEmitter_ID = Int(Rnd * 5000)
        uDAS.lFlag = Int(Rnd * 10)
        uDAS.lIFF = Int(Rnd * 4)
        uDAS.lOrigin_ID = Int(Rnd * 1000)
        uDAS.lParent_ID = Int(Rnd * 1000)
        uDAS.lSignal_ID = Int(Rnd * 10000)
        uDAS.lStatus = Int(Rnd * 5)
        uDAS.lTag = Int(Rnd * 256)
        uDAS.lTarget_ID = Int(Rnd * 1000)
        uDAS.sReport_Type = basCCAT.gaDAS_Rec_Type(Int(Rnd * 4))
        uDAS.sMsg_Type = frmMain.GetAlias("Message List", "CC_MSG" & iMsg_Type, "MESSAGE#" & iMsg_Type)
        '
        ' Add record to the database
        basDatabase.Add_Data_Record frmMain.GetNumber("Message ID", uDAS.sMsg_Type & "ID", 0 - iMsg_Type), uDAS
    Next lMsg
End Sub
'
' ROUTINE:  Simulate_Processing2
' AUTHOR:   Tom Elkins
' PURPOSE:  Routine for testing database interaction
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Sub Simulate_Processing2()
    Dim lMsg As Long            ' Message loop counter
    Dim iMsg_Type As Integer    ' Message type
    Dim lMax_Msg As Long        ' maximum number of messages
    Dim uDAS As DAS_MASTER_RECORD
    '
    ' Suppress error reporting
    On Error Resume Next
    '
    ' Delete table
    If basDatabase.bTable_Exists(guCurrent.DB, "Archive411_Data") Then guCurrent.DB.TableDefs.Delete "Archive411_Data"
    '
    ' Add a fake data table
    If basDatabase.bCreate_Data_Table(guCurrent.DB, 411) Then
        '
        '
        Set guCurrent.uArchive.rsData = guCurrent.DB.OpenRecordset(guCurrent.uSQL.sTable, dbOpenDynaset, dbAppendOnly)
        '
        ' Get the maximum number of messages
        lMax_Msg = frmMain.GetNumber("Miscellaneous Operations", "SIMULATE_PROCESSING", 0)
        '
        '
        frmArchive.barProgress.Max = lMax_Msg
        frmArchive.barProgress.Value = frmArchive.barProgress.Min
        '
        '
        basCCAT.WriteLogEntry " Starting Simulated Data Processing"
        '
        ' Generate random simulated messages
        For lMsg = 1 To lMax_Msg
            '
            '
            frmArchive.barProgress.Value = lMsg
            '
            ' Pick a message at random
            Randomize
            iMsg_Type = Int((Rnd * frmMain.GetNumber("Message List", "CC_MESSAGES", 1)) + 1)
            '
            ' Populate data table
            With guCurrent.uArchive.rsData
                .AddNew
                !Altitude = Rnd * 80000#
                !Bearing = Rnd * 360#
                !Elevation = Rnd * 180# - 90#
                !Frequency = Rnd * 2000000#
                !Heading = Rnd * 360#
                !Latitude = Rnd * 180# - 90#
                !Longitude = Rnd * 360# - 180#
                !PRI = Rnd * 100000#
                !Range = Rnd * 10000#
                !Report_Time = CDbl(Hour(Now) * 3600#) + CDbl(Minute(Now) * 60#) + CDbl(Second(Now))
                !Speed = Rnd * 1000#
                !XX = Rnd * 100000#
                !XY = Rnd * 100000#
                !YY = Rnd * 100000#
                !Common_ID = Int(Rnd * 25)
                !Emitter_ID = Int(Rnd * 5000)
                !Flag = Int(Rnd * 10)
                !IFF = Int(Rnd * 4)
                !Origin_ID = Int(Rnd * 1000)
                !Parent_ID = Int(Rnd * 1000)
                !Signal_ID = Int(Rnd * 10000)
                !Status = Int(Rnd * 5)
                !Tag = Int(Rnd * 256)
                !Target_ID = Int(Rnd * 1000)
                !Rpt_Type = basCCAT.gaDAS_Rec_Type(Int(Rnd * 4))
                .Update
            End With
        Next lMsg
        '
        '
        basCCAT.WriteLogEntry " Completed Simulated Data Processing -- " & lMax_Msg & " messages"
    End If
End Sub
'
' EVENT:    btnRestore_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Restores values on the form from the database
' TRIGGER:  User clicked on the "Restore" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnRestore_Click()
    '
    ' Restore the values from the database
    frmArchive.Populate_Form_With_Archive_Data guCurrent.iArchive
End Sub
'
' EVENT:    btnCancel_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Indicates to the calling routine that the form was cancelled
' TRIGGER:  User clicked on the "Cancel" button
' INPUT:    None
' OUTPUT:   None
' NOTES:    The Tag property of the OK button stores whether the user clicked on the
'           OK button (TRUE) or the Cancel button (FALSE)
Private Sub btnCancel_Click()
    '
    ' Tell the interface that nothing was changed
    frmArchive.btnOK.Tag = False
    '
    ' Remove the form
    frmArchive.Hide
End Sub
'
' EVENT:    btnOK_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Indicates to the calling routine that the form was accepted
' TRIGGER:  User clicked on the "OK" button
' INPUT:    None
' OUTPUT:   None
' NOTES:    The Tag property of the OK button stores whether the user clicked on the
'           OK button (TRUE) or the Cancel button (FALSE)
Private Sub btnOK_Click()
    '
    ' Tell the interface to save changes
    frmArchive.btnOK.Tag = True
    '
    ' Remove the form
    frmArchive.Hide
End Sub
'
' EVENT:    btnSave_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Presents the user with a file selector box for the archive file
' TRIGGER:  The user clicked on the "Save To..." button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnSave_Click()
    '
    ' Highlight the status light
    Me.shpStatus2.FillColor = vbYellow
    '
    ' Hide the status light for the third step
    Me.shpStatus3.FillColor = vbButtonFace
    '
    ' Disable the third step functions
    Me.lblStep3.Enabled = False
    Me.btnExtract.Enabled = False
    '
    ' Set the default labels for the file names
    Me.lblTapeTime.Enabled = True
    Me.lblTapeTime.Caption = "Archive Time stamp: N/A"
    Me.lblTapeTime.ForeColor = vbBlack
    Me.lblTapeTime.Enabled = True
    Me.lblTapeJunk.Caption = "Archive Header file: N/A"
    Me.lblTapeJunk.ForeColor = vbBlack
    Me.lblTapeJunk.Enabled = False
    Me.lblTapeData.Caption = "Archive Data file: N/A"
    Me.lblTapeData.ForeColor = vbBlack
    Me.lblTapeData.Enabled = False
    '
    ' Trap an error
    On Error GoTo BadFile
    '
    ' Display the file dialog
    With Me.dlgHelp
        '
        ' Use the archive file extensions in Raw2FiltRaw
        .DefaultExt = Raw2filtraw.sArchiveExt
        .DialogTitle = "Select destination for archive file"
        .FileName = frmArchive.txtName.Text & .DefaultExt
        .Filter = "COMPASS CALL Archives (*" & Raw2filtraw.sArchiveExt & ")|*" & Raw2filtraw.sArchiveExt
        .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
        '
        ' Allow the user to select the destination location for the archive files
        .ShowSave
        '
        ' Check if the user entered a name
        If .FileName <> "" Then
            Me.txtArchiveName.Text = .FileName
            '
            ' Highlight the current step indicator
            Me.shpStatus2.FillColor = &H8000&
            '
            ' Set the defaults for step 3 functions
            Me.shpStatus3.FillColor = vbBlack
            Me.lblStep3.Enabled = True
            Me.chkRewind.Enabled = True
            Me.btnExtract.Enabled = True
            Me.lblTapeTime.Enabled = True
            Me.lblTapeTime.Caption = "Archive Time stamp: N/A"
            Me.lblTapeTime.ForeColor = vbBlack
            Me.lblTapeJunk.Enabled = True
            Me.lblTapeJunk.Caption = "Archive Header file: N/A"
            Me.lblTapeJunk.ForeColor = vbBlack
            Me.lblTapeData.Enabled = True
            Me.lblTapeData.Caption = "Archive Data file: N/A"
            Me.lblTapeData.ForeColor = vbBlack
        End If
    End With
    '
    ' Disable error trapping
    On Error GoTo 0
Exit Sub
'
' File error
BadFile:
    '
    ' Inform the user of the error
    MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "attempting to set output file", , "Error Setting Output File"
    '
    ' Reset the step indicator
    Me.shpStatus2.FillColor = vbRed
    '
    ' Reset the error trap
    On Error GoTo 0
End Sub
'
'+v1.5
' EVENT:    btnStopTranslating_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Resets the processing flag so the system will terminate the translation phase
' TRIGGER:  The user clicked on the "Stop" button
' INPUT:    None
' OUTPUT:   None
' NOTES:    The global flag is checked by the message processing routines. If the flag is
'           reset, the translation process will be prematurely terminated.
Private Sub btnStopTranslating_Click()
    gbProcessing = False
End Sub
'-v1.5
'
' EVENT:    btnTapeInit_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Attempts to initialize the tape drive to read files
' TRIGGER:  The user clicked on the "Initialize Tape" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnTapeInit_Click()
    '
    ' Highlight the step status
    Me.shpStatus1.FillColor = vbYellow
    '
    ' Disable and lock the subsequent functions
    Me.lblArchiveName.Enabled = False
    Me.txtArchiveName.Locked = True
    Me.btnSave.Enabled = False
    Me.lblInstruction.Enabled = False
    Me.shpStatus2.FillColor = vbButtonFace
    Me.shpStatus3.FillColor = vbButtonFace
    Me.lblStep3.Enabled = False
    Me.btnExtract.Enabled = False
    Me.lblTapeTime.Caption = "Archive Time stamp: N/A"
    Me.lblTapeTime.ForeColor = vbBlack
    Me.lblTapeTime.Enabled = False
    Me.lblTapeJunk.Caption = "Archive Header file: N/A"
    Me.lblTapeJunk.ForeColor = vbBlack
    Me.lblTapeJunk.Enabled = False
    Me.lblTapeData.Caption = "Archive Data file: N/A"
    Me.lblTapeData.ForeColor = vbBlack
    Me.lblTapeData.Enabled = False
    '
    ' Trap errors
    On Error Resume Next
    '
    ' Attempt to initialize the tape
    If InitializeTape > 0 Then
'    If InitializeTape = 0 Then ' for testing
        '
        ' Trap the "Missing DLL" error
        If Err.Number = 53 Then
            '
            ' Inform the user
            '+v1.5
            'MsgBox "Cannot find the 'ReadTape.dll' library." & vbCr & "Please copy that file to the '" & App.Path & "' directory and try again.", vbCritical Or vbMsgBoxHelpButton, "DLL Error", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, IDH_Error53
            MsgBox "Cannot find the 'ReadTape.dll' library." & vbCr & "Please copy that file to the '" & App.Path & "' directory and try again.", vbCritical Or vbMsgBoxHelpButton, "DLL Error", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, basCCAT.lGetHelpID("Error_53")
            '-v1.5
        Else
            '
            ' enable the next step
            Me.lblArchiveName.Enabled = True
            Me.btnSave.Enabled = True
            Me.lblInstruction.Enabled = True
            Me.shpStatus1.FillColor = &H8000&
            Me.shpStatus2.FillColor = vbBlack
        End If
    Else
        '
        ' Indicate failed status
        Me.shpStatus1.FillColor = vbRed
    End If
    '
    ' Disable error trapping
    On Error GoTo 0
End Sub
'
' EVENT:    cmbClass_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Retrieves the numeric code for the selected classification
' TRIGGER:  The value of the classification changed
' INPUT:    None
' OUTPUT:   None
' NOTES:    The classification bit values are defined in the token file
Private Sub cmbClass_Change()
    '
    ' Check if an item was selected
    If frmArchive.cmbClass.ListIndex > 0 Then
        '
        ' Get the classification number from the text
        frmArchive.cmbClass.Tag = frmSecurity.GetNumber("Security values from text", "SECURITY_VAL_" & frmArchive.cmbClass.Text, 0)
        '
        ' Store the classification
        guArchive.sClass = frmArchive.cmbClass.Text
        '
        ' See if all the information is available to process
        frmArchive.Check_For_Valid_Info
    Else
        '
        ' Reset the default
        frmArchive.cmbClass.ListIndex = 0
        guArchive.sClass = ""
    End If
End Sub
'
' EVENT:    cmbClass_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Retrieves the numeric code for the selected classification
' TRIGGER:  The user selected a classification level from the list
' INPUT:    None
' OUTPUT:   None
' NOTES:    The classification bit values are defined in the token file
Private Sub cmbClass_Click()
    '
    ' Check if an item was selected
    If frmArchive.cmbClass.ListIndex > 0 Then
        '
        ' Get the classification number from the text
        frmArchive.cmbClass.Tag = frmSecurity.GetNumber("Security Values from text", "SECURITY_VAL_" & frmArchive.cmbClass.Text, 0)
        '
        ' Store the classification
        guArchive.sClass = frmArchive.cmbClass.Text
    Else
        frmArchive.tabOptions.Tabs(TAB_RAW + 1).Selected = True
    End If
    frmArchive.Check_For_Valid_Info
End Sub
'
' EVENT:    dtDate_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Computes the TSecs equivalent for the specified date
' TRIGGER:  User changed the archive date
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub dtDate_Change()
    '
    ' Compute the julian day, subtract 1 (1 Jan is day 0), and multiply by 86400 seconds/day
    frmArchive.dtDate.Tag = (DatePart("y", frmArchive.dtDate.Value) - 1) * 86400#
    guCurrent.uArchive.dOffset_Time = (DatePart("y", frmArchive.dtDate.Value) - 1) * 86400#
    guArchive.sDate = Format(frmArchive.dtDate.Value, "mm/dd/yyyy")
    '
    '+v1.5 Store archive date
    ' Strip the time off of the date and store it for use in time calculations
    guCurrent.uArchive.dtArchiveDate = Int(frmArchive.dtDate.Value)
    '-v1.5
    frmArchive.Check_For_Valid_Info
End Sub
'
' EVENT:    Form_KeyDown
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Intercepts certain keystrokes to allow the user to move from one tab to another
' TRIGGER:  User pressed a key on the keyboard while the form is showing.
' INPUT:    "iKeyCode" is the internal code for the key the user pressed
'           "iKeyState" is the current state of the Shift, Alt, and Ctrl keys
' OUTPUT:   None
' NOTES:
Private Sub Form_KeyDown(iKeyCode As Integer, iKeyState As Integer)
    Dim iTab As Integer
    '
    ' handle ctrl+tab to move to the next tab
    If iKeyState = vbCtrlMask And iKeyCode = vbKeyTab Then
        '
        ' Get the current tab
        iTab = tabOptions.SelectedItem.Index
        '
        ' Check if this is the last tab
        If iTab = tabOptions.Tabs.Count Then
            '
            'last tab so we need to wrap to tab 1
            Set tabOptions.SelectedItem = tabOptions.Tabs(1)
        Else
            '
            'increment the tab
            Set tabOptions.SelectedItem = tabOptions.Tabs(iTab + 1)
        End If
    End If
    '
    '+v1.5
    ' Set the help context
    Me.HelpContextID = basCCAT.lGetHelpID("Archive_Tab" & tabOptions.SelectedItem.Index)
    '-v1.5
End Sub
'
' EVENT:    Form_Load
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Prepares the form to be displayed
' TRIGGER:  The first time a control on the form is accessed
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    Dim iItem As Integer        ' Current item retrieved from the token list
    Dim sToken As String        ' Token prefix for message names
    Dim itmMsg As ListItem      ' New message item
    '
    ' Center the form
    frmArchive.Move (Screen.Width - frmArchive.Width) / 2, (Screen.Height - frmArchive.Height) / 2
    '
    ' Hide the frame boundaries
    frmArchive.fraPage(TAB_TAPE).BorderStyle = 0
    frmArchive.fraPage(TAB_RAW).BorderStyle = 0
    frmArchive.fraPage(TAB_MESSAGE).BorderStyle = 0
    frmArchive.fraPage(TAB_FILTERED).BorderStyle = 0
    '
    ' Force the raw archive tab to be selected
    frmArchive.tabOptions.Tabs(TAB_RAW + 1).Selected = True
    '
    ' Set the default media option
    frmArchive.optMedia(OPT_TAPE).Enabled = False
    frmArchive.optMedia(OPT_HD).Value = True
    frmArchive.optMedia(OPT_CD).Enabled = False
    '
    ' Disable the help button
    frmArchive.btnHelp.Enabled = True
    '
    '+v1.5
    ' Disable the stop translation button
    frmArchive.btnStopTranslating.Enabled = False
    '-v1.5
    '
    ' Disable the tape controls
    frmArchive.lblArchiveName.Caption = ""
    frmArchive.txtArchiveName.Text = ""
    frmArchive.btnSave.Enabled = True
    '
    ' Load the classification combo list
    frmArchive.cmbClass.AddItem "<Select classification>"
    frmArchive.cmbClass.AddItem "UNCLASSIFIED"
    frmArchive.cmbClass.AddItem "UNCLASSIFIED/SAR"
    frmArchive.cmbClass.AddItem "CONFIDENTIAL"
    frmArchive.cmbClass.AddItem "CONFIDENTIAL/SAR"
    frmArchive.cmbClass.AddItem "SECRET"
    frmArchive.cmbClass.AddItem "SECRET/SAR"
    frmArchive.cmbClass.AddItem "SECRET/SCI"
    frmArchive.cmbClass.AddItem "SECRET/SAR/SCI"
    frmArchive.cmbClass.AddItem "TOP SECRET"
    frmArchive.cmbClass.AddItem "TOP SECRET/SAR"
    frmArchive.cmbClass.AddItem "TOP SECRET/SCI"
    frmArchive.cmbClass.AddItem "TOP SECRET/SAR/SCI"
    '
    ' Load the message list
    For iItem = 1 To frmMain.GetNumber("Message List", "CC_MESSAGES", 0)
        '
        ' Get the message name and ID from the token file
        Set itmMsg = frmArchive.lvMessages.ListItems.Add(, "MSG" & frmMain.GetNumber("Message ID", frmMain.GetAlias("Message List", "CC_MSG" & iItem, "Message#" & iItem) & "ID", 0 - iItem), frmMain.GetAlias("Message List", "CC_MSG" & iItem, "Message" & iItem))
        itmMsg.Checked = True
        itmMsg.SubItems(1) = frmMain.GetNumber("Message ID", itmMsg.Text & "ID", 0 - iItem)
        itmMsg.SubItems(2) = frmMain.GetAlias("Message Descriptions", "CC_MSG_DESC" & itmMsg.SubItems(1), "UNKNOWN MESSAGE TYPE")
    Next iItem
    '
    ' Set the form controls
    frmArchive.lblProcessInfo.Caption = "Press TRANSLATE to start processing"
End Sub
'
' EVENT:    optMedia_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Adjust controls depending on which option is active
' TRIGGER:  User clicked on a Media type option button
'           A routine sets optMedia(?).Value to TRUE
' INPUT:    "iOption" is the index of the selected option button
' OUTPUT:   None
' NOTES:
Private Sub optMedia_Click(iOption As Integer)
    '
    ' Action is based on option
    Select Case iOption
        '
        ' Tape
        Case OPT_TAPE:
            '
            ' Disable the browse button
            frmArchive.btnBrowse.Enabled = False
            frmArchive.fraPage(TAB_TAPE).Enabled = True
        '
        ' Any other option
        Case Else:
            '
            ' Enable the browse button
            frmArchive.btnBrowse.Enabled = True
            frmArchive.fraPage(TAB_TAPE).Enabled = True
    End Select
End Sub
'
' EVENT:    tabOptions_Click
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Change the contents of the form based on the tab the user selected
' TRIGGER:  User clicked on a tab
' INPUT:    None
' OUTPUT:   None
' NOTES:    Tabs actually do nothing.  You put a container array on the form,
'           and put controls in each container.  When a tab is selected, its
'           index is used to make the container of the same index visible, and make
'           the other containers invisible.
Private Sub tabOptions_Click()
    Dim iTab As Integer     ' Current tab index
    '
    ' Loop through the tabs
    For iTab = 0 To tabOptions.Tabs.Count - 1
        '
        ' See if this is the selected tab
        If iTab = tabOptions.SelectedItem.Index - 1 Then
            '
            ' Make the controls for the tab visible
            frmArchive.fraPage(iTab).Left = 200
            DoEvents
        Else
            '
            ' Hide the controls for the other tabs
            frmArchive.fraPage(iTab).Left = -20000
        End If
    Next iTab
    '
    ' Special cases
    If tabOptions.SelectedItem.Index - 1 = TAB_TAPE Then
        '
        ' Disable all controls, and force user to walk through the steps
        frmArchive.shpStatus1.FillColor = vbBlack
        frmArchive.shpStatus2.FillColor = vbButtonFace
        frmArchive.lblArchiveName.Enabled = False
        frmArchive.lblInstruction.Enabled = False
        frmArchive.btnSave.Enabled = False
        frmArchive.shpStatus3.FillColor = vbButtonFace
        frmArchive.lblStep3.Enabled = False
        frmArchive.chkRewind.Enabled = False
        frmArchive.btnExtract.Enabled = False
        frmArchive.lblTapeTime.Enabled = False
        frmArchive.lblTapeJunk.Enabled = False
        frmArchive.lblTapeData.Enabled = False
    End If
End Sub
'
' ROUTINE:  Populate_Form_With_Archive_Data
' AUTHOR:   Tom Elkins
' PURPOSE:  Take the data values from the specified archive record and populate the
'           controls on the form
' INPUT:    "iArchive" is the archive record ID
' OUTPUT:   None
' NOTES:
Public Sub Populate_Form_With_Archive_Data(iArchive As Integer)
    Dim rsArchive As Recordset
    '
    ' Check for a valid archive ID
    If iArchive = 0 Then iArchive = guCurrent.iArchive
    '
    ' Log the event
    basCCAT.WriteLogEntry "ARCHIVE: Populate_Form_With_Archive_Data: Archive " & iArchive
    '
    ' Get the record for the specified archive
    Set rsArchive = guCurrent.DB.OpenRecordset("SELECT * FROM Archives WHERE ID = " & iArchive)
    '
    ' Check for data
    If rsArchive Is Nothing Then
        '
        ' Log the error
        basCCAT.WriteLogEntry "ARCHIVE: Populate_Form_With_Archive_Data: ERROR -- could not find the record for archive #" & iArchive
        '
        ' Something went wrong, inform the user
        MsgBox "Error -- could not find the record for archive #" & iArchive, , "Error Finding Record"
    Else
        '
        ' Display the archive end time (if available)
        If IsNull(rsArchive!End) Then
            frmArchive.lblEnd.Caption = ""
        Else
            frmArchive.lblEnd.Caption = rsArchive!End
        End If
        '
        ' Display the archive ID
        frmArchive.lblID.Caption = rsArchive!ID
        '
        ' Display the archive classification
        ' Map the security values to the combo box entry indices
        Select Case guCurrent.DB.Properties("Security").Value
            Case 0: frmArchive.cmbClass.ListIndex = 1
            Case 1: frmArchive.cmbClass.ListIndex = 2
            Case 4: frmArchive.cmbClass.ListIndex = 3
            Case 5: frmArchive.cmbClass.ListIndex = 4
            Case 8, 12: frmArchive.cmbClass.ListIndex = 5
            Case 9, 13: frmArchive.cmbClass.ListIndex = 6
            Case 10, 14: frmArchive.cmbClass.ListIndex = 7
            Case 11, 15: frmArchive.cmbClass.ListIndex = 8
            Case 16, 20, 24, 28: frmArchive.cmbClass.ListIndex = 9
            Case 17, 21, 25, 29: frmArchive.cmbClass.ListIndex = 10
            Case 18, 22, 26, 30: frmArchive.cmbClass.ListIndex = 11
            Case 19, 23, 27, 31: frmArchive.cmbClass.ListIndex = 12
            Case Else: frmArchive.cmbClass.ListIndex = 0
        End Select
        '
        ' Display the date the archive was last processed
        If IsNull(rsArchive!Processed) Then
            frmArchive.lblProcessed.Caption = "Never"
        Else
            frmArchive.lblProcessed.Caption = rsArchive!Processed
        End If
        '
        ' Display the archive start time
        If IsNull(rsArchive!Start) Then
            frmArchive.lblStart.Caption = ""
        Else
            frmArchive.lblStart.Caption = rsArchive!Start
        End If
        '
        ' Display the archive media type
        Select Case rsArchive!Media
            Case "TAPE"
                frmArchive.optMedia(OPT_TAPE).Value = True
            Case "HD"
                frmArchive.optMedia(OPT_HD).Value = True
            Case "CD", "DVD"
                frmArchive.optMedia(OPT_CD).Value = True
        End Select
        '
        ' Display the original archive file name
        If IsNull(rsArchive!Archive) Or rsArchive!Archive = "" Then
            frmArchive.txtFile.Text = ""
            frmArchive.btnFilter.Enabled = False
            frmArchive.dtDate.Enabled = False
        Else
            frmArchive.txtFile.Text = rsArchive!Archive
            frmArchive.btnFilter.Enabled = True
            frmArchive.cmbClass.Enabled = True
            frmArchive.dtDate.Enabled = True
        End If
        '
        ' Display the archive date
        If IsNull(rsArchive!Date) Then
            '
            ' No date value, so check for a file
            If frmArchive.txtFile.Text = "" Then
                '
                ' No archive file selected, so
                ' use the creation data of the database
                frmArchive.dtDate.Value = FileDateTime(guCurrent.sName)
            Else
                '
                ' Use the file date for the archive
                frmArchive.dtDate.Value = FileDateTime(frmArchive.txtFile.Text)
            End If
        Else
            '
            ' Use the stored value
            frmArchive.dtDate.Value = rsArchive!Date
        End If
        Call dtDate_Change
        '
        ' Display the name of the filtered file
        If IsNull(rsArchive!Analysis_File) Or rsArchive!Analysis_File = "" Then
            frmArchive.txtFilterFile.Text = ""
            frmArchive.fraPage(TAB_FILTERED).Enabled = False
            frmArchive.btnProcess.Enabled = False
            frmArchive.lblProcessInfo.Caption = ""
        Else
            frmArchive.txtFilterFile.Text = rsArchive!Analysis_File
            frmArchive.fraPage(TAB_FILTERED).Enabled = True
            frmArchive.btnProcess.Enabled = True
            frmArchive.lblProcessInfo.Caption = "Press TRANSLATE to process the file"
            frmArchive.cmbClass.Enabled = True
            frmArchive.dtDate.Enabled = True
        End If
        '
        ' Display the user-specified archive name
        If IsNull(rsArchive!Name) Then
            frmArchive.txtName.Text = ""
        Else
            frmArchive.txtName.Text = rsArchive!Name
        End If
        '
        ' Display the number of bytes processed and message count
        frmArchive.lblNumBytes.Caption = rsArchive!Bytes
        frmArchive.lblNumMsg.Caption = rsArchive!Messages
        '
        ' Close the record set
        rsArchive.Close
        '
        ' Reset the progress bar
        frmArchive.barProgress.Value = frmArchive.barProgress.Min
    End If
End Sub
'
' ROUTINE:  Save_Archive_Data
' AUTHOR:   Tom Elkins
' PURPOSE:  Take the values from the controls on the form and update the values in
'           the database
' INPUT:    "iArchive" is the archive record ID
' OUTPUT:   None
' NOTES:
Public Sub Save_Archive_Data(iArchive As Integer)
    Dim rsArchive As Recordset
    '
    '
    On Error GoTo ERR_HANDLER
    '
    ' Check for valid ID
    If iArchive = 0 Then iArchive = guCurrent.iArchive
    '
    ' Log the event
    basCCAT.WriteLogEntry "ARCHIVE: Save_Archive_Data: Archive " & guCurrent.iArchive
    '
    ' Get the record for the specified archive
    Set rsArchive = guCurrent.DB.OpenRecordset("SELECT * FROM Archives WHERE ID = " & guCurrent.iArchive)
    '
    ' Check for data
    If rsArchive Is Nothing Then
        '
        ' Log the error
        basCCAT.WriteLogEntry "          ERROR -- could not find the record for archive #" & guCurrent.iArchive
        '
        ' Something went wrong, inform the user
        MsgBox "Error -- could not find the record for archive #" & guCurrent.iArchive, , "Error Finding Record"
    Else
        '
        ' Prepare the record for changes
        rsArchive.Edit
        '
        ' Check for values and save
        If IsNull(rsArchive!End) And frmArchive.lblEnd.Caption <> "" Then rsArchive!End = frmArchive.lblEnd.Caption
        If frmArchive.lblEnd.Caption <> rsArchive!End Then rsArchive!End = frmArchive.lblEnd.Caption
        If frmArchive.lblProcessed.Caption <> rsArchive!Processed Then rsArchive!Processed = Format(frmArchive.lblProcessed.Caption, "mm/dd/yyyy hh:nn:ss")
        If Val(frmArchive.cmbClass.Tag) <> guCurrent.DB.Properties("Security").Value Then
            frmSecurity.Set_Classification frmArchive.cmbClass.Tag, gsARCHIVE
            guCurrent.DB.Properties("Security").Value = frmSecurity.Tag
        End If
        If IsNull(rsArchive!Date) Then rsArchive!Date = frmArchive.dtDate.Value
        If frmArchive.dtDate.Value <> rsArchive!Date Then
'            '+v1.5
'            If MsgBox("Archive date stamp has changed from " & rsArchive!Date & " to " & frmArchive.dtDate.Value & "." & vbCr & "Do you want to update the data records to the new date?", vbYesNo Or vbQuestion, "Archive Date Changed") = vbYes Then
'                basDatabase.ExecuteSQLAction "UPDATE Archive" & guCurrent.iArchive & "_Data SET Report_Time = Report_Time + " & CStr(Int(frmArchive.dtDate.Value) - Int(CDate(rsArchive!Date)))
'            End If
            rsArchive!Date = Int(frmArchive.dtDate.Value)
        End If
        '-v1.5
        If IsNull(rsArchive!Start) And frmArchive.lblStart.Caption <> "" Then rsArchive!Start = frmArchive.lblStart.Caption
        If frmArchive.lblStart.Caption <> rsArchive!Start Then rsArchive!Start = frmArchive.lblStart.Caption
        If IsNull(rsArchive!Archive) And frmArchive.txtFile.Text <> "" Then rsArchive!Archive = frmArchive.txtFile.Text
        If frmArchive.txtFile.Text <> rsArchive!Archive Then rsArchive!Archive = frmArchive.txtFile.Text
        If IsNull(rsArchive!Analysis_File) And frmArchive.txtFilterFile.Text <> "" Then rsArchive!Analysis_File = frmArchive.txtFilterFile.Text
        If frmArchive.txtFilterFile.Text <> rsArchive!Analysis_File Then rsArchive!Analysis_File = frmArchive.txtFilterFile.Text
        If IsNull(rsArchive!Name) And frmArchive.txtName.Text <> "" Then rsArchive!Name = frmArchive.txtName.Text
        '+v1.5.1 NEW
        'If frmArchive.txtName.Text <> rsArchive!Name Then rsArchive!Name = frmArchive.txtName.Text
        If frmArchive.txtName.Text <> rsArchive!Name Then
            guCurrent.DB.TableDefs(rsArchive!Name & basDatabase.TBL_SUMMARY).Name = frmArchive.txtName.Text & basDatabase.TBL_SUMMARY
            guCurrent.DB.TableDefs(rsArchive!Name & basDatabase.TBL_DATA).Name = frmArchive.txtName.Text & basDatabase.TBL_DATA
            frmMain.tvTreeView.Nodes(guCurrent.DB.Name & basDatabase.SEP_ARCHIVE & guCurrent.iArchive).Text = frmArchive.txtName.Text
            rsArchive!Name = frmArchive.txtName.Text
        End If
        '-v1.5.1 NEW
        If frmArchive.optMedia(OPT_TAPE).Value Then rsArchive!Media = "TAPE"
        If frmArchive.optMedia(OPT_HD).Value Then rsArchive!Media = "HD"
        If frmArchive.optMedia(OPT_CD).Value Then rsArchive!Media = "CD"
        rsArchive!Messages = CLng(Val(frmArchive.lblNumMsg.Caption))
        rsArchive!Bytes = CLng(Val(frmArchive.lblNumBytes.Caption))
        rsArchive.Update
        '
        ' Close the recordset
        rsArchive.Close
    End If
    Exit Sub

ERR_HANDLER:
    MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "while saving archive data.", vbOKOnly, "Operation Terminated"
    On Error GoTo 0
End Sub
'
' ROUTINE:  Edit_Archive_Properties
' AUTHOR:   Tom Elkins
' PURPOSE:  Allow the user to interact with the specified archive.
' INPUT:    "sID" is the key value for an Archive object
' OUTPUT:   None
' NOTES:
'Public Sub Edit_Archive_Properties(sID As String)
Public Function bEdit_Archive_Properties(sID As String) As Boolean
    Dim rsArchive As Recordset
    '
    ' Trap errors
    On Error GoTo ERR_HANDLER
    '
    ' Log the event
    basCCAT.WriteLogEntry "ARCHIVE: Edit_Archive_Properties: " & sID
    '
    ' Force the first tab to be selected
    frmArchive.tabOptions.Tabs(TAB_RAW + 1).Selected = True
    '
    ' Get the archive index
    guCurrent.iArchive = basCCAT.iExtract_ArchiveID(sID)
    '
    ' Populate the form controls with database values
    frmArchive.Populate_Form_With_Archive_Data guCurrent.iArchive
    '
    ' Reset the save indicator
    frmArchive.btnOK.Tag = False
    '
    ' Display the form
    frmArchive.Show vbModal
    '
    ' Check the OK button Tag property
    If frmArchive.btnOK.Tag Then
        '
        ' Save the changes to the database
        frmArchive.Save_Archive_Data guCurrent.iArchive
        '
        ' See if the Archive node exists
        If frmMain.bNode_Exists(sID) Then
            '
            ' Remove and replace the archive node on the tree
            frmMain.tvTreeView.Nodes.Remove sID
            frmMain.RefreshDisplay
'+v1.5
'            basDatabase.Add_Archive_Node guCurrent.DB.OpenRecordset("SELECT * FROM Archives WHERE ID = " & guCurrent.iArchive)
'            '
'            ' Find the archive item in the List if it exists
'            If Not frmMain.lvListView.FindItem(frmMain.tvTreeView.Nodes(sID).Text) Is Nothing Then
'                '
'                ' Item exists, update its name
'                frmMain.lvListView.ListItems(sID).Text = frmArchive.txtName.Text
'                '
'                ' Update its icon
'                If frmArchive.optMedia(OPT_TAPE).Value Then
'                    frmMain.lvListView.ListItems(sID).Icon = "TAPE"
'                    frmMain.lvListView.ListItems(sID).SmallIcon = "TAPE_CLOSED"
'                ElseIf frmArchive.optMedia(OPT_HD).Value Then
'                    frmMain.lvListView.ListItems(sID).Icon = "HD"
'                    frmMain.lvListView.ListItems(sID).SmallIcon = "HD_CLOSED"
'                Else
'                    frmMain.lvListView.ListItems(sID).Icon = "CD"
'                    frmMain.lvListView.ListItems(sID).SmallIcon = "CD_OPEN"
'                End If
'            End If
'            '
'            ' Update the archive node name.
'            frmMain.tvTreeView.Nodes(sID).Text = frmArchive.txtName.Text
'            '
'            ' Update the List View caption
'            frmMain.lblTitle(1).Caption = frmMain.tvTreeView.SelectedItem.FullPath
'            '
'            ' Update the node's icons
'            If frmArchive.optMedia(OPT_TAPE).Value Then
'                frmMain.tvTreeView.Nodes(sID).Image = "TAPE_CLOSED"
'                frmMain.tvTreeView.Nodes(sID).SelectedImage = "TAPE_OPEN"
'            ElseIf frmArchive.optMedia(OPT_HD).Value Then
'                frmMain.tvTreeView.Nodes(sID).Image = "HD_CLOSED"
'                frmMain.tvTreeView.Nodes(sID).SelectedImage = "HD_OPEN"
'            Else
'                frmMain.tvTreeView.Nodes(sID).Image = "CD_CLOSED"
'                frmMain.tvTreeView.Nodes(sID).SelectedImage = "CD_OPEN"
'            End If
'-v1.5
        Else
            '
            ' Node did not exist, add a new one
            basDatabase.Add_Archive_Node guCurrent.DB.OpenRecordset("SELECT * FROM Archives WHERE ID = " & guCurrent.iArchive)
        End If
    End If
    '
    ' Set the return property
    bEdit_Archive_Properties = frmArchive.btnOK.Tag
    '
    Exit Function
'
ERR_HANDLER:
    MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "While editing archive properties", vbOKOnly, "Operation Terminated"
    On Error GoTo 0
End Function
'
' ROUTINE:  Add_Archive
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds a new archive to the database
' INPUT:    None, the affected database is specified in guCurrent.DB
' OUTPUT:   None
' NOTES:
Public Sub Add_Archive()
    Dim rsArchive As Recordset  ' New archive record
    Dim iArchive As Integer     ' New archive identifier
    Dim nodArchive As Node      ' New archive node
    '
    ' Attempt to add a new archive to the database
    If basDatabase.bAdd_Archive_Record Then
        '
        ' Allow the user to enter information
        If frmArchive.bEdit_Archive_Properties(guCurrent.sName & SEP_ARCHIVE & guCurrent.iArchive) Then
            '
            ' Add a new tree view node
            basDatabase.Add_Archive_Node guCurrent.DB.OpenRecordset("SELECT * FROM Archives WHERE ID = " & guCurrent.iArchive)
        Else
            basCCAT.WriteLogEntry " ARCHIVE: ADD_ARCHIVE: New archive canceled by user"
            '
            ' Delete the archive
            basDatabase.Delete_Archive guCurrent.sName & SEP_ARCHIVE & guCurrent.iArchive
        End If
    Else
        MsgBox "Could not add archive", vbOKOnly, "Add Archive Failed"
    End If
End Sub
'
' ROUTINE:  Check_For_Valid_Info
' AUTHOR:   Tom Elkins
' PURPOSE:  Determine if the user entered enough information to continue processing
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Sub Check_For_Valid_Info()
    Dim bProceed As Boolean
    '
    ' Default is to not proceed
    bProceed = False
    '
    ' Check for file name
    If frmArchive.txtFile.Text <> "" Then
        frmArchive.dtDate.Enabled = True
        frmArchive.cmbClass.Enabled = True
        bProceed = True
    End If
    '
    ' Check for classification
    bProceed = (frmArchive.cmbClass.ListIndex > 0)
    '
    ' Enable the processing tab
    frmArchive.fraPage(TAB_FILTERED).Enabled = bProceed
    '
    ' Enable the filtering button
    frmArchive.btnFilter.Enabled = bProceed And (guArchive.iType <> FILE_FILTERED)
    '
    ' See if the selected file is already filtered
    If bProceed And (guArchive.iType = FILE_FILTERED) Then
        '
        ' Set the filtered file name
        frmArchive.txtFilterFile.Text = frmArchive.txtFile.Text
        '
        ' Remove the file name from the archive field
        frmArchive.txtFile.Text = ""
        '
        ' Bring up the Filtered file options
        frmArchive.tabOptions.Tabs(TAB_FILTERED + 1).Selected = True
        '
        ' Enable the translate button
        frmArchive.btnProcess.Enabled = True
    End If
End Sub
