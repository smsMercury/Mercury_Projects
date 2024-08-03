VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWizard 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compass Call Archive Wizard"
   ClientHeight    =   5250
   ClientLeft      =   1965
   ClientTop       =   1815
   ClientWidth     =   7155
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   Tag             =   "10"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Introduction Screen"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   0
      Left            =   -10000
      TabIndex        =   6
      Tag             =   "1000"
      Top             =   0
      Width           =   7155
      Begin VB.CheckBox chkHideIntro 
         Caption         =   "Don't show this intro page again"
         Height          =   315
         Left            =   2700
         MaskColor       =   &H00000000&
         TabIndex        =   16
         Tag             =   "1002"
         Top             =   3000
         Width           =   3810
      End
      Begin VB.Label lblInstructions 
         Caption         =   $"frmWizard.frx":0000
         Height          =   915
         Index           =   0
         Left            =   2700
         TabIndex        =   18
         Top             =   1350
         Width           =   4110
      End
      Begin VB.Label lblStep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the COMPASS CALL Archive Wizard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   0
         Left            =   2700
         TabIndex        =   7
         Tag             =   "1001"
         Top             =   210
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2040
         Index           =   0
         Left            =   210
         Picture         =   "frmWizard.frx":009E
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2115
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   2
      Left            =   -10000
      TabIndex        =   10
      Tag             =   "2002"
      Top             =   0
      Width           =   7155
      Begin VB.ComboBox cmbArchiveClass 
         Height          =   315
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   3105
         Width           =   2900
      End
      Begin VB.ComboBox cmbArchiveVersion 
         Height          =   315
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2610
         Width           =   2900
      End
      Begin VB.TextBox txtArchiveName 
         Height          =   330
         Left            =   3780
         TabIndex        =   34
         Text            =   "ArchiveName"
         Top             =   1620
         Width           =   2900
      End
      Begin MSComCtl2.DTPicker dtpArchiveDate 
         Height          =   330
         Left            =   3780
         TabIndex        =   35
         Top             =   2115
         Width           =   2900
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Format          =   61931520
         CurrentDate     =   37118
      End
      Begin VB.Label lblArcInfo 
         AutoSize        =   -1  'True
         Caption         =   "Classification"
         Height          =   195
         Index           =   3
         Left            =   2700
         TabIndex        =   33
         Top             =   3150
         Width           =   930
      End
      Begin VB.Label lblArcInfo 
         AutoSize        =   -1  'True
         Caption         =   "CCOS Version"
         Height          =   195
         Index           =   2
         Left            =   2700
         TabIndex        =   32
         Top             =   2655
         Width           =   990
      End
      Begin VB.Label lblArcInfo 
         AutoSize        =   -1  'True
         Caption         =   "Mission Date"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   31
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label lblArcInfo 
         AutoSize        =   -1  'True
         Caption         =   "Archive Name"
         Height          =   195
         Index           =   0
         Left            =   2700
         TabIndex        =   30
         Top             =   1665
         Width           =   990
      End
      Begin VB.Label lblInstructions 
         Caption         =   $"frmWizard.frx":D681
         Height          =   690
         Index           =   2
         Left            =   2700
         TabIndex        =   29
         Top             =   630
         Width           =   4155
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2 - Archive Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   2700
         TabIndex        =   11
         Tag             =   "2003"
         Top             =   210
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2220
         Index           =   2
         Left            =   210
         Picture         =   "frmWizard.frx":D720
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   3
      Left            =   -10000
      TabIndex        =   12
      Tag             =   "2004"
      Top             =   0
      Width           =   7155
      Begin VB.CommandButton btnMsgSel 
         Caption         =   "Invert Selection"
         Height          =   375
         Index           =   2
         Left            =   990
         TabIndex        =   41
         Top             =   3915
         Width           =   1320
      End
      Begin VB.CommandButton btnMsgSel 
         Caption         =   "Select None"
         Height          =   375
         Index           =   1
         Left            =   990
         TabIndex        =   40
         Top             =   3420
         Width           =   1320
      End
      Begin VB.CommandButton btnMsgSel 
         Caption         =   "Select All"
         Height          =   375
         Index           =   0
         Left            =   990
         TabIndex        =   39
         Top             =   2925
         Width           =   1320
      End
      Begin MSComctlLib.ListView lvMessages 
         Height          =   3030
         Left            =   2475
         TabIndex        =   46
         ToolTipText     =   "Select/Deselect messages to translate"
         Top             =   1350
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   5345
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
      Begin VB.Label lblInstructions 
         Caption         =   $"frmWizard.frx":192BD
         Height          =   690
         Index           =   3
         Left            =   2700
         TabIndex        =   38
         Top             =   630
         Width           =   4155
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3 - Message Selection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   2700
         TabIndex        =   13
         Tag             =   "2005"
         Top             =   210
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2580
         Index           =   3
         Left            =   210
         Picture         =   "frmWizard.frx":19350
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Finished!"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Index           =   4
      Left            =   -20000
      TabIndex        =   14
      Tag             =   "3000"
      Top             =   0
      Width           =   7155
      Begin MSComctlLib.ProgressBar barProgress 
         Height          =   375
         Left            =   45
         TabIndex        =   44
         Top             =   4005
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.ListBox lstToDo 
         Height          =   2310
         Left            =   2880
         Style           =   1  'Checkbox
         TabIndex        =   43
         Top             =   1395
         Width           =   4110
      End
      Begin VB.Label lblPctDone 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "% Complete"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   90
         TabIndex        =   45
         Top             =   3780
         Width           =   6945
      End
      Begin VB.Label lblInstructions 
         Caption         =   $"frmWizard.frx":1E13A
         Height          =   825
         Index           =   4
         Left            =   2880
         TabIndex        =   42
         Top             =   585
         Width           =   4155
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ready To Process The Archive!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   3105
         TabIndex        =   15
         Tag             =   "3001"
         Top             =   210
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   3075
         Index           =   5
         Left            =   210
         Picture         =   "frmWizard.frx":1E209
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2430
      End
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   4680
      Width           =   7155
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         Height          =   312
         Index           =   4
         Left            =   5910
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         Height          =   312
         Index           =   3
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         Height          =   312
         Index           =   2
         Left            =   3435
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "102"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   312
         Index           =   1
         Left            =   2250
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "101"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Help"
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "100"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   108
         X2              =   7012
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   108
         X2              =   7012
         Y1              =   24
         Y2              =   24
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Step 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Tag             =   "2000"
      Top             =   0
      Width           =   7155
      Begin VB.TextBox txtFileName 
         Height          =   330
         Left            =   3060
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   4110
         Width           =   3885
      End
      Begin VB.CommandButton btnSourceFile 
         Height          =   330
         Left            =   2700
         Picture         =   "frmWizard.frx":1EC0E
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4110
         Width           =   375
      End
      Begin VB.Frame fraTapeOpt 
         Caption         =   "Tape Options"
         Height          =   1005
         Left            =   2655
         TabIndex        =   23
         Top             =   3000
         Width           =   4290
         Begin VB.CheckBox chkTapeOpt 
            Caption         =   "Rewind the tape when finished"
            Height          =   195
            Index           =   1
            Left            =   270
            TabIndex        =   25
            Top             =   630
            Width           =   3885
         End
         Begin VB.CheckBox chkTapeOpt 
            Caption         =   "Save the raw archive to the hard disk"
            Height          =   195
            Index           =   0
            Left            =   270
            TabIndex        =   24
            Top             =   315
            Width           =   3165
         End
      End
      Begin VB.Frame fraSource 
         Caption         =   "Select Data Source"
         Height          =   1455
         Left            =   2655
         TabIndex        =   19
         Top             =   1395
         Width           =   4290
         Begin VB.OptionButton optSource 
            Caption         =   "Raytheon archive file"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   47
            Top             =   1080
            Width           =   2415
         End
         Begin VB.OptionButton optSource 
            Caption         =   "Read the raw data directly from the tape"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   810
            Width           =   3795
         End
         Begin VB.OptionButton optSource 
            Caption         =   "Filtered archive file - previously processed"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   540
            Width           =   3795
         End
         Begin VB.OptionButton optSource 
            Caption         =   "Raw archive file - The file read from the tape"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   270
            Width           =   3840
         End
      End
      Begin VB.Label lblSourceFileName 
         Alignment       =   1  'Right Justify
         Caption         =   "Save filtered data to what file? "
         Height          =   240
         Left            =   180
         TabIndex        =   28
         Top             =   4155
         Width           =   2445
      End
      Begin VB.Label lblInstructions 
         Caption         =   $"frmWizard.frx":1EF50
         Height          =   690
         Index           =   1
         Left            =   2700
         TabIndex        =   17
         Top             =   630
         Width           =   4155
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1 - Select Data Source"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   2700
         TabIndex        =   9
         Tag             =   "2001"
         Top             =   210
         Width           =   3960
      End
      Begin VB.Image imgStep 
         BorderStyle     =   1  'Fixed Single
         Height          =   2040
         Index           =   1
         Left            =   210
         Picture         =   "frmWizard.frx":1EFE6
         Stretch         =   -1  'True
         Top             =   210
         Width           =   2070
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   1305
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 2001 Mercury Solutions, Inc.
' FORM:     frmWizard
' AUTHOR:   Tom Elkins
' PURPOSE:  Walks the user through the steps required for processing an archive
' REVISIONS:
'   v1.6.0  TAE Original code
'   v1.6.1  TAE Corrected bug in tape file naming
'               Added verbose logging calls
'   v1.7.4  SPV Added set CCOSVersion for FileOps

Option Explicit
'
' Functions from Keith's tape library
'Private Declare Function tapeInit Lib "ReadTape.dll" Alias "InitializeTape" () As Long
'Private Declare Function tapePath Lib "ReadTape.dll" Alias "ParsePathStr" (ByVal sMyString As String) As String
'Private Declare Function tapeReadFile Lib "ReadTape.dll" Alias "ReadTapeFile" () As Long
'Private Declare Function tapeRewind Lib "ReadTape.dll" Alias "Rewind" () As Long
'Private Declare Function tapeEject Lib "ReadTape.dll" Alias "EjectTape" () As Long
'
' Constants
Const mintNUM_STEPS = 5             ' Number of steps in the wizard
'
' Source option constants
Const mintSRC_CCA = 0               ' Source is a raw archive (CCA) file
Const mintSRC_FLT = 1               ' Source is a filtered archive (FLT) file
Const mintSRC_TAPE = 2              ' Source is directly from tape
Const mintSRC_RAY = 3               ' Source is from Raytheon
'
' Tape option constants
Const mintTAPE_SAVE = 0             ' Save tape file to disk
Const mintTAPE_REW = 1              ' Rewind tape when finished
'
' Message selection buttons
Const mintSEL_ALL = 0               ' Select all messages
Const mintSEL_NONE = 1              ' Select no messages
Const mintSEL_INV = 2               ' Invert the selection
'
' Intro constants
Const mintINTRO_SHOW = vbUnchecked  ' Show the intro page
Const mintINTRO_HIDE = vbChecked    ' Hide the intro page
'
' Navigation buttons
Const mintBTN_HELP = 0              ' Help button
Const mintBTN_CANCEL = 1            ' Cancel button
Const mintBTN_BACK = 2              ' Back button
Const mintBTN_NEXT = 3              ' Next button
Const mintBTN_FINISH = 4            ' Finish button
'
' Wizard steps (used to grab the right frame control)
Const mintSTEP_INTRO = 0            ' Intro frame
Const mintSTEP_1 = 1                ' Steps
Const mintSTEP_2 = 2
Const mintSTEP_3 = 3
Const mintSTEP_FINISH = 4           ' Last step
'
' Navigation direction
Const mintDIR_NONE = 0              ' No direction
Const mintDIR_BACK = 1              ' User stepped back
Const mintDIR_NEXT = 2              ' User stepped forward

Const mstrFORM_TITLE = "COMPASS CALL Archive Wizard"
'
'module level vars
Private mintCurrent_Step As Integer ' The current step of the process
Private mblnFinish_OK As Boolean    ' Flag to show the Finish button
Private mintSource As Integer       ' Selected archive source
Private mblnTape_Save As Boolean    ' Flag to save the tape archive to disk
Private mblnTape_Rewind As Boolean  ' Flag to rewind the tape
Private mstrRawFile As String       ' Name of the raw file
Private mstrFilteredFile As String  ' Name of the filtered file
Private mblnCancelled As Boolean    ' Flag to indicate cancelled operation
Private clsReadTape As New ReadTape

'
' EVENT:    btnMsgSel_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Process message selection shortcut buttons
' INPUT:    intButton - The index of the button that was pressed
' OUTPUT:   None
' NOTES:
Private Sub btnMsgSel_Click(intButton As Integer)
    Dim pintMsg As Integer  ' Loop counter
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmWizard.btnMsgSel Click (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intButton
    End If
    '-v1.6.1
    '
    ' Loop through all of the messages in the list
    For pintMsg = 1 To Me.lvMessages.ListItems.Count
        '
        ' Check or uncheck based on the button selected:
        '   ALL - Check the item
        '   NONE - Uncheck the item
        '   INVERT - Set the opposite of the current state
        Me.lvMessages.ListItems(pintMsg).Checked = IIf(intButton = mintSEL_ALL, True, IIf(intButton = mintSEL_NONE, False, Not (Me.lvMessages.ListItems(pintMsg).Checked)))
    Next pintMsg
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.btnMsgSel Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnSourceFile_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Prompt the user to select the source file
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnSourceFile_Click()
    Dim pblnRaw As Boolean  ' A flag indicating if the file is a raw (TRUE) or filtered (FALSE) archive
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.btnSourceFile Click (Start)"
    '-v1.6.1
    '
    '
    ' Trap errors
    On Error GoTo Hell
    '
    ' Use control-level addressing
    With Me.dlgFile
        '
        ' Set the cancel button to generate an error
        .CancelError = True
        '
        ' Set up the file selection box depending on the source type
        Select Case mintSource
            '
            ' User wants to process a raw archive
            Case mintSRC_CCA:
                '
                ' Configure to open an existing raw archive file
                pblnRaw = True
                .DialogTitle = "Select the raw archive file to process"
                .Flags = cdlOFNFileMustExist
                .Filter = "Raw Compass Call archive (*.cca)|*.cca"
                .ShowOpen
            '
            ' User wants to process a filtered archive
            Case mintSRC_FLT:
                '
                ' Configure to open an existing filtered archive file
                pblnRaw = False
                .DialogTitle = "Select the filtered archive file to process"
                .Flags = cdlOFNFileMustExist
                .Filter = "Filtered Compass Call archive (*.flt)|*.flt"
                .ShowOpen
            '
            ' User wants to read the archive from the tape
            Case mintSRC_TAPE:
                '
                ' See if the user selected the option to save the file to disk
                If Me.chkTapeOpt(mintTAPE_SAVE).Value = vbChecked Then
                    '
                    ' Configure to save a raw archive file
                    pblnRaw = True
                    .DefaultExt = ".cca"
                    .Filter = "Raw Compass Call archive (*.cca)|*.cca"
                Else
                    '
                    ' Configure to save a filtered archive file
                    pblnRaw = False
                    .DefaultExt = ".flt"
                    .Filter = "Filtered Compass Call archive (*.flt)|*.flt"
                End If
                '
                ' Configure to save the file
                .DialogTitle = "Select the path and name for the saved archive"
                .Flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt
                .ShowSave
            Case mintSRC_RAY:
                ' Configure to open a Raytheon archive file
                pblnRaw = True
                .DialogTitle = "Select the archive file to process"
                .Flags = cdlOFNFileMustExist
                .Filter = "Raytheon archive (*.*)|*.*"
                .ShowOpen
                '
        End Select
        '
        ' See if user selected a file
        If .FileName <> "" Then
            '
            ' Save the file name, and allow the user to move to the next step
            Me.txtFileName.Text = .FileName
            Me.cmdNav(mintBTN_NEXT).Enabled = True
            '
            ' Save the filename
            If pblnRaw Then
                mstrRawFile = .FileName
                '
                ' Create the name for the filtered file
                mstrFilteredFile = Me.strDeriveFilteredFileName(mstrRawFile)
            Else
                mstrRawFile = ""
                mstrFilteredFile = .FileName
            End If
        End If
    End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.btnSourceFile Click (End)"
    '-v1.6.1
    '
    Exit Sub
'
'
Hell:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmWizard.btnSourceFile Click (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    Select Case Err.Number

        Case 32755: '(Cancel)
            '
            ' Remove the file text and disable the "Next" button
            Me.txtFileName.Text = ""
            Me.cmdNav(mintBTN_NEXT).Enabled = False
    End Select
End Sub
'
' EVENT:    chkHideIntro_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Save the user's preference for showing the intro screen
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkHideIntro_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.chkHideIntro Click (Start)"
    '-v1.6.1
    '
    ' Write the setting to the INI file
    basCCAT.WriteToken "Miscellaneous operations", "WizardIntro", CStr(Me.chkHideIntro.Value)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.chkHideIntro Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkTapeOpt_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Handles UI changes when certain options are selected
' INPUT:    intOption - The tape option selected
' OUTPUT:   None
' NOTES:
Private Sub chkTapeOpt_Click(intOption As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmWizard.chkTapeOpt Click (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intOption
    End If
    '-v1.6.1
    '
    '
    ' See if the "Save" option was selected
    If intOption = mintTAPE_SAVE Then
        '
        ' Reset the file selection, because this setting changes what type of file
        ' is saved.
        Me.txtFileName.Text = ""
        Me.dlgFile.FileName = ""
        Me.cmdNav(mintBTN_NEXT).Enabled = False
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.chkTapeOpt Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbArchiveClass_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the system-wide security classification based on the classification
'           selected
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbArchiveClass_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.cmbArchiveClass Click (Start)"
    '-v1.6.1
    '
    ' Get the coded numeric equivalent of the classification text
    Me.cmbArchiveClass.Tag = frmSecurity.lngGetNumber("Security values from text", "SECURITY_VAL_" & Me.cmbArchiveClass.Text, 0)
    '
    ' Update the "Next" button
    Me.cmdNav(mintBTN_NEXT).Enabled = (Len(Me.txtArchiveName.Text) > 0) And Me.cmbArchiveVersion.Text <> "" And Me.cmbArchiveClass.Text <> ""
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.cmbArchiveClass Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbArchiveVersion_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Update the UI when the user selects an archive version
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbArchiveVersion_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.cmbArchiveVersion Click (Start)"
    '-v1.6.1
    '
    ' Update the "Next" button
    If Me.cmbArchiveVersion.ListIndex > 0 Then
       Filtraw2Das.CCOSVersion = CSng(Val(basCCAT.GetAlias("Versions", "CCOS" & Me.cmbArchiveVersion.ListIndex, "0")))
       '+v1.7.4SPV
       FileOps.CCOSVersion = CSng(Val(basCCAT.GetAlias("Versions", "CCOS" & Me.cmbArchiveVersion.ListIndex, "0")))
       '-v1.7.4
    End If
    Me.cmdNav(mintBTN_NEXT).Enabled = (Len(Me.txtArchiveName.Text) > 0) And Me.cmbArchiveVersion.Text <> "" And Me.cmbArchiveClass.Text <> ""
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.cmbArchiveVersion Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmdNav_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Process the navigation action
' INPUT:    intButton - the button that was pressed
' OUTPUT:   None
' NOTES:
Private Sub cmdNav_Click(intButton As Integer)
    Dim pintNew_Step As Integer     ' The new step
    Dim plngHelp_Topic As Long      ' The index to the help file for the current step
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmWizard.cmdNav Click (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intButton
    End If
    '-v1.6.1
    '
    Select Case intButton
        '
        ' Help button
        Case mintBTN_HELP
            basCCAT.ShowHelp Me, basCCAT.IDH_GUI_WIZARD_INTRO + mintCurrent_Step
        '
        ' Cancel
        Case mintBTN_CANCEL
            '
            ' Stop archive processing
            Filtraw2Das.ContinueProcessing = False
            mblnCancelled = True
            Me.Hide
        '
        ' Back
        Case mintBTN_BACK
            '
            ' Go to the previous step
            pintNew_Step = mintCurrent_Step - 1
            SetStep pintNew_Step, mintDIR_BACK
        '
        ' Next
        Case mintBTN_NEXT
            '
            ' If we are at the first step, see if the source is NOT the tape
            If mintCurrent_Step = mintSTEP_1 And mintSource <> mintSRC_TAPE Then
                '
                ' Set the default archive date to the date of the selected file
                '+v1.6.1TE
                If Dir(Me.txtFileName.Text) <> "" Then
                    Me.dtpArchiveDate.Value = FileDateTime(Me.txtFileName.Text)
                Else
                    Me.dtpArchiveDate.Value = Now
                End If
                '-v1.6.1
            End If
            '
            ' Go to the next step
            pintNew_Step = mintCurrent_Step + 1
            SetStep pintNew_Step, mintDIR_NEXT
        '
        ' Finish
        Case mintBTN_FINISH
            '
            ' Process the archive
            Me.cmdNav(mintBTN_BACK).Enabled = False
            Me.cmdNav(mintBTN_FINISH).Enabled = False
            Me.ProcessArchive
            Unload Me
       
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.cmdNav Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Load
' AUTHOR:   Tom Elkins
' PURPOSE:  Initializes the form
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    Dim intStep As Integer          ' Loop counter for steps
    Dim blnShowIntro As Boolean     ' Flag to show the intro screen
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard Load (Start)"
    '-v1.6.1
    '
    ' Initialize the variables
    mblnFinish_OK = False
    mblnCancelled = False
    '
    ' Move the step frames out of the view
    For intStep = 0 To mintNUM_STEPS - 1
      fraStep(intStep).Left = -10000
    Next
    '
    ' Reset step 1 controls
    Me.fraTapeOpt.Enabled = False
    Me.chkTapeOpt(mintTAPE_SAVE).Enabled = False
    Me.chkTapeOpt(mintTAPE_REW).Enabled = False
    Me.lblSourceFileName.Caption = "Select file to open"
    Me.btnSourceFile.Enabled = False
    Me.txtFileName.Enabled = False
    Me.txtFileName.BackColor = vbButtonFace
    '
    ' Reset step 2 controls
    Me.txtArchiveName.Text = ""
    Me.dtpArchiveDate.Value = Date
    basCCAT.PopulateCCOSVersions Me.cmbArchiveVersion
    frmSecurity.PopulateClassification Me.cmbArchiveClass
    '
    ' Reset step 3 controls
    basCCAT.PopulateMessageList Me.lvMessages
    Me.btnMsgSel(mintSEL_ALL).Enabled = (Me.lvMessages.ListItems.Count > 0)
    Me.btnMsgSel(mintSEL_NONE).Enabled = (Me.lvMessages.ListItems.Count > 0)
    Me.btnMsgSel(mintSEL_INV).Enabled = (Me.lvMessages.ListItems.Count > 0)
    '
    ' Reset step 4 controls
    Me.lstToDo.Enabled = False
    Me.lstToDo.ForeColor = vbBlack
    '
    ' Determine 1st Step:
    If basCCAT.GetNumber("Miscellaneous operations", "WizardIntro", mintINTRO_SHOW) = mintINTRO_HIDE Then
        chkHideIntro.Value = vbChecked
        SetStep 1, mintDIR_NEXT
    Else
        SetStep 0, mintDIR_NONE
    End If
    '
    ' Assign help context
    Me.fraStep(mintSTEP_INTRO).HelpContextID = basCCAT.IDH_GUI_WIZARD_INTRO
    Me.fraStep(mintSTEP_1).HelpContextID = basCCAT.IDH_GUI_WIZARD_SOURCE
    Me.fraStep(mintSTEP_2).HelpContextID = basCCAT.IDH_GUI_WIZARD_INFORMATION
    Me.fraStep(mintSTEP_3).HelpContextID = basCCAT.IDH_GUI_WIZARD_MESSAGES
    Me.fraStep(mintSTEP_FINISH).HelpContextID = basCCAT.IDH_GUI_WIZARD_FINAL
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard Load (End)"
    '-v1.6.1
    '
End Sub
'
' METHOD:   SetStep
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets up the new step for the user
' INPUT:    intStep - The index of the new step
'           intDirection - the direction the user used to get to the step
'               mintDIR_NONE - the program set the step
'               mintDIR_BACK - the user used the "Back" button to get to the step
'               mintDIR_NEXT - the user used the "Next" button to get to the step
' OUTPUT:   None
' NOTES:
Private Sub SetStep(intStep As Integer, intDirection As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmWizard.SetStep (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intStep & ", " & intDirection
    End If
    '-v1.6.1
    '
    ' Disable the controls on the current step
    fraStep(mintCurrent_Step).Enabled = False
    '
    ' Move the new step into view
    fraStep(intStep).Left = 0
    '
    ' See if the current step is different than the new step
    If intStep <> mintCurrent_Step Then
        '
        ' Move the current step out of view
        fraStep(mintCurrent_Step).Left = -10000
    End If
    '
    ' Enable the controls on the new step
    fraStep(intStep).Enabled = True
    '
    ' Update the controls on the wizard form
    SetCaption intStep
    SetNavBtns intStep
    Me.HelpContextID = basCCAT.IDH_GUI_WIZARD_INTRO + intStep
    '
    ' Handle special processing for each step
    Select Case intStep
        '
        ' Intro processing
        Case mintSTEP_INTRO

        '
        ' Step 1
        Case mintSTEP_1
            '
            ' See if the "Next" button should be enabled or not
            Me.cmdNav(mintBTN_NEXT).Enabled = (Len(Me.txtFileName.Text) > 0) And Me.txtFileName.Enabled
      
        Case mintSTEP_2
            '
            ' See if the "Next" button should be enabled
            Me.cmdNav(mintBTN_NEXT).Enabled = (Len(Me.txtArchiveName.Text) > 0) And Me.cmbArchiveVersion.Text <> "" And Me.cmbArchiveClass.Text <> ""
        
        Case mintSTEP_3
      
        Case mintSTEP_FINISH
            '
            ' Enable the "Finish" button
            Me.cmdNav(mintBTN_FINISH).Enabled = True
            '
            ' Set up the To-Do list
            Me.lstToDo.Clear
            '
            ' See if the user wants to use the tape
            If Me.optSource(mintSRC_TAPE).Value Then
                '
                ' Add tape tasks
                Me.lstToDo.AddItem "1I -- Initialize the tape drive"
                Me.lstToDo.AddItem "1H -- Read header files from the tape"
                Me.lstToDo.AddItem "1A -- Read archive file from the tape"
                If mblnTape_Rewind Then Me.lstToDo.AddItem "1R -- Rewind and eject the tape"
            End If
            '
            ' Add database tasks
            Me.lstToDo.AddItem "2A -- Add archive information to database"
            Me.lstToDo.AddItem "2C -- Create Summary and Data tables"
            '
            ' If the raw archive was selected, add the filtering task
            If mintSource <> mintSRC_FLT Then Me.lstToDo.AddItem "3E -- Extract selected messages from the archive"
            '
            ' Add the translation and update tasks
            Me.lstToDo.AddItem "3P -- Parse and translate the filtered data"
            Me.lstToDo.AddItem "4U -- Update the user interface"
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmWizard.SetStep (End)"
    '-v1.6.1
    '
End Sub
'
' METHOD:   SetNavBtns
' AUTHOR:   Visual Basic
' PURPOSE:  Sets the state of the navigation buttons
' INPUT:    intStep - The index of the new step
' OUTPUT:   None
' NOTES:
Private Sub SetNavBtns(intStep As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmWizard.SetNavBtns (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intStep
    End If
    '-v1.6.1
    '
    '
    ' Save the current step
    mintCurrent_Step = intStep
    '
    ' Configure the navigation buttons based on the current step
    If mintCurrent_Step = 0 Then
        Me.cmdNav(mintBTN_BACK).Enabled = False
        Me.cmdNav(mintBTN_NEXT).Enabled = True
    ElseIf mintCurrent_Step = mintNUM_STEPS - 1 Then
        Me.cmdNav(mintBTN_NEXT).Enabled = False
        Me.cmdNav(mintBTN_BACK).Enabled = True
    Else
        Me.cmdNav(mintBTN_BACK).Enabled = True
        Me.cmdNav(mintBTN_NEXT).Enabled = True
    End If
    '
    ' set the state of the "Finish" button
    Me.cmdNav(mintBTN_FINISH).Enabled = mblnFinish_OK
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmWizard.SetNavBtns (End)"
    '-v1.6.1
    '
End Sub
'
' METHOD:   SetCaption
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets up the form caption to include the step
' INPUT:    intStep - The index of the new step
' OUTPUT:   None
' NOTES:
Private Sub SetCaption(intStep As Integer)
    Dim pstrStep As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmWizard.SetCaption (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intStep
    End If
    '-v1.6.1
    '
    '
    ' Set the string based on the step
    Select Case intStep
        Case mintSTEP_INTRO: pstrStep = "Introduction"
        Case mintSTEP_FINISH: pstrStep = "Final"
        Case Else: pstrStep = "Step " & intStep
    End Select
    '
    ' Set the caption
    Me.Caption = mstrFORM_TITLE & " - " & pstrStep & "(" & intStep & " of " & mintSTEP_FINISH & ")"
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmWizard.SetCaption (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    optSource_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configures other controls based on the type of source selected
' INPUT:    intSource - The selected source option
' OUTPUT:   None
' NOTES:
Private Sub optSource_Click(intSource As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmWizard.optSource Click (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intSource
    End If
    '-v1.6.1
    '
    ' Store the selected option
    mintSource = intSource
    '
    ' Enable/disable the tape options depending on the source
    Me.fraTapeOpt.Enabled = (intSource = mintSRC_TAPE)
    '
    ' For now, there is no way to NOT save the raw archive to disk, so
    ' check and disable the box
    Me.chkTapeOpt(mintTAPE_SAVE).Enabled = False
    Me.chkTapeOpt(mintTAPE_SAVE).Value = vbChecked
    Me.chkTapeOpt(mintTAPE_REW).Enabled = Me.fraTapeOpt.Enabled
    '
    ' Enable file selection, and set the prompt according to the source
    Me.lblSourceFileName.Enabled = True
    Me.btnSourceFile.Enabled = True
    Me.txtFileName.Enabled = True
    Me.txtFileName.BackColor = vbWindowBackground
    Select Case intSource
        Case mintSRC_CCA:
            Me.lblSourceFileName.Caption = "Select raw archive file"
        Case mintSRC_FLT:
            Me.lblSourceFileName.Caption = "Select filtered archive file"
        Case mintSRC_TAPE:
            Me.lblSourceFileName.Caption = "Select location for archive"
    End Select
    '
    ' Reset file selection
    Me.txtFileName.Text = ""
    Me.dlgFile.FileName = ""
    Me.cmdNav(mintBTN_NEXT).Enabled = False
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.optSource Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtArchiveName_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Configures controls when the archive name changes
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub txtArchiveName_Change()
    '
    ' See if the "Next" button should be enabled
    Me.cmdNav(mintBTN_NEXT).Enabled = (Len(Me.txtArchiveName.Text) > 0) And Me.cmbArchiveVersion.Text <> "" And Me.cmbArchiveClass.Text <> ""
End Sub
'
' METHOD:   blnInitializeTape
' AUTHOR:   Tom Elkins
' PURPOSE:  Initializes the tape drive
' INPUT:    None
' OUTPUT:   TRUE - if the tape initializes
'           FALSE - if there is no tape, no drive, or some other error
' NOTES:
Public Function blnInitializeTape() As Boolean
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnInitializeTape (Start)"
    '-v1.6.1
    '
    ' Trap errors
    On Error GoTo Hell
    '
    ' Initialize the tape
   ' blnInitializeTape = (tapeInit > 0)
   blnInitializeTape = clsReadTape.InitializeTape
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnInitializeTape (End)"
    '-v1.6.1
    '
    Exit Function
'
'
Hell:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmWizard.blnInitializeTape (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    Select Case Err.Number
    
        Case 53: ' File not found
            MsgBox "Cannot locate the library 'ReadTape.dll'" & _
                    vbCrLf & "Please locate the file, copy/move it " & _
                    "to the CCAT application directory, and restart the application.", _
                    vbOKOnly Or vbCritical, _
                    "Missing Tape Library", _
                    App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, _
                    basCCAT.IDH_ERR_53
    End Select
    blnInitializeTape = False
End Function
'
' METHOD:   blnReadTapeHeaderFiles
' AUTHOR:   Tom Elkins
' PURPOSE:  Reads the two header files found before the archive
' INPUT:    None
' OUTPUT:   TRUE - if the files were read
'           FALSE - if there was an error
' NOTES:
Public Function blnReadTapeHeaderFiles() As Boolean
    Dim pstrFile As String      ' Name of the output header files
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnReadTapeHeaderFiles (Start)"
    '-v1.6.1
    '
    ' Trap known errors
    On Error GoTo Hell
    '
    ' Assume failure
    blnReadTapeHeaderFiles = False
    '
    ' Get the filename for the first header file
    'pstrFile = tapePath(Me.txtFileName.Text & ".hdr1")
    '
    ' See if it was read
    If (clsReadTape.ReadTapeFile(Me.txtFileName.Text & ".hdr1") = "No Error") Then
        '
        ' Get the name of the second header file
        'pstrFile = tapePath(Me.txtFileName.Text & ".hdr2")
        '
        ' See if it was read
        If (clsReadTape.ReadTapeFile(Me.txtFileName.Text & ".hdr2") = "No Error") Then
            '
            ' Operation success
            blnReadTapeHeaderFiles = True
        End If
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnReadTapeHeaderFiles (End)"
    '-v1.6.1
    '
    Exit Function
'
'
Hell:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmWizard.blnReadTapeHeaderFiles (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
End Function
'
' METHOD:   blnReadTapeArchive
' AUTHOR:   Tom Elkins
' PURPOSE:  Reads the binary archive from the tape
' INPUT:    None
' OUTPUT:   TRUE - if the archive was read
'           FALSE - if there was an error
' NOTES:
Public Function blnReadTapeArchive() As Boolean
    Dim pstrFile As String      ' The name of the archive file
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnReadTapeArchive (Start)"
    '-v1.6.1
    '
    ' Trap known errors
    On Error GoTo Hell
    '
    ' Assume failure
    blnReadTapeArchive = False
    '
    ' Get the name of the archive
    '+v1.6.1TE
    'pstrFile = tapePath(Me.txtArchiveName.Text & "")
    'pstrFile = tapePath(Me.txtFileName.Text & "")
    
    '-v1.6.1
    '
    ' Read the file
    blnReadTapeArchive = clsReadTape.ReadTapeFile(Me.txtFileName.Text) = "No Error"

    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnReadTapeArchive (End)"
    '-v1.6.1
    '
    Exit Function
'
'
Hell:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmWizard.blnReadTapeArchive (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
End Function
'
' METHOD:   blnRewindEjectTape
' AUTHOR:   Tom Elkins
' PURPOSE:  Rewinds and ejects the tape
' INPUT:    None
' OUTPUT:   TRUE - if the tape was rewound and ejected
'           FALSE - if there was an error
' NOTES:
Public Function blnRewindEjectTape() As Boolean
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnRewindEjectTape (Start)"
    '-v1.6.1
    '
    ' Rewind and eject
    clsReadTape.RewindTape
    clsReadTape.EjectTape
    blnRewindEjectTape = True
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.blnRewindEjectTape (End)"
    '-v1.6.1
    '
End Function
'
' METHOD:   blnAddArchiveEntry
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds the entry to the archives table
' INPUT:    None
' OUTPUT:   TRUE - if the record was successfully added
'           FALSE - if there was an error
' NOTES:
Public Function blnAddArchiveEntry() As Boolean
    Dim prsArchives As Recordset    ' Entry in the Archives table
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnAddArchiveEntry (Start)"
    '-v1.6.1
    '
    ' Trap known errors
    On Error GoTo Hell
    '
    ' Assume failure
    blnAddArchiveEntry = False
    '
    ' Open the archives table
    Set prsArchives = basDatabase.guCurrent.DB.OpenRecordset(basDatabase.TBL_ARCHIVES)
    '
    ' Make sure there is a table
    If Not prsArchives Is Nothing Then
        '
        ' Add a new entry to the archives table
        prsArchives.AddNew
        '
        ' Populate the entry with the data provided in the wizard
        prsArchives("Name") = Me.txtArchiveName.Text
        prsArchives("Date") = Format(Me.dtpArchiveDate.Value, "mm/dd/yyyy")
        prsArchives("Archive") = Me.txtFileName.Text
        prsArchives("Analysis_File") = Me.cmbArchiveVersion.Text
        '
        ' Add the new record to the table
        prsArchives.Update
        '
        ' Close the table
        prsArchives.Close
        '
        ' Report success
        blnAddArchiveEntry = True
    End If
    '
    ' Remove the object
    Set prsArchives = Nothing
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnAddArchiveEntry (End)"
    '-v1.6.1
    '
    Exit Function
'
'
Hell:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmWizard.blnAddArchiveEntry (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
End Function
'
' METHOD:   strDeriveFilteredFileName
' AUTHOR:   Tom Elkins
' PURPOSE:  Changes the file extension from .CCA to .FLT
' INPUT:    strArchiveName - The name of the raw archive file
' OUTPUT:   The name of the filtered archive file
' NOTES:
Friend Function strDeriveFilteredFileName(strArchiveName As String) As String
    Dim pastrPath() As String   ' String array to hold the components of the filename
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : frmWizard.strDeriveFilteredFileName (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & strArchiveName
    End If
    '-v1.6.1
    '
    ' Parse the archive filename to produce the filtered file archive name
    pastrPath = Split(strArchiveName, ".")
    '
    ' Change the last component to the new extension
    pastrPath(UBound(pastrPath)) = "flt"
    '
    ' Re-combine the components
    strDeriveFilteredFileName = Join(pastrPath, ".")
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.strDeriveFilteredFileName (End)"
    '-v1.6.1
    '
End Function
'
' METHOD:   blnCreateArchiveTables
' AUTHOR:   Tom Elkins
' PURPOSE:  Creates the summary and data tables
' INPUT:    None
' OUTPUT:   TRUE - if the tables were successfully created
'           FALSE - if there was an error
' NOTES:
Public Function blnCreateArchiveTables() As Boolean
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnCreateArchiveTables (Start)"
    '-v1.6.1
    '
    ' Create the tables
    blnCreateArchiveTables = basDatabase.blnCreateSummaryTable(basDatabase.guCurrent.DB, Me.txtArchiveName.Text) And basDatabase.blnCreateDataTable(basDatabase.guCurrent.DB, Me.txtArchiveName.Text)
    '+v1.7BB
    blnCreateArchiveTables = blnCreateArchiveTables And basTOC.blnCreateTOCTable(basDatabase.guCurrent.DB, Me.txtArchiveName.Text) And basTOC.blnCreateVarStructTable(basDatabase.guCurrent.DB, Me.txtArchiveName.Text)
    blnCreateArchiveTables = blnCreateArchiveTables And basTOC.blnCreateProcDataTable(basDatabase.guCurrent.DB, Me.txtArchiveName.Text) And basTOC.blnCreateMessageTable(basDatabase.guCurrent.DB, Me.txtArchiveName.Text)
    
    Set guCurrent.uArchive.rsSummary = guCurrent.DB.OpenRecordset(Me.txtArchiveName.Text & basDatabase.TBL_SUMMARY, dbOpenDynaset)
    Set guCurrent.uArchive.rsTOC = guCurrent.DB.OpenRecordset(Me.txtArchiveName.Text & basDatabase.TBL_TOC, dbOpenDynaset)
    Set guCurrent.uArchive.rsVarStruct = guCurrent.DB.OpenRecordset(Me.txtArchiveName.Text & basDatabase.TBL_VAR_STRUCT, dbOpenDynaset)
    Set guCurrent.uArchive.rsMessage = guCurrent.DB.OpenRecordset(Me.txtArchiveName.Text & basDatabase.TBL_MESSAGE, dbOpenDynaset)
    guCurrent.sArchive = Me.txtArchiveName.Text
    '-v1.7BB    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmWizard.blnCreateArchiveTables (End)"
    '-v1.6.1
    '
End Function
'
' METHOD:   blnFilterArchive
' AUTHOR:   Tom Elkins
' PURPOSE:  Initiates the filtering process
' INPUT:    None
' OUTPUT:   TRUE - if the filtering completed
'           FALSE - if there was an error
' NOTES:
Public Function blnFilterArchive() As Boolean
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnFilterArchive (Start)"
    '-v1.6.1
    '
    ' Set the progress bar
    Me.barProgress.Min = 0
    Me.barProgress.Value = 0
    Me.barProgress.Max = FileLen(mstrRawFile)
    '
    ' Process the raw archive
    Dim rayArc As Boolean
    rayArc = False
    If mintSource = mintSRC_RAY Then
        rayArc = True
    End If
    
    blnFilterArchive = Raw2filtraw.blnFilterRawArchive(mstrRawFile, mstrFilteredFile, rayArc)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnFilterArchive (End)"
    '-v1.6.1
    '
End Function
'
' METHOD:   blnTranslateArchive
' AUTHOR:   Tom Elkins
' PURPOSE:  Initiates the translation process
' INPUT:    None
' OUTPUT:   TRUE - if the translation succeeded
'           FALSE - if there was an error
' NOTES:
Public Function blnTranslateArchive() As Boolean
    Dim prsArchive As Recordset     ' Record in the archives table
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnTranslateArchive (Start)"
    '-v1.6.1
    '
    ' Trap known errors
    On Error GoTo Hell
    '
    ' Assume failure
    blnTranslateArchive = False
    '
    ' Set the progress bar
    Me.barProgress.Min = 0
    Me.barProgress.Value = 0
    Me.barProgress.Max = FileLen(mstrFilteredFile)
    Me.lblPctDone = "Translating Archive - 0% Complete"
    '
    ' Reset the information about the archive
    guCurrent.uArchive.dOffset_Time = 0#
    guCurrent.uArchive.dtArchiveDate = Int(Me.dtpArchiveDate.Value)
    guCurrent.uArchive.dtEnd_Time = 0#
    guCurrent.uArchive.dtStart_Time = #1/1/9999#
    guCurrent.uArchive.lFile_Size = FileLen(mstrFilteredFile)
    guCurrent.uArchive.lNum_Bytes = 0
    guCurrent.uArchive.lNum_Messages = 0
    '
    ' Open the summary and data tables
    Set guCurrent.uArchive.rsData = guCurrent.DB.OpenRecordset(Me.txtArchiveName.Text & basDatabase.TBL_DATA, dbOpenDynaset)
    Set guCurrent.uArchive.rsSummary = guCurrent.DB.OpenRecordset(Me.txtArchiveName.Text & basDatabase.TBL_SUMMARY, dbOpenDynaset)
    Set guCurrent.uArchive.rsTOC = guCurrent.DB.OpenRecordset(Me.txtArchiveName.Text & basDatabase.TBL_TOC, dbOpenDynaset)
    Set guCurrent.uArchive.rsVarStruct = guCurrent.DB.OpenRecordset(Me.txtArchiveName.Text & basDatabase.TBL_VAR_STRUCT, dbOpenDynaset)
    Set guCurrent.uArchive.rsProcData = guCurrent.DB.OpenRecordset(Me.txtArchiveName.Text & basDatabase.TBL_PROC_DATA, dbOpenDynaset)
   '
    ' Process the filtered archive
    Filtraw2Das.ContinueProcessing = True
    blnTranslateArchive = Filtraw2Das.ProcFiltMain(mstrFilteredFile)
    '
    ' Update the archive entry
    Set prsArchive = guCurrent.DB.OpenRecordset("SELECT * FROM " & basDatabase.TBL_ARCHIVES & " WHERE Name = '" & Me.txtArchiveName.Text & "'")
    prsArchive.Edit
    prsArchive!Start = guCurrent.uArchive.dtStart_Time
    prsArchive!End = guCurrent.uArchive.dtEnd_Time
    prsArchive!Processed = Now
    prsArchive!Messages = guCurrent.uArchive.lNum_Messages
    prsArchive!Bytes = guCurrent.uArchive.lNum_Bytes
    prsArchive.Update
    Set prsArchive = Nothing
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmWizard.blnTranslateArchive (End)"
    '-v1.6.1
    '
    Exit Function
'
'
Hell:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmWizard.blnTranslateArchive (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
End Function
'
' PROPERTY: IsSelected
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets/Returns whether a particular message is selected for processing
' INPUT:    intMsg_ID - The message ID to be checked
' STATE:    TRUE - the specified message is selected to be processed
'           FALSE - the specified message is not selected
' NOTES:
Public Property Get IsSelected(intMsg_ID As Integer) As Boolean
    '
    ' Trap Errors
    On Error Resume Next
    '
    ' See if the specified message ID is selected
    IsSelected = Me.lvMessages.ListItems("MSG" & intMsg_ID).Checked
    '
    ' See if there was an error
    If Err Then
        '
        ' Log the error
        basCCAT.WriteLogEntry "Error #" & Err.Number & " - " & Err.Description & "; attempting to retrieve state of Message " & intMsg_ID
        '
        ' If the item was not in the list, return False
        If Err.Number = 35601 Then IsSelected = False
    End If
    '
    '
    On Error GoTo 0
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: frmWizard.IsSelected Get (" & intMsg_ID & ") = " & IsSelected
    '-v1.6.1
    '
End Property
'
Public Property Let IsSelected(intMsg_ID As Integer, blnSelected As Boolean)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "PROPERTY: frmWizard.IsSelected Let (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intMsg_ID & ", " & blnSelected
    End If
    '-v1.6.1
    '
    '
    ' Trap errors locally
    On Error Resume Next
    '
    ' Set the selected message to the specified state
    Me.lvMessages.ListItems("MSG" & intMsg_ID).Checked = blnSelected
    '
    ' Check for errors
    If Err Then
        '
        ' Log the error
        basCCAT.WriteLogEntry "Error #" & Err.Number & " - " & Err.Description & "; attempting to set Message " & intMsg_ID & " to " & blnSelected
    End If
    On Error GoTo 0
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY : frmWizard.IsSelected Let (End)"
    '-v1.6.1
    '
End Property
'
' METHOD:   ProcessArchive
' AUTHOR:   Tom Elkins
' PURPOSE:  Processes the items in the ToDo list
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Sub ProcessArchive()
    Dim pintTask As Integer         ' The current task
    Dim pblnContinue As Boolean     ' Flag to see if we continue processing
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmWizard.ProcessArchive (Start)"
    '-v1.6.1
    '
    ' Start
    pblnContinue = True
    '
    ' Loop through the tasks
    For pintTask = 0 To Me.lstToDo.ListCount - 1
        '
        ' Highlight the current task
        Me.lstToDo.ListIndex = pintTask
        '
        ' See if we should continue
        If pblnContinue Then
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmWizard.ProcessArchive (Task = " & Me.lstToDo.List(Me.lstToDo.ListIndex) & ")"
            '-v1.6.1
            '
            ' Perform highlighted task
            Select Case Left(Me.lstToDo.List(Me.lstToDo.ListIndex), 2)
                '
                Case "1I": ' Initialize the tape drive
                    '
                    '
                    pblnContinue = Me.blnInitializeTape
                    clsReadTape.RewindTape
                '
                '
                Case "1H": ' Read the header files
                    '
                    '
                    pblnContinue = Me.blnReadTapeHeaderFiles
                '
                '
                Case "1A": ' Read the archive file
                    '
                    '
                    pblnContinue = Me.blnReadTapeArchive
                '
                '
                Case "1R": ' Rewind the tape
                    '
                    ' This is not a fatal task, so just inform the user if there is
                    ' a problem
                    'If Not Me.blnRewindEjectTape Then MsgBox "Could not rewind and/or eject the tape." & vbCrLf & "You may try it manually.", vbOKOnly Or vbInformation, "Tape Problem"
                    'pblnContinue = True
                '
                '
                Case "2A": ' Add archive entry to database
                    pblnContinue = Me.blnAddArchiveEntry
                '
                '
                Case "2C": ' Create Data and Summary tables
                    pblnContinue = Me.blnCreateArchiveTables
                    'bb2004
                    If mintSource <> mintSRC_RAY Then
                        If pblnContinue Then FileOps.readHeader (Me.txtFileName.Text & ".hdr2") 'v1.7BB
                    End If
                '
                '
                Case "3E": ' Extract selected messages (filter)
                    pblnContinue = Me.blnFilterArchive
                '
                '
                Case "3P": ' Parse and translate the filtered archive
                    pblnContinue = Me.blnTranslateArchive
                    If Not pblnContinue Then
                       If MsgBox("Unable to Parse and Translate the filtered archive.  Continue Processing Tape?", vbYesNo) = vbYes Then
                            pblnContinue = True
                       End If
                    End If
                '
                '
                Case "4U": ' Update the user interface
                    'update name
                    Dim Tempstring() As String
                    Tempstring = Split(Me.txtFileName.Text, ".cca")
                    Me.txtFileName.Text = Tempstring(0) & "_nxt.cca"
                    mstrRawFile = Me.txtFileName.Text
                    mstrFilteredFile = Me.strDeriveFilteredFileName(mstrRawFile)

                    Me.txtArchiveName.Text = Me.txtArchiveName.Text & "_nxt"
                    If (pblnContinue = Me.blnReadTapeHeaderFiles) Then
                        If (MsgBox("More than one mission on tape.  Do you want to continue?", vbYesNo) = vbYes) Then
                            Me.barProgress.Value = 0
                            Me.lblPctDone = "Copying Next Archive"
                            pintTask = 1
                            Dim pTmp As Integer
                            For pTmp = 2 To Me.lstToDo.ListCount - 1
                               Me.lstToDo.Selected(pTmp) = False
                            Next pTmp
                        Else
                            If Not Me.blnRewindEjectTape Then MsgBox "Could not rewind and/or eject the tape." & vbCrLf & "You may try it manually.", vbOKOnly Or vbInformation, "Tape Problem"
                        End If
                    Else
                        If Not Me.blnRewindEjectTape Then MsgBox "Could not rewind and/or eject the tape." & vbCrLf & "You may try it manually.", vbOKOnly Or vbInformation, "Tape Problem"
                    End If
                   'pblnContinue = True
            End Select
            '
            ' Mark the task completed or not
            If Not mblnCancelled Then
                Me.lstToDo.Selected(pintTask) = pblnContinue
                Me.lstToDo.List(pintTask) = Me.lstToDo.List(pintTask) & IIf(pblnContinue, " -- COMPLETED", " -- FAILED")
            Else
                pblnContinue = False
            End If
            '
            ' Update the display
            DoEvents
        End If
    Next pintTask
    '
    ' Inform the user
    MsgBox IIf(pblnContinue, "All Tasks Completed!", "Processing Terminated"), vbOKOnly Or IIf(pblnContinue, vbInformation, vbCritical), IIf(pblnContinue, "Processing Completed", "Processing Terminated")
    '
    ' Refresh the display
    frmMain.RefreshDisplay
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmWizard.ProcessArchive (End)"
    '-v1.6.1
    '
End Sub
