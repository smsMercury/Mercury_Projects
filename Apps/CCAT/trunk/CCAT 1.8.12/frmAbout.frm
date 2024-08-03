VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HelpContextID   =   100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "About CCAT"
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   540
      HelpContextID   =   801
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      HelpContextID   =   103
      Left            =   6720
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   465
      WhatsThisHelpID =   103
      Width           =   1350
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      HelpContextID   =   105
      Left            =   6720
      TabIndex        =   1
      Tag             =   "&System Info..."
      Top             =   900
      WhatsThisHelpID =   105
      Width           =   1335
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5670
      Left            =   0
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   5670
      ScaleWidth      =   8250
      TabIndex        =   3
      Top             =   0
      Width           =   8250
      Begin VB.CheckBox chkVerbose 
         Caption         =   "Verbose Log"
         Height          =   240
         Left            =   6750
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   10
         Top             =   1350
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.Label lblClassification 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "lblClassification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   0
         WhatsThisHelpID =   101
         Width           =   8295
      End
      Begin VB.Label lblClassification 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "lblClassification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   5280
         WhatsThisHelpID =   101
         Width           =   8295
      End
      Begin VB.Label lblDisclaimer 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning: ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   705
         Left            =   90
         TabIndex        =   7
         Tag             =   "Warning: ..."
         Top             =   4425
         WhatsThisHelpID =   107
         Width           =   7935
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "App Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1170
         Left            =   2520
         TabIndex        =   6
         Tag             =   "App Description"
         Top             =   1560
         WhatsThisHelpID =   102
         Width           =   5535
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         Height          =   225
         Left            =   720
         TabIndex        =   5
         Tag             =   "Version"
         Top             =   870
         Width           =   5685
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   720
         TabIndex        =   4
         Tag             =   "Application Title"
         Top             =   480
         WhatsThisHelpID =   104
         Width           =   5895
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   2445
      Y2              =   2445
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 1999-2001, Mercury Solutions, Inc.
' FORM:     frmAbout
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Displays a brag box showing the user information about the program,
'           developers, and the user's computer.
' REVISIONS:
'   v1.3.0  TAE Added Keith's name to the credits for his tape control
'   v1.5.0  TAE Added context-sensitive help information to the form
'   v1.6.0  TAE Updated names to match programming convention
'   v1.6.1  TAE Added verbose logging calls
'               Added checkbox to toggle verbose logging
'
Option Explicit
'
' Registry Key Security Options...
Const mlngKEY_ALL_ACCESS = &H2003F
'
' Registry Key ROOT Types...
Const mlngHKEY_LOCAL_MACHINE = &H80000002   ' Registry key for the local machine
Const mintERROR_SUCCESS = 0                 ' Return code for success
Const mintREG_SZ = 1                        ' Unicode null-terminated string
Const mintREG_DWORD = 4                     ' 32-bit number
'
' Registry paths
Const mstrREG_KEY_SYSINFO_LOC = "SOFTWARE\Microsoft\Shared Tools Location"  ' Location of the SysInfo tool
Const mstrREG_VAL_SYSINFO_LOC = "MSINFO"                                    ' Value to look for
Const mstrREG_KEY_SYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const mstrREG_VAL_SYSINFO = "PATH"
'
' Function delcarations to the API
Private Declare Function advapiRegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function advapiRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function advapiRegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
'
'+v1.6.1TE
' EVENT:    chkVerbose_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Toggles the Verbose logging state
' TRIGGER:  User clicks the check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkVerbose_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmAbout.chkVerbose Click (Start)"
    '-v1.6.1
    '
    basCCAT.Verbose = (Me.chkVerbose.Value = vbChecked)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmAbout.chkVerbose Click (End) = " & basCCAT.Verbose
    '-v1.6.1
    '
End Sub
'-v1.6.1
'
'+v1.6TE
' EVENT:    Form_KeyDown
' AUTHOR:   Tom Elkins
' PURPOSE:  Traps the keypress event and launches the help file
' TRIGGER:  User presses a keyboard key
' INPUT:    "intKeyCode" is the key the user pressed
'           "intShift" is the state of the Alt, Shift, and Ctrl keys
' OUTPUT:   None
' NOTES:
'Private Sub Form_KeyDown(intKeyCode As Integer, intShift As Integer)
'    If intKeyCode = vbf1 Then
'        basCCAT.ShowHelp Me, basCCAT.IDH_GUI_ABOUT
'    End If
'End Sub
'-v1.6
'
' EVENT:    Form_Load
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Configure the controls of the form
' TRIGGER:  User clicks on the Help-->About menu item
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmAbout Load (Start)"
    '-v1.6.1
    '
    ' Use form-level addressing
    With frmAbout
        '
        ' Set the text of some of the labels
        .lblVersion.Caption = basCCAT.sGet_Version & " Copyright " & Chr(169) & " 1999-2001 " & App.CompanyName
        .lblTitle.Caption = "COMPASS CALL Archive Translator"
        .lblDescription.Caption = "Reads and processes COMPASS CALL binary archives, and outputs DAS files for analysis"
        .lblDisclaimer.Caption = "Warning: This product is for official U.S. Government business only, and is subject to export restrictions."
        '
        ' Position the background image at the upper left corner of the form
        .picBack.Left = 0
        .picBack.Top = 0
        '
        ' Size the form to match the picture dimensions
        .Height = .picBack.Height
        .Width = .picBack.Width
        '
        ' Set the classification banners to the appropriate colors and text
        .lblClassification(0).BackColor = frmSecurity.lngGetSecurityBackColor(frmSecurity.Tag)
        .lblClassification(0).ForeColor = frmSecurity.lngGetSecurityForeColor(frmSecurity.Tag)
        .lblClassification(0).Caption = frmSecurity.strGetSecurityText(frmSecurity.Tag)
        .lblClassification(1).BackColor = .lblClassification(0).BackColor
        .lblClassification(1).ForeColor = .lblClassification(0).ForeColor
        .lblClassification(1).Caption = .lblClassification(0).Caption
        '
        '+v1.6TE
        ''+v1.5
        '' Set the help context
        '.HelpContextID = basCCAT.lGetHelpID(Me.Name)
        .HelpContextID = basCCAT.IDH_GUI_ABOUT
        ''-v1.5
        '-v1.6
    End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmAbout Load (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmdSysInfo_Click
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Calls the System Information applet
' TRIGGER:  The user clicked on the "System Info" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmdSysInfo_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmAbout.cmdSysInfo Click (Start)"
    '-v1.6.1
    '
    ' Execute the System Information applet
    frmAbout.StartSysInfo
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmAbout.cmdSysInfo Click (End)"
    '-v1.6.1
End Sub
'
' EVENT:    cmdOK_Click
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Remove the form from memory
' TRIGGER:  The user clicked on the "OK" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmdOK_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmAbout.cmdOK Click"
    '-v1.6.1
    '
    ' Remove the form from memory
    Unload frmAbout
End Sub
'
' ROUTINE:  StartSysInfo
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Find and execute the System Info application using the system registry
' INPUT:    None
' OUTPUT:   None
' NOTES:    The System Info application provides a lot of useful information about
'           the host computer system.  This may be used to track down compatibility
'           issues if they ever arise.
Public Sub StartSysInfo()
    Dim pstrSysInfo_Path As String   ' Path to the System Info program
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmAbout.StartSysInfo (Start)"
    '-v1.6.1
    '
    ' Trap errors
    On Error GoTo SysInfoErr
    '
    ' Try To Get System Info Program Path\Name From Registry...
    If frmAbout.GetKeyValue(mlngHKEY_LOCAL_MACHINE, mstrREG_KEY_SYSINFO, mstrREG_VAL_SYSINFO, pstrSysInfo_Path) Then
    '
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf frmAbout.GetKeyValue(mlngHKEY_LOCAL_MACHINE, mstrREG_KEY_SYSINFO_LOC, mstrREG_VAL_SYSINFO_LOC, pstrSysInfo_Path) Then
        '
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(pstrSysInfo_Path & "\MSINFO32.EXE") <> "") Then
            '
            ' Update path
            pstrSysInfo_Path = pstrSysInfo_Path & "\MSINFO32.EXE"
        '
        ' Error - File Can Not Be Found...
        Else
            '
            ' Handle error
            GoTo SysInfoErr
        End If
    '
    ' Error - Registry Entry Can Not Be Found...
    Else
        '
        ' Handle error
        GoTo SysInfoErr
    End If
    '
    ' Execute the application
    Shell pstrSysInfo_Path, vbNormalFocus
    '
    ' Resume error reporting
    On Error GoTo 0
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmAbout.StartSysInfo (END)"
    '-v1.6.1
    '
    ' Leave the subroutine
    Exit Sub
'
' Handle the error
SysInfoErr:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmAbout.StartSysInfo"
    '-v1.6.1
    '
    ' Log the event
    basCCAT.WriteLogEntry "         SysInfo unavailable"
    '
    ' Inform the user the application was not available
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
    '
    ' Resume error reporting
    On Error GoTo 0
End Sub
'
' FUNCTION: GetKeyValue
' AUTHOR:   Visual Basic Application Wizard
' PURPOSE:  Search for a key in the system registry
' INPUT:    "lngKey_Root" is ?
'           "strKey_Name" is the name of a registry key
'           "strSub_Key_Ref" is the reference string under the key
' OUTPUT:   "strKey_Val" is the returned value of the registry key
'           A boolean value indicating that the registry key was found or not
' NOTES:
Public Function GetKeyValue(lngKey_Root As Long, strKey_Name As String, strSub_Key_Ref As String, ByRef strKey_Val As String) As Boolean
        Dim plngLoop As Long            ' Loop Counter
        Dim plngReturn_Code As Long     ' Return Code
        Dim plngKey As Long             ' Handle To An Open Registry Key
        Dim plngKey_Val_Type As Long    ' Data Type Of A Registry Key
        Dim pstrTemp_Val As String      ' Tempory Storage For A Registry Key Value
        Dim plngKey_Val_Size As Long    ' Size Of Registry Key Variable
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : frmAbout.GetKeyValue (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & lngKey_Root & ", " & strKey_Name & ", " & strSub_Key_Ref
    End If
    '-v1.6.1
    '
        '------------------------------------------------------------
        ' Open RegKey Under lngKey_Root {mlngHKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        plngReturn_Code = advapiRegOpenKeyEx(lngKey_Root, strKey_Name, 0, mlngKEY_ALL_ACCESS, plngKey) ' Open Registry Key
        '
        ' Trap error
        If (plngReturn_Code <> mintERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
        '
        ' Initialize variables
        pstrTemp_Val = String$(1024, 0)                             ' Allocate Variable Space
        plngKey_Val_Size = 1024                                       ' Mark Variable Size
        '
        '------------------------------------------------------------
        ' Retrieve Registry Key Value...
        '------------------------------------------------------------
        plngReturn_Code = advapiRegQueryValueEx(plngKey, strSub_Key_Ref, 0, plngKey_Val_Type, pstrTemp_Val, plngKey_Val_Size)    ' Get/Create Key Value
        '
        ' Check for errors
        If (plngReturn_Code <> mintERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
        '
        ' Retrieve a substring from the value
        pstrTemp_Val = VBA.Left(pstrTemp_Val, InStr(pstrTemp_Val, VBA.Chr(0)) - 1)
        '
        '------------------------------------------------------------
        ' Determine Key Value Type For Conversion...
        '------------------------------------------------------------
        Select Case plngKey_Val_Type                                  ' Search Data Types...
            '
            ' String
            Case mintREG_SZ                                             ' String Registry Key Data Type
                strKey_Val = pstrTemp_Val                                     ' Copy String Value
            '
            ' Double Word
            Case mintREG_DWORD                                          ' Double Word Registry Key Data Type
                For plngLoop = Len(pstrTemp_Val) To 1 Step -1                    ' Convert Each Bit
                        strKey_Val = strKey_Val + Hex(Asc(Mid(pstrTemp_Val, plngLoop, 1)))   ' Build Value Char. By Char.
                Next
                strKey_Val = Format$("&h" + strKey_Val)                     ' Convert Double Word To String
        End Select
        '
        ' Return
        GetKeyValue = True                                      ' Return Success
        plngReturn_Code = advapiRegCloseKey(plngKey)                                  ' Close Registry Key
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmAbout.GetKeyValue (End)"
    '-v1.6.1
    '
        Exit Function                                           ' Exit
'
' Handle errors
GetKeyError:    ' Cleanup After An Error Has Occured...
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmAbout.GetKeyValue"
    '-v1.6.1
    '
        strKey_Val = ""                                             ' Set Return Val To Empty String
        GetKeyValue = False                                     ' Return Failure
        plngReturn_Code = advapiRegCloseKey(plngKey)                                  ' Close Registry Key
End Function
'
' EVENT:    picIcon_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Display the names of the developers
' TRIGGER:  The user clicked on the icon
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub picIcon_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmAbout.picIcon Click (Start)"
    '-v1.6.1
    '
    MsgBox "CCAT Developed by Tom Elkins and Brad Brown" & vbCrLf & "Tape operations by Keith Gibby", vbOKOnly Or vbInformation, "Developers"
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmAbout.picIcon Click (End)"
    '-v1.6.1
    '
End Sub
