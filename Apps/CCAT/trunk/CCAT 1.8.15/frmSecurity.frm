VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSecurity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Classification Banner"
   ClientHeight    =   1410
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3330
   Icon            =   "frmSecurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlBanners 
      Left            =   1320
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   115
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":0442
            Key             =   "UNCLASSIFIED"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":09FE
            Key             =   "CONFIDENTIAL"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":2C5A
            Key             =   "CONFIDENTIALSAR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":56F6
            Key             =   "SECRET"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":5CB2
            Key             =   "SECRETSAR"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":626E
            Key             =   "SECRETSCI"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":682A
            Key             =   "SECRETSARSCI"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":6EEE
            Key             =   "TOPSECRET"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":8D2A
            Key             =   "TOPSECRETSAR"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":B08E
            Key             =   "TOPSECRETSCI"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":D4FA
            Key             =   "TOPSECRETSARSCI"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop Session"
      Height          =   375
      Left            =   2052
      TabIndex        =   1
      Top             =   936
      Width           =   1215
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   72
      TabIndex        =   0
      Top             =   936
      Width           =   1215
   End
   Begin VB.Label lblSCI 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "SCI"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   336
      Left            =   36
      TabIndex        =   3
      Top             =   540
      Width           =   3252
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblClass"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   108
      TabIndex        =   2
      Top             =   180
      Width           =   3168
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' FORM:     frmSecurity
' AUTHOR:   Tom Elkins
' PURPOSE:  Displays a classification banner and stores security-related routines
' NOTES:    The banner is shown if the classification level changes.  The databases
'           have a classification property, which will be checked when loaded.
'           If the classification changes, the security banner will be displayed,
'           showing the user the new level.  The user can accept the new level or
'           terminate the program.
' REVISIONS:
'   v1.3.0  TAE Replaced token control with INI routines
'   v1.4.0  TAE Added code to check for additional paths for the INI file
'           TAE Added code to allow the user to browse for a missing INI file
'           TAE Modified code to write a default INI file
'   v1.5.0  TAE Removed the "What's This" help because it was not working correctly with HTML help
'               Set the help context for the form so that the user can press F1 to get the
'               help file for this form
'   v1.6.0  TAE Added routine to populate a supplied combo box with the available classification list
'           TAE Updated variable naming convention
'   v1.6.1  TAE Added verbose logging calls
'               Modified classification combo box population to remove duplicates
Option Explicit
'
Private Declare Function knlGetINIString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal strSection As String, ByVal strKeyName As Any, ByVal strDefault As String, ByVal strReturnedString As String, ByVal lngSize As Long, ByVal strINIFileName As String) As Long
Private Declare Function knlGetININum Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal strSection As String, ByVal strKeyName As String, ByVal lngDefault As Long, ByVal strINIFileName As String) As Long
Private Declare Function knlPutINIString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal strSection As String, ByVal strKeyName As Any, ByVal strDefault As Any, ByVal strINIFileName As String) As Long
'
'+v1.6.1TE
Private Declare Function knlGetINIList Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal strSection As String, ByVal strBuffer As String, ByVal lngSize As Long, ByVal strINIFileName As String) As Long
'-v1.6.1
'
' Constants
Const mstrSECURITY_INI = "security.ini" ' INI file to use
'
Private mstrINI_Path As String        ' Stores the path to the INI file
'
' EVENT:    btnAccept_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Hide the form
' TRIGGER:  User clicks on the "Accept" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnAccept_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSecurity.btnAccept Click (Start)"
    '-v1.6.1
    '
    ' Log the event
    basCCAT.WriteLogEntry "INFO     : frmSecurity.btnAccept Click  (Current security level accepted)"
    '
    ' Hide the form
    frmSecurity.Hide
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSecurity.btnAccept Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnStop_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Terminate the program
' TRIGGER:  User clicks on the "Stop" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnStop_Click()
    Dim pintOption As Integer
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSecurity.btnStop Click (Start)"
    '-v1.6.1
    '
    ' Ask the user again
    pintOption = MsgBox("Are you sure you want to end this session?", vbYesNo, "Program Stop")
    '
    ' Terminate on a "yes" response
    If pintOption = vbYes Then
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : frmSecurity.btnStop Click (User requested a program halt)"
        frmSecurity.Hide
        Unload frmMain
        End
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSecurity.btnStop Click (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  InitializeSecurity
' AUTHOR:   Tom Elkins
' PURPOSE:  Initialize any security functions or values
' INPUT:    None
' OUTPUT:   None
' NOTES:    Currently the routine triggers the Form_Load event if the form was not
'           already loaded.  This will ensure that the token file was loaded.
Public Sub InitializeSecurity()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmSecurity.InitializeSecurity (Start)"
    '-v1.6.1
    '
    ' Cause a Form_Load event if it has not happened yet.
    frmSecurity.Refresh
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmSecurity.InitializeSecurity (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Activate
' AUTHOR:   Tom Elkins
' PURPOSE:  Set up the text, background color, and foreground color
' TRIGGER:  When frmSecurity becomes the active form.
' INPUT:    None
' OUTPUT:   None
' NOTES:    The current classification level is stored in the form's Tag property
'           as an integer.  The mapping from integer to classification colors and
'           text is laid out in the token file.
Private Sub Form_Activate()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSecurity Activate (Start)"
    '-v1.6.1
    '
    ' Use form-level addressing
    With frmSecurity
        '
        ' Set the background and foreground colors to the appropriate values
        ' based on the current classification level and the mapping in the
        ' token file.
        .BackColor = .lngGetSecurityBackColor(CInt(Val(.Tag)))
        .lblClass.ForeColor = .lngGetSecurityForeColor(CInt(Val(.Tag)))
        '
        ' Set the classification text
        .lblClass.Caption = .strGetSecurityText(CInt(Val(.Tag)))
        '
        ' See if the classification level requires an SCI banner
        If .Tag And .lngGetNumber("Classification bit masks", "BIT_SCI", 2) Then
            '
            ' Configure and make visible the SCI banner
            .lblSCI.Visible = True
            .lblSCI.Caption = .strGetAlias("Classification text", "SCI_TXT", "SCI")
            Replace .lblClass.Caption, "/SCI", ""
        Else
            '
            ' Hide the SCI banner
            .lblSCI.Visible = False
        End If
    End With
    '
    '+v1.6TE
    ''+v1.5
    '' Set the help context
    'frmSecurity.HelpContextID = basCCAT.lGetHelpID(frmSecurity.Name)
    frmSecurity.HelpContextID = basCCAT.IDH_GUI_SECURITY
    ''-v1.5
    '-v1.6
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSecurity Activate (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: lngGetSecurityBackColor
' AUTHOR:   Tom Elkins
' PURPOSE:  Returns an RGB color value corresponding to the background color for
'           the specified classification level
'           level
' INPUT:    "intLevel" is the classification level
' OUTPUT:   A long integer specifying a color in RGB binary format
' NOTES:    The mapping of classification level to banner color is specified
'           in the token file, which is DII COE compliant
Public Function lngGetSecurityBackColor(intLevel As Integer) As Long
    Dim plngColor As Long  ' Hold the temporary color value
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : frmSecurity.lngGetSecurityBackColor Click (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intLevel
    End If
    '-v1.6.1
    '
    '
    ' Use token control-level addressing
    With frmSecurity
        '
        ' By default, get the Unclassified color definition
        plngColor = .lngGetNumber("Classification banner colors", "RGB_UNCLASSIFIED", 2263842)
        '
        ' Check for the Confidential bit and change the color if it exists
        If intLevel And .lngGetNumber("Classification bit masks", "BIT_CONFIDENTIAL", 4) Then _
            plngColor = .lngGetNumber("Classification banner colors", "RGB_CONFIDENTIAL", 15453831)
        '
        ' Check for the Secret bit and change the color if it exists
        If intLevel And .lngGetNumber("Classification bit masks", "BIT_SECRET", 8) Then _
            plngColor = .lngGetNumber("Classification banner colors", "RGB_SECRET", 2895086)
        '
        ' Check for the Top Secret bit and change the color if it exists
        If intLevel And .lngGetNumber("Classification bit masks", "BIT_TOPSECRET", 16) Then _
            plngColor = .lngGetNumber("Classification banner colors", "RGB_TOPSECRET", 36095)
    End With
    '
    ' Return the color
    lngGetSecurityBackColor = plngColor
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSecurity.lngGetSecurityBackColor = " & plngColor & " (End)"
    '-v1.6.1
    '
End Function
'
' FUNCTION: lngGetSecurityForeColor
' AUTHOR:   Tom Elkins
' PURPOSE:  Returns an RGB color value corresponding to the text color for the
'           specified classification level
' INPUT:    "intLevel" is the requested classification level
' OUTPUT:   A long integer giving the RGB binary value for the text color
' NOTES:    The mapping of classification level to text color is specified
'           in the token file
Public Function lngGetSecurityForeColor(intLevel As Integer) As Long
    Dim plngColor As Long  ' Hold the temporary color value
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : frmSecurity.lngGetSecurityForeColor (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intLevel
    End If
    '-v1.6.1
    '
    ' Use token control-level addressing
    With frmSecurity
        '
        ' By default, set the text color to White (for Unclassified text)
        plngColor = vbWhite
        '
        ' Check for the Confidential bit and change the color if necessary
        If intLevel And .lngGetNumber("Classification bit masks", "BIT_CONFIDENTIAL", 4) Then _
            plngColor = vbBlack
        '
        ' Check for the Secret bit and change the color if necessary
        If intLevel And .lngGetNumber("Classifiction bit masks", "BIT_SECRET", 8) Then _
            plngColor = vbWhite
        '
        ' Check for the Top Secret bit and change the color if necessary
        If intLevel And .lngGetNumber("Classification bit masks", "BIT_TOPSECRET", 16) Then _
            plngColor = vbBlack
    End With
    '
    ' Return the text color
    lngGetSecurityForeColor = plngColor
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmSecurity.lngGetSecurityForeColor = " & plngColor & " (End)"
    '-v1.6.1
    '
End Function
'
' FUNCTION: strGetSecurityText
' AUTHOR:   Tom Elkins
' PURPOSE:  Returns the classification text description for the specified level
' INPUT:    "intLevel" is the requested classification level
' OUTPUT:   A string giving the classification text to be used in a banner
' NOTES:    The mapping of classification level to classification text is specified
'           in the token file
Public Function strGetSecurityText(intLevel As Integer) As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : frmSecurity.strGetSecurityText (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intLevel
    End If
    '-v1.6.1
    '
    ' Retrieve the classification text from the token file
    strGetSecurityText = frmSecurity.strGetAlias("Classification text", "SECURITY_TXT" & intLevel, 0)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmSecurity.strGetSecurityText (End)"
    '-v1.6.1
    '
End Function
'
' EVENT:    Form_Load
' AUTHOR:   Tom Elkins
' PURPOSE:  Load the token file
' TRIGGER:  The first time the form is referenced
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSecurity Load (Start)"
    '-v1.6.1
    '
    ' Initialize the security level (0 = Unclassified)
    If frmSecurity.Tag = "" Then frmSecurity.Tag = 0
    '
    ' Set default path to INI file
    mstrINI_Path = App.Path & DAS_TOKEN_PATH & mstrSECURITY_INI
    '
    ' Check for file existence
    If Dir(mstrINI_Path) = "" Then
        '
        ' Change to secondary path
        mstrINI_Path = App.Path & "\" & mstrSECURITY_INI
        '
        ' Check for file existence
        If Dir(mstrINI_Path) = "" Then
            '
            ' Ask user to specify path
            With frmMain.dlgCommonDialog
                '
                .CancelError = False
                .DialogTitle = "Find Security INI file"
                .FileName = ""
                .Filter = "Security settings file|security.ini"
                .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
                .InitDir = App.Path
                .ShowOpen
                '
                ' See if user found the file
                If .FileName <> "" Then
                    '
                    ' Copy the file to an appropriate place
                    FileCopy .FileName, mstrINI_Path
                Else
                    '
                    ' Create a default version of the file
                    frmSecurity.CreateSecurityTokenFile
                End If
            End With
        End If
    End If
    '
    ' Update the classification level
    frmSecurity.SetClassification frmSecurity.lngGetNumber("Classification", mstrSECURITY_INI & "_CLASS", 0), "Security INI File"
    '
    ' Assign help topics
    '
    '+v1.5
    ' Set help context
    'frmSecurity.btnAccept.WhatsThisHelpID = basCCAT.GetNumber("Help Map", "IDH_Security", 0)
    'frmSecurity.btnStop.WhatsThisHelpID = frmSecurity.btnAccept.WhatsThisHelpID
    'frmSecurity.lblClass.WhatsThisHelpID = frmSecurity.btnAccept.WhatsThisHelpID
    'frmSecurity.lblSCI.WhatsThisHelpID = frmSecurity.btnAccept.WhatsThisHelpID
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSecurity Load (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  SetClassification
' AUTHOR:   Tom Elkins
' PURPOSE:  Update the global classification level
' INPUT:    "intNew_Level" is the new classification level being added to the existing
'           system.
'           "strSource" is a string to tell the user what is driving the change
' OUTPUT:   None
' NOTES:    Security Levels and bit masks are defined in the token file.
'           Classification levels are ORed, not added.  If two level-6 files are loaded
'           into the translator, ORing the values results in a level-6 classification;
'           however, if they are added, we get level-12 classification, which is wrong
'           Classification level is stored in the Tag property of the form as an
'           integer
Public Sub SetClassification(intNew_Level As Integer, strSource As String)
    Dim pintOld_Level As Integer   ' Used to check the old level
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmSecurity.SetClassification (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intNew_Level & ", " & strSource
    End If
    '-v1.6.1
    '
    ' Store the old level
    pintOld_Level = Val(Trim(frmSecurity.Tag))
    '
    ' Log the event
    If pintOld_Level <> intNew_Level Then
        basCCAT.WriteLogEntry "INFO     : frmSecurity.SetClassification (Was " & frmSecurity.lngGetNumber("Classification Text", "SECURITY_TXT" & pintOld_Level, 0) & ")"
        basCCAT.WriteLogEntry "INFO     : frmSecurity.SetClassification (Now " & frmSecurity.lngGetNumber("Classification Text", "SECURITY_TXT" & intNew_Level, 0) & ")"
    End If
    '
    ' OR the new level with the old level
    frmSecurity.Tag = pintOld_Level Or intNew_Level
    '
    ' Check for a change in security level
    If frmSecurity.Tag <> pintOld_Level Then
        '
        ' Change the form title
        frmSecurity.Caption = strSource
        '
        ' Display the banner
        frmSecurity.Show vbModal
    End If
    '
    ' Update the image in the staus bar
    frmMain.sbStatusBar.Panels("SECURITY").Picture = frmSecurity.imlBanners.ListImages(frmSecurity.strGetAlias("Images", "SECURITYPIC" & frmSecurity.Tag, "UNCLASSIFIED")).Picture
    '
    ' Update the tooltip for the status bar
    frmMain.sbStatusBar.Panels("SECURITY").ToolTipText = "Current Security Level: " & frmSecurity.strGetAlias("Classification Text", "SECURITY_TXT" & frmSecurity.Tag, "UNCLASSIFIED")
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmSecurity.SetClassification (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  CreateSecurityTokenFile
' AUTHOR:   Tom Elkins
' PURPOSE:  Creates the security token file
' INPUT:    None
' OUTPUT:   None
' NOTES:    If the executable is moved to a new location without the token files, or
'           the token files are deleted, this routine will create the default security
'           token file for the program to use.
Public Sub CreateSecurityTokenFile()
    Dim pintFile As Integer    ' Token file handle
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmSecurity.CreateSecurityTokenFile (Start)"
    '-v1.6.1
    '
    '
    ' Log the event
    basCCAT.WriteLogEntry "INFO     : frmSecurity.CreateSecurityTokenFile (Token file did not exist)"
    '
    ' Assign the file identifier
    pintFile = FreeFile
    '
    ' Create the file
    Open mstrINI_Path For Output As pintFile
    '
    ' Write the contents
    Print #pintFile, ";Security Tokens file"
    Print #pintFile, ";These tokens establish the details for displaying classification banners"
    Print #pintFile, ";DO NOT MODIFY unless you are confident of altering these files."
    Print #pintFile, ";"
    Print #pintFile, "[File]"
    Print #pintFile, "; <filename>=<success message>=<1=true>"
    Print #pintFile, "security.ini_FILE=1"
    Print #pintFile, ";"
    Print #pintFile, "[Classification]"
    Print #pintFile, "; <filename>=<classification>=<classification ID>"
    Print #pintFile, "security.tok_CLASS=0"
    Print #pintFile, ";"
    Print #pintFile, "[Classification bit masks]"
    Print #pintFile, ";    0 = UNCLASSIFIED"
    Print #pintFile, ";   +1 = SAR"
    Print #pintFile, ";   +2 = SCI"
    Print #pintFile, ";   +4 = CONFIDENTIAL"
    Print #pintFile, ";   +8 = SECRET"
    Print #pintFile, ";   +16 = TOP SECRET"
    Print #pintFile, "BIT_UNCLASSIFIED=0"
    Print #pintFile, "BIT_SAR=1"
    Print #pintFile, "BIT_SCI=2"
    Print #pintFile, "BIT_CONFIDENTIAL=4"
    Print #pintFile, "BIT_SECRET=8"
    Print #pintFile, "BIT_TOPSECRET=16"
    Print #pintFile, ";"
    Print #pintFile, "[Classification banner colors]"
    Print #pintFile, "; Defined by DII UIS v3.0 (2/98)"
    Print #pintFile, "RGB_UNCLASSIFIED=2263842"
    Print #pintFile, "RGB_CONFIDENTIAL=15453831"
    Print #pintFile, "RGB_SECRET=2895086"
    Print #pintFile, "RGB_TOPSECRET=36095"
    Print #pintFile, ";"
    Print #pintFile, "[Classification Text]"
    Print #pintFile, "; Classification text -- text used to describe the level of classification"
    Print #pintFile, "SCI_TXT=SCI"
    Print #pintFile, "SECURITY_TXT0=UNCLASSIFIED"
    Print #pintFile, "SECURITY_TXT1=UNCLASSIFIED/HVSACO"
    Print #pintFile, "SECURITY_TXT2=UNCLASSIFIED/SCI"
    Print #pintFile, "SECURITY_TXT3=UNCLASSIFIED/SAR/SCI"
    Print #pintFile, "SECURITY_TXT4=CONFIDENTIAL"
    Print #pintFile, "SECURITY_TXT5=CONFIDENTIAL/SAR"
    Print #pintFile, "SECURITY_TXT6=CONFIDENTIAL/SCI"
    Print #pintFile, "SECURITY_TXT7=CONFIDENTIAL/SAR/SCI"
    Print #pintFile, "SECURITY_TXT8=SECRET"
    Print #pintFile, "SECURITY_TXT9=SECRET/SAR"
    Print #pintFile, "SECURITY_TXT10=SECRET/SCI"
    Print #pintFile, "SECURITY_TXT11=SECRET/SAR/SCI"
    Print #pintFile, "SECURITY_TXT12=SECRET"
    Print #pintFile, "SECURITY_TXT13=SECRET/SAR"
    Print #pintFile, "SECURITY_TXT14=SECRET/SCI"
    Print #pintFile, "SECURITY_TXT15=SECRET/SAR/SCI"
    Print #pintFile, "SECURITY_TXT16=TOP SECRET"
    Print #pintFile, "SECURITY_TXT17=TOP SECRET/SAR"
    Print #pintFile, "SECURITY_TXT18=TOP SECRET/SCI"
    Print #pintFile, "SECURITY_TXT19=TOP SECRET/SAR/SCI"
    Print #pintFile, "SECURITY_TXT20=TOP SECRET"
    Print #pintFile, "SECURITY_TXT21=TOP SECRET/SAR"
    Print #pintFile, "SECURITY_TXT22=TOP SECRET/SCI"
    Print #pintFile, "SECURITY_TXT23=TOP SECRET/SAR/SCI"
    Print #pintFile, "SECURITY_TXT24=TOP SECRET"
    Print #pintFile, "SECURITY_TXT25=TOP SECRET/SAR"
    Print #pintFile, "SECURITY_TXT26=TOP SECRET/SCI"
    Print #pintFile, "SECURITY_TXT27=TOP SECRET/SAR/SCI"
    Print #pintFile, "SECURITY_TXT28=TOP SECRET"
    Print #pintFile, "SECURITY_TXT29=TOP SECRET/SAR"
    Print #pintFile, "SECURITY_TXT30=TOP SECRET/SCI"
    Print #pintFile, "SECURITY_TXT31=TOP SECRET/SAR/SCI"
    Print #pintFile, ";"
    Print #pintFile, "[Images]"
    Print #pintFile, "; SECURITY_PIC_MAX=<image token>=<max number>"
    Print #pintFile, "; <image token><classification ID>=<image key>=<image index>"
    Print #pintFile, "SECURITY_PIC_MAX=31"
    Print #pintFile, "SECURITYPIC0=UNCLASSIFIED"
    Print #pintFile, "SECURITYPIC1=UNCLASSIFIED"
    Print #pintFile, "SECURITYPIC2=UNCLASSIFIED"
    Print #pintFile, "SECURITYPIC3=UNCLASSIFIED"
    Print #pintFile, "SECURITYPIC4=CONFIDENTIAL"
    Print #pintFile, "SECURITYPIC5=CONFIDENTIALSAR"
    Print #pintFile, "SECURITYPIC6=CONFIDENTIAL"
    Print #pintFile, "SECURITYPIC7=CONFIDENTIALSAR"
    Print #pintFile, "SECURITYPIC8=SECRET"
    Print #pintFile, "SECURITYPIC9=SECRETSAR"
    Print #pintFile, "SECURITYPIC10=SECRETSCI"
    Print #pintFile, "SECURITYPIC11=SECRETSARSCI"
    Print #pintFile, "SECURITYPIC12=SECRET"
    Print #pintFile, "SECURITYPIC13=SECRETSAR"
    Print #pintFile, "SECURITYPIC14=SECRETSCI"
    Print #pintFile, "SECURITYPIC15=SECRETSARSCI"
    Print #pintFile, "SECURITYPIC16=TOPSECRET"
    Print #pintFile, "SECURITYPIC20=TOPSECRET"
    Print #pintFile, "SECURITYPIC24=TOPSECRET"
    Print #pintFile, "SECURITYPIC28=TOPSECRET"
    Print #pintFile, "SECURITYPIC17=TOPSECRETSAR"
    Print #pintFile, "SECURITYPIC21=TOPSECRETSAR"
    Print #pintFile, "SECURITYPIC25=TOPSECRETSAR"
    Print #pintFile, "SECURITYPIC29=TOPSECRETSAR"
    Print #pintFile, "SECURITYPIC18=TOPSECRETSCI"
    Print #pintFile, "SECURITYPIC22=TOPSECRETSCI"
    Print #pintFile, "SECURITYPIC26=TOPSECRETSCI"
    Print #pintFile, "SECURITYPIC30=TOPSECRETSCI"
    Print #pintFile, "SECURITYPIC19=TOPSECRETSARSCI"
    Print #pintFile, "SECURITYPIC23=TOPSECRETSARSCI"
    Print #pintFile, "SECURITYPIC27=TOPSECRETSARSCI"
    Print #pintFile, "SECURITYPIC31=TOPSECRETSARSCI"
    Print #pintFile, ";"
    Print #pintFile, "[Security Values from text]"
    Print #pintFile, ";SECURITY_VAL_<Text>=<Classification text>=<Security ;>"
    Print #pintFile, "SECURITY_VAL_UNCLASSIFIED=0"
    Print #pintFile, "SECURITY_VAL_UNCLASSIFIED/SAR=1"
    Print #pintFile, "SECURITY_VAL_UNCLASSIFIED/HVSACO=1"
    Print #pintFile, "SECURITY_VAL_CONFIDENTIAL=4"
    Print #pintFile, "SECURITY_VAL_CONFIDENTIAL/SAR=5"
    Print #pintFile, "SECURITY_VAL_SECRET=8"
    Print #pintFile, "SECURITY_VAL_SECRET/SAR=9"
    Print #pintFile, "SECURITY_VAL_SECRET/SCI=10"
    Print #pintFile, "SECURITY_VAL_SECRET/SAR/SCI=11"
    Print #pintFile, "SECURITY_VAL_TOP SECRET=16"
    Print #pintFile, "SECURITY_VAL_TOP SECRET/SAR=17"
    Print #pintFile, "SECURITY_VAL_TOP SECRET/SCI=18"
    Print #pintFile, "SECURITY_VAL_TOP SECRET/SAR/SCI=19"
    '
    ' Close the file
    Close pintFile
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmSecurity.CreateSecurityTokenFile (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: GetNumber
' AUTHOR:   Tom Elkins
' PURPOSE:  Replaces the token file method to use INI files
' INPUT:    "strSection" is the [<name>] portion of the INI file
'           "strToken" is the string to be replaced
'           "lngDefault_Value" is the number to be returned if the token does not exist
' OUTPUT:   "GetNumber" is the number found in the INI file for the specified token
' NOTES:
Public Function lngGetNumber(strSection As String, strToken As String, lngDefault_Value As Long) As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : frmSecurity.lngGetNumber (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & strSection & ", " & strToken & ", " & lngDefault_Value
    End If
    '-v1.6.1
    '
    lngGetNumber = knlGetININum(strSection, strToken, lngDefault_Value, mstrINI_Path)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmSecurity.lngGetNumber (End)"
    '-v1.6.1
    '
End Function
'
' FUNCTION: strGetAlias
' AUTHOR:   Tom Elkins
' PURPOSE:  Replaces the token file method to use INI files
' INPUT:    "strSection" is the [<name>] portion of the INI file
'           "strToken" is the string to be replaced
'           "strDefault_Value" is the string to be returned if the token does not exist
' OUTPUT:   "strGetAlias" is the string found in the INI file for the specified token
' NOTES:
Public Function strGetAlias(strSection As String, strToken As String, strDefault_Value As String) As String
    Dim pstrToken As String        ' Return string buffer
    Dim plngToken_Len As Long       ' Length of return string
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmSecurity.strGetAlias(" & strSection & ", " & strToken & ", " & strDefault_Value & ")"
    '-v1.6.1
    '
    '
    ' Set the buffer to all spaces
    pstrToken = String(255, " ")
    '
    ' Get the return string for the token
    plngToken_Len = knlGetINIString(strSection, strToken, strDefault_Value, pstrToken, 255, mstrINI_Path)
    '
    ' Return the specified number of characters from the buffer
    strGetAlias = Left(pstrToken, plngToken_Len)
End Function
'
'+v1.6TE
' ROUTINE:  PopulateClassifications
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds the known security classifications to a specified combo box
' INPUT:    "cmbList" is the combo box that will be populated
' OUTPUT:   None - the combo box is altered directly
' NOTES:    This will ensure that all classification combo boxes will have the same
'           entries and values.
Public Sub PopulateClassification(cmbList As ComboBox)
    Dim pintList As Integer
    Dim pstrVal As String
    '
    '+v1.6.1TE
    Dim pastrValues() As String
    '-v1.6.1
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmSecurity.PopulateClassification (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & cmbList.Name
    End If
    '-v1.6.1
    '
    ' Initialize the list counter and add a default blank entry to the list
    pintList = 0
    cmbList.AddItem ""
    '
    '+v1.6.1TE
    ' Get the list from the INI file
    pstrVal = String(32000, " ")
    knlGetINIList "Security values from text", pstrVal, CLng(32000), mstrINI_Path
    pastrValues = Split(pstrVal, Chr(0))
    '
    ' Loop through the entries
    For pintList = LBound(pastrValues) To UBound(pastrValues) - 2
        '
        If pastrValues(pintList) <> "" Then cmbList.AddItem Mid(Trim(pastrValues(pintList)), 14, InStr(1, pastrValues(pintList), "=") - 14)
    Next pintList
    ''
    '' Get the first security value from the INI file
    'pstrVal = frmSecurity.strGetAlias("Classification Text", "SECURITY_TXT" & pintList, "DONE")
    ''
    '' Loop until there are no more entries
    'While pstrVal <> "DONE"
    '    '
    '    ' Add the entry to the list
    '    cmbList.AddItem pstrVal
    '    '
    '    ' Update the counter
    '    pintList = pintList + 1
    '    '
    '    ' Get the next entry
    '    pstrVal = frmSecurity.strGetAlias("Classification Text", "SECURITY_TXT" & pintList, "DONE")
    'Wend
    '-v1.6.1
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmSecurity.PopulateClassification (End)"
    '-v1.6.1
    '
End Sub
'-v1.6
'
