VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Quick Conversion"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2925
   HelpContextID   =   858
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   2925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtComp 
      Height          =   285
      Index           =   3
      Left            =   1995
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   345
      Width           =   855
   End
   Begin VB.TextBox txtComp 
      Height          =   285
      Index           =   2
      Left            =   1365
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   345
      Width           =   540
   End
   Begin VB.TextBox txtComp 
      Height          =   285
      Index           =   1
      Left            =   735
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   345
      Width           =   540
   End
   Begin VB.TextBox txtComp 
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   345
      Width           =   540
   End
   Begin VB.TextBox txtFloat 
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   975
      Width           =   2745
   End
   Begin VB.Label lblComp 
      Alignment       =   2  'Center
      Caption         =   "TSecs"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   12
      Top             =   750
      Width           =   2820
   End
   Begin VB.Label lblComp 
      Alignment       =   2  'Center
      Caption         =   "Seconds"
      Height          =   195
      Index           =   3
      Left            =   1995
      TabIndex        =   11
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblComp 
      Alignment       =   2  'Center
      Caption         =   "Min"
      Height          =   195
      Index           =   2
      Left            =   1395
      TabIndex        =   10
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblComp 
      Alignment       =   2  'Center
      Caption         =   "Hour"
      Height          =   195
      Index           =   1
      Left            =   765
      TabIndex        =   9
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblComp 
      Alignment       =   2  'Center
      Caption         =   "Day"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   8
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblColon 
      AutoSize        =   -1  'True
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1905
      TabIndex        =   7
      Top             =   315
      Width           =   90
   End
   Begin VB.Label lblColon 
      AutoSize        =   -1  'True
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1275
      TabIndex        =   6
      Top             =   315
      Width           =   90
   End
   Begin VB.Label lblColon 
      AutoSize        =   -1  'True
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   645
      TabIndex        =   4
      Top             =   330
      Width           =   90
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' FORM:     frmConvert
' AUTHOR:   Tom Elkins
' PURPOSE:  Interface to do quick conversion calculations for time and angle measurements
' REVISIONS:
'   v1.4.0  TAE Created
'   v1.5.0  TAE Added context-sensitive help
'   v1.6.0  TAE Modified names to meet programming convention
'   v1.6.1  TAE Added verbose logging calls
Option Explicit
'
' Constants
Const mintDAY = 0                       ' Day component box index
Const mintDEG_HR = 1                    ' Deg/Hour component box index
Const mintMINUTE = 2                    ' Minute component box index
Const mintSECOND = 3                    ' Second component box index
Const mintFLOAT = 4                     ' Floating point total box index
Const mdblSEC_PER_DAY = 86400#          ' Seconds in a day
Const mdblSEC_PER_HR = 3600#            ' Seconds in an hour
Const mdblSEC_PER_DEG = 3600#           ' Seconds in a degree
Const mdblSEC_PER_MIN = 60#             ' Seconds in a minute
Const mdblMIN_PER_DEG = 60#             ' Minutes in a degree
'
' Local variables
Private mblnConvert_Time As Boolean     ' True if converting time values
Private mblnUse_Components As Boolean   ' True if the components are being entered
'
' PROPERTY: InTimeMode
' AUTHOR:   Tom Elkins
' PURPOSE:  Set/Retrieve the state of the mblnConvert_Time Flag
' INPUT:    LET: "blnMode" is True for time conversion, False for angle conversion
' OUTPUT:   GET: "InTimeMode" is True for time conversion, False for angle conversion
' NOTES:
Public Property Let InTimeMode(blnMode As Boolean)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: frmConvert.InTimeMode Let (Start) = " & blnMode
    '-v1.6.1
    '
    mblnConvert_Time = blnMode
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: frmConvert.InTimeMode Let (End)"
    '-v1.6.1
    '
End Property
'
Public Property Get InTimeMode() As Boolean
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: frmConvert.InTimeMode Get (Start)"
    '-v1.6.1
    '
    InTimeMode = mblnConvert_Time
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: frmConvert.InTimeMode Get (End)"
    '-v1.6.1
    '
End Property
'
' EVENT:    Form_Activate
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the interface for the correct mode
' TRIGGER:  Any time the form is called
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Activate()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert Activate (Start)"
    '-v1.6.1
    '
    ' Check the mode
    If frmConvert.InTimeMode Then
        '
        ' Configure the interface for time conversion
        ' Show day component
        frmConvert.lblComp(mintDAY).Caption = "Day"
        frmConvert.txtComp(mintDAY).Visible = True
        frmConvert.lblColon(mintDAY).Visible = True
        '
        ' Change labels to time values
        frmConvert.lblComp(mintDEG_HR).Caption = "Hour"
        frmConvert.lblComp(mintFLOAT).Caption = "TSecs"
    Else
        '
        ' Configure the interface for angle conversion
        ' No day component, so hide it
        frmConvert.lblComp(mintDAY).Caption = ""
        frmConvert.txtComp(mintDAY).Text = ""
        frmConvert.txtComp(mintDAY).Visible = False
        frmConvert.lblColon(mintDAY).Visible = False
        '
        ' Change labels to degrees
        frmConvert.lblComp(mintDEG_HR).Caption = "Deg"
        frmConvert.lblComp(mintFLOAT).Caption = "Degrees"
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert Activate (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Load
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets the defaults for the form
' TRIGGER:  The first time the form is called
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    Dim ptxtBox As TextBox   ' Reference to all the component text boxes on the form
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert Load (Start)"
    '-v1.6.1
    '
    ' Loop through all the component text boxes on the form
    For Each ptxtBox In frmConvert.txtComp
        '
        ' Clear the contents
        ptxtBox.Text = ""
    Next ptxtBox
    '
    ' Clear the floating point box
    frmConvert.txtFloat.Text = ""
    '
    ' Set the labels for the minutes and seconds (same for either conversion mode)
    frmConvert.lblComp(mintMINUTE).Caption = "Min"
    frmConvert.lblComp(mintSECOND).Caption = "Seconds"
    '
    '+v1.5
    ' Set the help context
    frmConvert.HelpContextID = basCCAT.IDH_GUI_CONVERT
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert Load (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtComp_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Computes the floating point version of the components
' TRIGGER:  User entered text in any of the component text boxes
' INPUT:    "intField" is the reference to the specific component text box modified
' OUTPUT:   None
' NOTES:
Private Sub txtComp_Change(intField As Integer)
    Dim pdblResult As Double   ' The computed floating point result
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmConvert.txtComp Change (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intField
    End If
    '-v1.6.1
    '
    ' See if the user was editing the component box or if it was done by the
    ' conversion calculations
    If mblnUse_Components Then
        '
        ' Compute the floating point value using the components
        '           [               Hours/Degrees                    ]   [             Minutes/Arcminutes               ]   [           Seconds/Arcseconds         ]
        pdblResult = (Val(frmConvert.txtComp(mintDEG_HR).Text) * mdblSEC_PER_HR) + (Val(frmConvert.txtComp(mintMINUTE).Text) * mdblSEC_PER_MIN) + Val(frmConvert.txtComp(mintSECOND).Text)
        '
        ' Result is now in seconds/arcseconds
        '
        ' Check to see if we are in time conversion mode
        If frmConvert.InTimeMode Then
            '
            ' Add Julian day to the total                 seconds + [   Days - 1 (no day 0)   * seconds/day]
            If Val(frmConvert.txtComp(mintDAY).Text) > 0 Then pdblResult = pdblResult + ((Val(frmConvert.txtComp(mintDAY).Text) - 1) * mdblSEC_PER_DAY)
        Else
            '
            ' We are in degree mode, so we convert from arcseconds to degrees
            ' by dividing by 3600 arcseconds/degree
            pdblResult = pdblResult / mdblSEC_PER_DEG
        End If
        '
        ' Put the result in the floating point text box
        frmConvert.txtFloat.Text = pdblResult
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert.txtComp Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtComp_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Set the component edit flag and highlight the text to be modified
' TRIGGER:  User entered one of the component text boxes
' INPUT:    "intField" is the reference to the specific text box being edited
' OUTPUT:   None
' NOTES:
Private Sub txtComp_GotFocus(intField As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmConvert.txtComp GotFocus (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intField
    End If
    '-v1.6.1
    '
    ' Set the component editing flag
    mblnUse_Components = True
    '
    ' Highlight the text in the box
    frmConvert.txtComp(intField).SelStart = 0
    frmConvert.txtComp(intField).SelLength = Len(frmConvert.txtComp(intField).Text)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert.txtComp GotFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtFloat_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Computes the components of the floating point value specified
' TRIGGER:  User modified the contents of the floating point box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub txtFloat_Change()
    Dim pdblResult As Double   ' Floating point result
    Dim pintVal As Integer     ' computed component
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert.txtFloat Change (Start)"
    '-v1.6.1
    '
    '
    ' Make sure we are not in component edit mode
    If Not mblnUse_Components Then
        '
        ' Copy the current contents of the floating point box
        pdblResult = CDbl(Val(frmConvert.txtFloat.Text))
        '
        ' See if we are in time conversion mode
        If frmConvert.InTimeMode Then
            '
            ' Compute the Day value
            ' The "\" means integer divide
            '      seconds \ 86400 seconds/day
            pintVal = pdblResult \ mdblSEC_PER_DAY
            '
            ' Subtract the day component from the time value
            pdblResult = pdblResult - (pintVal * mdblSEC_PER_DAY)
            '
            ' Add 1 to the day number (there is no day 0)
            frmConvert.txtComp(mintDAY).Text = pintVal + 1
            '
            ' Compute the Hour value
            '      seconds \ 3600 seconds/hour
            pintVal = pdblResult \ mdblSEC_PER_HR
            '
            ' Subtract the hour component from the time value
            pdblResult = pdblResult - (pintVal * mdblSEC_PER_HR)
            '
            ' Display the hour value
            frmConvert.txtComp(mintDEG_HR).Text = pintVal
            '
            ' Compute the Minute value
            '      seconds \ 60 seconds/minute
            pintVal = pdblResult \ mdblSEC_PER_MIN
            '
            ' Subtract the minute component from the time value
            pdblResult = pdblResult - (pintVal * mdblSEC_PER_MIN)
            '
            ' Display the minute value
            frmConvert.txtComp(mintMINUTE).Text = pintVal
            '
            ' Display the seconds value
            frmConvert.txtComp(mintSECOND).Text = pdblResult
        Else
            ' We are in degree conversion mode
            ' Ensure no day component
            frmConvert.txtComp(mintDAY).Text = 0
            '
            ' Extract the degrees
            pintVal = Int(pdblResult)
            '
            ' Display the degrees
            frmConvert.txtComp(mintDEG_HR).Text = pintVal
            '
            ' Subtract the degrees and convert to minutes
            pdblResult = (pdblResult - pintVal) * mdblMIN_PER_DEG
            '
            ' Extract the minutes
            pintVal = Int(pdblResult)
            '
            ' Display the minutes
            frmConvert.txtComp(mintMINUTE).Text = pintVal
            '
            ' Subtract the minutes and convert to seconds
            pdblResult = (pdblResult - pintVal) * mdblSEC_PER_MIN
            '
            ' Display the seconds
            frmConvert.txtComp(mintSECOND).Text = pdblResult
        End If
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert.txtFloat Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtFloat_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets the component edit flag to false and highlights the value
' TRIGGER:  User entered the floating point box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub txtFloat_GotFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert.txtFloat GotFocus (Start)"
    '-v1.6.1
    '
    ' Set the component edit flag
    mblnUse_Components = False
    '
    ' Highlight the contents of the box
    frmConvert.txtFloat.SelStart = 0
    frmConvert.txtFloat.SelLength = Len(frmConvert.txtFloat.Text)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmConvert.txtFloat GotFocus (End)"
    '-v1.6.1
    '
End Sub
