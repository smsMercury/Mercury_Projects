VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5700
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   5670
      ScaleWidth      =   8250
      TabIndex        =   0
      Top             =   0
      Width           =   8280
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Licensed To: United States Government"
         Height          =   195
         Left            =   5205
         TabIndex        =   7
         Tag             =   "LicenseTo"
         Top             =   0
         Width           =   2850
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DAS - Data Analysis System"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Tag             =   "CompanyProduct"
         Top             =   360
         Width           =   4845
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPASS CALL Archive Translator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   2235
         TabIndex        =   5
         Tag             =   "Product"
         Top             =   1080
         Width           =   4815
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Tag             =   "Warning"
         Top             =   4800
         Width           =   6855
      End
      Begin VB.Label lblPlatform 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "for Windows 95, 98, NT, and 2000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   3
         Tag             =   "Platform"
         Top             =   4365
         Width           =   4725
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
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
         Left            =   5640
         TabIndex        =   2
         Tag             =   "Version"
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developed by Mercury Solutions, Inc"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Tag             =   "Company"
         Top             =   5280
         Width           =   2610
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' FORM:     frmSplash
' AUTHOR:   Visual Basic Application Wizard
'           Tom Elkins
' PURPOSE:  Identifies the software and gives the user something to look at while the
'           main program is loading.
' REVISIONS:
'   v1.0.0  TAE Original code
'   v1.5.0  TAE Added context-sensitive help information
'   v1.6.0  TAE Updated splash screen contents
'   v1.6.1  TAE Added verbose logging calls
Option Explicit
'
' EVENT:    Form_Load
' AUTHOR:   Tom Elkins
' PURPOSE:  Size the form to the image, and set up the captions
' TRIGGER:  frmMain.Form_Load
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSplash Load (Start)"
    '-v1.6.1
    '
    '
    ' The image was loaded into a picture box at design time.  The picture box
    ' is positioned at the top left of the form, and the form is resized to the
    ' dimensions of the picture box.
    frmSplash.picLogo.Left = 0
    frmSplash.picLogo.Top = 0
    frmSplash.Width = frmSplash.picLogo.Width
    frmSplash.Height = frmSplash.picLogo.Height
    '
    ' Set the form border to fixed dialog
    frmSplash.BorderStyle = vbFixedDouble
    '
    ' Set up the captions
    frmSplash.lblCompany.Caption = "Copyright " & Chr(169) & " 1999-" & Year(Now) & " Mercury Solutions, Inc."
    frmSplash.lblCompanyProduct.Caption = "DAS - Data Analysis System"
    frmSplash.lblLicenseTo.Caption = "Licensed to: The United States Government"
    '
    '+v1.6TE
    'frmSplash.lblPlatform.Caption = "for Windows 95, 98, and NT"
    frmSplash.lblPlatform.Caption = "for Windows 95, 98, NT, and 2000"
    '-v1.6
    frmSplash.lblProductName.Caption = "COMPASS CALL Archive Translator"
    frmSplash.lblVersion.Caption = basCCAT.sGet_Version
    frmSplash.lblWarning.Caption = "Warning: This software is for official U.S. Government use only and is subject to export controls"
    '
    '+v1.5
    ' Set help context
    Me.HelpContextID = basCCAT.IDH_GUI_SPLASH
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmSplash Load (End)"
    '-v1.6.1
    '
End Sub
