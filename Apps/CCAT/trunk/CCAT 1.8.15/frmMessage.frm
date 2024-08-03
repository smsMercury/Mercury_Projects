VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Properties"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   1785
      TabIndex        =   6
      Top             =   1725
      Width           =   945
   End
   Begin VB.Frame fraDesc 
      Caption         =   "Description"
      Height          =   1230
      Left            =   120
      TabIndex        =   4
      Top             =   465
      Width           =   4380
      Begin VB.Label lblDesc 
         Caption         =   "lblDesc"
         Height          =   885
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4185
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblName"
      Height          =   255
      Left            =   2445
      TabIndex        =   3
      Top             =   150
      Width           =   2055
   End
   Begin VB.Label lblTxtName 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   1890
      TabIndex        =   2
      Top             =   165
      Width           =   420
   End
   Begin VB.Label lblID 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblID"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   135
      Width           =   1215
   End
   Begin VB.Label lblTxtID 
      AutoSize        =   -1  'True
      Caption         =   "ID"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   165
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' FORM:     frmMessage
' AUTHOR:   Tom Elkins
' PURPOSE:  Display information about a selected message
' REVISIONS:
'   v1.0.0  TAE Original code
'   v1.5.0  TAE Added context-sensitive help to the form
'           TAE Changed the way message properties are retrieved
'   v1.6.1  TAE Added verbose logging calls
'
'
Option Explicit
'
' ROUTINE:  DisplayMessageProperties
' AUTHOR:   Tom Elkins
' PURPOSE:  Displays some information about the selected message
' INPUT:    strMsg - Key for the selected message node/item
' OUTPUT:   None
' NOTES:
Public Sub DisplayMessageProperties(strMsg As String)
    Dim prsMsg As Recordset
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmMessage.DisplayMessageProperties (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & strMsg
    End If
    '-v1.6.1
    '
    ' Extract the message identifier
    frmMessage.lblID.Caption = basCCAT.iExtract_MessageID(strMsg)
    '
    '+v1.5
    ''
    '' Open a recordset to the Archive Summary table
    'Set prsMsg = guCurrent.DB.OpenRecordset("SELECT * FROM Archive" & basCCAT.iExtract_ArchiveID(strMsg) & "_Summary WHERE MSG_ID = " & frmMessage.lblID.Caption)
    ''
    '' Populate the form
    'frmMessage.lblDesc.Caption = prsMsg!Description
    'frmMessage.lblName.Caption = prsMsg!Message
    frmMessage.lblName.Caption = basCCAT.GetAlias("Message Names", "CC_MSGID" & Me.lblID.Caption, "(UNKNOWN)")
    frmMessage.lblDesc.Caption = basCCAT.GetAlias("Message Descriptions", "CC_MSG_DESC" & Me.lblID.Caption, "(UNKNOWN)")
    ''
    '' Close the recordset
    'prsMsg.Close
    '-v1.5
    '
    ' Show the form
    frmMessage.Show vbModal
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmMessage.DisplayMessageProperties (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnOK_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Remove the form from the screen
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnOK_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMessage.btnOK Click (Start)"
    '-v1.6.1
    '
    ' Remove the form
    frmMessage.Hide
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMessage.btnOK Click (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' EVENT:    Form_Load
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets up the form
' TRIGGER:  The first time the form is used
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMessage Load (Start)"
    '-v1.6.1
    '
    ' Set help context
    frmMessage.HelpContextID = basCCAT.IDH_GUI_MESSAGE
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMessage Load (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
