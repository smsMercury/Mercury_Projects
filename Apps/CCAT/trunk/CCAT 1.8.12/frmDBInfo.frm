VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDBInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Info"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   HelpContextID   =   600
   Icon            =   "frmDBInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbClass 
      Height          =   315
      Left            =   1080
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2520
      Width           =   1725
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   180
      TabIndex        =   8
      Top             =   2940
      Width           =   1095
   End
   Begin VB.TextBox txtDescription 
      Height          =   1035
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmDBInfo.frx":0442
      Top             =   1380
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   -2147483646
      CalendarTitleForeColor=   -2147483639
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   22806531
      CurrentDate     =   36202
      MinDate         =   35796
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1020
      TabIndex        =   0
      Text            =   "txtName"
      Top             =   60
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      CalendarTitleBackColor=   -2147483646
      CalendarTitleForeColor=   -2147483639
      CustomFormat    =   "MM/dd/yyyy"
      Format          =   22806531
      CurrentDate     =   36202
      MinDate         =   35796
   End
   Begin VB.Label lblClass 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Classification"
      Height          =   195
      Left            =   105
      TabIndex        =   11
      Top             =   2565
      Width           =   915
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Description"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1140
      Width           =   2535
   End
   Begin VB.Label lblEnd 
      Alignment       =   2  'Center
      Caption         =   "End Date"
      Height          =   195
      Left            =   1500
      TabIndex        =   6
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      Caption         =   "Start Date"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmDBInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' FORM:     frmDBInfo
' AUTHOR:   Tom Elkins
' PURPOSE:  Display the contents of a database's Info table and allow the user
'           to modify the values.
' REVISIONS:
'   v1.3.4  TAE Added ability to modify database-level classification
'   v1.5.0  TAE Removed the What's This Help because it was not working with the HTML help file
'               Instead, pressing F1 on the form will bring up the help file to the appropriate page
'   v1.6.1  TAE Added verbose logging calls
Option Explicit
'
' EVENT:    btnAccept_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Mark that the user accepts the changes to the Info table
' TRIGGER:  User clicks on the "Accept" button
' INPUT:    None
' OUTPUT:   None
' NOTES:    The Tag property of the Accept button is used to indicate whether the
'           user accepts or rejects the changes to the Info table.  If "True", the
'           user accepts the changes
Private Sub btnAccept_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmDBInfo.btnAccept Click (Start)"
    '-v1.6.1
    '
    ' Store the accept in the tag property
    frmDBInfo.btnAccept.Tag = True
    '
    ' Remove the form
    frmDBInfo.Hide
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmDBInfo.btnAccept Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnCancel_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Mark that the user cancels the changes to the Info table
' TRIGGER:  User clicks on the "Cancel" button
' INPUT:    None
' OUTPUT:   None
' NOTES:    The Tag property of the Accept button is used to indicate whether the
'           user accepts or rejects the changes to the Info table.  If "False", the
'           user rejects the changes
Private Sub btnCancel_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmDBInfo.btnCancel Click (Start)"
    '-v1.6.1
    '
    ' Store the cancel in the tag property
    frmDBInfo.btnAccept.Tag = False
    '
    ' Remove the form
    frmDBInfo.Hide
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmDBInfo.btnCancel Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Load
' AUTHOR:   Tom Elkins
' PURPOSE:  Initialize the controls on the form
' TRIGGER:  The first time the form is addressed
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmDBInfo Load (Start)"
    '-v1.6.1
    '
    ' Clear the acceptance value
    frmDBInfo.btnAccept.Tag = False
    '
    ' Populate classification list
    frmDBInfo.cmbClass.AddItem "UNCLASSIFIED"
    frmDBInfo.cmbClass.AddItem "UNCLASSIFIED/SAR"
    frmDBInfo.cmbClass.AddItem "CONFIDENTIAL"
    frmDBInfo.cmbClass.AddItem "CONFIDENTIAL/SAR"
    frmDBInfo.cmbClass.AddItem "SECRET"
    frmDBInfo.cmbClass.AddItem "SECRET/SAR"
    frmDBInfo.cmbClass.AddItem "SECRET/SCI"
    frmDBInfo.cmbClass.AddItem "SECRET/SAR/SCI"
    frmDBInfo.cmbClass.AddItem "TOP SECRET"
    frmDBInfo.cmbClass.AddItem "TOP SECRET/SAR"
    frmDBInfo.cmbClass.AddItem "TOP SECRET/SCI"
    frmDBInfo.cmbClass.AddItem "TOP SECRET/SAR/SCI"
    '
    ' Assign the help topics to the controls
    '+v1.5
    'frmDBInfo.btnAccept.WhatsThisHelpID = IDH_DBInfoAccept
    'frmDBInfo.btnCancel.WhatsThisHelpID = IDH_DBInfoCancel
    'frmDBInfo.dtpEnd.WhatsThisHelpID = IDH_DBInfoDates
    'frmDBInfo.dtpStart.WhatsThisHelpID = IDH_DBInfoDates
    'frmDBInfo.txtDescription.WhatsThisHelpID = IDH_DBInfoDescription
    'frmDBInfo.txtName.WhatsThisHelpID = IDH_DBInfoName
    frmDBInfo.HelpContextID = basCCAT.IDH_GUI_DBINFO
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmDBInfo Load (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: blnEditInfoTable
' AUTHOR:   Tom Elkins
' PURPOSE:  Set the values of the form's controls to the current values in the
'           Info table of the specified database. Allow the user to make changes,
'           and either change the values in the table or keep the old values.
' INPUT:    "dbCurrent" is the currently opened database
' OUTPUT:   "TRUE" if the contents of the Info table were modified
'           "FALSE" if the contents of the Info table were not modified
Public Function blnEditInfoTable(dbCurrent As Database) As Boolean
    Dim prsInfo As Recordset     ' Pointer to Info record
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : frmDBInfo.blnEditInfoTable (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS:" & dbCurrent.Name
    End If
    '
    '' Log the event
    'basCCAT.WriteLogEntry "DBINFO: blnEditInfoTable: " & dbCurrent.Name
    '-v1.6.1
    '
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmDBInfo.blnEditInfoTable (Opening Info table)"
    '-v1.6.1
    '
    ' Open the Info table
    Set prsInfo = dbCurrent.OpenRecordset("Info")
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmDBInfo.blnEditInfoTable (Populating values)"
    '-v1.6.1
    '
    ' Populate the name box with the record value
    If IsNull(prsInfo!Name) Then
        frmDBInfo.txtName.Text = Mid(frmMain.dlgCommonDialog.FileTitle, 1, InStr(1, frmMain.dlgCommonDialog.FileTitle, ".mdb") - 1)
    Else
        frmDBInfo.txtName.Text = prsInfo!Name
    End If
    '
    ' Populate the date selectors
    If IsNull(prsInfo!Start) Then
        frmDBInfo.dtpStart.Value = Format(Now, "mm/dd/yy")
    Else
        frmDBInfo.dtpStart.Value = prsInfo!Start
    End If
    If IsNull(prsInfo!end) Then
        frmDBInfo.dtpEnd.Value = Format(Now, "mm/dd/yy")
    Else
        frmDBInfo.dtpEnd.Value = prsInfo!end
    End If
    '
    ' Populate the description field
    If IsNull(prsInfo!Description) Then
        frmDBInfo.txtDescription.Text = ""
    Else
        frmDBInfo.txtDescription.Text = prsInfo!Description
    End If
    '
    ' Set the classification level
    Select Case dbCurrent.Properties("Security").Value
        Case 0: frmDBInfo.cmbClass.ListIndex = 0
        Case 1: frmDBInfo.cmbClass.ListIndex = 1
        Case 4: frmDBInfo.cmbClass.ListIndex = 2
        Case 5: frmDBInfo.cmbClass.ListIndex = 3
        Case 8, 12: frmDBInfo.cmbClass.ListIndex = 4
        Case 9, 13: frmDBInfo.cmbClass.ListIndex = 5
        Case 10, 14: frmDBInfo.cmbClass.ListIndex = 6
        Case 11, 15: frmDBInfo.cmbClass.ListIndex = 7
        Case 16, 20, 24, 28: frmDBInfo.cmbClass.ListIndex = 8
        Case 17, 21, 25, 29: frmDBInfo.cmbClass.ListIndex = 9
        Case 18, 22, 26, 30: frmDBInfo.cmbClass.ListIndex = 10
        Case 19, 23, 27, 31: frmDBInfo.cmbClass.ListIndex = 11
        Case Else: frmDBInfo.cmbClass.ListIndex = 12
    End Select
    frmDBInfo.cmbClass.Tag = frmDBInfo.cmbClass.ListIndex
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmDBInfo.blnEditInfoTable (Displaying form)"
    '-v1.6.1
    '
    ' Display the form modally.  This pauses the routine until the form is closed
    ' by either the "Accept" or "Cancel" buttons.
    frmDBInfo.Show vbModal
    '
    ' Determine if the user pressed "Accept" or "Cancel".  The Accept button's tag
    ' property holds the answer.  "True" if the Accept button was pressed,
    ' "False" if the Cancel button was pressed.
    blnEditInfoTable = frmDBInfo.btnAccept.Tag
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmDBInfo.blnEditInfoTable (Accept = " & blnEditInfoTable & ")"
    '-v1.6.1
    '
    If frmDBInfo.btnAccept.Tag Then
        '
        ' Log the event
        basCCAT.WriteLogEntry "          Saving Changes"
        '
        ' Validate information
        If frmDBInfo.dtpEnd.Value < frmDBInfo.dtpStart.Value Then
            frmDBInfo.dtpEnd.Tag = frmDBInfo.dtpStart.Value
            frmDBInfo.dtpStart.Value = frmDBInfo.dtpEnd.Value
            frmDBInfo.dtpEnd.Value = frmDBInfo.dtpEnd.Tag
        End If
        '
        ' User accepts the changes, so update the record
        prsInfo.Edit
        '
        ' Place the current values in the record
        prsInfo!Name = frmDBInfo.txtName.Text
        prsInfo!Start = frmDBInfo.dtpStart.Value
        prsInfo!end = frmDBInfo.dtpEnd.Value
        prsInfo!Description = frmDBInfo.txtDescription.Text
        '
        ' Update the database
        prsInfo.Update
        '
        ' Set the classification level
        ' Check if the user downgraded the level
        If frmDBInfo.cmbClass.ListIndex < Val(frmDBInfo.cmbClass.Tag) Then
            '
            ' Warn the user of downgrading
            If MsgBox("You are attempting to DOWNGRADE the classification of this database!" & vbCr & "Are you ABSOLUTELY SURE that the data in this database" & vbCr & "can LEGALLY be downgraded to the specified level?", vbExclamation Or vbYesNo, "SECURITY WARNING") = vbYes Then
                '
                '+v1.6.1TE
                If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmDBInfo.blnEditInfoTable (Downgrading database classification)"
                '-v1.6.1
                '
                '
                ' User wants to downgrade, so get the classification values
                ' Force the current classification to be the lower value
                frmSecurity.Tag = frmSecurity.lngGetNumber("Security Values from text", "SECURITY_VAL_" & frmDBInfo.cmbClass.Text, 0)
                '
                ' Formally set the classification value
                frmSecurity.SetClassification Val(frmSecurity.Tag), "User downgraded Database Information"
                '
                ' Display the classification banner
                frmSecurity.Show vbModal, Me
            Else
                frmDBInfo.cmbClass.ListIndex = Val(frmDBInfo.cmbClass.Tag)
            End If
        '
        ' Check if the user upgraded the level
        ElseIf frmDBInfo.cmbClass.ListIndex > Val(frmDBInfo.cmbClass.Tag) Then
            '
            ' Set the new level
            frmSecurity.SetClassification frmSecurity.lngGetNumber("Security Values from text", "SECURITY_VAL_" & frmDBInfo.cmbClass.Text, 0), "User upgraded database information"
        End If
        '
        ' Remember the current properties
        dbCurrent.Properties("Security").Value = frmSecurity.Tag
        '
        ' Suppress error reporting
        On Error Resume Next
        '
        ' Update the database node
        frmMain.tvTreeView.Nodes(dbCurrent.Name).Text = frmDBInfo.txtName.Text
        '
        ' Update the List View caption
        frmMain.lblTitle(1).Caption = frmMain.tvTreeView.SelectedItem.FullPath
        '
        ' Update the list view item
        If Not frmMain.lvListView.ListItems(dbCurrent.Name) Is Nothing Then
            frmMain.lvListView.ListItems(dbCurrent.Name).Text = frmDBInfo.txtName.Text
            frmMain.lvListView.ListItems(dbCurrent.Name).SubItems(1) = frmDBInfo.dtpStart.Value
            frmMain.lvListView.ListItems(dbCurrent.Name).SubItems(2) = frmDBInfo.dtpEnd.Value
            frmMain.lvListView.ListItems(dbCurrent.Name).SubItems(3) = frmDBInfo.txtDescription.Text
        End If
        '
        ' Restore error reporting
        On Error GoTo 0
    End If
    '
    ' Close the table
    prsInfo.Close
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmDBInfo.blnEditInfoTable (End)"
    '-v1.6.1
    '
End Function
