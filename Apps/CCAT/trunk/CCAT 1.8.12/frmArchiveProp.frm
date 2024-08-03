VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmArchiveProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archive Properties"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   ControlBox      =   0   'False
   HelpContextID   =   250
   Icon            =   "frmArchiveProp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnHelp 
      Caption         =   "Help"
      Height          =   420
      Left            =   2925
      TabIndex        =   7
      Top             =   2070
      Width           =   960
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   1575
      TabIndex        =   6
      Top             =   2040
      Width           =   960
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   420
      Left            =   180
      TabIndex        =   5
      Top             =   2070
      Width           =   960
   End
   Begin VB.TextBox txtArchiveName 
      Height          =   330
      Left            =   1125
      TabIndex        =   0
      Text            =   "ArchiveName"
      Top             =   90
      Width           =   2900
   End
   Begin VB.ComboBox cmbArchiveVersion 
      Height          =   315
      Left            =   1125
      TabIndex        =   2
      Text            =   "cmbArchiveVersion"
      Top             =   1080
      Width           =   2900
   End
   Begin VB.ComboBox cmbArchiveClass 
      Height          =   315
      Left            =   1125
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1575
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker dtpArchiveDate 
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Top             =   585
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Format          =   19267584
      CurrentDate     =   37118
   End
   Begin VB.Label lblArcInfo 
      AutoSize        =   -1  'True
      Caption         =   "Archive Name"
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   10
      Top             =   135
      Width           =   990
   End
   Begin VB.Label lblArcInfo 
      AutoSize        =   -1  'True
      Caption         =   "Mission Date"
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   9
      Top             =   630
      Width           =   945
   End
   Begin VB.Label lblArcInfo 
      AutoSize        =   -1  'True
      Caption         =   "CCOS Version"
      Height          =   195
      Index           =   2
      Left            =   45
      TabIndex        =   8
      Top             =   1125
      Width           =   990
   End
   Begin VB.Label lblArcInfo 
      AutoSize        =   -1  'True
      Caption         =   "Classification"
      Height          =   195
      Index           =   3
      Left            =   45
      TabIndex        =   3
      Top             =   1620
      Width           =   930
   End
End
Attribute VB_Name = "frmArchiveProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 2001 Mercury Solutions, Inc.
' FORM:     frmArchiveProp
' AUTHOR:   Tom Elkins
' PURPOSE:  Allows the user to view/change some properties of the archive
' REVISIONS:
'   v1.6.0  TAE Original code
'   v1.6.1  TAE Added verbose logging calls
Option Explicit
'
' Module-level variables
Private mrsArchive As Recordset     ' Current archive's record
'
' EVENT:    btnCancel_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Reset the form's controls back to their original values
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnCancel_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.btnCancel Click (Start)"
    '-v1.6.1
    '
    ' The original values were stored in the "Tag" properties of each control
    Me.txtArchiveName.Text = Me.txtArchiveName.Tag
    Me.dtpArchiveDate.Value = Me.dtpArchiveDate.Tag
    Me.cmbArchiveClass.ListIndex = Me.cmbArchiveClass.Tag
    Me.cmbArchiveVersion.ListIndex = Me.cmbArchiveVersion.Tag
    '
    ' Hide the form
    Me.Hide
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.btnCancel Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnHelp_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Displays the help file for this form
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnHelp_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.btnHelp Click (Start)"
    '-v1.6.1
    '
    basCCAT.ShowHelp Me, basCCAT.IDH_GUI_ARCHIVE_PROPERTIES
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.btnHelp Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnSave_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Processes the changes and updates the database accordingly
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnSave_Click()
    Dim plngDelta As Long   ' Computed difference in time
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.btnSave Click (Start)"
    '-v1.6.1
    '
    ' Edit the record
    mrsArchive.Edit
    '
    ' See if the name changed
    If Me.txtArchiveName.Text <> Me.txtArchiveName.Tag Then
        '
        ' Confirm
        If MsgBox("The archive name was changed.  Are you sure you want all of the records and tables renamed?", vbYesNo Or vbQuestion, "Archive Name Changed") = vbYes Then
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmArchiveProp.btnSave (Renaming tables)"
            '-v1.6.1
            '
            ' Rename the data and summary tables
            guCurrent.DB.TableDefs(mrsArchive!Name & basDatabase.TBL_DATA).Name = Me.txtArchiveName.Text & basDatabase.TBL_DATA
            guCurrent.DB.TableDefs(mrsArchive!Name & basDatabase.TBL_SUMMARY).Name = Me.txtArchiveName.Text & basDatabase.TBL_SUMMARY
            '
            ' Save the new name in the archives table
            mrsArchive!Name = Me.txtArchiveName.Text
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmArchiveProp.btnSave (Tables successfully renamed)"
            '-v1.6.1
            '
        End If
    End If
    '
    ' See if the date changed
    If Me.dtpArchiveDate.Value <> Me.dtpArchiveDate.Tag Then
        '
        ' Check the database version
        If guCurrent.fVersion <= 3 Then
            '
            ' Confirm
            If MsgBox("The archive date changed." & vbCrLf & "All exported data will be referenced to this date." & vbCrLf & _
                "Are you sure you want to change the date?", vbYesNo Or vbQuestion, "Archive Date Changed") = vbYes Then
                '
                ' Change the archive reference date (v1.4 and earlier)
                mrsArchive!Date = Me.dtpArchiveDate.Value
            End If
        Else
            '
            ' Confirm
            If MsgBox("The archive date was changed." & vbCrLf & "Every data record will be changed to have the new date, and every related table will be modified to show the new date." & vbCrLf & "This operation could take a while." & vbCrLf & vbCrLf & "Do you want to update the date stamp on every data record and all tables?", vbYesNo Or vbQuestion, "Archive Date Changed") = vbYes Then
                '
                '+v1.6.1TE
                If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmArchiveProp.btnSave (Changing record dates)"
                '-v1.6.1
                '
                ' Change the mouse
                Me.MousePointer = vbHourglass
                '
                ' Compute the time delta
                plngDelta = CLng(CDate(Int(Me.dtpArchiveDate.Value)) - CDate(Me.dtpArchiveDate.Tag))
                '
                ' Update the fields in the archives table
                mrsArchive!Date = Me.dtpArchiveDate.Value
                mrsArchive!Start = CDate(CDbl(mrsArchive!Start) + plngDelta)
                mrsArchive!end = CDate(CDbl(mrsArchive!end) + plngDelta)
                '
                ' Update the records in the summary table
                guCurrent.DB.Execute "UPDATE [" & Me.txtArchiveName.Text & basDatabase.TBL_SUMMARY & "] Set First = First + " & plngDelta
                guCurrent.DB.Execute "UPDATE [" & Me.txtArchiveName.Text & basDatabase.TBL_SUMMARY & "] Set Last = Last + " & plngDelta
                '
                ' Update the records in the data table
                guCurrent.DB.Execute "UPDATE [" & Me.txtArchiveName.Text & basDatabase.TBL_DATA & "] Set ReportTime = ReportTime + " & plngDelta
                '
                ' Change the mouse back
                Me.MousePointer = vbDefault
                '
                '+v1.6.1TE
                If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmArchiveProp.btnSave (Record dates successfuly changed)"
                '-v1.6.1
                '
            End If
        End If
    End If
    '
    ' See if the version changed
    If Me.cmbArchiveVersion.ListIndex <> Me.cmbArchiveVersion.Tag Then
        '
        ' Confirm
        If MsgBox("The archive version was changed.  Do you want to update the archive record?", vbYesNo Or vbQuestion, "Archive Version Changed") = vbYes Then
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmArchiveProp.btnSave (Updating archive version)"
            '-v1.6.1
            '
            mrsArchive!Analysis_File = Me.cmbArchiveVersion.Text
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmArchiveProp.btnSave (Archive version successfully updated)"
            '-v1.6.1
            '
        End If
    End If
    '
    ' See if the classification changed
    If Me.cmbArchiveClass.ListIndex <> Me.cmbArchiveClass.Tag Then
        '
        ' Confirm
        If MsgBox("The archive classification was changed.  Do you want to update the database?", vbYesNo Or vbQuestion, "Archive Classification Changed") = vbYes Then
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmArchiveProp.btnSave (Updating archive classification)"
            '-v1.6.1
            '
            frmSecurity.SetClassification frmSecurity.lngGetNumber("Security values from text", "SECURITY_VAL_" & Me.cmbArchiveClass.Text, 0), gsARCHIVE
            guCurrent.DB.Properties("Security").Value = frmSecurity.Tag
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmArchiveProp.btnSave (Archive classification successfully updated)"
            '-v1.6.1
            '
        End If
    End If
    '
    ' Save changes
    mrsArchive.Update
    Me.Hide
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmArchiveProp.btnSave Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbArchiveClass_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Update controls when the user changes the archive classification
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbArchiveClass_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.cmbArchiveClass Click (Start)"
    '-v1.6.1
    '
    Me.SetButtonState
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.cmbArchiveClass Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbArchiveVersion_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Update controls when the user changes the version
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbArchiveVersion_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.cmbArchiveVersion Click (Start)"
    '-v1.6.1
    '
    Me.SetButtonState
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.cmbArchiveVersion Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    dtpArchiveDate_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Update controls when the user changes the date
' INPUT:    None
' OUTPUT:   None
' NOTES:    This event is triggered if the user changes a date component with the
'           arrow keys
Private Sub dtpArchiveDate_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.dtpArchiveDate Change (Start)"
    '-v1.6.1
    '
    Me.SetButtonState
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.dtpArchiveDate Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    dtpArchiveDate_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Update controls when the user changes the date
' INPUT:    None
' OUTPUT:   None
' NOTES:    This event is triggered if the user changes the date by clicking on the
'           calendar
Private Sub dtpArchiveDate_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.dtpArchiveDate Click (Start)"
    '-v1.6.1
    '
    Me.SetButtonState
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.dtpArchiveDate Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Load
' AUTHOR:   Tom Elkins
' PURPOSE:  Initialize the controls on the form
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp Load (Start)"
    '-v1.6.1
    '
    ' Initialize the controls
    Me.txtArchiveName.Text = ""
    Me.dtpArchiveDate.Value = Date
    Me.cmbArchiveClass.Clear
    Me.cmbArchiveVersion.Clear
    '
    ' Populate the combo boxes
    basCCAT.PopulateCCOSVersions Me.cmbArchiveVersion
    frmSecurity.PopulateClassification Me.cmbArchiveClass
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp Load (End)"
    '-v1.6.1
    '
End Sub
'
' METHOD:   EditArchiveInfo
' AUTHOR:   Tom Elkins
' PURPOSE:  Get the current information about the selected archive
' INPUT:    strArchiveName - The name of the archive to be viewed
' OUTPUT:   The new name of the archive
' NOTES:
Public Function EditArchiveInfo(strArchiveName As String) As String
    Dim pblnFound As Boolean    ' Flag to indicate that the specified archive was found
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : frmArchiveProp.EditArchiveInfo (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & strArchiveName
    End If
    '-v1.6.1
    '
    '
    ' Set up the recordset
    Set mrsArchive = guCurrent.DB.OpenRecordset(basDatabase.TBL_ARCHIVES, dbOpenDynaset)
    '
    ' Search for the record with the specified archive
    pblnFound = False
    While Not mrsArchive.EOF And Not pblnFound
        pblnFound = (mrsArchive!Name Like strArchiveName)
        If Not pblnFound Then mrsArchive.MoveNext
    Wend
    '
    ' See if it was found
    If Not pblnFound Then
        MsgBox "Could not find the archive named '" & strArchiveName & "' in the database!", vbOKOnly Or vbInformation, "No Such Archive"
    Else
        '
        ' Populate the controls with the current data
        ' Save the values in the Tag property (for comparison)
        Me.txtArchiveName.Text = mrsArchive!Name
        Me.txtArchiveName.Tag = Me.txtArchiveName.Text
        Me.dtpArchiveDate.Value = mrsArchive!Date
        Me.dtpArchiveDate.Tag = Me.dtpArchiveDate.Value
        '
        ' Display the archive classification
        ' Map the security values to the combo box entry indices
        Me.cmbArchiveClass.ListIndex = guCurrent.DB.Properties("Security").Value + 1
        Me.cmbArchiveClass.Tag = Me.cmbArchiveClass.ListIndex
        '
        ' Extract the archive version
        If Left(mrsArchive!Analysis_File, 4) = "CCOS" Then Me.cmbArchiveVersion.Text = mrsArchive!Analysis_File
        Me.cmbArchiveVersion.Tag = Me.cmbArchiveVersion.ListIndex
        '
        ' Enable the action buttons
        Me.SetButtonState
        '
        ' Show the form
        Me.Show vbModal
        '
        ' Return the name
        EditArchiveInfo = Me.txtArchiveName.Text
        Unload Me
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : frmArchiveProp.EditArchiveInfo (End)"
    '-v1.6.1
    '
End Function
'
' EVENT:    Form_Unload
' AUTHOR:   Tom Elkins
' PURPOSE:  Clean up the objects before destroying the form
' INPUT:    None
' OUTPUT:   intCancel - 0 (default) the form will be destroyed
'               non-zero, the form will not be destroyed
' NOTES:
Private Sub Form_Unload(intCancel As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmArchiveProp Unload (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intCancel
    End If
    '-v1.6.1
    '
    ' Destroy the record object
    Set mrsArchive = Nothing
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp Unload (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtArchiveName_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates controls when the user changes the name of the archive
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub txtArchiveName_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.txtArchiveName Change (Start)"
    '-v1.6.1
    '
    Me.SetButtonState
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmArchiveProp.txtArchiveName Change (End)"
    '-v1.6.1
    '
End Sub
'
' METHOD:   SetButtonState
' AUTHOR:   Tom Elkins
' PURPOSE:  Enables/disables the Save button depending on whether the archive information changed
' INPUT:    None
' OUTPUT:   None
' NOTES:
Friend Sub SetButtonState()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmArchiveProp.SetButtonState (Start)"
    '-v1.6.1
    '
    ' Enable the button only if at least one of the properties changed
    Me.btnSave.Enabled = (Me.txtArchiveName.Text <> Me.txtArchiveName.Tag) Or (CStr(Me.dtpArchiveDate.Value) <> Me.dtpArchiveDate.Tag) Or (Val(Me.cmbArchiveClass.Tag) <> Me.cmbArchiveClass.ListIndex) Or (Val(Me.cmbArchiveVersion.Tag) <> Me.cmbArchiveVersion.ListIndex)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmArchiveProp.SetButtonState (End)"
    '-v1.6.1
    '
End Sub
