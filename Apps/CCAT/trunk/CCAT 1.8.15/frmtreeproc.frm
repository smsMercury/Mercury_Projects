VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmtreeproc 
   Caption         =   "Process Messages"
   ClientHeight    =   10260
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   10153.85
   ScaleMode       =   0  'User
   ScaleWidth      =   10851.38
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Multiple Message List"
      Height          =   4455
      Left            =   7680
      TabIndex        =   9
      Top             =   5640
      Width           =   3135
      Begin VB.CheckBox HexCheck2 
         Caption         =   "HEX Dump"
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   3720
         Width           =   735
      End
      Begin VB.CommandButton ProcessCommand 
         Caption         =   "Process and Save Multiple Messages"
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   5741
         View            =   2
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Message Template"
      Height          =   4455
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   7455
      Begin MSComctlLib.TreeView tvVarStruct 
         Height          =   4095
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7223
         _Version        =   393217
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Message Scratchpad"
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   10695
      Begin VB.OptionButton rbOutput 
         Caption         =   "Special Processing"
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   15
         Top             =   5040
         Width           =   1815
      End
      Begin VB.OptionButton rbOutput 
         Caption         =   "HEX Dump"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   14
         Top             =   4680
         Width           =   1575
      End
      Begin VB.OptionButton rbOutput 
         Caption         =   "Normal CSV"
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   13
         Top             =   4320
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton CancelCommand 
         Caption         =   "Cancel/Exit"
         Height          =   615
         Left            =   9000
         TabIndex        =   3
         Top             =   4560
         Width           =   1575
      End
      Begin VB.ComboBox SelTimeCombo 
         Height          =   315
         ItemData        =   "frmtreeproc.frx":0000
         Left            =   360
         List            =   "frmtreeproc.frx":0002
         TabIndex        =   6
         Top             =   4800
         Width           =   3015
      End
      Begin VB.CommandButton SaveCommand 
         Caption         =   "Save Scratchpad"
         Height          =   615
         Left            =   5640
         TabIndex        =   1
         Top             =   4560
         Width           =   1575
      End
      Begin VB.CommandButton ClearCommand 
         Caption         =   "Clear Scratchpad"
         Height          =   615
         Left            =   7320
         TabIndex        =   2
         Top             =   4560
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   7011
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmtreeproc.frx":0004
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Select Time To Process a Single Message"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   4440
         Width           =   3015
      End
   End
   Begin VB.Menu mnuMsgTemplate 
      Caption         =   "Message Template"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu mnuTempSelAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuTempClearAll 
         Caption         =   "&Clear All"
      End
   End
End
Attribute VB_Name = "frmtreeproc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lNumNodes As Long
Dim lCurNode As Long
Dim lPrevNode As Long


Private Sub CancelCommand_Click()
   Unload Me
End Sub

Private Sub ClearCommand_Click()
    RichTextBox1.Text = ""
End Sub

Private Sub Form_Activate()

    If (basTOC.Init_Msg_Proc_Structs = True) Then
        basTOC.Init_Msg_Proc
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    basTOC.Close_Msg_Proc_Structs
    tvVarStruct.Nodes.Clear
    Me.Hide

End Sub

Private Sub mnuTempClearAll_Click()
   Dim i As Integer
   For i = 1 To tvVarStruct.Nodes.Count
      tvVarStruct.Nodes(i).Checked = False
   Next i
   
   basTOC.SetAllChecks False
End Sub

Private Sub mnuTempSelAll_Click()
   Dim i As Integer
   For i = 1 To tvVarStruct.Nodes.Count
      tvVarStruct.Nodes(i).Checked = True
   Next i
   basTOC.SetAllChecks True
End Sub

Private Sub ProcessCommand_Click()
    Dim i As Long
    Dim strNewFile As String
    Dim iFile As Integer

    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.SaveCommand_Click Click (Start)"
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
        .Filter = "Comma-delimited Text (*.csv)|*.csv|Text (*.txt)|*.txt"
        '
        ' Set the flags
        .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
        '
        ' Set the default file type
        '.FilterIndex = guExport.iFile_Type
        '
        ' Give it to the user
        .ShowSave
        '
        ' Save the filename
        strNewFile = .FileName
    End With
    '
    ' Check for a filename
    If Len(strNewFile) > 0 Then
        ' Log the event
        basCCAT.WriteLogEntry "frmtreeproc: SaveCommand_Click: Output File = " & strNewFile
        '
        MousePointer = vbHourglass
        iFile = FreeFile
        Open strNewFile For Output As #iFile
        For i = 1 To ListView1.ListItems.Count
            If (ListView1.ListItems(i).Selected = True) Then
                Print #iFile, , basTOC.Proc_one_TOC_entry(ListView1.ListItems(i).Tag, frmtreeproc.rbOutput)
            End If
        Next i
        Close #iFile
        MousePointer = vbDefault
     End If
End Sub

Private Sub SaveCommand_Click()
'
' EVENT:    SaveCommand_Click
' AUTHOR:   Brad Brown
' PURPOSE:  Saves data to an external file
' INPUT:    None
' OUTPUT:   None
' NOTES:


   Dim strNewFile As String

    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.SaveCommand_Click Click (Start)"
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
        .Filter = "Comma-delimited Text (*.csv)|*.csv|Text (*.txt)|*.txt"
        '
        ' Set the flags
        .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
        '
        ' Set the default file type
        '.FilterIndex = guExport.iFile_Type
        '
        ' Give it to the user
        .ShowSave
        '
        ' Save the filename
        strNewFile = .FileName
    End With
    '
    ' Check for a filename
    If Len(strNewFile) > 0 Then
        ' Log the event
        basCCAT.WriteLogEntry "frmtreeproc: SaveCommand_Click: Output File = " & strNewFile
        '
        RichTextBox1.SaveFile strNewFile, rtfText
        
     End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmtreeproc: SaveCommand_Click (End)"
    '-v1.6.1
    '
End Sub


Private Sub SelTimeCombo_Click()

    MousePointer = vbHourglass
    frmtreeproc.RichTextBox1.Text = frmtreeproc.RichTextBox1.Text & basTOC.Proc_one_TOC_entry(frmtreeproc.SelTimeCombo.ItemData(frmtreeproc.SelTimeCombo.ListIndex), frmtreeproc.rbOutput) & vbNewLine
    MousePointer = vbDefault
        
    
End Sub


Private Sub tvVarStruct_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbRightButton Then
        '
        ' Display the popup menu
        frmtreeproc.PopupMenu mnuMsgTemplate, , x, (y + Frame2.Top)
    End If

End Sub

Private Sub tvVarStruct_NodeCheck(ByVal Node As MSComctlLib.Node)
   Dim lIndex As Long
   Dim blnChecked As Boolean
   
   lIndex = CLng(Node.Index)
   blnChecked = Node.Checked
   basTOC.SetChecks lIndex, blnChecked
End Sub
