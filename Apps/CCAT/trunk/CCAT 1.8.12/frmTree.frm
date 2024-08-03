VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmTree 
   Caption         =   "VarStruct Form"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox SelMsgCombo 
      Height          =   315
      Left            =   3960
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   6240
      Width           =   2535
   End
   Begin VB.CommandButton LoopVarButton 
      Caption         =   "Select Loop Variable"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   4800
      Width           =   2535
   End
   Begin VB.TextBox LoopVarText 
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5520
      Width           =   2535
   End
   Begin VB.ComboBox DasFldCombo 
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox FldLabelText 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   6240
      Width           =   2415
   End
   Begin VB.ComboBox ConvTypeCombo 
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton CancelCommand 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton ApplyCommand 
      Caption         =   "APPLY"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton OKCommand 
      Caption         =   "OK"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   7080
      Width           =   1815
   End
   Begin MSComctlLib.TreeView tvVarStruct 
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6800
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Select Message"
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label loopVarLabel 
      Caption         =   "Select Loop Variable then press ""OK"" or ""Cancel""."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "DAS Field"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Field Label"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Conversion Type"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4560
      Width           =   1815
   End
End
Attribute VB_Name = "frmTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type EditEntry
    lVarStructId As Long
    lConvType As Variant
    sFldLabel As Variant
    lDasFld As Variant
    lRecPtrFld As Variant
    lLoopNode As Long
End Type

Public Enum DAS_MASTER_ENUM
   iNone
   dReportTime
   sMsg_Type
   sReport_Type
   sOrigin
   lOrigin_ID
   lTarget_ID
   dLatitude
   dLongitude
   dAltitude
   dHeading
   dSpeed
   sParent
   lParent_ID
   sAllegiance
   lIFF
   sEmitter
   lEmitter_ID
   sSignal
   lSignal_ID
   dFrequency
   dPRI
   lStatus
   lTag
   lFlag
   lCommon_ID
   dRange
   dBearing
   dElevation
   dXX
   dXY
   dYY
   sSupplemental
End Enum

Dim lVSid() As EditEntry
Dim lNumNodes As Long
Dim lCurNode As Long
Dim lPrevNode As Long
Dim lLoopNode As Long
Dim bselectFlag As Boolean


Private Sub LoopVarButton_Click()

    bselectFlag = True
    frmTree.loopVarLabel.Visible = True
    frmTree.LoopVarButton.Visible = False
    frmTree.LoopVarText.Visible = False
    frmTree.Label1.Visible = False
    frmTree.Label2.Visible = False
    frmTree.Label3.Visible = False
    frmTree.ConvTypeCombo.Visible = False
    frmTree.FldLabelText.Visible = False
    frmTree.DasFldCombo.Visible = False
    frmTree.ApplyCommand.Visible = False
    frmTree.Label4.Visible = False
    frmTree.SelMsgCombo.Visible = False
    
End Sub

Private Sub OKCommand_Click()
'Save changes to varStruct table and exit
    Dim i As Integer
    Dim rsVarStruct As Recordset
    
    If bselectFlag Then
        LoopVarText.Text = lVSid(lLoopNode).sFldLabel
        lVSid(lCurNode).lRecPtrFld = lVSid(lLoopNode).lVarStructId
        
        frmTree.loopVarLabel.Visible = False
        frmTree.LoopVarButton.Visible = True
        frmTree.LoopVarText.Visible = True
        frmTree.Label1.Visible = True
        frmTree.Label2.Visible = True
        frmTree.Label3.Visible = True
        frmTree.ConvTypeCombo.Visible = True
        frmTree.FldLabelText.Visible = True
        frmTree.DasFldCombo.Visible = True
        frmTree.ApplyCommand.Visible = True
        frmTree.Label4.Visible = True
        frmTree.SelMsgCombo.Visible = True

        bselectFlag = False

    Else
        If (MsgBox("Are you sure you want to SAVE your changes and EXIT?", vbOKCancel) = vbOK) Then
            Call SaveChanges
            frmTree.Hide
            tvVarStruct.Nodes.Clear
        End If
    End If
    
End Sub
Private Sub SaveChanges()
    Dim i As Integer
    Dim rsVarStruct As Recordset
            
            For i = 1 To lNumNodes - 1
                'save changes
                Set rsVarStruct = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_VarStruct", dbOpenDynaset)
                rsVarStruct.FindFirst "varStructID = " & lVSid(i).lVarStructId
                rsVarStruct.Edit
                If i = lCurNode Then
                    rsVarStruct!fieldlabel = FldLabelText.Text
                    If (ConvTypeCombo.ListIndex < 1) Then
                        rsVarStruct!ConvType = 0
                    Else
                        rsVarStruct!ConvType = ConvTypeCombo.ListIndex
                    End If
                    If DasFldCombo.ListIndex < 1 Then
                        DasFldCombo.ListIndex = 0
                    Else
                        rsVarStruct!DasField = DasFldCombo.ListIndex
                    End If
                    rsVarStruct!MultiRecPtr = lVSid(i).lRecPtrFld
                Else
                    rsVarStruct!fieldlabel = lVSid(i).sFldLabel
                    rsVarStruct!ConvType = lVSid(i).lConvType
                    rsVarStruct!DasField = lVSid(i).lDasFld
                    rsVarStruct!MultiRecPtr = lVSid(i).lRecPtrFld
                End If
                rsVarStruct.Update
                'rsVarStruct.Close
            Next i
End Sub
Private Sub ApplyCommand_Click()
    If (MsgBox("Are you sure you want to SAVE your changes?", vbOKCancel) = vbOK) Then
      'save changes
      Call SaveChanges
    End If

End Sub

Private Sub CancelCommand_Click()
    
    If bselectFlag Then
        LoopVarText.Text = lVSid(lLoopNode).sFldLabel
        lVSid(lCurNode).lRecPtrFld = lVSid(lLoopNode).lVarStructId
        
        frmTree.loopVarLabel.Visible = False
        frmTree.LoopVarButton.Visible = True
        frmTree.LoopVarText.Visible = True
        frmTree.Label1.Visible = True
        frmTree.Label2.Visible = True
        frmTree.Label3.Visible = True
        frmTree.ConvTypeCombo.Visible = True
        frmTree.FldLabelText.Visible = True
        frmTree.DasFldCombo.Visible = True
        frmTree.ApplyCommand.Visible = True
        frmTree.Label4.Visible = True
        frmTree.SelMsgCombo.Visible = True
      
        bselectFlag = False
    Else
        tvVarStruct.Nodes.Clear
        frmTree.Hide
    End If

End Sub

Private Sub Tree_Create(ByVal msgid As Integer)
    Dim rsVarStruct As Recordset
    Dim iMaxNode As Integer
    Dim varOldLine As Variant
    Dim varNewLine As Variant
    Dim nTreeNodes() As Node
    Dim boolDifferent As Boolean
    Dim iLevel As Integer
    Dim iLastIndex As Integer
    Dim iUnion As Integer

    Set rsVarStruct = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_VarStruct", dbOpenDynaset)
    rsVarStruct.FindFirst "MsgId = " & msgid
    If (rsVarStruct.NoMatch = False) Then
        tvVarStruct.Nodes.Clear
        iMaxNode = -1
        varOldLine = Split(" ", ".")
        lNumNodes = 1
        lCurNode = 0
        lPrevNode = 0
        iUnion = 0
        While (Not (rsVarStruct.NoMatch))
            If ((rsVarStruct!DataType <> "STRUCT END") And (rsVarStruct!DataType <> "UNION END")) Then
                varNewLine = Split(rsVarStruct!fieldname, ".")
                If (UBound(varNewLine) > iMaxNode) Then
                    iMaxNode = UBound(varNewLine)
                    ReDim Preserve nTreeNodes(0 To iMaxNode)
                End If
                ReDim Preserve lVSid(0 To lNumNodes)
                lVSid(lNumNodes).lVarStructId = rsVarStruct!varStructID
                lVSid(lNumNodes).lConvType = rsVarStruct!ConvType
                lVSid(lNumNodes).lDasFld = rsVarStruct!DasField
                lVSid(lNumNodes).sFldLabel = rsVarStruct!fieldlabel
                lVSid(lNumNodes).lRecPtrFld = rsVarStruct!MultiRecPtr
                lNumNodes = lNumNodes + 1
                boolDifferent = False
                For iLevel = 0 To UBound(varNewLine)
                    If (iLevel > UBound(varOldLine)) Then
                        boolDifferent = True
                    ElseIf (varNewLine(iLevel) <> varOldLine(iLevel)) Then
                        boolDifferent = True
                    End If
                    If (boolDifferent = True) Then
                        If (iLevel = 0) Then
                            Set nTreeNodes(iLevel) = tvVarStruct.Nodes.Add(, , , varNewLine(iLevel))
                        Else
                            Set nTreeNodes(iLevel) = tvVarStruct.Nodes.Add(nTreeNodes((iLevel - 1)).Index, tvwChild, , varNewLine(iLevel))
                            If (rsVarStruct!DataType = "UNION BEGIN") Then
                                iUnion = iUnion + 1
                            End If
                            If (iUnion = 1) Then
                                nTreeNodes(iLevel).ForeColor = vbRed
                            ElseIf (iUnion = 2) Then
                                nTreeNodes(iLevel).ForeColor = vbBlue
                            End If

                           nTreeNodes(iLevel).EnsureVisible
                            If (rsVarStruct!MultiEntry = 1) Then
                                nTreeNodes(iLevel).Bold = True
                            End If
                        End If
                        'iLastIndex = nTreeNodes.Index
                    End If
                Next iLevel
                varOldLine = varNewLine
            ElseIf (rsVarStruct!DataType = "UNION END") Then
               iUnion = iUnion - 1
            End If
            
            rsVarStruct.FindNext "MsgId = " & msgid
        Wend
    End If
    rsVarStruct.Close

End Sub

Private Sub Form_Activate()
    Dim rsVarStruct As Recordset
    Dim rsMsgTable As Recordset
    Dim iMaxNode As Integer
    Dim varOldLine As Variant
    Dim varNewLine As Variant
    Dim nTreeNodes() As Node
    Dim boolDifferent As Boolean
    Dim iLevel As Integer
    Dim iLastIndex As Integer
    
    FldLabelText.Text = ""
    ConvTypeCombo.Text = ""
    DasFldCombo.Text = ""
    LoopVarText.Text = ""
    bselectFlag = False
    
    'Clear combo boxes
    frmTree.ConvTypeCombo.Clear
    frmTree.DasFldCombo.Clear
    frmTree.SelMsgCombo.Clear

    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_ProcMsg_Record(" & iMsg_ID & ")"
    '
    Tree_Create guCurrent.iMessage

    ' Populate Conversion Type list
    frmTree.ConvTypeCombo.AddItem "(None)"
    frmTree.ConvTypeCombo.AddItem "AllegToString"
    frmTree.ConvTypeCombo.AddItem "Bam16ToDegree"
    frmTree.ConvTypeCombo.AddItem "Bam32ToDegree"
    frmTree.ConvTypeCombo.AddItem "DegreeToBam32"
    frmTree.ConvTypeCombo.AddItem "FreqConv"
    frmTree.ConvTypeCombo.AddItem "GetOriginID"
    frmTree.ConvTypeCombo.AddItem "HBFuncToString"
    frmTree.ConvTypeCombo.AddItem "HBIndexToString"
    frmTree.ConvTypeCombo.AddItem "LMBSigToLong"
    frmTree.ConvTypeCombo.AddItem "LMBSigToString"
    frmTree.ConvTypeCombo.AddItem "LongToDouble"
    frmTree.ConvTypeCombo.AddItem "RunmodeToString"
    frmTree.ConvTypeCombo.AddItem "XmtrstatToString"
    frmTree.ConvTypeCombo.AddItem "XYToLatLon"
    frmTree.ConvTypeCombo.AddItem "Hex Dump"

    'Populate DAS Field list
    frmTree.DasFldCombo.AddItem "(None)"
    frmTree.DasFldCombo.AddItem "dReportTime"
    frmTree.DasFldCombo.AddItem "sMsg_Type"
    frmTree.DasFldCombo.AddItem "sReport_Type"
    frmTree.DasFldCombo.AddItem "sOrigin"
    frmTree.DasFldCombo.AddItem "lOrigin_ID"
    frmTree.DasFldCombo.AddItem "lTarget_ID"
    frmTree.DasFldCombo.AddItem "dLatitude"
    frmTree.DasFldCombo.AddItem "dLongitude"
    frmTree.DasFldCombo.AddItem "dAltitude"
    frmTree.DasFldCombo.AddItem "dHeading"
    frmTree.DasFldCombo.AddItem "dSpeed"
    frmTree.DasFldCombo.AddItem "sParent"
    frmTree.DasFldCombo.AddItem "lParent_ID"
    frmTree.DasFldCombo.AddItem "sAllegiance"
    frmTree.DasFldCombo.AddItem "lIFF"
    frmTree.DasFldCombo.AddItem "sEmitter"
    frmTree.DasFldCombo.AddItem "lEmitter_ID"
    frmTree.DasFldCombo.AddItem "sSignal"
    frmTree.DasFldCombo.AddItem "lSignal_ID"
    frmTree.DasFldCombo.AddItem "dFrequency"
    frmTree.DasFldCombo.AddItem "dPRI"
    frmTree.DasFldCombo.AddItem "lStatus"
    frmTree.DasFldCombo.AddItem "lTag"
    frmTree.DasFldCombo.AddItem "lFlag"
    frmTree.DasFldCombo.AddItem "lCommon_ID"
    frmTree.DasFldCombo.AddItem "dRange"
    frmTree.DasFldCombo.AddItem "dBearing"
    frmTree.DasFldCombo.AddItem "dElevation"
    frmTree.DasFldCombo.AddItem "dXX"
    frmTree.DasFldCombo.AddItem "dXY"
    frmTree.DasFldCombo.AddItem "dYY"
    frmTree.DasFldCombo.AddItem "sSupplemental"
    
    'Populate the Select Message Combo box
    Set rsMsgTable = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_Message", dbOpenDynaset)
         
    While (Not (rsMsgTable.EOF))
        frmTree.SelMsgCombo.AddItem rsMsgTable!Msg_Name
        rsMsgTable.MoveNext
    Wend

    rsMsgTable.FindFirst "Msg_Id = " & guCurrent.iMessage
    If (rsMsgTable.NoMatch = False) Then
        frmTree.SelMsgCombo.Text = rsMsgTable!Msg_Name
    End If

End Sub



Private Sub SelMsgCombo_Click()
   Dim rsVarStruct As Recordset

    Set rsVarStruct = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_Message", dbOpenDynaset)
     rsVarStruct.FindFirst "Msg_Name = '" & SelMsgCombo.Text & "'"
   If (rsVarStruct.NoMatch = False) Then
        Tree_Create rsVarStruct!Msg_id
    End If
End Sub

Private Sub tvVarStruct_NodeClick(ByVal Node As MSComctlLib.Node)

    If bselectFlag Then
        lLoopNode = Node.Index
    Else
        lPrevNode = lCurNode
        If lPrevNode > 0 Then
            lVSid(lPrevNode).sFldLabel = FldLabelText.Text
            lVSid(lPrevNode).lConvType = ConvTypeCombo.ListIndex
            lVSid(lPrevNode).lDasFld = DasFldCombo.ListIndex
            lVSid(lPrevNode).lRecPtrFld = lVSid(lCurNode).lRecPtrFld
        End If
            
        lCurNode = Node.Index
        If IsNull(lVSid(lCurNode).sFldLabel) Then
            FldLabelText.Text = ""
        Else
            FldLabelText.Text = lVSid(lCurNode).sFldLabel
            frmTree.ConvTypeCombo.ListIndex = lVSid(lCurNode).lConvType
            frmTree.DasFldCombo.ListIndex = lVSid(lCurNode).lDasFld
            If lVSid(lCurNode).lRecPtrFld <> 0 Then
                For i = 1 To lNumNodes
                    If lVSid(lCurNode).lRecPtrFld = lVSid(i).lVarStructId Then
                        frmTree.LoopVarText.Text = lVSid(i).sFldLabel
                        i = lNumNodes
                    End If
                Next i
            Else
                frmTree.LoopVarText = ""
            End If
        End If
    End If
            
End Sub
