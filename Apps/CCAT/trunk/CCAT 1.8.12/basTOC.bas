Attribute VB_Name = "basTOC"

Global Const TBL_MESSAGE = "_Message"
Global Const TBL_PROC_DATA = "_ProcData"
Global Const TBL_VAR_STRUCT = "_VarStruct"
Global Const TBL_TOC = "_TOC"
'Global Const SEP_TOC_MSG = "^"
Dim pintStart As Integer
Dim tmpMsgData() As Byte
Dim rsTOC As Recordset
Dim rsVarStruct As Recordset
Dim rsProcData As Recordset
Dim strTOCTime As String
Dim tmpByteMsgMax As Integer
Dim iCurrLevel As Integer
Dim iLevelArray(0 To 10) As Integer
Dim strText As String
Dim strText2 As String
Dim lStrIndex As Long
Dim tmpCheck() As Byte
Dim tmpMsgHdr As Msg_Hdr
    



'
'+v1.7BB
' ROUTINE:  blnCreateTOCTable
' AUTHOR:   Brad Brown
' PURPOSE:  Creates the default, blank archive Table of Contents table in the specified database
' INPUT:    "dbCurrent" is the currently selected database
'           "strArchive" is the name of the new archive
' OUTPUT:   True if the table was created
'           False if the table was not created
' NOTES:    TOC tables contain results from processing an archive.
'               MsgTOCIndex is the name of a message in the archive
'               MsgId is the numeric identifier of the message type
'               Time is the time of the first occurance of this message in the archive
'               MsgSize is the size (in bytes) of this message in the archive
'               MsgOffset is the offset (in bytes) of this message in the archive
'               MsgProc is whether this message has been processed yet
Public Function blnCreateTOCTable(dbCurrent As Database, strArchive As String) As Boolean
    Dim tblTOC As TableDef  ' New TOC table
    '
    ' Trap errors
    On Error GoTo Hell
    '
    ' Set the default return value
    blnCreateTOCTable = False
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.blnCreateTOCTable (Start)"
    basCCAT.WriteLogEntry "ARGUMENTS: DB = " & dbCurrent.Name & ", Archive = " & strArchive
    '
    ' Create the table
    Set tblTOC = dbCurrent.CreateTableDef(strArchive & TBL_TOC)
    '
    ' Use table-level addressing
    With tblTOC
        '
        ' Add the fields
        .Fields.Append .CreateField("MsgTOCIndex", dbLong)
        .Fields.Append .CreateField("MsgId", dbLong)
        .Fields.Append .CreateField("Time", dbDate)
        .Fields.Append .CreateField("RawTime", dbDouble)
        .Fields.Append .CreateField("MsgSize", dbLong)
        .Fields.Append .CreateField("MsgOffset", dbLong)
        '.Fields.Append .CreateField("MsgProc", dbBoolean)
        ' Set the field attribute to allow null strings

        '.Fields("Description").AllowZeroLength = True
        
    End With
    '
    ' Add table to database
    dbCurrent.TableDefs.Append tblTOC
    '
    ' Set the return value
    blnCreateTOCTable = True
    '

    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.blnCreateTOCTable (End)"

    '
    ' Leave
    Exit Function
'
' Error handler
Hell:
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    ' Set the return value to false
    blnCreateTOCTable = False
    '
    ' Log the error
    basCCAT.WriteLogEntry "ERROR    : basDatabase.blnCreateTOCTable (Error #" & Err.Number & " - " & Err.Description & ")"
    '
    ' Inform the user
    MsgBox "Error #" & Err.Number & vbCr & Err.Description, vbOKOnly, "Error Creating TOC Table"
End Function
'-v1.7BB
'
'+v1.7BB
' ROUTINE:  blnCreateVarStructTable
' AUTHOR:   Brad Brown
' PURPOSE:  Creates the default, blank archive Table of Contents table in the specified database
' INPUT:    "dbCurrent" is the currently selected database
'           "strArchive" is the name of the new archive
' OUTPUT:   True if the table was created
'           False if the table was not created
' NOTES:    TOC tables contain results from processing an archive.
'               MsgTOCIndex is the name of a message in the archive
'               MsgId is the numeric identifier of the message type
'               Time is the time of the first occurance of this message in the archive
'               MsgSize is the size (in bytes) of this message in the archive
'               MsgOffset is the offset (in bytes) of this message in the archive
'               MsgProc is whether this message has been processed yet
Public Function blnCreateVarStructTable(dbCurrent As Database, strArchive As String) As Boolean
    Dim tblVarStruct As TableDef  ' New Var Struct table
    '
    ' Trap errors
    On Error GoTo Hell
    '
    ' Set the default return value
    blnCreateVarStructTable = False
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.blnCreateVarStructTable (Start)"
    basCCAT.WriteLogEntry "ARGUMENTS: DB = " & dbCurrent.Name & ", Archive = " & strArchive
    '
    ' Create the table
    Set tblVarStruct = dbCurrent.CreateTableDef(strArchive & TBL_VAR_STRUCT)
    '
    ' Use table-level addressing
    With tblVarStruct
        '
        ' Add the fields
        .Fields.Append .CreateField("VarStructID", dbLong)
        .Fields.Append .CreateField("MsgId", dbLong)
        .Fields.Append .CreateField("FieldName", dbText, 255)
        .Fields.Append .CreateField("FieldSize", dbLong)
        .Fields.Append .CreateField("DataType", dbText, 50)
        .Fields.Append .CreateField("ConvType", dbLong)
        .Fields.Append .CreateField("FieldLabel", dbText, 50)
        .Fields.Append .CreateField("DASField", dbLong)
        .Fields.Append .CreateField("MultiEntry", dbLong)
        .Fields.Append .CreateField("MultiRecPtr", dbLong)
        .Fields.Append .CreateField("StructLevel", dbLong)
        ' Set the field attribute to allow null strings

        .Fields("FieldName").AllowZeroLength = True
        .Fields("DataType").AllowZeroLength = True
        .Fields("FieldLabel").AllowZeroLength = True

        
    End With
    '
    ' Add table to database
    dbCurrent.TableDefs.Append tblVarStruct
    '
    ' Set the return value
    blnCreateVarStructTable = True
    '

    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.blnCreateVarStructTable (End)"

    '
    ' Leave
    Exit Function
'
' Error handler
Hell:
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    ' Set the return value to false
    blnCreateVarStructTable = False
    '
    ' Log the error
    basCCAT.WriteLogEntry "ERROR    : basDatabase.blnCreateVarStructTable (Error #" & Err.Number & " - " & Err.Description & ")"
    '
    ' Inform the user
    MsgBox "Error #" & Err.Number & vbCr & Err.Description, vbOKOnly, "Error Creating Var Struct Table"
End Function
'-v1.7BB'
'+v1.7BB
' ROUTINE:  blnCreateProcDataTable
' AUTHOR:   Brad Brown
' PURPOSE:  Creates the default, blank archive Table of Contents table in the specified database
' INPUT:    "dbCurrent" is the currently selected database
'           "strArchive" is the name of the new archive
' OUTPUT:   True if the table was created
'           False if the table was not created
' NOTES:    TOC tables contain results from processing an archive.
'               MsgTOCIndex is the name of a message in the archive
'               MsgId is the numeric identifier of the message type
'               Time is the time of the first occurance of this message in the archive
'               MsgSize is the size (in bytes) of this message in the archive
'               MsgOffset is the offset (in bytes) of this message in the archive
'               MsgProc is whether this message has been processed yet
Public Function blnCreateProcDataTable(dbCurrent As Database, strArchive As String) As Boolean
    Dim tblProcData As TableDef  ' New TOC table
    '
    ' Trap errors
    On Error GoTo Hell
    '
    ' Set the default return value
    blnCreateProcDataTable = False
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.blnCreateProcDataTable (Start)"
    basCCAT.WriteLogEntry "ARGUMENTS: DB = " & dbCurrent.Name & ", Archive = " & strArchive
    '
    ' Create the table
    Set tblProcData = dbCurrent.CreateTableDef(strArchive & TBL_PROC_DATA)
    '
    ' Use table-level addressing
    With tblProcData
        '
        ' Add the fields
        '.Fields.Append .CreateField("MsgTOCIndex", dbLong)
        .Fields.Append .CreateField("VarStructID", dbLong)
        .Fields.Append .CreateField("RawValue", dbText, 25)
       ' .Fields.Append .CreateField("ProcValue", dbText, 25)
        ' Set the field attribute to allow null strings

        .Fields("RawValue").AllowZeroLength = True
        '.Fields("ProcValue").AllowZeroLength = True
        
    End With
    '
    ' Add table to database
    dbCurrent.TableDefs.Append tblProcData
    '
    ' Set the return value
    blnCreateProcDataTable = True
    '

    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.blnCreateProcDataTable (End)"

    '
    ' Leave
    Exit Function
'
' Error handler
Hell:
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    ' Set the return value to false
    blnCreateProcDataTable = False
    '
    ' Log the error
    basCCAT.WriteLogEntry "ERROR    : basDatabase.blnCreateProcDataTable (Error #" & Err.Number & " - " & Err.Description & ")"
    '
    ' Inform the user
    MsgBox "Error #" & Err.Number & vbCr & Err.Description, vbOKOnly, "Error Creating Proc Data Table"
End Function
'-v1.7BB'
'+v1.7BB
' ROUTINE:  blnCreateMessageTable
' AUTHOR:   Brad Brown
' PURPOSE:  Creates the default, blank archive Table of Contents table in the specified database
' INPUT:    "dbCurrent" is the currently selected database
'           "strArchive" is the name of the new archive
' OUTPUT:   True if the table was created
'           False if the table was not created
' NOTES:    TOC tables contain results from processing an archive.
'               MsgTOCIndex is the name of a message in the archive
'               MsgId is the numeric identifier of the message type
'               Time is the time of the first occurance of this message in the archive
'               MsgSize is the size (in bytes) of this message in the archive
'               MsgOffset is the offset (in bytes) of this message in the archive
'               MsgProc is whether this message has been processed yet
Public Function blnCreateMessageTable(dbCurrent As Database, strArchive As String) As Boolean
    Dim tblMessage As TableDef  ' New Message table
    '
    ' Trap errors
    On Error GoTo Hell
    '
    ' Set the default return value
    blnCreateMessageTable = False
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.blnCreateMessageTable (Start)"
    basCCAT.WriteLogEntry "ARGUMENTS: DB = " & dbCurrent.Name & ", Archive = " & strArchive
    '
    ' Create the table
    Set tblMessage = dbCurrent.CreateTableDef(strArchive & TBL_MESSAGE)
    '
    ' Use table-level addressing
    With tblMessage
        '
        ' Add the fields
        .Fields.Append .CreateField("Msg_ID", dbLong)
        .Fields.Append .CreateField("Msg_Name", dbText, 50)
        .Fields.Append .CreateField("Select_Msg", dbBoolean)
        .Fields.Append .CreateField("Proc_Msg", dbBoolean)
        .Fields.Append .CreateField("DAS_Msg", dbBoolean)

        ' Set the field attribute to allow null strings
        
    End With
    '
    ' Add table to database
    dbCurrent.TableDefs.Append tblMessage
    '
    ' Set the return value
    blnCreateMessageTable = True
    '

    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.blnCreateMessageTable (End)"

    '
    ' Leave
    Exit Function
'
' Error handler
Hell:
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    ' Set the return value to false
    blnCreateMessageTable = False
    '
    ' Log the error
    basCCAT.WriteLogEntry "ERROR    : basDatabase.blnCreateMessageTable (Error #" & Err.Number & " - " & Err.Description & ")"
    '
    ' Inform the user
    MsgBox "Error #" & Err.Number & vbCr & Err.Description, vbOKOnly, "Error Creating Message Table"
End Function
'-v1.7BB'
' ROUTINE:  Add_Summary_Record
' AUTHOR:   Brad Brown
' PURPOSE:  Updates the Summary Table
' INPUT:    "iMsg_ID" is the ID of the message
' OUTPUT:   None
Public Sub Add_Summary_Record(iMsg_ID As Integer, dReportTime As Double)
 
    Dim dtMsg_Time As Date      ' Extracted message date and time

    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Summary_Record(" & iMsg_ID & ")"
    '
    ' Trap errors
    On Error GoTo ERR_HANDLER
    '
    ' Update the current info structure
    guCurrent.iMessage = iMsg_ID
    '

    dtMsg_Time = DateAdd("s", dReportTime, 0#)
    '
    ' Search for an existing record for this message ID
    guCurrent.uArchive.rsSummary.FindFirst "MSG_ID = " & guCurrent.iMessage
    '
    ' Look for a match
    If guCurrent.uArchive.rsSummary.NoMatch Then
        '
        ' No match, add a new record to the summary table
        guCurrent.uArchive.rsSummary.AddNew
        guCurrent.uArchive.rsSummary!Message = basTOC.Get_Message_Name(iMsg_ID)

        guCurrent.uArchive.rsSummary!MSG_ID = guCurrent.iMessage
        guCurrent.uArchive.rsSummary!Count = 1

        guCurrent.uArchive.rsSummary!First = Format(dtMsg_Time, "mm/dd/yyyy hh:nn:ss")
        guCurrent.uArchive.rsSummary!Last = Format(dtMsg_Time, "mm/dd/yyyy hh:nn:ss")
        '-v1.5
        '
        guCurrent.uArchive.rsSummary!Description = basCCAT.GetAlias("Message Descriptions", "CC_MSG_DESC" & iMsg_ID, "UNKNOWN")
        guCurrent.uArchive.rsSummary.Update
    Else
        guCurrent.uArchive.rsSummary.Edit
        guCurrent.uArchive.rsSummary!Count = guCurrent.uArchive.rsSummary!Count + 1
        '
        If dtMsg_Time < CDate(guCurrent.uArchive.rsSummary!First) Then guCurrent.uArchive.rsSummary!First = Format(dtMsg_Time, "mm/dd/yyyy hh:nn:ss")
        If dtMsg_Time > CDate(guCurrent.uArchive.rsSummary!Last) Then guCurrent.uArchive.rsSummary!Last = Format(dtMsg_Time, "mm/dd/yyyy hh:nn:ss")
        '-v1.5
        '
        guCurrent.uArchive.rsSummary.Update
        guCurrent.sMessage = guCurrent.uArchive.rsSummary!Message
    End If
    '
    ' Update the archive form details
    If dtMsg_Time < guCurrent.uArchive.dtStart_Time Then guCurrent.uArchive.dtStart_Time = dtMsg_Time
    If dtMsg_Time > guCurrent.uArchive.dtEnd_Time Then guCurrent.uArchive.dtEnd_Time = dtMsg_Time

    Exit Sub
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basDatabase.Add_Summary_Record (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    Err.Raise Err.Number, "CCAT:Add_Summary_Record", Err.Description
End Sub


Public Sub ProcDumpDataElement(lVSid As Long)
    Dim tmpLoopCount As Long
    Dim tmpVarSize As Long
    Dim pbytTemp As Byte
    Dim pintTemp As Integer
    Dim plngTemp As Long
    Dim pvarTempRaw As Variant
    Dim pvarTempProc As Variant
    Dim strLevelArray(0 To 10) As String
    Dim varNewLine As Variant
    Dim iLevel As Integer
    Dim strLevel As String
    Dim iLocalLevel As Integer
    Dim strValue As String
    Dim bytTempArray() As Byte
    Dim i As Integer
    Dim lTempLen As Long
    
        If (rsVarStruct!MultiEntry) Then
            tmpLoopCount = rsVarStruct!MultiEntry
        Else
            tmpLoopCount = 0
        End If
        If ((tmpLoopCount = 1) And (rsVarStruct!MultiRecPtr <> 0)) Then
            'get the value of the multirecptr address
            rsProcData.FindLast "VarStructID = " & rsVarStruct!MultiRecPtr
            If (rsProcData!RawValue > 0) Then
                tmpLoopCount = rsProcData!RawValue
            End If
        End If
        
        varNewLine = Split(rsVarStruct!fieldname, ".")
        iCurrLevel = UBound(varNewLine)
        For iLevel = 0 To UBound(varNewLine)
            varNewLine(iLevel) = Replace(varNewLine(iLevel), " 0]", iLevelArray(iLevel) & "]", , 1, vbTextCompare)
            strLevelArray(iLevel + 1) = strLevelArray(iLevel) & varNewLine(iLevel) & "."
        Next iLevel
            
        If ((rsVarStruct!DataType <> "STRUCT BEGIN") And (rsVarStruct!DataType <> "UNION BEGIN")) Then
            If (tmpLoopCount = 0) Then
                tmpVarSize = rsVarStruct!FieldSize
                If ((tmpVarSize + pintStart - 1) <= tmpByteMsgMax) Then
                    Select Case (tmpVarSize)
                        Case 1
                            Call CopyMemory(pbytTemp, tmpMsgData(pintStart), tmpVarSize)
                            pvarTempRaw = pbytTemp
                            pintStart = pintStart + tmpVarSize
                        Case 2
                            Call CopyMemory(pintTemp, tmpMsgData(pintStart), tmpVarSize)
                            pvarTempRaw = agSwapBytes%(pintTemp)
                            pintStart = pintStart + tmpVarSize
                        Case 4
                            Call CopyMemory(plngTemp, tmpMsgData(pintStart), tmpVarSize)
                            pvarTempRaw = agSwapWords&(plngTemp)
                            pintStart = pintStart + tmpVarSize
                        Case Else
                            Call CopyMemory(plngTemp, tmpMsgData(pintStart), tmpVarSize)
                            pvarTempRaw = plngTemp
                            pintStart = pintStart + tmpVarSize
                    End Select
                    
                    If (rsVarStruct!DataType = "CHAR") Then
                        pvarTempProc = Str(pvarTempRaw)
                    ElseIf (rsVarStruct!ConvType) Then
                       pvarTempProc = ProcFormatVal(rsVarStruct!ConvType, pvarTempRaw)
                       strValue = " , " & pvarTempProc
                    Else
                        strValue = ""
                    End If
                    
    
                    rsProcData.AddNew
                    rsProcData!RawValue = pvarTempRaw
                    rsProcData!varStructID = rsVarStruct!varStructID
                    'rsProcData!MsgTOCIndex = rsTOC!MsgTOCIndex
                    rsProcData.Update
                     If (tmpCheck(lVSid) <> 0) Then
                        strText = strTOCTime & "  , " & strLevelArray(UBound(varNewLine) + 1) & " , " & pvarTempRaw & " , " & rsVarStruct!fieldlabel & strValue & vbNewLine
                        lTempLen = Len(strText)
                        Mid$(strText2, lStrIndex, lTempLen) = strText
                        lStrIndex = lStrIndex + lTempLen
                     End If
                End If
            Else
                If ((rsVarStruct!FieldSize + pintStart - 1) <= tmpByteMsgMax) Then
                    tmpVarSize = rsVarStruct!FieldSize / tmpLoopCount
                    ReDim bytTempArray(1 To tmpLoopCount)
                    For i = LBound(bytTempArray) To UBound(bytTempArray)
                        bytTempArray(i) = 0
                    Next i
    
                    For multicount = 1 To tmpLoopCount
                        strLevel = "[" & multicount - 1 & "]"
                        Select Case (tmpVarSize)
                            Case 1
                                Call CopyMemory(pbytTemp, tmpMsgData(pintStart), tmpVarSize)
                                pvarTempRaw = pbytTemp
                                pintStart = pintStart + tmpVarSize
                            Case 2
                                Call CopyMemory(pintTemp, tmpMsgData(pintStart), tmpVarSize)
                                pvarTempRaw = agSwapBytes%(pintTemp)
                                pintStart = pintStart + tmpVarSize
                            Case 4
                                Call CopyMemory(plngTemp, tmpMsgData(pintStart), tmpVarSize)
                                pvarTempRaw = agSwapWords&(plngTemp)
                                pintStart = pintStart + tmpVarSize
                            Case Else
                                Call CopyMemory(plngTemp, tmpMsgData(pintStart), tmpVarSize)
                                pvarTempRaw = plngTemp
                                pintStart = pintStart + tmpVarSize
                            End Select
                        
                    If (rsVarStruct!DataType = "CHAR") Then
                        bytTempArray(multicount) = pvarTempRaw
                        If (multicount = tmpLoopCount) Then
                            pvarTempProc = Trim(StrConv(bytTempArray, vbUnicode))
                            rsProcData.AddNew
                            rsProcData!RawValue = pvarTempRaw
                            rsProcData!varStructID = rsVarStruct!varStructID
                            'rsProcData!MsgTOCIndex = rsTOC!MsgTOCIndex
                            rsProcData.Update
                            If (tmpCheck(lVSid) <> 0) Then
                              strText = strTOCTime & "  , " & strLevelArray(UBound(varNewLine) + 1) & " , " & pvarTempProc & " , " & rsVarStruct!fieldlabel & vbNewLine
                              lTempLen = Len(strText)
                              Mid$(strText2, lStrIndex, lTempLen) = strText
                              lStrIndex = lStrIndex + lTempLen
                           End If
                        End If
                    Else
                        If (rsVarStruct!ConvType) Then
                            pvarTempProc = ProcFormatVal(rsVarStruct!ConvType, pvarTempRaw)
                            strValue = " , " & pvarTempProc
                        Else
                            strValue = ""
                        End If
                        rsProcData.AddNew
                        rsProcData!RawValue = pvarTempRaw
                        rsProcData!varStructID = rsVarStruct!varStructID
                        'rsProcData!MsgTOCIndex = rsTOC!MsgTOCIndex
                        rsProcData.Update
                        If (tmpCheck(lVSid) <> 0) Then
                           strText = strTOCTime & "  , " & strLevelArray(UBound(varNewLine) + 1) & strLevel & " , " & pvarTempRaw & " , " & rsVarStruct!fieldlabel & strValue & vbNewLine
                           lTempLen = Len(strText)
                           Mid$(strText2, lStrIndex, lTempLen) = strText
                           lStrIndex = lStrIndex + lTempLen
                        End If
                    End If
                    Next multicount
                End If
            End If
        ElseIf (rsVarStruct!DataType = "UNION BEGIN") Then
            ProcDumpUnion pintStart, lVSid
        Else 'STRUCT BEGIN
            If tmpLoopCount = 0 Then
                ProcDumpStruct (lVSid)
            Else 'ARRAY of STRUCTS
                iLocalLevel = iCurrLevel
                For multicount = 1 To tmpLoopCount
                    iLevelArray(iLocalLevel) = multicount - 1
                    ProcDumpStruct (lVSid)
                Next multicount
            End If
        End If


End Sub

Public Function ProcFormatVal(ConvType As Long, Dataval As Variant) As Variant

    Select Case ConvType
'        Case 1     ' AllegToString
'            ProcFormatVal = Filtraw2Das.modAllegToString((CByte(Dataval))

        Case 2      ' Bam16ToDeg
            ProcFormatVal = Filtraw2Das.modBam16ToDeg(CLng(Dataval))
            
        Case 3      ' Bam32ToDeg
            ProcFormatVal = Filtraw2Das.modBam32ToDeg(CLng(Dataval))
            
        Case 4      'DegreeToBam32
            ProcFormatVal = Filtraw2Das.modDegToBam32(Dataval)
            
        Case 5     'FreqConv
            ProcFormatVal = Filtraw2Das.modFreqConv(CLng(Dataval))

        Case 6     ' HBFuncToString
            ProcFormatVal = Filtraw2Das.sGetOriginID(CLng(Dataval))
            
        Case 7     ' HBFuncToString
            ProcFormatVal = Filtraw2Das.modHBFuncToString(CInt(Dataval))

        Case 8     ' HBIndexToString
            ProcFormatVal = Filtraw2Das.modHbIndexToString(CLng(Dataval))

        Case 9     ' Hex Dump
            ProcFormatVal = Hex(Dataval)

'        Case 10     'LMBSigToLong
'            ProcFormatVal = Filtraw2Das.modLMBSigToLng((CLng(Dataval))

'        Case 11     'LMBSigToString
'            ProcFormatVal = Filtraw2Das.modLmbSigToString(CLng(Dataval))

        Case 12    'LongToDouble
            ProcFormatVal = Filtraw2Das.modLong2Dbl(CInt(Dataval))

        Case 13    ' RunmodeToString
            ProcFormatVal = Filtraw2Das.modRunmodeToString(CInt(Dataval))

        Case 14    ' XmtrstatToString
            ProcFormatVal = Filtraw2Das.modXmtrstatToString(CInt(Dataval))

'        Case 15    ' XYToLatLon
'           ProcFormatVal = Filtraw2Das.modXYToLatLon((CLng(Dataval))

'        Case 16
    End Select
End Function
Public Sub ProcDumpStruct(lVSid As Long)

    rsVarStruct.FindFirst "varStructID = " & lVSid
    rsVarStruct.MoveNext
    
    While (rsVarStruct!DataType <> "STRUCT END")
        ProcDumpDataElement (rsVarStruct!varStructID)
        rsVarStruct.MoveNext
    Wend
End Sub

Public Sub ProcDumpUnion(ByVal iCurrIndex As Integer, lVSid As Long)
    Dim iEndIndex As Integer
    
    rsVarStruct.FindFirst "varStructID = " & lVSid
    rsVarStruct.MoveNext
    
    While (rsVarStruct!DataType <> "UNION END")
        pintStart = iCurrIndex
        ProcDumpDataElement (rsVarStruct!varStructID)
        rsVarStruct.MoveNext
        'store largest end pintstart
        If (pintStart > iEndIndex) Then
            iEndIndex = pintStart
        End If
    Wend
    pintStart = iEndIndex
End Sub

' ROUTINE:  ProcDump_MsgHdr
' AUTHOR:   Shaun Vogel
' PURPOSE:  Process MsgHdr
' INPUT:    none
' OUTPUT:   Returns a string containing output of a message header
Public Sub ProcDump_MsgHdr()
    Dim lTempLen As Long
    Dim lVSid As Long
    
    'Process msg header
    lVSid = rsVarStruct!varStructID
    If tmpCheck(lVSid + 1) = 0 Then
        Exit Sub
    End If
    If tmpCheck(lVSid + 2) <> 0 Then
        pvarTempRaw = agSwapBytes%(tmpMsgHdr.iMsgLength)
        strText = strTOCTime & ", " & "Length, " & pvarTempRaw & vbNewLine
    End If
    If tmpCheck(lVSid + 3) <> 0 Then
        pvarTempRaw = agSwapBytes%(tmpMsgHdr.iMsgId)
        strText = strText & strTOCTime & ", " & "ID, " & pvarTempRaw & vbNewLine
    End If
    If tmpCheck(lVSid + 4) <> 0 Then
        pvarTempRaw = agSwapBytes%(tmpMsgHdr.iMsgTo)
        strText = strText & strTOCTime & ", " & "To, " & pvarTempRaw & vbNewLine
    End If
    If tmpCheck(lVSid + 5) <> 0 Then
        pvarTempRaw = agSwapBytes%(tmpMsgHdr.iMsgFrom)
        strText = strText & strTOCTime & ", " & "From, " & pvarTempRaw & vbNewLine
    End If
    lTempLen = Len(strText)
    Mid$(strText2, lStrIndex, lTempLen) = strText
    lStrIndex = lStrIndex + lTempLen

End Sub

' ROUTINE:  ProcDump_ExtData_TrackRec
' AUTHOR:   Shaun Vogel
' PURPOSE:  Special processing of mtdpsextrsp TrackRec data
' INPUT:    none
' OUTPUT:
Public Sub ProcDump_ExtData_TrackRec(iLength As Integer, abytBuffer() As Byte)
    Dim j As Integer
    Dim trkRec As ExtData_TrackRec
    
    Call CopyMemory(trkRec, abytBuffer(1), iLength)
    pvarTempRaw = agSwapBytes%(trkRec.iRecType)
    strText = strText & strTOCTime & ", " & "TrackRec.iRecType, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapBytes%(trkRec.iValidDataFlag)
    strText = strText & strTOCTime & ", " & "TrackRec.ValidDataFlag, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapBytes%(trkRec.iTan)
    strText = strText & strTOCTime & ", " & "TrackRec.Tan, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapBytes%(trkRec.iIff)
    strText = strText & strTOCTime & ", " & "TrackRec.IFF, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapWords&(trkRec.lXval)
    strText = strText & strTOCTime & ", " & "TrackRec.Xval, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapWords&(trkRec.lYval)
    strText = strText & strTOCTime & ", " & "TrackRec.Yval, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapWords&(trkRec.lZalt)
    strText = strText & strTOCTime & ", " & "TrackRec.Zalt, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapBytes%(trkRec.iXdot)
    strText = strText & strTOCTime & ", " & "TrackRec.Xdot, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapBytes%(trkRec.iYdot)
    strText = strText & strTOCTime & ", " & "TrackRec.Ydot, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapBytes%(trkRec.iZdot)
    strText = strText & strTOCTime & ", " & "TrackRec.Zdot, " & pvarTempRaw & vbNewLine
    For j = 1 To 8
        pvarTempRaw = agSwapBytes%(trkRec.iAuxData(j))
        strText = strText & strTOCTime & ", " & "TrackRec.AuxData[" & j & "], " & pvarTempRaw & vbNewLine
    Next j

End Sub


' ROUTINE:  ProcDump_Mtdpsextrsp
' AUTHOR:   Shaun Vogel
' PURPOSE:  Special processing of mtdpsextrsp message
' INPUT:    none
' OUTPUT:   Returns a string containing output of mtdpsextrsp message
Public Function ProcDump_Mtdpsextrsp() As String
    Dim tmpDpsExtRspMsg As Mtdpsextrsp
    Dim tmpExtRecType As ExtRecType
    Dim tmpTrackRec As ExtData_TrackRec
    Dim tmpGciRec As ExtData_GciRec
    Dim tmpUnkRec As ExtData_UnkRec
    Dim tmpSpecRec As ExtData_SpecRec
    Dim tmpSpecRec6 As Class6_SpecRec
    Dim tmpSpecRec7a As Class7_ZRV_SpecRec
    Dim tmpSpecRec7b As Class7_RTV_SpecRec
    Dim tmpspecrec8a As Class8var1_3SpecRec
    Dim tmpspecrec8b As Class8var2_4SpecRec
    Dim tmpspecrec9 As Class9_SpecRec
    Dim tmpspecrec12 As Class12_SpecRec
    Dim tmpspecrec13 As Class13_SpecRec
    Dim tmpSpecRec14 As Class14_SpecRec
    Dim tmpspecrec16 As Class16_SpecRec
    Dim tmpspecrec17 As Class17_SpecRec
    Dim tmpSpecRec18 As Class18_SpecRec
    Dim tmpspecrec19 As Class19_SpecRec
    Dim tmpSpecRec22 As Class22_SpecRec
    Dim lTempLen As Long
    Dim lTempDataLen As Long
    Dim tmpintStart As Long
    Dim tmpNumRecs As Long
    Dim recnum As Long
    Dim tmpRecType As Integer
    Dim tmpClassNum As Integer
    Dim tmpVarNum As Integer
    Dim tmpUsageNum As Byte
    Dim tmpdata() As Byte
    Dim lVSid As Long
    
    
    Call CopyMemory(tmpDpsExtRspMsg, tmpMsgData(pintStart), LenB(tmpDpsExtRspMsg))
    
    'Process msg header
    Call CopyMemory(tmpMsgHdr, tmpMsgData(pintStart), LenB(tmpMsgHdr))
    ProcDump_MsgHdr
    
    'Start lVSid at uRspData struct
    lVSid = rsVarStruct!varStructID + 7
    strText = ""
    'Process msg DpsRespData
    If tmpCheck(lVSid + 1) <> 0 Then
        pvarTempRaw = agSwapBytes%(tmpDpsExtRspMsg.uRspData.iSignalID)
        strText = strText & strTOCTime & ", " & "SignalId, " & pvarTempRaw & vbNewLine
    End If
    If tmpCheck(lVSid + 2) <> 0 Then
        pvarTempRaw = (tmpDpsExtRspMsg.uRspData.bytError)
        strText = strText & strTOCTime & ", " & "Error, " & pvarTempRaw & vbNewLine
    End If
    If tmpCheck(lVSid + 3) <> 0 Then
        pvarTempRaw = (tmpDpsExtRspMsg.uRspData.bytRadioType)
        strText = strText & strTOCTime & ", " & "RadioType, " & pvarTempRaw & vbNewLine
    End If
    If tmpCheck(lVSid + 4) <> 0 Then
        pvarTempRaw = (tmpDpsExtRspMsg.uRspData.bytChanNum)
        strText = strText & strTOCTime & ", " & "ChanNum, " & pvarTempRaw & vbNewLine
    End If
    If tmpCheck(lVSid + 5) <> 0 Then
        pvarTempRaw = agSwapWords&(tmpDpsExtRspMsg.uRspData.lFreq)
        strText = strText & strTOCTime & ", " & "Frequency, " & pvarTempRaw & vbNewLine
    End If
    pvarTempRaw = (tmpDpsExtRspMsg.uRspData.uSigType.bytClass)
    tmpClassNum = pvarTempRaw
    strText = strText & strTOCTime & ", " & "SigType.Class, " & pvarTempRaw & vbNewLine
    pvarTempRaw = (tmpDpsExtRspMsg.uRspData.uSigType.bytVariant)
    tmpVarNum = pvarTempRaw
    strText = strText & strTOCTime & ", " & "SigType.Variant, " & pvarTempRaw & vbNewLine
    pvarTempRaw = (tmpDpsExtRspMsg.uRspData.uUsage)
    tmpUsageNum = pvarTempRaw
    strText = strText & strTOCTime & ", " & "Usage, " & pvarTempRaw & vbNewLine
    pvarTempRaw = (tmpDpsExtRspMsg.uRspData.bytLastPacket)
    strText = strText & strTOCTime & ", " & "LastPacket, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapBytes%(tmpDpsExtRspMsg.uRspData.uTime.iDayOfYear)
    strText = strText & strTOCTime & ", " & "Time.DayOfYear, " & pvarTempRaw & vbNewLine
    pvarTempRaw = (tmpDpsExtRspMsg.uRspData.uTime.bytHighUsecs)
    strText = strText & strTOCTime & ", " & "Time.HighUsecs, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapWords&(tmpDpsExtRspMsg.uRspData.uTime.lUsecs)
    strText = strText & strTOCTime & ", " & "Time.Usecs, " & pvarTempRaw & vbNewLine
    pvarTempRaw = (tmpDpsExtRspMsg.uRspData.bytSigPresent)
    strText = strText & strTOCTime & ", " & "SigPresent, " & pvarTempRaw & vbNewLine
    pvarTempRaw = (tmpDpsExtRspMsg.uRspData.bytDataStatus)
    strText = strText & strTOCTime & ", " & "DataStatus, " & pvarTempRaw & vbNewLine
    pvarTempRaw = (tmpDpsExtRspMsg.uRspData.bytDataType)
    strText = strText & strTOCTime & ", " & "DataType, " & pvarTempRaw & vbNewLine
    pvarTempRaw = (tmpDpsExtRspMsg.uRspData.bytDataStatus)
    strText = strText & strTOCTime & ", " & "DataStatus, " & pvarTempRaw & vbNewLine
    pvarTempRaw = agSwapBytes%(tmpDpsExtRspMsg.uRspData.iDataLength)
    strText = strText & strTOCTime & ", " & "DataLength, " & pvarTempRaw & vbNewLine
    
    tmpintStart = pintStart + LenB(tmpDpsExtRspMsg) + LenB(tmpExtRecType)
    
    'Process ExtRecType data
    If (pvarTempRaw > 0) Then
        Call CopyMemory(tmpExtRecType, tmpMsgData(pintStart + LenB(tmpDpsExtRspMsg)), LenB(tmpExtRecType))
        pvarTempRaw = agSwapBytes%(tmpExtRecType.iDisc)
        strText = strText & strTOCTime & ", " & "ExtRecType.Disc, " & pvarTempRaw & vbNewLine
        pvarTempRaw = agSwapBytes%(tmpExtRecType.iNumRecs)
        tmpNumRecs = pvarTempRaw
        strText = strText & strTOCTime & ", " & "ExtRecType.NumRecs, " & pvarTempRaw & vbNewLine
        pvarTempRaw = agSwapBytes%(tmpExtRecType.iSigFrmFlag)
        strText = strText & strTOCTime & ", " & "ExtRecType.SigFrameFlag, " & pvarTempRaw & vbNewLine
        'pvarTempRaw = agSwapBytes%(tmpExtRecType.iRecType)
        'strText = strText & strTOCTime & ", " & "ExtRecType.iRecType, " & pvarTempRaw & vbNewLine
        For recnum = 1 To tmpNumRecs
            Call CopyMemory(tmpRecType, tmpMsgData(tmpintStart), 2)
            pvarTempRaw = agSwapBytes%(tmpRecType)
        
            Select Case pvarTempRaw 'RecType
                Case 1 'Track Rec
                    Call CopyMemory(tmpTrackRec, tmpMsgData(tmpintStart), LenB(tmpTrackRec))
                    tmpintStart = tmpintStart + LenB(tmpTrackRec)
                    pvarTempRaw = agSwapBytes%(tmpTrackRec.iRecType)
                    strText = strText & strTOCTime & ", " & "TrackRec.iRecType, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpTrackRec.iValidDataFlag)
                    strText = strText & strTOCTime & ", " & "TrackRec.ValidDataFlag, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpTrackRec.iTan)
                    strText = strText & strTOCTime & ", " & "TrackRec.Tan, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpTrackRec.iIff)
                    strText = strText & strTOCTime & ", " & "TrackRec.IFF, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapWords&(tmpTrackRec.lXval)
                    strText = strText & strTOCTime & ", " & "TrackRec.Xval, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapWords&(tmpTrackRec.lYval)
                    strText = strText & strTOCTime & ", " & "TrackRec.Yval, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapWords&(tmpTrackRec.lZalt)
                    strText = strText & strTOCTime & ", " & "TrackRec.Zalt, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpTrackRec.iXdot)
                    strText = strText & strTOCTime & ", " & "TrackRec.Xdot, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpTrackRec.iYdot)
                    strText = strText & strTOCTime & ", " & "TrackRec.Ydot, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpTrackRec.iZdot)
                    strText = strText & strTOCTime & ", " & "TrackRec.Zdot, " & pvarTempRaw & vbNewLine
                    For i = 1 To 8
                        pvarTempRaw = agSwapBytes%(tmpTrackRec.iAuxData(i))
                        strText = strText & strTOCTime & ", " & "TrackRec.AuxData[" & i & "], " & pvarTempRaw & vbNewLine
                    Next i
                    pvarTempRaw = agSwapBytes%(tmpTrackRec.iParityFlags)
                    strText = strText & strTOCTime & ", " & "TrackRec.ParityFlags, " & pvarTempRaw & vbNewLine

                Case 2 'GCI Rec
                    Call CopyMemory(tmpGciRec, tmpMsgData(tmpintStart), LenB(tmpGciRec))
                    tmpintStart = tmpintStart + LenB(tmpGciRec)
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iRecType)
                    strText = strText & strTOCTime & ", " & "GciRec.iRecType, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iValidDataFlag)
                    strText = strText & strTOCTime & ", " & "GciRec.ValidDataFlag, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iAddress)
                    strText = strText & strTOCTime & ", " & "GciRec.Address, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iCourse)
                    strText = strText & strTOCTime & ", " & "GciRec.Course, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iSpeed)
                    strText = strText & strTOCTime & ", " & "GciRec.Speed, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapWords&(tmpGciRec.lAltitude)
                    strText = strText & strTOCTime & ", " & "GciRec.Altitude, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapWords&(tmpGciRec.lRange)
                    strText = strText & strTOCTime & ", " & "GciRec.Range, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iRangeInd)
                    strText = strText & strTOCTime & ", " & "GciRec.RangeInd, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iAzimuth)
                    strText = strText & strTOCTime & ", " & "GciRec.Azimuth, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iElevation)
                    strText = strText & strTOCTime & ", " & "GciRec.Elevation, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iCloseRate)
                    strText = strText & strTOCTime & ", " & "GciRec.CloseRate, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iCloseRate)
                    strText = strText & strTOCTime & ", " & "GciRec.CloseRate, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iAspect)
                    strText = strText & strTOCTime & ", " & "GciRec.Aspect, " & pvarTempRaw & vbNewLine
                    For i = 1 To 6
                        pvarTempRaw = agSwapBytes%(tmpGciRec.iAuxData(i))
                        strText = strText & strTOCTime & ", " & "GciRec.AuxData[" & i & "], " & pvarTempRaw & vbNewLine
                    Next i
                    pvarTempRaw = agSwapBytes%(tmpGciRec.iParityFlags)
                    strText = strText & strTOCTime & ", " & "GciRec.ParityFlags, " & pvarTempRaw & vbNewLine
                
                Case 3 'Unknown
                    Call CopyMemory(tmpUnkRec, tmpMsgData(tmpintStart), LenB(tmpUnkRec))
                    tmpintStart = tmpintStart + LenB(tmpUnkRec)
                    pvarTempRaw = agSwapBytes%(tmpUnkRec.iRecType)
                    strText = strText & strTOCTime & ", " & "UnkRec.iRecType, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpUnkRec.iFrmLen)
                    strText = strText & strTOCTime & ", " & "UnkRec.iFrmLen, " & pvarTempRaw & vbNewLine
                    pvarTempRaw = agSwapBytes%(tmpUnkRec.iUnkData)
                    strText = strText & strTOCTime & ", " & "UnkRec.UnkData, " & pvarTempRaw & vbNewLine
                    
                Case 4 'Special Report
                    Select Case tmpClassNum
                        Case 6
                            Call CopyMemory(tmpSpecRec6, tmpMsgData(tmpintStart), LenB(tmpSpecRec6))
                            tmpintStart = tmpintStart + LenB(tmpSpecRec6)
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec6.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iDataStruct)
                            strText = strText & strTOCTime & ", " & "SpecRec6.DataStruct, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec6.NumWords, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iPDisc)
                            strText = strText & strTOCTime & ", " & "SpecRec6.PDisc, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iSDisc)
                            strText = strText & strTOCTime & ", " & "SpecRec6.SDisc, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iBatAddr)
                            strText = strText & strTOCTime & ", " & "SpecRec6.BatAddr, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iTanA)
                            strText = strText & strTOCTime & ", " & "SpecRec6.TanA, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iTanB)
                            strText = strText & strTOCTime & ", " & "SpecRec6.TanB, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iDinome)
                            strText = strText & strTOCTime & ", " & "SpecRec6.Dinome, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iSector)
                            strText = strText & strTOCTime & ", " & "SpecRec6.Sector, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iAlt)
                            strText = strText & strTOCTime & ", " & "SpecRec6.Alt, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iXval)
                            strText = strText & strTOCTime & ", " & "SpecRec6.Xval, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iYval)
                            strText = strText & strTOCTime & ", " & "SpecRec6.Yval, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iBearing)
                            strText = strText & strTOCTime & ", " & "SpecRec6.Bearing, " & pvarTempRaw & ", " & pvarTempRaw * 0.176 & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec6.iDist)
                            strText = strText & strTOCTime & ", " & "SpecRec6.Dist, " & pvarTempRaw & ", " & pvarTempRaw * 100 & " M" & vbNewLine
                          
                        Case 7
                            Select Case tmpUsageNum
                                Case 3, 4, 5 'ZRV
                                    Call CopyMemory(tmpSpecRec7a, tmpMsgData(tmpintStart), LenB(tmpSpecRec7a))
                                    tmpintStart = tmpintStart + LenB(tmpSpecRec7a)
                                    pvarTempRaw = agSwapBytes%(tmpSpecRec7a.iRecType)
                                    strText = strText & strTOCTime & ", " & "SpecRec7.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                                    pvarTempRaw = agSwapBytes%(tmpSpecRec7b.iNumWords)
                                    strText = strText & strTOCTime & ", " & "SpecRec7.NumWords, " & pvarTempRaw & vbNewLine
                                    pvarTempRaw = agSwapBytes%(tmpSpecRec7a.iDisc)
                                    strText = strText & strTOCTime & ", " & "SpecRec7.Disc, " & pvarTempRaw & vbNewLine
                                    pvarTempRaw = agSwapBytes%(tmpSpecRec7a.iSubAddr)
                                    strText = strText & strTOCTime & ", " & "SpecRec7.SubAddr, " & pvarTempRaw & vbNewLine
                                    For i = 1 To 14
                                        pvarTempRaw = agSwapBytes%(tmpSpecRec7a.iSpecfld(i))
                                        strText = strText & strTOCTime & ", " & "SpecRec7.Specfld(" & i & "), " & pvarTempRaw & vbNewLine
                                    Next i
                                Case Else
                                    Call CopyMemory(tmpSpecRec7b, tmpMsgData(tmpintStart), LenB(tmpSpecRec7b))
                                    tmpintStart = tmpintStart + LenB(tmpSpecRec7b)
                                    pvarTempRaw = agSwapBytes%(tmpSpecRec7b.iRecType)
                                    strText = strText & strTOCTime & ", " & "SpecRec7.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                                    pvarTempRaw = agSwapBytes%(tmpSpecRec7b.iDataStruct)
                                    strText = strText & strTOCTime & ", " & "SpecRec7.DataStruct, " & pvarTempRaw & vbNewLine
                                    pvarTempRaw = agSwapBytes%(tmpSpecRec7b.iNumWords)
                                    strText = strText & strTOCTime & ", " & "SpecRec7.NumWords, " & pvarTempRaw & vbNewLine
                                    pvarTempRaw = agSwapBytes%(tmpSpecRec7b.iDisc)
                                    strText = strText & strTOCTime & ", " & "SpecRec7.Disc, " & pvarTempRaw & vbNewLine
                                    pvarTempRaw = agSwapBytes%(tmpSpecRec7b.iSubAddr)
                                    strText = strText & strTOCTime & ", " & "SpecRec7.SubAddr, " & pvarTempRaw & vbNewLine
                                    For i = 1 To 14
                                        pvarTempRaw = agSwapBytes%(tmpSpecRec7b.iSpecfld(i))
                                        strText = strText & strTOCTime & ", " & "SpecRec7.Specfld(" & i & "), " & pvarTempRaw & vbNewLine
                                    Next i
                            End Select
                        
                        Case 8
                            If (tmpVarNum = 1) Or (tmpVarNum = 3) Then
                                Call CopyMemory(tmpspecrec8a, tmpMsgData(tmpintStart), LenB(tmpspecrec8a))
                                tmpintStart = tmpintStart + LenB(tmpspecrec8a)
                                pvarTempRaw = agSwapBytes%(tmpspecrec8a.iRecType)
                                strText = strText & strTOCTime & ", " & "SpecRec8.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8a.iDataStruct)
                                strText = strText & strTOCTime & ", " & "SpecRec8.DataStruct, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpSpecRec8.iNumWords)
                                strText = strText & strTOCTime & ", " & "SpecRec8.NumWords, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8a.iXscrn)
                                strText = strText & strTOCTime & ", " & "SpecRec8.Xscrn, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8a.iYscrn)
                                strText = strText & strTOCTime & ", " & "SpecRec8.Yscrn, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8a.iSymAmp1)
                                strText = strText & strTOCTime & ", " & "SpecRec8.SymAmp1, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8a.iSymAmp2)
                                strText = strText & strTOCTime & ", " & "SpecRec8.SymAmp2, " & pvarTempRaw & vbNewLine
                                
                            Else 'tmpVarNum = 2 or 4
                                Call CopyMemory(tmpspecrec8b, tmpMsgData(tmpintStart), LenB(tmpspecrec8b))
                                tmpintStart = tmpintStart + LenB(tmpspecrec8b)
                                pvarTempRaw = agSwapBytes%(tmpspecrec8b.iRecType)
                                strText = strText & strTOCTime & ", " & "SpecRec8.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8b.iDataStruct)
                                strText = strText & strTOCTime & ", " & "SpecRec8.DataStruct, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8b.iNumWords)
                                strText = strText & strTOCTime & ", " & "SpecRec8.NumWords, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8b.iSpecFlds(1))
                                strText = strText & strTOCTime & ", " & "SpecRec8.Specfld(1), " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8b.iSpecFlds(2))
                                strText = strText & strTOCTime & ", " & "SpecRec8.Specfld(2), " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8b.iSpecFlds(3))
                                strText = strText & strTOCTime & ", " & "SpecRec8.Specfld(3), " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8b.iSpecFlds(4))
                                strText = strText & strTOCTime & ", " & "SpecRec8.Specfld(4), " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec8b.iSpecFlds(5))
                                strText = strText & strTOCTime & ", " & "SpecRec8.Specfld(5), " & pvarTempRaw & vbNewLine
                            
                            End If
                        Case 9
                            Call CopyMemory(tmpspecrec9, tmpMsgData(tmpintStart), LenB(tmpspecrec9))
                            tmpintStart = tmpintStart + LenB(tmpspecrec9)
                            pvarTempRaw = agSwapBytes%(tmpspecrec9.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec9.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec9.iDataStruct)
                            strText = strText & strTOCTime & ", " & "SpecRec9.DataStruct, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec9.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec9.NumWords, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec9.iData)
                            strText = strText & strTOCTime & ", " & "SpecRec9.Data, " & pvarTempRaw & vbNewLine
                        Case 12
                            Call CopyMemory(tmpspecrec12, tmpMsgData(tmpintStart), LenB(tmpspecrec12))
                            tmpintStart = tmpintStart + LenB(tmpspecrec12)
                            pvarTempRaw = agSwapBytes%(tmpspecrec12.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec12.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec12.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec12.NumWords, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec12.iDataStruct)
                            strText = strText & strTOCTime & ", " & "SpecRec12.DataStruct, " & pvarTempRaw & vbNewLine
                            
                            strText = strText & strTOCTime & ", " & "SpecRec12.sText, " & tmpspecrec12.sText & vbNewLine
                            
                            'For i = 1 To 65
                                'pvarTempRaw = agSwapBytes%(tmpspecrec12.sText(i))
                                'strText = strText & strTOCTime & ", " & "SpecRec12.sText(" & i & "), " & pvarTempRaw & vbNewLine
                            'Next i
                        Case 13
                            Call CopyMemory(tmpspecrec13, tmpMsgData(tmpintStart), LenB(tmpspecrec13))
                            tmpintStart = tmpintStart + LenB(tmpspecrec13)
                            pvarTempRaw = agSwapBytes%(tmpspecrec13.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec13.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec13.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec13.NumWords, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec13.iDataStruct)
                            strText = strText & strTOCTime & ", " & "SpecRec13.DataStruct, " & pvarTempRaw & vbNewLine
                            strText = strText & strTOCTime & ", " & "SpecRec13.sText, " & tmpspecrec13.sText & vbNewLine

                        Case 14
                            Call CopyMemory(tmpSpecRec14, tmpMsgData(tmpintStart), LenB(tmpSpecRec14))
                            tmpintStart = tmpintStart + LenB(tmpSpecRec14)
                            pvarTempRaw = agSwapBytes%(tmpSpecRec14.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec14.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec14.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec14.SpecialData.NumWords, " & pvarTempRaw & vbNewLine
                            For i = 1 To 8
                                pvarTempRaw = agSwapBytes%(tmpSpecRec14.uCmds(i).iTanPar)
                                strText = strText & strTOCTime & ", " & "SpecRec14.SpecialData.uCmds(" & i & ").TanPar, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpSpecRec14.uCmds(i).iTan)
                                strText = strText & strTOCTime & ", " & "SpecRec14.SpecialData.uCmds(" & i & ").Tan, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpSpecRec14.uCmds(i).iCmd)
                                strText = strText & strTOCTime & ", " & "SpecRec14.SpecialData.uCmds(" & i & ").Cmd, " & pvarTempRaw & vbNewLine
                            Next i
                            For i = 1 To 8
                                pvarTempRaw = agSwapBytes%(tmpSpecRec14.uRsps(i).iRspPar)
                                strText = strText & strTOCTime & ", " & "SpecRec14.SpecialData.uRsps(" & i & ").RspPar, " & pvarTempRaw & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpSpecRec14.uRsps(i).iRsp)
                                strText = strText & strTOCTime & ", " & "SpecRec14.SpecialData.uRsps(" & i & ").Rsp, " & pvarTempRaw & vbNewLine
                            Next i
                        Case 16
                            Call CopyMemory(tmpspecrec16, tmpMsgData(tmpintStart), LenB(tmpspecrec16))
                            tmpintStart = tmpintStart + LenB(tmpspecrec16)
                            pvarTempRaw = agSwapBytes%(tmpspecrec16.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec16.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec16.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec16.NumWords, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec16.iResponse)
                            strText = strText & strTOCTime & ", " & "SpecRec16.Response, " & pvarTempRaw & vbNewLine
                        Case 17
                            Call CopyMemory(tmpspecrec17, tmpMsgData(tmpintStart), LenB(tmpspecrec17))
                            tmpintStart = tmpintStart + LenB(tmpspecrec17)
                            pvarTempRaw = agSwapBytes%(tmpspecrec17.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec17.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec17.iDataStruct)
                            strText = strText & strTOCTime & ", " & "SpecRec17.DataStruct, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec17.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec17.NumWords, " & pvarTempRaw & vbNewLine
                            For i = 1 To 3
                                pvarTempRaw = agSwapBytes%(tmpspecrec17.uSubRec(i).iAangle)
                                strText = strText & strTOCTime & ", " & "SpecRec17.SubRec.iAangle, " & pvarTempRaw & ", " & pvarTempRaw * 1.5 & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec17.uSubRec(i).iEangle)
                                strText = strText & strTOCTime & ", " & "SpecRec17.SubRec.iEangle, " & pvarTempRaw & ", " & pvarTempRaw * 1.5 & vbNewLine
                                pvarTempRaw = agSwapBytes%(tmpspecrec17.uSubRec(i).iLcmd)
                                strText = strText & strTOCTime & ", " & "SpecRec17.SubRec.iLcmd, " & pvarTempRaw & vbNewLine
                            Next i

                        Case 18
                            Call CopyMemory(tmpSpecRec18, tmpMsgData(tmpintStart), LenB(tmpSpecRec18))
                            tmpintStart = tmpintStart + LenB(tmpSpecRec18)
                            pvarTempRaw = agSwapBytes%(tmpSpecRec18.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec18.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec18.iDataStruct)
                            strText = strText & strTOCTime & ", " & "SpecRec18.DataStruct, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec18.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec18.NumWords, " & pvarTempRaw & vbNewLine
                            For i = 1 To 3
                                pvarTempRaw = agSwapBytes%(tmpSpecRec18.iData(i))
                                strText = strText & strTOCTime & ", " & "SpecRec18.Data(" & i & "), " & pvarTempRaw & vbNewLine
                            Next i
                        Case 19
                            Call CopyMemory(tmpspecrec19, tmpMsgData(tmpintStart), LenB(tmpspecrec19))
                            tmpintStart = tmpintStart + LenB(tmpspecrec19)
                            pvarTempRaw = agSwapBytes%(tmpspecrec19.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec19.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec19.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec19.NumWords, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec19.iMsgType)
                            strText = strText & strTOCTime & ", " & "SpecRec19.MsgType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec19.iTanA)
                            strText = strText & strTOCTime & ", " & "SpecRec19.TanA, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec19.iTanB)
                            strText = strText & strTOCTime & ", " & "SpecRec19.TanB, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec19.iXval)
                            strText = strText & strTOCTime & ", " & "SpecRec19.Xval, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpspecrec19.iYval)
                            strText = strText & strTOCTime & ", " & "SpecRec19.Yval, " & pvarTempRaw & vbNewLine
                        Case 22
                            Call CopyMemory(tmpSpecRec22, tmpMsgData(tmpintStart), LenB(tmpSpecRec22))
                            tmpintStart = tmpintStart + LenB(tmpSpecRec22)
                            pvarTempRaw = agSwapBytes%(tmpSpecRec22.iRecType)
                            strText = strText & strTOCTime & ", " & "SpecRec22.SpecialData.RecType, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec22.iNumWords)
                            strText = strText & strTOCTime & ", " & "SpecRec22.SpecialData.NumWords, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec22.iSync1)
                            strText = strText & strTOCTime & ", " & "SpecRec22.SpecialData.Sync1, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec22.iSync2)
                            strText = strText & strTOCTime & ", " & "SpecRec22.SpecialData.Sync2, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec22.iSync3)
                            strText = strText & strTOCTime & ", " & "SpecRec22.SpecialData.Sync3, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec22.iSync4)
                            strText = strText & strTOCTime & ", " & "SpecRec22.SpecialData.Sync4, " & pvarTempRaw & vbNewLine
                            pvarTempRaw = agSwapBytes%(tmpSpecRec22.iAck)
                            strText = strText & strTOCTime & ", " & "SpecRec22.SpecialData.Ack, " & pvarTempRaw & vbNewLine
                    End Select
            End Select
        Next recnum
    End If
    
    lTempLen = Len(strText)
    Mid$(strText2, lStrIndex, lTempLen) = strText
    lStrIndex = lStrIndex + lTempLen
    
    ProcDump_Mtdpsextrsp = Left$(strText2, lStrIndex)
End Function


' ROUTINE:  Proc_one_TOC_entry
' AUTHOR:   Brad Brown
' PURPOSE:  Updates the ProcMsg Table
' INPUT:    "iMsg_ID" is the ID of the message
' OUTPUT:   None
Public Function Proc_one_TOC_entry(lTOCIndex As Long, msgOutput As Object) As String
Dim temp As Integer
Dim tmpTime As String

    strText2 = Space(400000) 'BBnew
    lStrIndex = 1             'BBnew
    
    strText = "" & vbNewLine
    rsTOC.FindFirst "MsgTOCIndex = " & lTOCIndex
    strTOCTime = rsTOC!Time
    tmpByteMsgMax = rsTOC!MsgSize
    'Dimension array to data size
    ReDim tmpMsgData(1 To tmpByteMsgMax)
    Get giRawInputFile, rsTOC!MsgOffset, tmpMsgData()        ' get the message data
    pintStart = LBound(tmpMsgData)
    
    'Normal CSV output
    If msgOutput(0).Value Then
        rsVarStruct.FindFirst "MsgId = " & guCurrent.iMessage
        While ((pintStart < tmpByteMsgMax) And Not (rsVarStruct.NoMatch))
            ProcDumpDataElement (rsVarStruct!varStructID)
            rsVarStruct.FindNext "MsgId = " & guCurrent.iMessage
        Wend
        Proc_one_TOC_entry = Left$(strText2, lStrIndex)
    'HEX Dump output
    ElseIf msgOutput(1).Value Then
        For i = pintStart To UBound(tmpMsgData)
            Proc_one_TOC_entry = Proc_one_TOC_entry + " " + Hex(tmpMsgData(i))
            If (i Mod 10 = 0) Then
                Proc_one_TOC_entry = Proc_one_TOC_entry + vbCrLf
            End If
        Next i
        Proc_one_TOC_entry = Proc_one_TOC_entry + vbNewLine
    'Special Processing output
    Else
        rsVarStruct.FindFirst "MsgId = " & guCurrent.iMessage
        Select Case guCurrent.iMessage
            Case MTDPSEXTRSPID
                Proc_one_TOC_entry = ProcDump_Mtdpsextrsp
            Case Else
                MsgBox ("No special processing for this message.")
        End Select
    End If
    
    Exit Function
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basDatabase.Proc_one_TOC_entry(Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    Err.Raise Err.Number, "CCAT:Proc_one_TOC_entry", Err.Description
End Function

' ROUTINE:  Add_TOC_Record
' AUTHOR:   Brad Brown
' PURPOSE:  Updates the Table of Contents Table
' INPUT:    "iMsg_ID" is the ID of the message
' OUTPUT:   None
Public Sub Add_TOC_Record(uTOCMsg As Toc_Record)
    Dim dtMsg_Time As Date      ' Extracted message date and time
 
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_TOC_Record(" & iMsg_ID & ")"
    '
    ' Trap errors
    On Error GoTo ERR_HANDLER
    dtMsg_Time = DateAdd("s", uTOCMsg.dTimeStamp, 0#)
        '  add a new record to the summary table
        guCurrent.uArchive.rsTOC.AddNew
        guCurrent.uArchive.rsTOC!MsgTOCIndex = uTOCMsg.lMsgCount
        guCurrent.uArchive.rsTOC!msgid = uTOCMsg.iMsgId
        guCurrent.uArchive.rsTOC!Time = Format(dtMsg_Time, "mm/dd/yyyy hh:nn:ss")
        guCurrent.uArchive.rsTOC!RawTime = uTOCMsg.dTimeStamp
        guCurrent.uArchive.rsTOC!MsgSize = uTOCMsg.iMsgSize
        guCurrent.uArchive.rsTOC!MsgOffset = uTOCMsg.iStartByte
        guCurrent.uArchive.rsTOC.Update
        

    Exit Sub
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basDatabase.Add_TOC_Record (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    Err.Raise Err.Number, "CCAT:Add_TOC_Record", Err.Description
End Sub

'
' ROUTINE:  Add_TOC_Node
' AUTHOR:   Brad Brown
' PURPOSE:  Add a node in the TreeView to correspond to the specified Message record
' INPUT:    "sArchive" is the archive node's key
'           "rsMsg" is the current Message record
' OUTPUT:   None
' NOTES:    The Key value for a message node is A#M, where
'               A is the Archive node's key
'               # is a separator string
'               M is the message ID
Public Sub Add_TOC_Node(sArchive As String, rsMsg As Recordset)
    Dim nodMsg As Node          ' New message node
    '
    '+v1.7BB
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_TOC_Node (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sArchive
    End If

    '
    ' Check for the existence of the archive node
    If frmMain.blnNodeExists(sArchive) Then
        '
        If Not frmMain.blnNodeExists(sArchive & TBL_TOC) Then
            '
            On Error Resume Next
            Set nodMsg = frmMain.tvTreeView.Nodes.Add(sArchive, tvwChild, sArchive & TBL_TOC, "TOC Messages", "ClosedBook", "OpenBook")
            nodMsg.Sorted = True
        End If
        ' Check for the existence of the message node
        If Not frmMain.blnNodeExists(sArchive & SEP_TOC_MSG & rsMsg!MSG_ID) Then
            '
            ' Log the event
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : basDatabase.Add_TOC_Node (" & sArchive & SEP_TOC_MSG & rsMsg!MSG_ID & ")"
            '
            ' Create a new node with the following properties
            '   Relative:       specified archive node
            '   Relationship:   Child
            '   Key:            See NOTES above
            '   Text:           Message name
            Set nodMsg = frmMain.tvTreeView.Nodes.Add(sArchive & TBL_TOC, tvwChild, sArchive & SEP_TOC_MSG & rsMsg!MSG_ID, rsMsg!Message)
            '
            ' Set the icons
            nodMsg.Sorted = True
            nodMsg.Image = "MSG_CLOSED"
            nodMsg.SelectedImage = "MSG_OPEN"
            '
            ' Save the node type
            nodMsg.Tag = gsTOCMSG
        End If
    End If
    '

    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_TOC_Node (End)"
    '-v1.7BB
    '
End Sub

'
' ROUTINE:  Display_TOCMsg_Details
' AUTHOR:   Brad Brown
' PURPOSE:  Display a message's data in the ListView
' INPUT:    "nodMsg" is the selected message node from the TreeView
' OUTPUT:   None
' NOTES:
Public Sub Display_TOCMsg_Details()
    Dim sToken As String
    Dim lTokenLen As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_TOCMsg_Details (Start)"
    End If
    '-v1.6.1
    '
    ' Assign the database to the data control
    frmMain.Data1.DatabaseName = guCurrent.sName
    '
    ' Position the data grid to fit the list view space
    frmMain.grdData.Left = frmMain.lvListView.Left
    frmMain.grdData.Top = frmMain.lvListView.Top
    frmMain.grdData.Width = frmMain.lvListView.Width
    frmMain.grdData.Height = frmMain.lvListView.Height
    '
    ' Hide the list view and show the grid
    frmMain.lvListView.Visible = False
    frmMain.grdData.Visible = True
    '
    ' Create the table name, which is ARCHIVE<Archive ID>_DATA
    '+v1.6TE
    guCurrent.uSQL.sTable = guCurrent.sArchive & TBL_TOC
    '-v1.6
    '
    ' Create a default SQL query
    ' Get a custom field list from the token file
    guCurrent.uSQL.sFields = "MsgId, Time, MsgSize, MsgOffset"
    '
    ' Set the filter to extract records for the selected message
    guCurrent.uSQL.sFilter = "MsgId = " & guCurrent.iMessage
    '
    ' Remove the sort order
    guCurrent.uSQL.sOrder = ""
    '
    ' If the data table exists, execute the query
    '+v1.5
    If bTable_Exists(guCurrent.DB, guCurrent.uSQL.sTable) Then basDatabase.QueryData basDatabase.sCreate_SQL
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_TOCMsg_Details (End)"
    '-v1.6.1
    '
End Sub

'
' ROUTINE:  Display_ProcMsg_Details
' AUTHOR:   Brad Brown
' PURPOSE:  Display a message's data in the ListView
' INPUT:    "nodMsg" is the selected message node from the TreeView
' OUTPUT:   None
' NOTES:
Public Sub Display_ProcMsg_Details()
    Dim sToken As String
    Dim lTokenLen As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_ProcMsg_Details (Start)"
        'basCCAT.WriteLogEntry "ARGUMENTS: " & nodMsg.Text & " [" & nodMsg.Key & "]"
    End If
    '-v1.6.1
    '
    ' Assign the database to the data control
    frmMain.Data1.DatabaseName = guCurrent.sName
    '
    ' Position the data grid to fit the list view space
    frmMain.grdData.Left = frmMain.lvListView.Left
    frmMain.grdData.Top = frmMain.lvListView.Top
    frmMain.grdData.Width = frmMain.lvListView.Width
    frmMain.grdData.Height = frmMain.lvListView.Height
    '
    ' Hide the list view and show the grid
    frmMain.lvListView.Visible = False
    frmMain.grdData.Visible = True
    '
    ' Create the table name, which is ARCHIVE<Archive ID>_DATA
    '+v1.6TE
    guCurrent.uSQL.sTable = guCurrent.sArchive & TBL_PROC_DATA
    '-v1.6
    '
    ' Create a default SQL query
    ' Get a custom field list from the token file
    guCurrent.uSQL.sFields = "MsgId, Time, MsgSize, MsgOffset"
    '
    ' Set the filter to extract records for the selected message
    guCurrent.uSQL.sFilter = "MsgId = " & guCurrent.iMessage
    '
    ' Remove the sort order
    guCurrent.uSQL.sOrder = ""
    '
    ' If the data table exists, execute the query
    '+v1.5
    If bTable_Exists(guCurrent.DB, guCurrent.uSQL.sTable) Then basDatabase.QueryData basDatabase.sCreate_SQL
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_TOCMsg_Details (End)"
    '-v1.6.1
    '
End Sub

' ROUTINE:  Get_Message_Name
' AUTHOR:   Brad Brown
' PURPOSE:
' INPUT:    "iMsg_ID" is the ID of the message
' OUTPUT:   None
Public Function Get_Message_Name(iMsg_ID As Integer) As String
    Dim rsMessage As Recordset  ' Pointer to records in the Message table
 
        Set rsMessage = guCurrent.DB.OpenRecordset(guCurrent.sArchive & TBL_MESSAGE, dbOpenDynaset)

    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basTOC.Get_Message_Name(" & iMsg_ID & ")"
    '
    ' Trap errors
    On Error GoTo ERR_HANDLER
    '
    
    rsMessage.FindFirst "Msg_ID = " & iMsg_ID
    '
    ' Look for a match
    If rsMessage.NoMatch Then
        Get_Message_Name = basCCAT.GetAlias("Message names", "CC_MSGID" & iMsg_ID, "MESSAGE#" & iMsg_ID)
    Else
        Get_Message_Name = UCase(rsMessage!MSG_NAME)
    End If
    rsMessage.Close
    Exit Function
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basTOC.Get_Message_Name (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    Err.Raise Err.Number, "CCAT:Get_Message_Name", Err.Description
End Function
' ROUTINE:  Get_Message_Selection
' AUTHOR:   Brad Brown
' PURPOSE:
' INPUT:    "iMsg_ID" is the ID of the message
' OUTPUT:   None
Public Function Get_Message_Selection(rsMessage As Recordset, iMsg_ID As Integer) As Boolean
 

    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basTOC.Get_Message_Selection(" & iMsg_ID & ")"
    '
    ' Trap errors
    On Error GoTo ERR_HANDLER
    '
    rsMessage.FindFirst "Msg_ID = " & iMsg_ID
    '
    ' Look for a match
    If rsMessage.NoMatch Then
        Get_Message_Selection = False
    Else
        Get_Message_Selection = rsMessage!Select_Msg
    End If

    Exit Function
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basTOC.Get_Message_Selection (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    Err.Raise Err.Number, "CCAT:Get_Message_Name", Err.Description
End Function

Public Function Init_Msg_Proc_Structs() As Boolean
    Dim rsTable As Recordset
    Dim sInfileName As String
    Dim i As Integer
    
    iCurrLevel = 0
    For i = LBound(iLevelArray) To UBound(iLevelArray)
        iLevelArray(i) = 0
    Next i
    Set rsTable = guCurrent.DB.OpenRecordset("SELECT * FROM " & TBL_ARCHIVES & " WHERE Name = '" & guCurrent.sArchive & "'")
    If Not rsTable.NoMatch Then sInfileName = rsTable!Archive
    '
    ' Look for a match
    Init_Msg_Proc_Structs = False
    If (Open_Specified_File(sInfileName, giRawInputFile, Binary_Read) = False) Then
            ' print out error message and exit
        MsgBox "Inputfile error", vbExclamation, "File error"
    Else
        Set rsVarStruct = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_VarStruct", dbOpenDynaset)
        Set rsTOC = guCurrent.DB.OpenRecordset(guCurrent.sArchive & TBL_TOC, dbOpenDynaset)
        Set rsProcData = guCurrent.DB.OpenRecordset(guCurrent.sArchive & TBL_PROC_DATA, dbOpenDynaset)
        Init_Msg_Proc_Structs = True
    End If

End Function

Public Sub Close_Msg_Proc_Structs()

    ' Trap errors
    On Error GoTo ERR_HANDLER

    Close #giRawInputFile
    rsVarStruct.Close
    rsTOC.Close
    Exit Sub
    
ERR_HANDLER:
        '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basTOC.Close_Msg_Proc_Structs (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    'Err.Raise Err.Number, "CCAT:Get_Message_Name", Err.Description


End Sub

Public Sub Init_Msg_Proc()
    Dim iMaxNode As Integer
    Dim varOldLine As Variant
    Dim varNewLine As Variant
    Dim nTreeNodes() As Node
    Dim boolDifferent As Boolean
    Dim iLevel As Integer
    Dim iLastIndex As Integer
    Dim rsTOCTime As Recordset
    Dim itmX As ListItem
    Dim iUnion As Integer
    Dim lStartIndex As Long
    
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_ProcMsg_Record(" & iMsg_ID & ")"
    rsVarStruct.FindFirst "MsgId = " & guCurrent.iMessage
    If (rsVarStruct.NoMatch = False) Then
        frmtreeproc.tvVarStruct.Nodes.Clear
        'check if special processing required
        If (guCurrent.iMessage = 127) Then
            frmtreeproc.rbOutput(2).Visible = True
            frmtreeproc.rbOutput(2).Value = True
        Else
            frmtreeproc.rbOutput(2).Visible = False
            frmtreeproc.rbOutput(0).Value = True
        End If
        iMaxNode = -1
        varOldLine = Split(" ", ".")
        lNumNodes = 1
        lCurNode = 0
        lPrevNode = 0
        iUnion = 0
        lStartIndex = rsVarStruct!varStructID
        While (Not (rsVarStruct.NoMatch))
         ReDim tmpCheck(lStartIndex To rsVarStruct!varStructID)
            If ((rsVarStruct!DataType <> "STRUCT END") And (rsVarStruct!DataType <> "UNION END")) Then
                varNewLine = Split(rsVarStruct!fieldname, ".")
                If (UBound(varNewLine) > iMaxNode) Then
                    iMaxNode = UBound(varNewLine)
                    ReDim Preserve nTreeNodes(0 To iMaxNode)
                End If
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
                            Set nTreeNodes(iLevel) = frmtreeproc.tvVarStruct.Nodes.Add(, , , varNewLine(iLevel))
                        Else
                            Set nTreeNodes(iLevel) = frmtreeproc.tvVarStruct.Nodes.Add(nTreeNodes((iLevel - 1)).Index, tvwChild, , varNewLine(iLevel))
                            If (rsVarStruct!DataType = "UNION BEGIN") Then
                                iUnion = iUnion + 1
                            End If

                            If (rsVarStruct!MultiEntry = 1) Then
                                nTreeNodes(iLevel).Bold = True
                            End If
                            If (iUnion = 1) Then
                                nTreeNodes(iLevel).ForeColor = vbRed
                            ElseIf (iUnion = 2) Then
                                nTreeNodes(iLevel).ForeColor = vbBlue
                            End If
                            nTreeNodes(iLevel).EnsureVisible
                            frmtreeproc.tvVarStruct.Nodes(lNumNodes - 1).Checked = True
                            frmtreeproc.tvVarStruct.Nodes(lNumNodes - 1).Tag = rsVarStruct!varStructID
                        End If
                     End If
                Next iLevel
                varOldLine = varNewLine
            ElseIf (rsVarStruct!DataType = "UNION END") Then
               iUnion = iUnion - 1
            End If
            rsVarStruct.FindNext "MsgId = " & guCurrent.iMessage
        Wend
        For lStartIndex = LBound(tmpCheck) To UBound(tmpCheck)
           tmpCheck(lStartIndex) = 1
        Next lStartIndex
    End If
    
    On Error GoTo ERR_HANDLER

    rsTOC.FindFirst "MsgId = " & guCurrent.iMessage
    frmtreeproc.SelTimeCombo.Clear
    frmtreeproc.ListView1.ListItems.Clear
    While (Not (rsTOC.NoMatch))
        With frmtreeproc.SelTimeCombo
            .AddItem rsTOC!Time
            .ItemData(.NewIndex) = rsTOC!MsgTOCIndex
        End With
        With frmtreeproc.ListView1
            Set itmX = .ListItems.Add(, , rsTOC!Time)
            itmX.Tag = rsTOC!MsgTOCIndex
        End With
        rsTOC.FindNext "MsgId = " & guCurrent.iMessage
    Wend
    Exit Sub
    
ERR_HANDLER:
    MsgBox ("Error listing all times")
    
End Sub

Public Sub SetChecks(lNode As Long, blnChecked As Boolean)

    If (lNode <> 1) Then
        With frmtreeproc.tvVarStruct
           If (blnChecked) Then
              tmpCheck(.Nodes.Item(lNode).Tag) = 1
           Else
              tmpCheck(.Nodes.Item(lNode).Tag) = 0
           End If
        End With
    End If
   
End Sub

Public Sub SetAllChecks(blnChecked As Boolean)
   Dim i As Long
   
   With frmtreeproc.tvVarStruct
      If (blnChecked) Then
         For i = LBound(tmpCheck) To UBound(tmpCheck)
            tmpCheck(i) = 1
         Next i
      Else
         For i = LBound(tmpCheck) To UBound(tmpCheck)
            tmpCheck(i) = 0
         Next i
      End If
   End With
End Sub

Function agSwapBytes%(ByVal src%)
    Dim tmpByteArray(1 To 2) As Byte
    Dim tmpByteArray2(1 To 2) As Byte
    Dim i As Integer
    
    Call CopyMemory(tmpByteArray(1), src, 2)
    tmpByteArray2(2) = tmpByteArray(1)
    tmpByteArray2(1) = tmpByteArray(2)
    Call CopyMemory(agSwapBytes%, tmpByteArray2(1), 2)
End Function

Function agSwapWords&(ByVal src&)
    Dim tmpByteArray(1 To 4) As Byte
    Dim tmpByteArray2(1 To 4) As Byte
    Dim i As Integer
    
    Call CopyMemory(tmpByteArray(1), src, 4)
    tmpByteArray2(4) = tmpByteArray(1)
    tmpByteArray2(3) = tmpByteArray(2)
    tmpByteArray2(2) = tmpByteArray(3)
    tmpByteArray2(1) = tmpByteArray(4)
    Call CopyMemory(agSwapWords&, tmpByteArray2(1), 4)
End Function




