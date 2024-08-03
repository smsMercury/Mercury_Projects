Attribute VB_Name = "Raw2filtraw"
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' Author:           Brad Brown
' Filename:         Raw2filtraw.BAS
' Classification:   Unclassified
' Purpose:          This module is the main for the task that converts a raw archive file
'                   into a raw file that has been filtered by message type. Only the applicable
'                   message types will be copied to the new raw file. In addition the archive
'                   header information will be removed leaving only message header information
'                   and message body. This function additionally will populate a Table of
'                   contents file that contains a guide to every occurance of a desired message.
'                   It will also generate a summary file that contains a general summary (count,
'                   first time, last time) of each applicable message type.
'
'Inputs:
'Ouputs:
' Version
'   v1.6.0  TAE Removed the code that updated the progress bar while filtering
'           TAE Modified code to use the Wizard instead of the Archive Options form
'           TAE Modified code to initialize the summary table from the INI file
'           TAE Replaced Proc_Raw_Main with a function that returns a boolean status
'           TAE Replaced Proc_Raw_File with a function that returns a boolean status
'   v1.6.1  TAE Added verbose logging calls
'

Public Type Summary_Record
    iMsgCount As Long
    dTimeFirst As Double
    dTimeLast As Double
End Type 'Summary_Record

Public Type Toc_Record
    iMsgId As Integer
    lMsgCount As Long
    dTimeStamp As Double
    iStartByte As Long
    iMsgSize As Integer
End Type 'Toc_Record

Public Enum File_Action
    Binary_Read = 0
    Binary_Write = 1
    Text_Read = 2
    Text_Write = 3
End Enum

'Raytheon archive header
Private Type Csc_Msg_Hdr                    '   16 bytes fixed
    lMsgLength As Long                      '   4 bytes
    iFromId As Integer                      '   2 bytes
    iToId As Integer                        '   2 bytes
    iToSocket As Integer                    '   2 bytes
    iMsgType As Integer                     '   2 bytes
    lPad As Long                            '   4 bytes
End Type

' Archive header
Private Type Ray_Arc_Hdr                    '   4 bytes Fixed
    uCscMsgHdr As Csc_Msg_Hdr               '   16 bytes
    lTimestamp As Double                    '   4 bytes
    msgTime As TodTime
End Type 'Ray_Arc_Hdr
    
'Private Type Arc_Hdr
'    lTimestamp As Long
'End Type 'Arc_Hdr

Public giRawInputFile As Integer
Public giRawOutputFile As Integer
Global Const sArchiveExt = ".cca"
Global Const sFilteredExt = ".flt"


Public iJdate As Integer
Public dTimeZero As Double

Global Const iMsgSizeMax = 4096   'temporary size allocation for all CC messages
Global Const iMsgIdMax = 32767   'temporary size allocation for all CC messages
Global Const iTocStep = 100     ' this will be redimensioned later
Public iTocCount As Long     ' Count of number of messages in TOC
Public iTocStart As Long     ' Starting location in filtered file
Public iTocHiWater As Long   ' Point that we need to add more entries
Public gaSummary(1 To iMsgIdMax) As Summary_Record

'
' Module name:      Init_Summary_Table
' Classification:   Unclassified
' Purpose:
' Inputs:
' Outputs:
'
Public Sub Init_Summary_Table()
    '
    '+v1.6TE
    Dim plngMsg_Count As Long
    Dim pstrMsg_Name As String
    Dim pintMsg_ID As Integer
    '-v1.6
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : Raw2filtraw.Init_Summary_Table (Start)"
    '-v1.6.1
    '
    ' initialize to -1 indicating that the message is not to be used
    For iCount = 1 To iMsgIdMax
        gaSummary(iCount).iMsgCount = -1
    Next
    '
    ' Grab file telling us which messages to use
    '
    '+v1.6TE
    For plngMsg_Count = 1 To basCCAT.GetNumber("Message List", "CC_Messages", 0)
        pstrMsg_Name = basCCAT.GetAlias("Message List", "CC_MSG" & plngMsg_Count, "")
        pintMsg_ID = basCCAT.GetNumber("Message ID", pstrMsg_Name & "ID", -9999)
        '
        ' See if we have a valid message
        If pstrMsg_Name <> "" And pintMsg_ID > 0 Then
            '
            ' See if it was selected by the user
            If frmWizard.IsSelected(pintMsg_ID) Then gaSummary(pintMsg_ID).iMsgCount = 0
        End If
    Next plngMsg_Count

 
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : Raw2filtraw.Init_Summary_Table (End)"
    '-v1.6.1
    '

End Sub 'Init_Summary_Table

'
' Module name:      Open_Specified_File
' Classification:   Unclassified
' Purpose:
' Inputs:
' Outputs:
'
Public Function Open_Specified_File(sFileName As String, iFilename As Integer, eAction As File_Action) As Boolean
    Dim iAttr As Integer
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.Open_Specified_File (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: File = " & sFileName & ", #" & iFilename & ", Action = " & eAction
    End If
    '-v1.6.1
    '

    On Error GoTo OpenError
    iFilename = FreeFile

    If Dir(sFileName) <> "" Then
        iAttr = GetAttr(sFileName)
        If (iAttr And vbReadOnly) Then SetAttr sFileName, iAttr Xor vbReadOnly
    End If

    Select Case eAction
        Case Binary_Write
            Open sFileName For Binary Access Write As #iFilename
        Case Binary_Read
            Open sFileName For Binary Access Read As #iFilename
        Case Text_Write
            Open sFileName For Output As #iFilename
        Case Text_Read
            Open sFileName For Input As #iFilename
        Case Else
            GoTo OpenError
    End Select
    
    
    On Error Resume Next
    Open_Specified_File = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.Open_Specified_File (End)"
    '-v1.6.1
    '
    'Return
    Exit Function

OpenError:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : Raw2filtraw.Open_Specified_File (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '

    Open_Specified_File = False
    'Return


End Function ' Open_Specified_File
'
' FUNCTION: GetJdate
' AUTHOR:   Brad Brown
' PURPOSE:  Extract system JDate from message buffer
' INPUT:    iMsgLength = the length of the buffer in bytes
'           bytBuff() = the buffer
' OUTPUT:   The JDate
' NOTES:
Public Function GetJdate(iMsgLength As Integer, bytBuff() As Byte) As Integer
Dim iStart As Integer
Dim uMttimesyn As Mttimesyn
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.GetJdate (Start)"
    '-v1.6.1
    '

    iStart = LBound(bytBuff)
    Call CopyMemory(uMttimesyn.uSystime, bytBuff(iStart), iMsgLength)

   GetJdate = uSystime.iBeginRfosDate
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.GetJdate (End) = " & GetJdate
    '-v1.6.1
    '

End Function
'
' FUNCTION: GetTimeZero
' AUTHOR:   Brad Brown
' PURPOSE:
' INPUT:    iMsgLength = The length of the buffer in bytes
'           bytBuff = the buffer
' OUTPUT:   TRUE =
'           FALSE =
' NOTES:
Public Function GetTimeZero(iMsgLength As Integer, bytBuff() As Byte) As Boolean
Dim iStart As Integer
Dim uMtsetacqsmode As Mtsetacqsmode
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.GetTimeZero (Start)"
    '-v1.6.1
    '

    iStart = LBound(bytBuff)
    Call CopyMemory(uMtsetacqsmode.bytSubMode, bytBuff(iStart), iMsgLength)

    If (uMtsetacqsmode.bytSubMode = 3) Then ' is this the start of set environment
        GetTimeZero = True
    Else
        GetTimeZero = False
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.GetTimeZero (End) = " & GetTimeZero
    '-v1.6.1
    '

End Function
'
'+v1.6TE
' FUNCTION: blnFilterRawArchive
' AUTHOR:   Tom Elkins
' PURPOSE:  Copy of Proc_Raw_Main -- returns a boolean indicating success
' INPUT:    sInfileName = The input file name
'           sOutfileName = The output file name
'           rayArc = whether or not we are dealing with a Ratheon archive or not
' OUTPUT:   TRUE = Filtering was successful
'           FALSE = Filtering failed
' NOTES:
Public Function blnFilterRawArchive(sInfileName As String, sOutfileName As String, rayArc As Boolean) As Boolean
    Dim pintRetVal As Integer
    Dim pblnProcess_Filt As Boolean
    Dim pstrTmp_Filename As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.blnFilterRawArchive (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: Input file = " & sInfileName & ", Output file = " & sOutfileName
    End If
    '-v1.6.1
    '
    pblnProcess_Filt = False
    '
    ' Open Input binary file (Raw)
    If (Open_Specified_File(sInfileName, giRawInputFile, Binary_Read) = False) Then
        '
        ' print out error message and exit
        MsgBox "Inputfile error", vbExclamation, "File error"
    Else
        '
        ' Open Output binary file (Filtered Raw)
        If (Open_Specified_File(sOutfileName, giRawOutputFile, Binary_Write) = False) Then
            '
            ' print out error message and exit
            MsgBox "Outputfile error", vbExclamation, "File error"
        Else
            '
            ' Initialize message array (summary table)
            ' Initialize table of contents
            ' Process messages
   
            Init_Summary_Table
            'Init_TOC
            If rayArc Then
                pblnProcess_Filt = blnProcessRayFile
            Else
                pblnProcess_Filt = blnProcessRawFile
            End If
            '
            ' need to do something with the summary table
            ' and with the TOC. Put them in access???
        End If
    End If
    '
    ' Close both opened files
    Close #giRawInputFile
    Close #giRawOutputFile
    '
    'v1.6TE
    'If (pblnProcess_Filt = True) Then
    '    '
    '    On Error Resume Next
    '    frmArchive.barProgress.Value = 0
    '    frmArchive.barProgress.Max = FileLen(sOutfileName)
    '    guCurrent.uArchive.lFile_Size = FileLen(sOutfileName)
    '    On Error GoTo 0
    '    '
    '    ProcFiltMain sOutfileName
    'End If
    blnFilterRawArchive = pblnProcess_Filt
    '-v1.6
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.blnFilterRawArchive (End) = " & blnFilterRawArchive
    '-v1.6.1
    '
    
End Function 'blnFilterRawArchive
'
'+v1.6TE
' FUNCTION: blnProcessRawFile
' AUTHOR:   Tom Elkins
' PURPOSE:  Copy of Proc_Raw_File -- returns boolean status
' INPUT:    none
' OUTPUT:   TRUE = Processing complete
'           FALSE = Processing failed
' NOTES:
Public Function blnProcessRawFile() As Boolean
    Dim tmpArcHdr As Arc_Hdr
    Dim tmpMsgHdr As Msg_Hdr
    Dim tmpByteCount As Integer
    Dim tmpMsgData() As Byte
    Dim lSwapTimestamp As Long
    Dim iSwapMsgId As Integer
    Dim iSwapMsgSize As Integer
    Const lBlocksize As Long = 32768
    Dim lBytesRemain As Long
    Dim iMinSize As Integer
    Dim dTimeStamp As Double
    Dim uTOCMsg As Toc_Record
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.blnProcessRawFile (Start)"
    '-v1.6.1
    '

    On Error GoTo Hell
    blnProcessRawFile = False
    '
    lBytesRemain = lBlocksize
    iMinSize = (LenB(tmpArcHdr) + LenB(tmpMsgHdr))
    iTocStart = 1
    iTocCount = 0
    While Not EOF(giRawInputFile)
    
        If (lBytesRemain < iMinSize) Then
            If (lBytesRemain > 0) Then
                ReDim tmpMsgData(1 To lBytesRemain)
                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
                iTocStart = iTocStart + lBytesRemain
                basCCAT.WriteLogEntry "ERROR    : Raw2filtraw.blnProcessRawFile (Skipping " & lBytesRemain & " bytes - no header room)"
            End If
            lBytesRemain = lBlocksize
            '
            ' Take advantage of this interruption to update the form
            DoEvents
            '
        End If
        
        Get giRawInputFile, , tmpArcHdr
        Get giRawInputFile, , tmpMsgHdr
        iTocStart = iTocStart + LenB(tmpArcHdr)

        '
        On Error Resume Next
        '
        '+v1.6TE
        'frmArchive.barProgress.Value = Loc(giRawInputFile)
        'frmArchive.lblProcessInfo.Caption = "Filtering File - " & Int(frmArchive.barProgress.Value / (frmArchive.barProgress.Max / 100#)) & "% Complete"
        frmWizard.barProgress.Value = Loc(giRawInputFile)
        frmWizard.lblPctDone.Caption = "Filtering File - " & Int(frmWizard.barProgress.Value / (frmWizard.barProgress.Max / 100#)) & "% Complete"
        '-v1.6
        DoEvents
        On Error GoTo 0
        '
        lBytesRemain = lBytesRemain - iMinSize
        lSwapTimestamp = agSwapWords&(tmpArcHdr.lTimestamp)
        iSwapMsgId = agSwapBytes%(tmpMsgHdr.iMsgId)
        iSwapMsgSize = agSwapBytes%(tmpMsgHdr.iMsgLength)
        
        If (iSwapMsgId = MTARCFILLERID) Then
            If (lBytesRemain > 0) Then
                ReDim tmpMsgData(1 To lBytesRemain)
                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
 '               iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
            End If
            iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
 '           basCCAT.WriteLogEntry "INFO     : Raw2filtraw.blnProcessRawFile (MTARCFILLERID: Skipping " & lBytesRemain & " bytes)"
            lBytesRemain = lBlocksize
        ElseIf ((iSwapMsgId < 1) Or (iSwapMsgId > iMsgIdMax)) Then
            If (lBytesRemain > 0) Then
                ReDim tmpMsgData(1 To lBytesRemain)
                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
 '               iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
            End If
            iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
            basCCAT.WriteLogEntry "ERROR    : Raw2filtraw.blnProcessRawFile (PROC_RAW_FILE: Bad Message ID: " & iSwapMsgId & " Skipping " & lBytesRemain & " bytes)"
            lBytesRemain = lBlocksize
        ElseIf ((lSwapTimestamp <= 0) Or (lSwapTimestamp > 864000#)) Then
            If (lBytesRemain > 0) Then
                ReDim tmpMsgData(1 To lBytesRemain)
                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
 '               iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
            End If
            iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
            basCCAT.WriteLogEntry "ERROR    : Raw2filtraw.blnProcessRawFile (Bad Message Time: " & lSwapTimestamp & " Skipping " & lBytesRemain & " bytes)"
            lBytesRemain = lBlocksize
        'ElseIf ((iSwapMsgSize <= 8) Or (iSwapMsgSize > iMsgSizeMax)) Then
        ElseIf ((iSwapMsgSize < 8) Or (iSwapMsgSize > iMsgSizeMax)) Then
            If (lBytesRemain > 0) Then
                ReDim tmpMsgData(1 To lBytesRemain)
                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
 '               iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
            End If
            iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
            basCCAT.WriteLogEntry "ERROR    : Raw2filtraw.blnProcessRawFile (Bad Message Size: " & iSwapMsgSize & " Skipping " & lBytesRemain & " bytes)"
            lBytesRemain = lBlocksize
        Else
            tmpByteCount = iSwapMsgSize - 8
            If (tmpByteCount > 0) Then
                'Dimension array to data size
                ReDim tmpMsgData(1 To tmpByteCount)
                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
			End If
			lBytesRemain = lBytesRemain - tmpByteCount						  
			dTimeStamp = lSwapTimestamp / 10#
			iTocCount = iTocCount + 1               'v1.7B
			uTOCMsg.dTimeStamp = dTimeStamp         'v1.7B
			uTOCMsg.lMsgCount = iTocCount
			uTOCMsg.iMsgId = iSwapMsgId             'v1.7B
			uTOCMsg.iMsgSize = iSwapMsgSize         'v1.7B
			uTOCMsg.iStartByte = iTocStart          'v1.7B
			iTocStart = iTocStart + iSwapMsgSize    'v1.7B
			Call basTOC.Add_TOC_Record(uTOCMsg)
			Call basTOC.Add_Summary_Record(uTOCMsg.iMsgId, uTOCMsg.dTimeStamp)
            ' is it a message that we want
			If ((tmpByteCount > 0) And (gaSummary(iSwapMsgId).iMsgCount >= 0)) Then
				Put giRawOutputFile, , tmpArcHdr
				Put giRawOutputFile, , tmpMsgHdr
				Put giRawOutputFile, , tmpMsgData()
				' is this the first occurance of this message
                If (gaSummary(iSwapMsgId).iMsgCount = 0) Then
                    gaSummary(iSwapMsgId).dTimeFirst = dTimeStamp
					basCCAT.WriteLogEntry "INFO		: Raw2filtraw.blnProcessRawFile(First Occurance of Message ID " & iSwapMsgId & ")"	
                End If ' message that we want
				' set last occurance time
				gaSummary(iSwapMsgId).dTimeLast = dTimeStamp
				'Increment message count
				gaSummary(iSwapMsgId).iMsgCount = gaSummary(iSwapMsgId).iMsgCount + 1
			End if ' message that we want	
            'End If ' byte count > 0
        End If 'multiple check
    Wend
    blnProcessRawFile = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.blnProcessRawFile (End)"
    '-v1.6.1
    '
    Exit Function
'
'
Hell:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : Raw2filtraw.blnProcessRawFile (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
End Function
'
'+v1.6TE
' FUNCTION: blnProcessRayFile
' AUTHOR:   Tom Elkins
' PURPOSE:  Copy of Proc_Ray_File -- returns boolean status
' INPUT:    none
' OUTPUT:   TRUE = Processing complete
'           FALSE = Processing failed
' NOTES:
Public Function blnProcessRayFile() As Boolean
    Dim tmpRayHdr As Ray_Arc_Hdr
    Dim tmpCscHdr As Csc_Msg_Hdr
    Dim tmpArcHdr As Arc_Hdr
    Dim tmpMsgHdr As Msg_Hdr
    Dim tmpByteCount As Integer
    Dim tmpMsgData() As Byte
    Dim tmpFudge() As Byte
    Dim lSwapTimestamp As Long
    Dim iSwapMsgId As Integer
    Dim iSwapMsgSize As Integer
    Const iRfosMsgType1 As Integer = 1014
    Const iRfosMsgType2 As Integer = 5
    Dim iFudge As Integer
    Dim irfosMsg As Integer
    Dim lBytesRemain As Long
    Dim iMinSize As Integer
    Dim dTimeStamp As Double
    Dim uTOCMsg As Toc_Record
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.blnProcessRawFile (Start)"
    '-v1.6.1
    '

    On Error GoTo Hell
    blnProcessRayFile = False
    '
    lBytesRemain = 0
    iMinSize = (LenB(tmpRayHdr) + LenB(tmpCscHdr))
    iTocStart = 1
    iTocCount = 0
    While Not EOF(giRawInputFile) And Not blnProcessRayFile
        '
        ' Take advantage of this interruption to update the form
        DoEvents
        '
        Get giRawInputFile, , tmpRayHdr
        Dim ltemptoc As Long
        ltemptoc = iTocStart
        iTocStart = iTocStart + LenB(tmpRayHdr)
        lSwapTimestamp = Filtraw2Das.LongTOD(tmpRayHdr.msgTime)
        tmpArcHdr.lTimestamp = agSwapWords&(lSwapTimestamp * 10)
        
        lBytesRemain = agSwapWords&(tmpRayHdr.uCscMsgHdr.lMsgLength) - LenB(tmpRayHdr)
        irfosMsg = agSwapBytes%(tmpRayHdr.uCscMsgHdr.iMsgType)
        If lBytesRemain >= LenB(tmpCscHdr) Then
            If irfosMsg = iRfosMsgType1 Then
                ReDim tmpMsgData(1 To lBytesRemain)
                Get giRawInputFile, , tmpMsgData
                iTocStart = iTocStart + lBytesRemain
                GoTo EndWhile
            Else
                Get giRawInputFile, , tmpCscHdr
                iTocStart = iTocStart + LenB(tmpCscHdr)
                irfosMsg = agSwapBytes%(tmpCscHdr.iMsgType)
                lBytesRemain = lBytesRemain - LenB(tmpCscHdr)
                If irfosMsg <> iRfosMsgType2 Then
                    ReDim tmpMsgData(1 To lBytesRemain)
                    Get giRawInputFile, , tmpMsgData
                    iTocStart = iTocStart + lBytesRemain
                    GoTo EndWhile
                End If
            End If
            Get giRawInputFile, , tmpMsgHdr
                lBytesRemain = lBytesRemain - LenB(tmpMsgHdr)
            '
            On Error Resume Next
            
            frmWizard.barProgress.Value = Loc(giRawInputFile)
            frmWizard.lblPctDone.Caption = "Filtering File - " & Int(frmWizard.barProgress.Value / (frmWizard.barProgress.Max / 100#)) & "% Complete"
            '-v1.6
            DoEvents
            On Error GoTo 0
            '
            'lSwapTimestamp = agSwapWords&(tmpArcHdr.lTimestamp)
            iSwapMsgId = agSwapBytes%(tmpMsgHdr.iMsgId)
            iSwapMsgSize = agSwapBytes%(tmpMsgHdr.iMsgLength)
        
            If ((iSwapMsgId < 1) Or (iSwapMsgId > iMsgIdMax)) Then
                If (lBytesRemain > 0) Then
                    ReDim tmpMsgData(1 To lBytesRemain)
                    Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
                End If
                iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
                basCCAT.WriteLogEntry "ERROR     : Raw2filtraw.blnProcessRawFile (MTARCFILLERID: Skipping " & lBytesRemain & " bytes)"
            ElseIf ((iSwapMsgSize <= 8) Or (iSwapMsgSize > iMsgSizeMax)) Then
            If (lBytesRemain > 0) Then
                ReDim tmpMsgData(1 To lBytesRemain)
                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
            End If
            iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
            basCCAT.WriteLogEntry "ERROR    : Raw2filtraw.blnProcessRawFile (Bad Message Size: " & iSwapMsgSize & " Skipping " & lBytesRemain & " bytes)"
            lBytesRemain = lBlocksize
        Else
            If (lBytesRemain > 0) Then
                'Dimension array to data size
                Dim tmpHex As String
                Dim i As Integer
                Dim tmpHdr(1 To LenB(tmpMsgHdr)) As Byte
                
                Call CopyMemory(tmpHdr(1), tmpMsgHdr, LenB(tmpMsgHdr))
                tmpByteCount = iSwapMsgSize - 8
                iFudge = lBytesRemain - tmpByteCount
                If (iFudge > 0) Then
                    ReDim tmpMsgData(1 To tmpByteCount)
                    ReDim tmpFudge(1 To iFudge)
                    Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
                    Get giRawInputFile, , tmpFudge
                Else
                    tmpByteCount = lBytesRemain
                    ReDim tmpMsgData(1 To tmpByteCount)
                    Get giRawInputFile, , tmpMsgData()
                End If
                ReDim tmpMsgData(1 To tmpByteCount)
                ReDim tmpFudge(1 To iFudge)
                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
                Get giRawInputFile, , tmpFudge
           
                dTimeStamp = lSwapTimestamp
                iTocCount = iTocCount + 1               'v1.7B
                uTOCMsg.dTimeStamp = dTimeStamp         'v1.7B
                uTOCMsg.lMsgCount = iTocCount
                uTOCMsg.iMsgId = iSwapMsgId             'v1.7B
                uTOCMsg.iMsgSize = iSwapMsgSize         'v1.7B
                uTOCMsg.iStartByte = iTocStart          'v1.7B
                iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
                Call basTOC.Add_TOC_Record(uTOCMsg)
                Call basTOC.Add_Summary_Record(uTOCMsg.iMsgId, uTOCMsg.dTimeStamp)
                ' is it a message that we want
                If (gaSummary(iSwapMsgId).iMsgCount >= 0) Then
                    Put giRawOutputFile, , tmpArcHdr
                    Put giRawOutputFile, , tmpMsgHdr
                    Put giRawOutputFile, , tmpMsgData()
                    ' is this the first occurance of this message
                    If (gaSummary(iSwapMsgId).iMsgCount = 0) Then
                        gaSummary(iSwapMsgId).dTimeFirst = dTimeStamp
                        basCCAT.WriteLogEntry "INFO     : Raw2filtraw.blnProcessRawFile (First Occurance of Message ID " & iSwapMsgId & ")"
                    End If
                    ' set last occurance time
                    gaSummary(iSwapMsgId).dTimeLast = dTimeStamp
                    'Increment message count
                    gaSummary(iSwapMsgId).iMsgCount = gaSummary(iSwapMsgId).iMsgCount + 1
                End If ' message that we want
            End If ' byte count > 0
        End If 'multiple check
    Else
        blnProcessRayFile = True
    End If
EndWhile:
    Wend
    blnProcessRayFile = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : Raw2filtraw.blnProcessRawFile (End)"
    '-v1.6.1
    '
    Exit Function
'
'
Hell:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : Raw2filtraw.blnProcessRawFile (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
End Function

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



