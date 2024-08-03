Attribute VB_Name = "Filtraw2Das"
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' Author:           Brad Brown
' Filename:         Filtraw2Das.BAS
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
' Revisions:
'   v1.3.5  BDB Added HB ID Mapping from AID/SigID to index
'   v1.4.0  TAE Added MsgHdr.From to Origin_ID field
'           TAE Added error trapping to stop processing on errors without bombing
'           TAE Tied some message boxes to help file and added help button
'   v1.5.0  TAE Trap custom error to terminate translation process
'           TAE Added context-sensitive help for some error messages
'           TAE Modified the RCV mapping to be consistent
'           TAE Added a way to check RCV components in the INI file if the combined RCV does not exist
'           BDB Modified code to translate frequency information from MTSDALARM, MTDFALARM, and MTDFSDALARM messages
'           BDB Modified code to translate LOB information from MTLOBSETRSLT messages
'           TAE Added error handling to all message translation routines
'   v1.6.0  BDB Added check for CCOS version to process message differences
'           BDB Added LOB quality to DAS Flag field
'           BDB Added routine to process the MTSSERROR message
'           TAE Added property to store CCOS version
'           TAE Added routine to translate ORIGIN_ID code to Subsystem
'           TAE Redesigned early-termination methodology
'           TAE Modified processing routines to return a boolean completion status
'           TAE Removed extraneous variable declarations
'           TAE Modified code to use the new Archive Wizard instead of the Archive Options form
'           TAE Modified error trap to allow the user to terminate the entire process
'           TAE Added error trap to handle the Subscript out of range error for MTSDALARM to allow
'               the user to delete existing reports and switch to the alternate processing method,
'               or to delete existing reports and ignore the message for the rest of the processing.
'           TAE Changed ECHMOD use of Origin ID and put the Ech Type in the Status field
'   v1.8.0  SPV Modified for Block 35 changes.  Modified messages include:  mtsigupd, mthbactrep, mtanaalarm,
'               mtanarslt, mtrunmode, mthbxmtrstat.
'
Public giFiltInputFile As Integer       ' Input file ID
'
'+v1.6TE
'Public giSigOutputFile As Integer
'Public giLobOutputFile As Integer
'Public giStfOutputFile As Integer
'Public giMtfOutputFile As Integer
'-v1.6
Public gdFreq(0 To 6000) As Double      ' Internal frequency table for matching SigID
Public glEmitter(0 To 6000) As Long     ' Internal sigID table
Public gsEmitter(0 To 6000) As String   ' Internal source table
Public guPlanAreaCtr As LatLon          ' Planning area center  BB 2002
Public gvPlanAreaLat As Variant
Public gvPlanAreaLon As Variant
Public guNavPositionLat As Double          ' Latest Nav Position  BB 2002
Public guNavPositionLon As Double          ' Latest Nav Position  BB 2002

'
'+v1.6TE
Private mblnContinue As Boolean         ' Flag to terminate the processing
Private msngVersion As Single           ' Stores the CCOS version

Global libGeo As Geodesy

'Declare Sub Cart2Geo Lib "DASGeodesyLib.dll" Alias "Cartesian_To_Geodetic" (X As Variant, Y As Variant, Z As Variant, _
'            Origin_Latitude As Variant, Origin_Longitude As Variant, Origin_Altitude As Variant, _
'            ByRef Latitude As Variant, ByRef Longitude As Variant, ByRef Altitude As Variant)
'-v1.6
'
' Module name:      ProcFiltMain
' Author:           Brad Brown
' Classification:   Unclassified
' Purpose:
' Inputs:           strInFileName  : Name of filtered raw data input file
' Outputs:          TRUE if the file was filtered completely
'                   FALSE if the file was not filtered completely
'+v1.6TE
'Sub ProcFiltMain(strInFileName As String)
Public Function ProcFiltMain(strInFileName As String) As Boolean
'    Dim iRetVal As Integer
'-v1.6
    '
    ' Open Input binary file (Filtered Raw)

    If (Open_Specified_File(strInFileName, giFiltInputFile, Binary_Read) = False) Then
        '
        ' Print out error message and exit
        '+v1.6TE
        'iRetVal = MsgBox("Inputfile error", vbExclamation, "File error")
        MsgBox "Input File Error", vbExclamation, "File error"
        '-v1.6
    Else
        ' Process messages
        '+v1.6TE
        'Proc_Filt_File
        ProcFiltMain = Proc_Filt_File
        '-v1.6
    End If
    '
    ' Close opened file
    Close #giFiltInputFile

End Function 'ProcFiltMain
'
' Module name:      Proc_Filt_File
' Author:           Brad Brown
' Classification:   Unclassified
' Purpose:
' Inputs:
' Outputs:          TRUE if the file was filtered
'                   FALSE if there was an error
'+v1.6TE
'Public Sub Proc_Filt_File()
Public Function Proc_Filt_File() As Boolean
'-v1.6
    Dim pusrTmp_Msg_Hdr As Msg_Hdr
    Dim pintTmp_Byte_Count As Integer
    Dim pusrTmp_Arc_Hdr As Arc_Hdr
    Dim pabytTmp_Msg_Data() As Byte
    Dim pabytTmp_Msg_Data2() As Byte
    Dim plngTmp_Msg_Count As Long
    'Dim pusrMTSIGUPD As Mtsigupd
    Dim pdblTime As Double
    Dim pintSwap_Msg_ID As Integer
    Dim pintSwap_Msg_Size As Integer

    
    '
    ' Initialize message count
    plngTmp_Msg_Count = 0
    gvPlanAreaLat = CVar(basCCAT.GetAlias("Planning Area", "LAT", 0)) ' BB 2002
    gvPlanAreaLon = CVar(basCCAT.GetAlias("Planning Area", "LON", 0)) ' BB 2002
    guPlanAreaCtr.lLat = modDegToBam32(gvPlanAreaLat)
    guPlanAreaCtr.lLon = modDegToBam32(gvPlanAreaLon)
    Set libGeo = New Geodesy
    '
    ' Trap errors
    On Error GoTo ERR_HANDLER
    '
    '+v1.6TE
    ' Loop while we can
    'While (Not EOF(giFiltInputFile))
    While (Not EOF(giFiltInputFile)) And mblnContinue
    '-v1.6
        plngTmp_Msg_Count = plngTmp_Msg_Count + 1
        '
        ' Read the archive header and message header
        Get giFiltInputFile, , pusrTmp_Arc_Hdr
        Get giFiltInputFile, , pusrTmp_Msg_Hdr
        '
        ' Swap bytes to make sense
        pintSwap_Msg_ID = agSwapBytes%(pusrTmp_Msg_Hdr.iMsgId)
        pintSwap_Msg_Size = agSwapBytes%(pusrTmp_Msg_Hdr.iMsgLength)
        pdblTime = agSwapWords&(pusrTmp_Arc_Hdr.lTimestamp)
        '
        ' Convert time to seconds
        pdblTime = pdblTime / 10
        '
        ' Store the message size
        pintTmp_Byte_Count = pintSwap_Msg_Size
        '
        ' See if there is a valid message size
        If (pintTmp_Byte_Count > 8) Then
            '
            'Dimension array to data size
            ReDim pabytTmp_Msg_Data2(1 To (pintTmp_Byte_Count - 8))
            ReDim pabytTmp_Msg_Data(1 To pintTmp_Byte_Count)
            '
            ' need to read it even if we don't write it
            Get giFiltInputFile, , pabytTmp_Msg_Data2()
            Call CopyMemory(pabytTmp_Msg_Data(1), pusrTmp_Msg_Hdr, LenB(pusrTmp_Msg_Hdr))
            Call CopyMemory(pabytTmp_Msg_Data(9), pabytTmp_Msg_Data2(1), pintTmp_Byte_Count - 8)
            '
            ' is it the message that we want
            On Error Resume Next
            '
            ' See if the message is selected for processing.
            ' The Messages tab of the Archive form allows the user to enable or
            ' disable individual messages.  If a message is disabled, the program
            ' will not process it.
            '
            '+v1.6TE
            'If Not frmArchive.lvMessages.ListItems("MSG" & pintSwap_Msg_ID).Checked Then pintSwap_Msg_ID = 0
            If Not frmWizard.IsSelected(pintSwap_Msg_ID) Then pintSwap_Msg_ID = 0
            '
            'On Error GoTo 0
            '-v1.6
            On Error GoTo ERR_HANDLER
            '
            ' If the message ID is one of interest, process it
            Select Case (pintSwap_Msg_ID)
                Case MTDEFPMAID
                    Call modDasMtdefpma(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTSIGALARMID
                    '
                    Call modDasMtsigalarm(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTANAALARMID
                    '
                    '+v1.8.0SPV
                    If msngVersion >= 3 Then
                        Call modDasMtanaalarm3_0(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    Else
                        Call modDasMtanaalarm(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    End If
                    '-v1.8.0
                Case MTHBDYNRSPID
                    '
                    Call modDasMthbdynrsp(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTSDALARMID
                    '
                    '+v1.6BBTE
                    If msngVersion >= 2.3 Then
                        Call modDasMtsdalarm2_3(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    Else
                    '-v1.6
                        Call modDasMtsdalarm(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    '
                    '+v1.6TE
                    End If
                    '-v1.6
                Case MTHBACTREPID
                    '
                    '+v1.8.0SPV
                    If msngVersion >= 3 Then
                        Call modDasMthbactrep3_0(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    Else
                        Call modDasMthbactrep(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    End If
                    '-vv1.8.0
                Case MTDFALARMID
                    Call modDasMtdfalarm(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTRFSTATUSID
                    'Call modDasMtrfstatus(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTDFSDALARMID
                    Call modDasMtdfsdalarm(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTLOBUPDID
                    Call modDasMtlobupd(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTANARSLTID
                    '
                    '+v1.8.0SPV
                    If msngVersion >= 3 Then
                        Call modDasMtanarslt3_0(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    Else
                        Call modDasMtanarslt(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    End If
                    '-v1.8.0
                Case MTSIGUPDID
                    '
                    '+v1.8.0SPV
                    If msngVersion >= 3 Then
                        Call modDasMtsigupd3_0(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    Else
                        Call modDasMtsigupd(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    End If
                    '-v1.8.0
                    '+1.8.7SPV
                Case MTLOBRSLTID
                    Call modDasMtlobrslt(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    '+1.8.7SPV
                Case MTHBLOBUPDID
                    Call modDasMthblobupd(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTULDDATAID
                    Call modDasMtulddata(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTLOBSETRSLTID
                    Call modDasMtlobsetrslt(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTFIXRSLTID
                    Call modDasMtfixrslt(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTHBSIGUPDID
                    Call modDasMthbsigupd(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTHBGSREPID
                    '
                    Call modDasMthbgsrep(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTHBSEMISTATID
                    'Call modDasMthbsemistat(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTHBSELASKID
                    Call modDasMthbselask(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTJAMSTATID
                    Call modDasMtjamstat(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTRUNMODEID
                    '
                    '+v1.8.0SPV
                    If msngVersion >= 3 Then
                        Call modDasMtrunmode3_0(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    Else
                        Call modDasMtrunmode(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    End If
                    '-v1.8.0
                Case MTHBSELJAMID
                    Call modDasMthbseljam(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTHBXMTRSTATID
                    '
                    '+v1.8.0SPV
                    If msngVersion >= 3 Then
                        Call modDasMthbxmtrstat3_0(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    Else
                        Call modDasMthbxmtrstat(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    End If
                    '-v1.8.0
                Case MTNAVREPID
                    Call modDasMtnavrep(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTLHCORRELATEID
                    Call modDasMtlhcorrelate(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTLHECHMODID
                    Call modDasMtlhechmod(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTLHTRACKUPDID
                    Call modDasMtlhtrackupd(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTLHTRACKREPID
                    '
                    '+v1.8.0SPV
                    If msngVersion >= 3 Then
                        Call modDasMtlhtrackrep3_0(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    Else
                        Call modDasMtlhtrackrep(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    End If
                    '-v1.8.0
                Case MTLHBEARINGONLYID
                    '-v1.8.12
                    Call modDasMtlhtrackrep3_0(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTSETACQSMODEID
                    '
                    '+v1.6BB
                    'Call modDasMtsetacqsmode(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    '-v1.6
                '
                '+v1.6BB
                Case MTSSERRORID
                    Call modDasMtsserror(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                '-v1.6
                'BB 2002
                Case MTPLANAREAID
                    Call modDasMtplanarea(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTANAREQID
                    Call modDasMtanareq(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                 'BB 2002
                Case MTDFFLGSID
                    Call modDasMtdfflgs(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                 'BB 2002
                  'BB 2003
                Case MTRRSLTID
                    Call modDasMtrrslt(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                 'BB 2003
                Case MTHBGSUPSTATID
                    Call modDasMthbgsupstat(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTHBGSDELID
                    Call modDasMthbgsdel(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTHBASSREPID
                    Call modDasMthbassrep(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                Case MTSYSCONFIGID
                    '+v1.8.2SPV
                    If msngVersion >= 3 Then
                        Call modDasMtsysconfig3_0(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    Else
                        Call modDasMtsysconfig(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                    End If
                    '-v1.8.2
                Case MTENVSTATID
                    Call modDasMtenvstat(pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data)
                '
                '+v1.8.10 TE
                Case MTCNTTRGID: modDASMTCNTTRG pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data
                Case MTTXRFCONFID: modDASMTTXRFCONF pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data
'                Case MTTGTLISTUPDID: modDasMTTGTLISTUPD pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data
'                Case MTTGTLISTUPDID: modMTTGTLISTUPD.Process_MTTGTLISTUPD pdblTime, pintTmp_Byte_Count, pabytTmp_Msg_Data
                '-v1.8.10 TE
                '
               Case Else
                  Debug.Print
                    'Call modDasBoot
            End Select
        End If 'byte count > 0
    Wend
    Set libGeo = Nothing
    '
    '+v1.6TE
    Proc_Filt_File = mblnContinue
    '-v1.6
Exit Function
'
'
ERR_HANDLER:
    Select Case Err.Number
        Case 3026: ' Out of disk space
            '+v1.5
            ' Display help file page for the disk full error
            'MsgBox "The disk is full, please remove any files you don't need," & vbCr & "or move the database to another disk and re-process.", vbOKOnly Or vbMsgBoxHelpButton, "Disk Full", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, IDH_Error3026
            MsgBox "The disk is full, please remove any files you don't need," & vbCr & "or move the database to another disk and re-process.", vbOKOnly Or vbMsgBoxHelpButton, "Disk Full", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, basCCAT.IDH_ERR_3026
            '-v1.5
        Case 53: ' Can't find file
            Dim sFile As String
            sFile = Trim(Mid(Err.Description, InStr(1, Err.Description, ":") + 1))
            If Dir(App.Path & "\" & sFile) <> "" Then
                '+v1.5
                'MsgBox "Copy the file '" & App.Path & "\" & sFile & "'" & vbCr & "to your WINNT\SYSTEM32 directory and re-process", vbOKOnly Or vbMsgBoxHelpButton, "File missing", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, IDH_Error53
                MsgBox "Copy the file '" & App.Path & "\" & sFile & "'" & vbCr & "to your WINNT\SYSTEM32 directory and re-process", vbOKOnly Or vbMsgBoxHelpButton, "File missing", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, basCCAT.IDH_ERR_53
                '-v1.5
            Else
                '+v1.5
                'MsgBox "Cannot find the file '" & sFile & "'." & vbCr & "This file is required for processing.", vbOKOnly Or vbMsgBoxHelpButton, "Missing File", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, IDH_Error53
                MsgBox "Cannot find the file '" & sFile & "'." & vbCr & "This file is required for processing.", vbOKOnly Or vbMsgBoxHelpButton, "Missing File", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, basCCAT.IDH_ERR_53
                '-v1.5
            End If
        '
        '+v1.6TE
        ''+v1.5
        '' Trap the custom error to establish translation termination
        'Case vbObjectError + 911:
        '    MsgBox "User prematurely terminated the translation process", vbOKOnly Or vbMsgBoxHelpButton, "Translation Incomplete", App.HelpFile, lHelpID
        ''-v1.5
        '-v1.6
        '
        Case 49:
            Resume Next
            
        Case Else:
            MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "From modDAS" & basCCAT.GetAlias("Message Names", "CC_MSGID" & pintSwap_Msg_ID, "UNKNOWN") & " while translating archive", vbOKOnly, "Translation error"
    End Select
    On Error GoTo 0
End Function
'
'
' ROUTINE:  modDasMtdefpma
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTDEFPMA message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtdefpma(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTDEFPMA As Mtdefpma
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintLoop As Integer
    Dim pintParent As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTDEFPMA, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTDEFPMA"
50  pusrDAS_Rec.sReport_Type = "GEO"
60  For pintLoop = 1 To pusrMTDEFPMA.uPMA.bytNumberOfVertices
70      pusrDAS_Rec.lOrigin_ID = pintLoop
80      pusrDAS_Rec.sOrigin = "PMA" & Str(pintLoop)
90      If pintLoop = 1 Then
100         pintParent = pusrMTDEFPMA.uPMA.bytNumberOfVertices
        Else
110         pintParent = pintLoop - 1
        End If
120     pusrDAS_Rec.lParent_ID = pintParent
130     pusrDAS_Rec.sParent = "PMA" & pintParent
140     pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrMTDEFPMA.uPMA.uPmaVertices(pintLoop).lLat))
150     pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrMTDEFPMA.uPMA.uPmaVertices(pintLoop).lLon))
160     Call Add_Data_Record(MTDEFPMAID, pusrDAS_Rec)
       'Call Process_MTDEFPMA(pusrDAS_Rec)
    Next pintLoop
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtdefpma"
    '
    ' Process the error
    Select Case plngErr_Num
        
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtdefpma", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTDEFPMA message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTDEFPMA", App.HelpFile, basCCAT.IDH_TRANSLATE_MTDEFPMA)
                Case vbAbort:
                    mblnContinue = False

                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtsigalarm
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTSIGALARM message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtsigalarm(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTSIGALARM As Mtsigalarm
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Sig As Integer
    Dim pintLoop As Integer
    Dim pusrTmp_Alarm As SignalAlarmData
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTSIGALARM)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTSIGALARM, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTSIGALARM"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTSIGALARM.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrMTSIGALARM.uAlarm(0).uData.uData.lCtrFreq))
    '
    '+v1.6TE
100 pusrDAS_Rec.lFlag = CLng(agSwapBytes%(pusrMTSIGALARM.uAlarm(0).iSignalStatus))
110 pusrDAS_Rec.sSupplemental = "BW:" & modFreqConv(agSwapWords&(pusrMTSIGALARM.uAlarm(0).uData.uData.lBandwidth)) & ",PK:" & modFreqConv(agSwapWords&(pusrMTSIGALARM.uAlarm(0).uData.uData.lPeakFrequency))
120 pusrDAS_Rec.dPRI = CDbl(modFreqConv(agSwapWords&(pusrMTSIGALARM.uAlarm(0).uData.uData.iAmplitude)))
130 pusrDAS_Rec.lStatus = CLng(pusrMTSIGALARM.uAlarm(0).uData.bytRfStatus)
    '-v1.6
    '
140 Call Add_Data_Record(MTSIGALARMID, pusrDAS_Rec)
    'Call Process_MTSIGALARM(pusrDAS_Rec)
150 pintNum_Sig = agSwapBytes%(pusrMTSIGALARM.iNumAlarms)
    'basCCAT.Double_Check MTSIGALARMID, intMsg_Length, pintStruct_Length, pintNum_Sig, LenB(pusrTmp_Alarm)
160 If pintNum_Sig > 1 Then
170     For pintLoop = 1 To (pintNum_Sig - 1)
180         Call CopyMemory(pusrTmp_Alarm, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Alarm) * (pintLoop - 1))), LenB(pusrTmp_Alarm))
190         pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrTmp_Alarm.uData.uData.lCtrFreq))
            '
            '+v1.6TE
200         pusrDAS_Rec.lFlag = CLng(agSwapBytes%(pusrTmp_Alarm.iSignalStatus))
210         pusrDAS_Rec.sSupplemental = "BW:" & modFreqConv(agSwapWords&(pusrTmp_Alarm.uData.uData.lBandwidth)) & ",PK:" & modFreqConv(agSwapWords&(pusrTmp_Alarm.uData.uData.lPeakFrequency))
220         pusrDAS_Rec.dPRI = CDbl(modFreqConv(agSwapWords&(pusrTmp_Alarm.uData.uData.iAmplitude)))
230         pusrDAS_Rec.lStatus = CLng(pusrTmp_Alarm.uData.bytRfStatus)
            '-v1.6
            '
240         Call Add_Data_Record(MTSIGALARMID, pusrDAS_Rec)
            'Call Process_MTSIGALARM(pusrDAS_Rec)
        Next pintLoop
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsigalarm"
    '
    ' Process the error
    Select Case plngErr_Num
        
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtsigalarm", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSIGALARM message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSIGALARM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSIGALARM)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtanaalarm
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTANAALARM message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtanaalarm(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTANAALARM As Mtanaalarm
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pbytChan As Byte
    Dim pintStart As Integer
    Dim pintTemp_SigID As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTANAALARM, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTANAALARM"
50  pusrDAS_Rec.sReport_Type = "SIG"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTANAALARM.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
80  pbytChan = LBound(pusrMTANAALARM.uChannelData)

90  With pusrMTANAALARM
100     pintTemp_SigID = agSwapBytes%(.iSigID)
110     If (pintTemp_SigID = -1) Then
120         pusrDAS_Rec.sSignal = "NEW_SIG"
        Else
130         pusrDAS_Rec.sSignal = "EXISTING_SIG"
        End If
140     pusrDAS_Rec.lSignal_ID = pintTemp_SigID
150     pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(.lFrequency))
160     pusrDAS_Rec.lStatus = .uChannelData(pbytChan).bytChannelActive
170     pusrDAS_Rec.sEmitter = modLmbSigToString(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
180     pusrDAS_Rec.lEmitter_ID = modLMBSigToLng(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
190     pusrDAS_Rec.lTag = .uChannelData(pbytChan).uSignalType.bytVariant
195     pusrDAS_Rec.sSupplemental = "Usage : " & .uChannelData(pbytChan).bytSigUsage
200     Call Add_Data_Record(MTANAALARMID, pusrDAS_Rec)
    End With
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtanaalarm"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtanaalarm", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTANAALARM message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTANAALARM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTANAALARM)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtanaalarm3_0
' AUTHOR:   Shaun Vogel
' PURPOSE:  Translate the MTANAALARM message to DAS data structure for blk35
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:    New for version 1.8.0
'
Public Sub modDasMtanaalarm3_0(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTANAALARM As Mtanaalarm3_0
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pbytChan As Byte
    Dim pintStart As Integer
    Dim pintTemp_SigID As Integer
    '
    On Error GoTo Hell
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTANAALARM, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTANAALARM"
50  pusrDAS_Rec.sReport_Type = "SIG"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTANAALARM.uMsgHdr.iMsgFrom)
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
80  pbytChan = LBound(pusrMTANAALARM.uChannelData)

90  With pusrMTANAALARM
100     pintTemp_SigID = agSwapBytes%(.iSigID)
110     If (pintTemp_SigID = -1) Then
120         pusrDAS_Rec.sSignal = "NEW_SIG"
        Else
130         pusrDAS_Rec.sSignal = "EXISTING_SIG"
        End If
140     pusrDAS_Rec.lSignal_ID = pintTemp_SigID
150     pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(.lFrequency))
160     pusrDAS_Rec.lStatus = .uChannelData(pbytChan).bytChannelActive
170     pusrDAS_Rec.sEmitter = modLmbSigToString(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
180     pusrDAS_Rec.lEmitter_ID = modLMBSigToLng(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
190     pusrDAS_Rec.lTag = .uChannelData(pbytChan).uSignalType.bytVariant
195     pusrDAS_Rec.sSupplemental = "Usage : " & .uChannelData(pbytChan).bytSigUsage
200     Call Add_Data_Record(MTANAALARMID, pusrDAS_Rec)
        'Call Process_MTANAALARM(pusrDAS_Rec)
    End With
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtanaalarm3_0"
    '
    ' Process the error
    Select Case plngErr_Num
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTANAALARM message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTANAALARM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTANAALARM)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
End Sub
'
'
' ROUTINE:  modDasMthbdynrsp
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBDYNRSP message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbdynrsp(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBDYNRSP As Mthbdynrsp
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Resp As Integer
    Dim pintLoop As Integer
    Dim pusrTmp_Resp As HbdynRspRec
    'Dim lHbIndex As Long
    Dim pintHBID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBDYNRSP)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBDYNRSP, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBDYNRSP"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBDYNRSP.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrMTHBDYNRSP.uHbDynRsp(0).bytSignal), CInt(gbytNOT_DEFINED))
100 pintHBID = modGetHBID(CLng(pusrMTHBDYNRSP.uHbDynRsp(0).bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
110 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHBID, "UNKNOWN")
120 pusrDAS_Rec.lEmitter_ID = pusrDAS_Rec.lSignal_ID
130 pusrDAS_Rec.lStatus = pusrMTHBDYNRSP.uHbDynRsp(0).uHbDynRsp(0).bytResponse
140 pusrDAS_Rec.lTag = pusrMTHBDYNRSP.uHbDynRsp(0).uHbDynRsp(0).uHbDynRec.bytOnList
150 pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrMTHBDYNRSP.uHbDynRsp(0).uHbDynRsp(0).uHbDynRec.lFrequency))
160 pusrDAS_Rec.sSupplemental = modHBFuncToString(Int(pusrMTHBDYNRSP.uHbDynRsp(0).bytFunction)) & " " & basCCAT.GetAlias("Hbopt", "HBOPT" & pusrMTHBDYNRSP.uHbDynRsp(0).uHbDynRsp(0).uHbDynRec.bytOption, "UNKNOWN") & " " & pusrMTHBDYNRSP.uHbDynRsp(0).uHbDynRsp(0).uHbDynRec.bytOption

170 Call Add_Data_Record(MTHBDYNRSPID, pusrDAS_Rec)
    'Call Process_MTHBDYNRSP(pusrDAS_Rec)
   
180 pintNum_Resp = agSwapBytes%(pusrMTHBDYNRSP.iNumOfRsps)
190 If pintNum_Resp > 1 Then
200     For pintLoop = 1 To (pintNum_Resp - 1)
210         Call CopyMemory(pusrTmp_Resp, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Resp) * (pintLoop - 1))), LenB(pusrTmp_Resp))
220         pusrDAS_Rec.lStatus = pusrTmp_Resp.uHbDynRsp(0).bytResponse
230         pusrDAS_Rec.lTag = pusrTmp_Resp.uHbDynRsp(0).uHbDynRec.bytOnList
240         pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrTmp_Resp.uHbDynRsp(0).uHbDynRec.lFrequency))
250         pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrTmp_Resp.bytSignal), CInt(gbytNOT_DEFINED))
260         pintHBID = modGetHBID(CLng(pusrTmp_Resp.bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
270         pusrDAS_Rec.lEmitter_ID = pusrDAS_Rec.lSignal_ID
280         pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHBID, "UNKNOWN")
290         pusrDAS_Rec.sSupplemental = modHBFuncToString(Int(pusrTmp_Resp.bytFunction)) & " " & basCCAT.GetAlias("Hbopt", "HBOPT" & pusrTmp_Resp.uHbDynRsp(0).uHbDynRec.bytOption, "UNKNOWN") & " " & pusrTmp_Resp.uHbDynRsp(0).uHbDynRec.bytOption

300         Call Add_Data_Record(MTHBDYNRSPID, pusrDAS_Rec)
            'Call Process_MTHBDYNRSP(pusrDAS_Rec)
        Next pintLoop
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbdynrsp"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbdynrsp", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBDYNRSP message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBDYNRSP", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBDYNRSP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtsdalarm
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTSDALARM message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtsdalarm(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTSDALARM As Mtsdalarm
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Sig As Integer
    Dim pintNum_Freqs As Integer                                            'SCR 14
    Dim i As Integer, j As Integer, k As Integer                            'SCR 14
    Dim pusrTmp_Alarm As ShortDurationAlarmData
    Dim pdblTmp_Acq_Freq As Double                                          'SCR 14
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTSDALARM)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTSDALARM, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTSDALARM"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTSDALARM.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  With pusrMTSDALARM.uAlarm(0)
100     pdblTmp_Acq_Freq = modFreqConv(agSwapWords&(.uAlarm.lCtrFreq))        'SCR 14
110     pusrDAS_Rec.dFrequency = pdblTmp_Acq_Freq                                 'SCR 14
120     pusrDAS_Rec.sSupplemental = ""
130     Call Add_Data_Record(MTSDALARMID, pusrDAS_Rec)
140     pusrDAS_Rec.sSupplemental = "Acq Freq : " & pdblTmp_Acq_Freq              'SCR 14
150     pusrDAS_Rec.sReport_Type = "VEC"

160     For i = 1 To .bytNumberOfLobs
170         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uDeferredLobs(i).uLobData.uLobData.uAcLoc.lLat))
180         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uDeferredLobs(i).uLobData.uLobData.uAcLoc.lLon))
190         pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.uDeferredLobs(i).uLobData.uLobData.iTrueBearing))
200         pusrDAS_Rec.lStatus = .uDeferredLobs(i).uLobData.uLobData.bytInOutPma
            '
            '+v1.6BB
210         pusrDAS_Rec.lFlag = .uDeferredLobs(i).uLobData.uLobData.bytQualFactor
            '-v1.6
220         pintNum_Freqs = agSwapBytes%(.uDeferredLobs(i).uLobData.iNumberOfFrequencys) 'SCR 14
230         For k = 1 To pintNum_Freqs                                      'SCR 14
240             pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(.uDeferredLobs(i).uLobData.lFrequency(k))) 'SCR 14
250             Call Add_Data_Record(MTSDALARMID, pusrDAS_Rec)
            Next k                                                      'SCR 14
        Next i
    End With
   
260 pintNum_Sig = agSwapBytes%(pusrMTSDALARM.iNumAlarms)
    'basCCAT.Double_Check MTSDALARMID, intMsg_Length, pintStruct_Length, pintNum_Sig, LenB(pusrTmp_Alarm)
270 If pintNum_Sig > 1 Then

280     For i = 1 To (pintNum_Sig - 1)
290         pusrDAS_Rec.dLatitude = 0
300         pusrDAS_Rec.dLongitude = 0
310         pusrDAS_Rec.dBearing = 0
320         pusrDAS_Rec.lStatus = 0
330         Call CopyMemory(pusrTmp_Alarm, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Alarm) * (i - 1))), LenB(pusrTmp_Alarm))
340         pdblTmp_Acq_Freq = modFreqConv(agSwapWords&(pusrTmp_Alarm.uAlarm.lCtrFreq))    'SCR 14
350         pusrDAS_Rec.dFrequency = pdblTmp_Acq_Freq                                     'SCR 14
360         pusrDAS_Rec.sReport_Type = "SIG"
370         pusrDAS_Rec.sSupplemental = ""
380         Call Add_Data_Record(MTSDALARMID, pusrDAS_Rec)
390         pusrDAS_Rec.sSupplemental = "Acq Freq : " & pdblTmp_Acq_Freq                  'SCR 14
400         pusrDAS_Rec.sReport_Type = "VEC"
410         For j = 1 To pusrTmp_Alarm.bytNumberOfLobs
420             pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.uAcLoc.lLat))
430             pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.uAcLoc.lLon))
440             pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.iTrueBearing))
450             pusrDAS_Rec.lStatus = pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.bytInOutPma
                '
                '+v1.6BB
460             pusrDAS_Rec.lFlag = pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.bytQualFactor
                '-v1.6
470             pintNum_Freqs = agSwapBytes%(pusrTmp_Alarm.uDeferredLobs(j).uLobData.iNumberOfFrequencys) 'SCR 14
480             For k = 1 To pintNum_Freqs                                          'SCR 14
490                 pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrTmp_Alarm.uDeferredLobs(j).uLobData.lFrequency(k))) 'SCR 14
500                 Call Add_Data_Record(MTSDALARMID, pusrDAS_Rec)
                Next k                                                          'SCR 14
            Next j
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsdalarm"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtsdalarm", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        '+v1.6TE
        Case 9: ' Subscript out of range
            Select Case MsgBox("Error processing the MTSDALARM message.  This usually indicates that the wrong CCOS version was entered." & vbCrLf & vbCrLf & "You may choose to ABORT the processing and start again, RETRY the processing using an alternate processing routine, or IGNORE subsequent MTSDALARM messages and continue processing.", vbAbortRetryIgnore Or vbCritical, "MTSDALARM Processing Error")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    basDatabase.ExecuteSQLAction "DELETE FROM [" & frmWizard.txtArchiveName.Text & "_Data] WHERE Msg_Type = 'MTSDALARM'"
                    basDatabase.ExecuteSQLAction "DELETE FROM [" & frmWizard.txtArchiveName.Text & "_Summary] WHERE Message = 'MTSDALARM'"
                    msngVersion = 2.3
                    modDasMtsdalarm2_3 dblTime, intMsg_Length, abytBuffer
                
                Case vbIgnore:
                    basDatabase.ExecuteSQLAction "DELETE FROM [" & frmWizard.txtArchiveName.Text & "_Data] WHERE Msg_Type = 'MTSDALARM'"
                    basDatabase.ExecuteSQLAction "DELETE FROM [" & frmWizard.txtArchiveName.Text & "_Summary] WHERE Message = 'MTSDALARM'"
                    frmWizard.IsSelected(MTSDALARMID) = False
            End Select
        '-v1.6
        
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSDALARM message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSDALARM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSDALARM)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'+v1.6BB
' ROUTINE:  modDasMtsdalarm2_3
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTSDALARM message to DAS data structure for
'           version 2.3 of CCOS
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtsdalarm2_3(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTSDALARM As Mtsdalarm2_3
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Sig As Integer
    Dim pintNum_Freqs As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim pusrTmp_Alarm As ShortDurationAlarmData2_3
    Dim pdblTmp_Acq_Freq As Double
    '
    '
    On Error GoTo Hell
    '
10  pintStruct_Length = LenB(pusrMTSDALARM)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTSDALARM, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTSDALARM"
60  pusrDAS_Rec.sReport_Type = "SIG"

70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTSDALARM.uMsgHdr.iMsgFrom)
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
90  With pusrMTSDALARM.uAlarm(0)
100     pdblTmp_Acq_Freq = modFreqConv(agSwapWords&(.uAlarm.lCtrFreq))
110     pusrDAS_Rec.dFrequency = pdblTmp_Acq_Freq
120     pusrDAS_Rec.sSupplemental = ""
130     Call Add_Data_Record(MTSDALARMID, pusrDAS_Rec)
140     pusrDAS_Rec.sSupplemental = "Acq Freq : " & pdblTmp_Acq_Freq
150     pusrDAS_Rec.sReport_Type = "VEC"

160     For i = 1 To .bytNumberOfLobs
170         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uDeferredLobs(i).uLobData.uLobData.uAcLoc.lLat))
180         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uDeferredLobs(i).uLobData.uLobData.uAcLoc.lLon))
190         pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.uDeferredLobs(i).uLobData.uLobData.iTrueBearing))
200         pusrDAS_Rec.lStatus = .uDeferredLobs(i).uLobData.uLobData.bytInOutPma
210         pusrDAS_Rec.lFlag = .uDeferredLobs(i).uLobData.uLobData.bytQualFactor
220         pintNum_Freqs = agSwapBytes%(.uDeferredLobs(i).uLobData.iNumberOfFrequencys)
230         For k = 1 To pintNum_Freqs
240             pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(.uDeferredLobs(i).uLobData.lFrequency(k)))
250             Call Add_Data_Record(MTSDALARMID, pusrDAS_Rec)
                'Call Process_MTSDALARM(pusrDAS_Rec)
            Next k
        Next i
    End With
   
260 pintNum_Sig = agSwapBytes%(pusrMTSDALARM.iNumAlarms)
270 If pintNum_Sig > 1 Then

280     For i = 1 To (pintNum_Sig - 1)
290         pusrDAS_Rec.dLatitude = 0
300         pusrDAS_Rec.dLongitude = 0
310         pusrDAS_Rec.dBearing = 0
320         pusrDAS_Rec.lStatus = 0
330         Call CopyMemory(pusrTmp_Alarm, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Alarm) * (i - 1))), LenB(pusrTmp_Alarm))
340         pdblTmp_Acq_Freq = modFreqConv(agSwapWords&(pusrTmp_Alarm.uAlarm.lCtrFreq))
350         pusrDAS_Rec.dFrequency = pdblTmp_Acq_Freq
360         pusrDAS_Rec.sReport_Type = "SIG"
370         pusrDAS_Rec.sSupplemental = ""
380         Call Add_Data_Record(MTSDALARMID, pusrDAS_Rec)
390         pusrDAS_Rec.sSupplemental = "Acq Freq : " & pdblTmp_Acq_Freq
400         pusrDAS_Rec.sReport_Type = "VEC"
410         For j = 1 To pusrTmp_Alarm.bytNumberOfLobs
420             pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.uAcLoc.lLat))
430             pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.uAcLoc.lLon))
440             pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.iTrueBearing))
450             pusrDAS_Rec.lStatus = pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.bytInOutPma
460             pusrDAS_Rec.lFlag = pusrTmp_Alarm.uDeferredLobs(j).uLobData.uLobData.bytQualFactor
470             pintNum_Freqs = agSwapBytes%(pusrTmp_Alarm.uDeferredLobs(j).uLobData.iNumberOfFrequencys)
480             For k = 1 To pintNum_Freqs
490                 pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrTmp_Alarm.uDeferredLobs(j).uLobData.lFrequency(k))) 'SCR 14
500                 Call Add_Data_Record(MTSDALARMID, pusrDAS_Rec)
                    'Call Process_MTSDALARM(pusrDAS_Rec)
                Next k
            Next j
        Next i
    End If
    '
    '
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsdalarm2_3"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtsdalarm2_3", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        '+v1.6TE
        Case 9: ' Subscript out of range
            Select Case MsgBox("Error processing the MTSDALARM message.  This usually indicates that the wrong CCOS version was entered." & vbCrLf & vbCrLf & "You may choose to ABORT the processing, change the version, and start again, RETRY the processing using an alternate processing routine, or IGNORE subsequent MTSDALARM messages and continue processing.", vbAbortRetryIgnore Or vbCritical, "MTSDALARM Processing Error")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    basDatabase.ExecuteSQLAction "DELETE FROM [" & frmWizard.txtArchiveName.Text & "_Data] WHERE Msg_Type = 'MTSDALARM'"
                    basDatabase.ExecuteSQLAction "DELETE FROM [" & frmWizard.txtArchiveName.Text & "_Summary] WHERE Message = 'MTSDALARM'"
                    msngVersion = 2.2
                    modDasMtsdalarm dblTime, intMsg_Length, abytBuffer
                
                Case vbIgnore:
                    basDatabase.ExecuteSQLAction "DELETE FROM [" & frmWizard.txtArchiveName.Text & "_Data] WHERE Msg_Type = 'MTSDALARM'"
                    basDatabase.ExecuteSQLAction "DELETE FROM [" & frmWizard.txtArchiveName.Text & "_Summary] WHERE Message = 'MTSDALARM'"
                    frmWizard.IsSelected(MTSDALARMID) = False
            End Select
        '-v1.6

        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSDALARM message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSDALARM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSDALARM)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbactrep
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBACTREP message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbactrep(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBACTREP As Mthbactrep
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Act As Integer
    Dim i As Integer
    Dim pusrTmp_Act As ActRec
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBACTREP)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBACTREP, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBACTREP"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBACTREP.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrMTHBACTREP.uAct(0).bytSignal), CInt(pusrMTHBACTREP.uAct(0).bytChannel))
100 pintHB_ID = modGetHBID(CLng(pusrMTHBACTREP.uAct(0).bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
110 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrMTHBACTREP.uAct(0).bytChannel))
120 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrMTHBACTREP.uAct(0).bytChannel
130 pusrDAS_Rec.lStatus = pusrMTHBACTREP.uAct(0).bytStatus
140 pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrMTHBACTREP.uAct(0).lFreq))
150 pusrDAS_Rec.sSupplemental = modHBFuncToString(Int(pusrMTHBACTREP.uAct(0).bytFunction)) & " " & basCCAT.GetAlias("Hbopt", "HBOPT" & pusrMTHBACTREP.uAct(0).bytOption, "UNKNOWN") & " " & pusrMTHBACTREP.uAct(0).bytOption
160 Call Add_Data_Record(MTHBACTREPID, pusrDAS_Rec)
   
170 pintNum_Act = agSwapBytes%(pusrMTHBACTREP.iNumSigs)
180 If pintNum_Act > 1 Then
190     For i = 1 To (pintNum_Act - 1)
200         Call CopyMemory(pusrTmp_Act, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Act) * (i - 1))), LenB(pusrTmp_Act))
210         pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrTmp_Act.bytSignal), CInt(pusrTmp_Act.bytChannel))
220         pintHB_ID = modGetHBID(CLng(pusrTmp_Act.bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
230         pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrTmp_Act.bytChannel))
240         pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrTmp_Act.bytChannel
250         pusrDAS_Rec.lStatus = pusrTmp_Act.bytStatus
260         pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrTmp_Act.lFreq))
270         pusrDAS_Rec.sSupplemental = modHBFuncToString(Int(pusrTmp_Act.bytFunction)) & " " & basCCAT.GetAlias("Hbopt", "HBOPT" & pusrTmp_Act.bytOption, "UNKNOWN") & " " & pusrTmp_Act.bytOption
280         Call Add_Data_Record(MTHBACTREPID, pusrDAS_Rec)
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbactrep"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbactrep", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBACTREP message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBACTREP", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBACTREP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbactrep3_0
' AUTHOR:   Shaun Vogel
' PURPOSE:  Translate the MTHBACTREP message to DAS data structure for Blk35
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbactrep3_0(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBACTREP As Mthbactrep3_0
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Act As Integer
    Dim i As Integer
    Dim pusrTmp_Act As ActRec3_0
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBACTREP)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBACTREP, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBACTREP"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBACTREP.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrMTHBACTREP.uAct(0).bytSignal), CInt(pusrMTHBACTREP.uAct(0).bytChannel))
100 pintHB_ID = modGetHBID(CLng(pusrMTHBACTREP.uAct(0).bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
110 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrMTHBACTREP.uAct(0).bytChannel))
120 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrMTHBACTREP.uAct(0).bytChannel
130 pusrDAS_Rec.lStatus = pusrMTHBACTREP.uAct(0).bytStatus
140 pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrMTHBACTREP.uAct(0).lFreq))
150 pusrDAS_Rec.sSupplemental = modHBFuncToString(Int(pusrMTHBACTREP.uAct(0).bytFunction)) & " " & basCCAT.GetAlias("Hbopt", "HBOPT" & pusrMTHBACTREP.uAct(0).bytOption, "UNKNOWN") & " " & pusrMTHBACTREP.uAct(0).bytOption
160 Call Add_Data_Record(MTHBACTREPID, pusrDAS_Rec)
    'Call Process_MTHBACTREP(pusrDAS_Rec)
   
170 pintNum_Act = agSwapBytes%(pusrMTHBACTREP.iNumSigs)
180 If pintNum_Act > 1 Then
190     For i = 1 To (pintNum_Act - 1)
200         Call CopyMemory(pusrTmp_Act, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Act) * (i - 1))), LenB(pusrTmp_Act))
210         pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrTmp_Act.bytSignal), CInt(pusrTmp_Act.bytChannel))
220         pintHB_ID = modGetHBID(CLng(pusrTmp_Act.bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
230         pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrTmp_Act.bytChannel))
240         pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrTmp_Act.bytChannel
250         pusrDAS_Rec.lStatus = pusrTmp_Act.bytStatus
260         pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrTmp_Act.lFreq))
270         pusrDAS_Rec.sSupplemental = modHBFuncToString(Int(pusrTmp_Act.bytFunction)) & " " & basCCAT.GetAlias("Hbopt", "HBOPT" & pusrTmp_Act.bytOption, "UNKNOWN") & " " & pusrTmp_Act.bytOption
280         Call Add_Data_Record(MTHBACTREPID, pusrDAS_Rec)
            'Call Process_MTHBACTREP(pusrDAS_Rec)
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbactrep3_0"
    '
    ' Process the error
    Select Case plngErr_Num
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBACTREP message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBACTREP", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBACTREP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtdfalarm
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTDFALARM message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtdfalarm(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTDFALARM As Mtdfalarm
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Sig As Integer
    Dim i As Integer, j As Integer, k As Integer                         'SCR 14
    Dim pusrTmp_Alarm As DfAlarmData
    Dim pdblTmp_Acq_Freq As Double                                             'SCR 14
    Dim pintNum_Freqs As Integer                                             'SCR 14
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTDFALARM)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTDFALARM, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTDFALARM"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTDFALARM.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  With pusrMTDFALARM.uAlarm(0)
100     pdblTmp_Acq_Freq = modFreqConv(agSwapWords&(.uAlarm.uData.lCtrFreq)) 'SCR 14
110     pusrDAS_Rec.dFrequency = pdblTmp_Acq_Freq                              'SCR 14
120     pusrDAS_Rec.sSupplemental = ""
130     Call Add_Data_Record(MTDFALARMID, pusrDAS_Rec)
        'Call Process_MTDFALARM(pusrDAS_Rec)
140     pusrDAS_Rec.sSupplemental = "Acq Freq : " & pdblTmp_Acq_Freq           'SCR 14
150     pusrDAS_Rec.sReport_Type = "VEC"
160     For i = 1 To .bytNumberOfLobs
170         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uLobs(i).uLobData.uAcLoc.lLat))
180         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uLobs(i).uLobData.uAcLoc.lLon))
190         pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.uLobs(i).uLobData.iTrueBearing))
200         pusrDAS_Rec.lStatus = .uLobs(i).uLobData.bytInOutPma
            '
            '+v1.6BB
210         pusrDAS_Rec.lFlag = .uLobs(i).uLobData.bytQualFactor
            '-v1.6
220         pintNum_Freqs = agSwapBytes%(.uLobs(i).iNumberOfFrequencys)   'SCR 14
230         pusrDAS_Rec.lTag = pintNum_Freqs
            For k = 1 To pintNum_Freqs                                    'SCR 14
240             pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(.uLobs(i).lFrequency(k))) 'SCR 14
250             Call Add_Data_Record(MTDFALARMID, pusrDAS_Rec)
                'Call Process_MTDFALARM(pusrDAS_Rec)
            Next k                                                    'SCR 14
        Next i
   End With
   
260 pintNum_Sig = agSwapBytes%(pusrMTDFALARM.iNumAlarms)
270 If pintNum_Sig > 1 Then
280     For i = 1 To (pintNum_Sig - 1)
290         pusrDAS_Rec.dLatitude = 0
300         pusrDAS_Rec.dLongitude = 0
310         pusrDAS_Rec.dBearing = 0
320         pusrDAS_Rec.lStatus = 0
330         Call CopyMemory(pusrTmp_Alarm, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Alarm) * (i - 1))), LenB(pusrTmp_Alarm))
340         pdblTmp_Acq_Freq = modFreqConv(agSwapWords&(pusrTmp_Alarm.uAlarm.uData.lCtrFreq)) 'SCR 14
350         pusrDAS_Rec.dFrequency = pdblTmp_Acq_Freq                              'SCR 14
360         pusrDAS_Rec.sSupplemental = ""
370         pusrDAS_Rec.sReport_Type = "SIG"
380         Call Add_Data_Record(MTDFALARMID, pusrDAS_Rec)
            'Call Process_MTDFALARM(pusrDAS_Rec)
390         pusrDAS_Rec.sSupplemental = "Acq Freq : " & pdblTmp_Acq_Freq           'SCR 14
400         pusrDAS_Rec.sReport_Type = "VEC"
410         For j = 1 To pusrTmp_Alarm.bytNumberOfLobs
420             pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrTmp_Alarm.uLobs(j).uLobData.uAcLoc.lLat))
430             pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrTmp_Alarm.uLobs(j).uLobData.uAcLoc.lLon))
440             pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(pusrTmp_Alarm.uLobs(j).uLobData.iTrueBearing))
450             pusrDAS_Rec.lStatus = pusrTmp_Alarm.uLobs(j).uLobData.bytInOutPma
                '
                '+v1.6BB
460             pusrDAS_Rec.lFlag = pusrTmp_Alarm.uLobs(j).uLobData.bytQualFactor
                '-v1.6
470             pintNum_Freqs = agSwapBytes%(pusrTmp_Alarm.uLobs(j).iNumberOfFrequencys)   'SCR 14
                pusrDAS_Rec.lTag = pintNum_Freqs
480             For k = 1 To pintNum_Freqs                                    'SCR 14
490                 pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrTmp_Alarm.uLobs(j).lFrequency(k))) 'SCR 14
500                 Call Add_Data_Record(MTDFALARMID, pusrDAS_Rec)
                    'Call Process_MTDFALARM(pusrDAS_Rec)
                Next k                                                    'SCR 14
            Next j
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtdfalarm"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtdfalarm", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTDFALARM message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTDFALARM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTDFALARM)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtrfstatus
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTRFSTATUS message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtrfstatus(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTRFSTATUS As Mtrfstatus
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTRFSTATUS)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTRFSTATUS, abytBuffer(pintStart), pintStruct_Length)
40  pusrDAS_Rec.sReport_Type = "EVT"
    
50  pusrDAS_Rec.dReportTime = dblTime
60  pusrDAS_Rec.sMsg_Type = "MTRFSTATUS"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTRFSTATUS.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
   
90  Call Add_Data_Record(MTRFSTATUSID, pusrDAS_Rec)
    'Call Process_MTRFSTATUS(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtrfstatus"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtrfstatus", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTRFSTATUS message", vbAbortRetryIgnore Or vbExclamation, "Error Translating MTRFSTATUS")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtdfsdalarm
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTDFSDALARM message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtdfsdalarm(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTDFSDALARM As Mtdfsdalarm
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Sig As Integer
    Dim i As Integer, j As Integer, k As Integer                         'SCR 14
    Dim pusrTmp_Alarm As DfShortDurationResultsData
    Dim pdblTmp_Acq_Freq As Double                                             'SCR 14
    Dim pintNum_Freqs As Integer                                             'SCR 14
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTDFSDALARM)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTDFSDALARM, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTDFSDALARM"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTDFSDALARM.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  With pusrMTDFSDALARM.uAlarm(0)
100     pdblTmp_Acq_Freq = modFreqConv(agSwapWords&(.lAcqCenterFrequency)) 'SCR 14
110     pusrDAS_Rec.dFrequency = pdblTmp_Acq_Freq                              'SCR 14
120     pusrDAS_Rec.sSupplemental = ""
130     Call Add_Data_Record(MTDFSDALARMID, pusrDAS_Rec)
        'Call Process_MTDFSDALARM(pusrDAS_Rec)
140     pusrDAS_Rec.sSupplemental = "Acq Freq : " & pdblTmp_Acq_Freq           'SCR 14
150     pusrDAS_Rec.sReport_Type = "VEC"
160     For i = 1 To .bytNumberOfLobs
170         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uLobs(i).uLobData.uAcLoc.lLat))
180         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uLobs(i).uLobData.uAcLoc.lLon))
190         pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.uLobs(i).uLobData.iTrueBearing))
200         pusrDAS_Rec.lStatus = .uLobs(i).uLobData.bytInOutPma
            '
            '+v1.6BB
210         pusrDAS_Rec.lFlag = .uLobs(i).uLobData.bytQualFactor
            '-v1.6
220         pintNum_Freqs = agSwapBytes%(.uLobs(i).iNumberOfFrequencys)   'SCR 14
230         For k = 1 To pintNum_Freqs                                    'SCR 14
240             pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(.uLobs(i).lFrequency(k))) 'SCR 14
250             Call Add_Data_Record(MTDFSDALARMID, pusrDAS_Rec)          'SCR 14
                'Call Process_MTDFSDALARM(pusrDAS_Rec)
            Next k                                                    'SCR 14
        Next i
   End With
   
260 pintNum_Sig = agSwapBytes%(pusrMTDFSDALARM.iNumAlarms)
270 If pintNum_Sig > 1 Then
280     For i = 1 To (pintNum_Sig - 1)
290         pusrDAS_Rec.dLatitude = 0
300         pusrDAS_Rec.dLongitude = 0
310         pusrDAS_Rec.dBearing = 0
320         pusrDAS_Rec.lStatus = 0
330         Call CopyMemory(pusrTmp_Alarm, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Alarm) * (i - 1))), LenB(pusrTmp_Alarm))
340         pusrDAS_Rec.sReport_Type = "SIG"
350         pdblTmp_Acq_Freq = modFreqConv(agSwapWords&(pusrTmp_Alarm.lAcqCenterFrequency)) 'SCR 14
360         pusrDAS_Rec.dFrequency = pdblTmp_Acq_Freq                              'SCR 14
370         pusrDAS_Rec.sSupplemental = ""
380         Call Add_Data_Record(MTDFSDALARMID, pusrDAS_Rec)
            'Call Process_MTDFSDALARM(pusrDAS_Rec)
390         pusrDAS_Rec.sSupplemental = "Acq Freq : " & pdblTmp_Acq_Freq           'SCR 14
400         pusrDAS_Rec.sReport_Type = "VEC"
410         For j = 1 To pusrTmp_Alarm.bytNumberOfLobs
420             pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrTmp_Alarm.uLobs(j).uLobData.uAcLoc.lLat))
430             pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrTmp_Alarm.uLobs(j).uLobData.uAcLoc.lLon))
440             pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(pusrTmp_Alarm.uLobs(j).uLobData.iTrueBearing))
450             pusrDAS_Rec.lStatus = pusrTmp_Alarm.uLobs(j).uLobData.bytInOutPma
                '
                '+v1.6BB
460             pusrDAS_Rec.lFlag = pusrTmp_Alarm.uLobs(j).uLobData.bytQualFactor
                '-v1.6
470             pintNum_Freqs = agSwapBytes%(pusrTmp_Alarm.uLobs(j).iNumberOfFrequencys)   'SCR 14
480             For k = 1 To pintNum_Freqs                                    'SCR 14
490                 pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrTmp_Alarm.uLobs(j).lFrequency(k))) 'SCR 14
500                 Call Add_Data_Record(MTDFSDALARMID, pusrDAS_Rec)          'SCR 14
                    'Call Process_MTDFSDALARM(pusrDAS_Rec)
                Next k                                                    'SCR 14
            Next j
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtdfsdalarm"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtdfsdalarm", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTDFSDALARM message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTDFSDALARM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTDFSDALARM)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtlobupd
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTLOBUPD message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtlobupd(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTLOBUPD As Mtlobupd
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim i As Integer
    Dim pintTemp_SigID As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTLOBUPD, abytBuffer(pintStart), intMsg_Length)
   
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTLOBUPD"
50  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTLOBUPD.uMsgHdr.iMsgFrom)
    '+v1.6TE
60  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
70  pintTemp_SigID = agSwapBytes%(pusrMTLOBUPD.iSigID)
80  If (pintTemp_SigID = -1) Then
90      pusrDAS_Rec.sSignal = "NEW_SIG"
    Else
100     pusrDAS_Rec.sSignal = "EXISTING_SIG"
    End If
110 pusrDAS_Rec.lSignal_ID = pintTemp_SigID
    
120 If ((pintTemp_SigID >= LBound(gdFreq)) And (pintTemp_SigID <= UBound(gdFreq))) Then
130     pusrDAS_Rec.dFrequency = gdFreq(pintTemp_SigID)
140     pusrDAS_Rec.lEmitter_ID = glEmitter(pintTemp_SigID)
150     pusrDAS_Rec.sEmitter = gsEmitter(pintTemp_SigID)
    End If

160 pusrDAS_Rec.sReport_Type = "VEC"
170 For i = 1 To pusrMTLOBUPD.bytNumLobs
180     pusrDAS_Rec.lStatus = pusrMTLOBUPD.uLobData(i).bytInOutPma
        '
        '+v1.6BB
190     pusrDAS_Rec.lFlag = pusrMTLOBUPD.uLobData(i).bytQualFactor
        '-v1.6
200     pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrMTLOBUPD.uLobData(i).uAcLoc.lLat))
210     pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrMTLOBUPD.uLobData(i).uAcLoc.lLon))
220     pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(pusrMTLOBUPD.uLobData(i).iTrueBearing))
230     Call Add_Data_Record(MTLOBUPDID, pusrDAS_Rec)
        'Call Process_MTLOBUPD(pusrDAS_Rec)
    Next i
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtlobupd"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtlobupd", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTLOBUPD message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTLOBUPD", App.HelpFile, basCCAT.IDH_TRANSLATE_MTLOBUPD)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtanarslt
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTANARSLT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtanarslt(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTANARSLT As Mtanarslt
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim i As Integer
    Dim pbytChan As Byte
    Dim pintSigID As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTANARSLT, abytBuffer(pintStart), intMsg_Length)
30  pbytChan = LBound(pusrMTANARSLT.uChannelData)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTANARSLT"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTANARSLT.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
90  With pusrMTANARSLT
100     pintSigID = agSwapBytes%(.iSigID)
110     pusrDAS_Rec.lSignal_ID = pintSigID
120     If (pintSigID = -1) Then
130         pusrDAS_Rec.sSignal = "NEW_SIG"
        Else
140         pusrDAS_Rec.sSignal = "EXISTING_SIG"
        End If
150     pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(.lFreq))
160     pusrDAS_Rec.sEmitter = modLmbSigToString(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
170     pusrDAS_Rec.lEmitter_ID = modLMBSigToLng(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
180     pusrDAS_Rec.lTag = .uChannelData(pbytChan).uSignalType.bytVariant
185     pusrDAS_Rec.sSupplemental = "Usage : " & .uChannelData(pbytChan).bytSigUsage
190     pusrDAS_Rec.lParent_ID = .bytPassFreq
200     pusrDAS_Rec.lTarget_ID = .bytSignalPresentDf
210     pusrDAS_Rec.lStatus = .bytSignalPresentAna
220     pusrDAS_Rec.lFlag = agSwapBytes%(.iRequestorID)
     
    End With
        
224 Call Add_Data_Record(MTANARSLTID, pusrDAS_Rec)
    
228 pusrDAS_Rec.sReport_Type = "VEC"
230 For i = 1 To pusrMTANARSLT.bytNumLobs
240     With pusrMTANARSLT.uLobData(i).uLobData
250         pusrDAS_Rec.lStatus = .bytInOutPma
            '
            '+v1.6BB
260         pusrDAS_Rec.lFlag = .bytQualFactor
            '-v1.6
270         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uAcLoc.lLat))
280         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uAcLoc.lLon))
290         pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.iTrueBearing))
        End With
300     Call Add_Data_Record(MTANARSLTID, pusrDAS_Rec)
    Next i
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtanarslt"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtanarslt", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTANARSLTID message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTANARSLT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTANARSLT)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtanarslt3_0
' AUTHOR:   Shaun Vogel
' PURPOSE:  Translate the MTANARSLT message to DAS data structure for Blk35
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtanarslt3_0(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTANARSLT As Mtanarslt3_0
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim i As Integer
    Dim pbytChan As Byte
    Dim pintSigID As Integer
    '
    On Error GoTo Hell
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTANARSLT, abytBuffer(pintStart), intMsg_Length)
30  pbytChan = LBound(pusrMTANARSLT.uChannelData)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTANARSLT"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTANARSLT.uMsgHdr.iMsgFrom)
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
90  With pusrMTANARSLT
100     pintSigID = agSwapBytes%(.iSigID)
110     pusrDAS_Rec.lSignal_ID = pintSigID
120     If (pintSigID = -1) Then
130         pusrDAS_Rec.sSignal = "NEW_SIG"
        Else
140         pusrDAS_Rec.sSignal = "EXISTING_SIG"
        End If
150     pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(.lFreq))
160     pusrDAS_Rec.sEmitter = modLmbSigToString(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
170     pusrDAS_Rec.lEmitter_ID = modLMBSigToLng(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
180     pusrDAS_Rec.lTag = .uChannelData(pbytChan).uSignalType.bytVariant
185     pusrDAS_Rec.sSupplemental = "Usage : " & .uChannelData(pbytChan).bytSigUsage
190     pusrDAS_Rec.lParent_ID = .bytPassFreq
200     pusrDAS_Rec.lTarget_ID = .bytSignalPresentDf
210     pusrDAS_Rec.lStatus = .bytSignalPresentAna
220     pusrDAS_Rec.lFlag = agSwapBytes%(.iRequestorID)
     
    End With
        
224 Call Add_Data_Record(MTANARSLTID, pusrDAS_Rec)
    'Call Process_MTANARSLT(pusrDAS_Rec)
228 pusrDAS_Rec.sReport_Type = "VEC"
230 For i = 1 To pusrMTANARSLT.bytNumLobs
240     With pusrMTANARSLT.uLobData(i).uLobData
250         pusrDAS_Rec.lStatus = .bytInOutPma
260         pusrDAS_Rec.lFlag = .bytQualFactor
270         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uAcLoc.lLat))
280         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uAcLoc.lLon))
290         pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.iTrueBearing))
        End With
300     Call Add_Data_Record(MTANARSLTID, pusrDAS_Rec)
        'Call Process_MTANARSLT(pusrDAS_Rec)
    Next i
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtanarslt3_0"
    '
    ' Process the error
    Select Case plngErr_Num
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTANARSLTID message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTANARSLT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTANARSLT)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtsigupd
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTSIGUPD message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtsigupd(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTSIGUPD As Mtsigupd
    Dim i As Integer
    Dim pbytChan As Byte
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintSigID As Integer
    Dim pdblFreq As Double
    Dim plngEmitter As Long
    Dim pstrEmitter As String
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
   'If (dblTime >= (4882890 Mod 86400)) Then
   '   Debug.Print
   'End If
   
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTSIGUPD, abytBuffer(pintStart), intMsg_Length)
     
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTSIGUPD"
50  pusrDAS_Rec.sReport_Type = "SIG"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTSIGUPD.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
80  pbytChan = LBound(pusrMTSIGUPD.uSig.uSigData.uChannelData)

90  With pusrMTSIGUPD.uSig.uSigData
        '
        '+v1.8.11 TE
        pusrDAS_Rec.lTarget_ID = agSwapBytes%(.iEchid)
        '-v1.8.11 TE
        '
100     pintSigID = agSwapBytes%(.iSignum)
110     pdblFreq = modFreqConv(agSwapWords&(.lFreq))
120     plngEmitter = modLMBSigToLng(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
130     pstrEmitter = modLmbSigToString(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
140     If ((pintSigID >= LBound(gdFreq)) And (pintSigID <= UBound(gdFreq))) Then
150         gdFreq(pintSigID) = pdblFreq
160         glEmitter(pintSigID) = plngEmitter
170         gsEmitter(pintSigID) = pstrEmitter
        End If
180     pbytChan = LBound(.uChannelData)
190     If (pintSigID = -1) Then
200         pusrDAS_Rec.sSignal = "NEW_SIG"
        Else
210         pusrDAS_Rec.sSignal = "EXISTING_SIG"
        End If
220     pusrDAS_Rec.lSignal_ID = pintSigID
230     pusrDAS_Rec.dFrequency = pdblFreq
240     pusrDAS_Rec.lStatus = .uChannelData(pbytChan).bytChannelActive
250     pusrDAS_Rec.sEmitter = modLmbSigToString(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
260     pusrDAS_Rec.lEmitter_ID = plngEmitter
270     pusrDAS_Rec.lTag = .uChannelData(pbytChan).uSignalType.bytVariant
275     pusrDAS_Rec.sSupplemental = "Usage : " & .uChannelData(pbytChan).bytSigUsage
280     pusrDAS_Rec.sAllegiance = modAllegToString(.bytAllegiance, pusrDAS_Rec.lIFF)
290     pusrDAS_Rec.lFlag = .bytRespOpr
300     If (pusrDAS_Rec.lFlag = 255) Then
310         pusrDAS_Rec.lFlag = -1
        End If

    End With
320 Call Add_Data_Record(MTSIGUPDID, pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsigupd"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtsigupd", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSIGUPD message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSIGUPD", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSIGUPD)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtsigupd3_0
' AUTHOR:   Shaun Vogel
' PURPOSE:  Translate the MTSIGUPD message to DAS data structure for Blk35
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtsigupd3_0(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTSIGUPD As Mtsigupd3_0
    Dim i As Integer
    Dim pbytChan As Byte
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintSigID As Integer
    Dim pdblFreq As Double
    Dim plngEmitter As Long
    Dim pstrEmitter As String
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
   'If (dblTime >= (4882890 Mod 86400)) Then
   '   Debug.Print
   'End If
   
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTSIGUPD, abytBuffer(pintStart), intMsg_Length)
     
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTSIGUPD"
50  pusrDAS_Rec.sReport_Type = "SIG"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTSIGUPD.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
80  pbytChan = LBound(pusrMTSIGUPD.uSig.uSigData.uChannelData)
    
90  With pusrMTSIGUPD.uSig.uSigData
        '
        '+v1.8.11 TE
        pusrDAS_Rec.lTarget_ID = agSwapBytes%(.iEchid)
        '-v1.8.11 TE
        '

100     pintSigID = agSwapBytes%(.iSignum)
110     pdblFreq = modFreqConv(agSwapWords&(.lFreq))
120     plngEmitter = modLMBSigToLng(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
130     pstrEmitter = modLmbSigToString(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
140     If ((pintSigID >= LBound(gdFreq)) And (pintSigID <= UBound(gdFreq))) Then
150         gdFreq(pintSigID) = pdblFreq
160         glEmitter(pintSigID) = plngEmitter
170         gsEmitter(pintSigID) = pstrEmitter
        End If
180     pbytChan = LBound(.uChannelData)
190     If (pintSigID = -1) Then
200         pusrDAS_Rec.sSignal = "NEW_SIG"
        Else
210         pusrDAS_Rec.sSignal = "EXISTING_SIG"
        End If
220     pusrDAS_Rec.lSignal_ID = pintSigID
230     pusrDAS_Rec.dFrequency = pdblFreq
240     pusrDAS_Rec.lStatus = .uChannelData(pbytChan).bytChannelActive
250     pusrDAS_Rec.sEmitter = modLmbSigToString(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
260     pusrDAS_Rec.lEmitter_ID = plngEmitter
270     pusrDAS_Rec.lTag = .uChannelData(pbytChan).uSignalType.bytVariant
275     pusrDAS_Rec.sSupplemental = "Usage : " & .uChannelData(pbytChan).bytSigUsage
280     pusrDAS_Rec.sAllegiance = modAllegToString(.bytAllegiance, pusrDAS_Rec.lIFF)
290     pusrDAS_Rec.lFlag = .bytRespOpr
300     If (pusrDAS_Rec.lFlag = 255) Then
310         pusrDAS_Rec.lFlag = -1
        End If

    End With
320 Call Add_Data_Record(MTSIGUPDID, pusrDAS_Rec)
    'Call Process_MTSIGUPD(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsigupd3_0"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtsigupd", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSIGUPD message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSIGUPD", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSIGUPD)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthblobupd
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBLOBUPD message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthblobupd(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBLOBUPD As Mthblobupd
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_LOBs As Integer
    Dim i As Integer
    Dim j As Integer
    Dim pusrTmp_LOBs As HbLobTblEntry
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim plngTemp_Lat As Long
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBLOBUPD)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBLOBUPD, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBLOBUPD"
60  pusrDAS_Rec.sReport_Type = "VEC"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBLOBUPD.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
90  With pusrMTHBLOBUPD.uHblobEntry(0).uHblobData
100     For i = 1 To 10
110         pusrDAS_Rec.sSupplemental = pusrDAS_Rec.sSupplemental & modHBIndexToLng(CInt(.uContrib(i).bytSignal), CInt(.uContrib(i).bytChannel)) & ","
        Next i
120     For i = 1 To 8
130         plngTemp_Lat = agSwapWords&(.uOwnShip(i).lLat)
140         If (plngTemp_Lat <> -1) Then
150             pusrDAS_Rec.dLatitude = modBam32ToDeg(plngTemp_Lat)
160             pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uOwnShip(i).lLon))
170             pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.iTrackBearing(i)))
180             pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(.uContrib(1).bytSignal), CInt(.uContrib(1).bytChannel))
190             pintHB_ID = modGetHBID(CLng(.uContrib(1).bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
200             pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(.uContrib(1).bytChannel))
210             pusrDAS_Rec.dFrequency = CDbl(pusrDAS_Rec.lEmitter_ID)
220             pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & .uContrib(1).bytChannel
230             Call Add_Data_Record(MTHBLOBUPDID, pusrDAS_Rec)
                'Call Process_MTHBLOBUPD(pusrDAS_Rec)
            End If
        Next i
    End With
   
240 pintNum_LOBs = agSwapBytes%(pusrMTHBLOBUPD.iRecsNMsg)
250 If pintNum_LOBs > 1 Then
260     For j = 1 To (pintNum_LOBs - 1)
270         Call CopyMemory(pusrTmp_LOBs, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_LOBs) * (j - 1))), LenB(pusrTmp_LOBs))
280         pusrDAS_Rec.sSupplemental = ""
290         For i = 1 To 10
300             pusrDAS_Rec.sSupplemental = pusrDAS_Rec.sSupplemental & modHBIndexToLng(CInt(pusrTmp_LOBs.uHblobData.uContrib(i).bytSignal), CInt(pusrTmp_LOBs.uHblobData.uContrib(i).bytChannel)) & ","
            Next i
310         For i = 1 To 8
320             plngTemp_Lat = agSwapWords&(pusrTmp_LOBs.uHblobData.uOwnShip(i).lLat)
330             If (plngTemp_Lat <> -1) Then
340                 pusrDAS_Rec.dLatitude = modBam32ToDeg(plngTemp_Lat)
350                 pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrTmp_LOBs.uHblobData.uOwnShip(i).lLon))
360                 pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(pusrTmp_LOBs.uHblobData.iTrackBearing(i)))
370                 pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrTmp_LOBs.uHblobData.uContrib(1).bytSignal), CInt(pusrTmp_LOBs.uHblobData.uContrib(1).bytChannel))
380                 pintHB_ID = modGetHBID(CLng(pusrTmp_LOBs.uHblobData.uContrib(1).bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
390                 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrTmp_LOBs.uHblobData.uContrib(1).bytChannel))
400                 pusrDAS_Rec.dFrequency = CDbl(pusrDAS_Rec.lEmitter_ID)
410                 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrTmp_LOBs.uHblobData.uContrib(1).bytChannel
420                 Call Add_Data_Record(MTHBLOBUPDID, pusrDAS_Rec)
                    'Call Process_MTHBLOBUPD(pusrDAS_Rec)
                End If
            Next i
        Next j
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthblobupd"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthblobupd", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBLOBUPD message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBLOBUPD", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBLOBUPD)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtulddata
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTDEFPMA message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtulddata(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTULDDATA As Mtulddata
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pusrEldbEntry As EldbEntry
    Dim pusrLobTable As LobTableEntry
    Dim pusrLobList As LobListRecord
    Dim iNumLobs As Integer
    Dim uLoc As LatLon
    Dim lType As Long
    Dim lPmaBitMask As Long
    Dim lQualBitMask As Long
    Dim lLobBitMap As Long
    Dim pintSigID As Integer
    Dim uTmpPos As XyCopy
    Dim pusrVarRecHdr As VarRecHdr
    Dim pdblTempTime As Double

    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTULDDATA)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTULDDATA, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTULDDATA"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTULDDATA.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
    lType = agSwapWords&(pusrMTULDDATA.lType)
    
    If (lType = 24) Then ' DL_DF_ELDB_TABLE
    ElseIf (lType = 25) Then ' DL_DF_LOB_TABLE
    ElseIf (lType = 26) Then ' DL_DF_LOB_LIST
        Call CopyMemory(pusrVarRecHdr, abytBuffer(pintStart + pintStruct_Length), LenB(pusrVarRecHdr))
        Call CopyMemory(pusrLobList, abytBuffer(pintStart + pintStruct_Length + LenB(pusrVarRecHdr)), LenB(pusrLobList))
        lPmaBitMask = &H800000
        lQualBitMask = &H3F0000
        pusrDAS_Rec.sReport_Type = "VEC"
        iNumLobs = agSwapBytes%(pusrLobList.iNumberOfLobs)
        If ((iNumLobs >= LBound(pusrLobList.uLobs)) And (iNumLobs <= UBound(pusrLobList.uLobs))) Then
            For i = 1 To iNumLobs
                pintSigID = agSwapBytes%(pusrLobList.uLobs(i).iFixAssoc)
                If ((pintSigID >= LBound(gdFreq)) And (pintSigID <= UBound(gdFreq))) Then
                    pdblTempTime = modLong2Dbl(agSwapWords&(pusrLobList.uLobs(i).uTod.lUsecs))
                    pdblTempTime = pdblTempTime + (pusrLobList.uLobs(i).uTod.bytHighUsecs * 4294967296#)
                    pusrDAS_Rec.dReportTime = CDbl(pdblTempTime / 1000000)
                    pusrDAS_Rec.dFrequency = gdFreq(pintSigID)
                    pusrDAS_Rec.lEmitter_ID = glEmitter(pintSigID)
                    pusrDAS_Rec.sEmitter = gsEmitter(pintSigID)
                    pusrDAS_Rec.lSignal_ID = pintSigID
                    pusrDAS_Rec.sSignal = Str(pintSigID)
                    lLobBitMap = agSwapWords&(pusrLobList.uLobs(i).uFlags.lLobBitMap)
                    If (lLobBitMap And lQualBitMask) Then
                        pusrDAS_Rec.lFlag = 1
                    End If
                    If (lLobBitMap And lPmaBitMask) Then
                        pusrDAS_Rec.lStatus = 1
                    End If
                    Call CopyMemory(uTmpPos.dX, agSwapWords&(pusrLobList.uLobs(i).uPosition.dX), 4)
                    Call CopyMemory(uTmpPos.dY, agSwapWords&(pusrLobList.uLobs(i).uPosition.dY), 4)
                    uTmpPos.dX = uTmpPos.dX * libGeo.Meters_Per_NM
                    uTmpPos.dY = uTmpPos.dY * libGeo.Meters_Per_NM
                    uLoc = modXYToLatLon(uTmpPos)
                    pusrDAS_Rec.dLatitude = modBam32ToDeg(uLoc.lLat)
                    pusrDAS_Rec.dLongitude = modBam32ToDeg(uLoc.lLon)
                    pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(pusrLobList.uLobs(i).iTrueBearing))
                    Call Add_Data_Record(MTULDDATAID, pusrDAS_Rec)
                    'Call Process_MTULDDATA(pusrDAS_Rec)
                End If
            Next i
        End If
     End If
    
   
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtulddata"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtulddata", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTULDDATA message", vbAbortRetryIgnore Or vbExclamation, "Error Translating MTULDDATA")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtlobrslt
' AUTHOR:   Shaun Vogel
' PURPOSE:  Translate the MTLOBRSLT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtlobrslt(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTLOBRSLT As Mtlobrslt
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_LOBs As Integer
    Dim pintSignal_ID As Integer
    Dim i As Integer
    Dim pusrTmp_LOBs As LobData
    '
    On Error GoTo Hell
 
10  pintStruct_Length = LenB(pusrMTLOBRSLT)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTLOBRSLT, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTLOBRSLT"
60  pusrDAS_Rec.sReport_Type = "VEC"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTLOBRSLT.uMsgHdr.iMsgFrom)
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)

pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrMTLOBRSLT.lFreq))
'pusrDAS_Rec.dBandwidth = modFreqConv(agSwapWords&(pusrMTLOBRSLT.lBandwidth))
'pusrDAS_Rec.lStatus = pusrMTLOBRSLT.bytLobStatus


160 pintNum_LOBs = agSwapBytes%(pusrMTLOBRSLT.iNumRec)
170 If pintNum_LOBs >= 1 Then                           'S
180     With pusrMTLOBRSLT.uLobData(0).uLobs
190         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uAcLoc.lLat))
200         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uAcLoc.lLon))
210         pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.iTrueBearing))
220         pusrDAS_Rec.lStatus = .bytInOutPma
            '
            '+v1.6BB
230         pusrDAS_Rec.lFlag = .bytQualFactor
            '-v1.6
240         Call Add_Data_Record(MTLOBRSLTID, pusrDAS_Rec)
            'Call Process_MTLOBRSLT(pusrDAS_Rec)
        End With
250     If pintNum_LOBs > 1 Then
260         For i = 1 To (pintNum_LOBs - 1)
270             Call CopyMemory(pusrTmp_LOBs, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_LOBs) * (i - 1))), LenB(pusrTmp_LOBs))
280             pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrTmp_LOBs.uLobs.uAcLoc.lLat))
290             pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrTmp_LOBs.uLobs.uAcLoc.lLon))
300             pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(pusrTmp_LOBs.uLobs.iTrueBearing))
310             pusrDAS_Rec.lStatus = pusrTmp_LOBs.uLobs.bytInOutPma
                '
                '+v1.6BB
320             pusrDAS_Rec.lFlag = pusrTmp_LOBs.uLobs.bytQualFactor
                '-v1.6
330             Call Add_Data_Record(MTLOBRSLTID, pusrDAS_Rec)
                'Call Process_MTLOBRSLT(pusrDAS_Rec)
            Next i
        End If
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtlobrslt"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtlobrslt", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTLOBRSLT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTLOBRSLT")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtlobsetrslt
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTLOBSETRSLT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtlobsetrslt(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTLOBSETRSLT As Mtlobsetrslt
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_LOBs As Integer
    Dim pintSignal_ID As Integer
    Dim i As Integer
    Dim pusrTmp_LOBs As LobPacket
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTLOBSETRSLT)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTLOBSETRSLT, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTLOBSETRSLT"
60  pusrDAS_Rec.sReport_Type = "VEC"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTLOBSETRSLT.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
90  pintSignal_ID = agSwapBytes%(pusrMTLOBSETRSLT.iSigID)
100 pusrDAS_Rec.lSignal_ID = pintSignal_ID
110 pusrDAS_Rec.sSignal = Str(pintSignal_ID)
120 If ((pintSignal_ID >= LBound(gdFreq)) And (pintSignal_ID <= UBound(gdFreq))) Then
130     pusrDAS_Rec.dFrequency = gdFreq(pintSignal_ID)
140     pusrDAS_Rec.lEmitter_ID = glEmitter(pintSignal_ID)
150     pusrDAS_Rec.sEmitter = gsEmitter(pintSignal_ID)
    End If

160 pintNum_LOBs = agSwapBytes%(pusrMTLOBSETRSLT.iNumRec)
170 If pintNum_LOBs >= 1 Then                           'SCR 14
180     With pusrMTLOBSETRSLT.uLobs(0)
190         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uAcLoc.lLat))
200         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uAcLoc.lLon))
210         pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.iTrueBearing))
220         pusrDAS_Rec.lStatus = .bytInOutPma
            '
            '+v1.6BB
230         pusrDAS_Rec.lFlag = .bytQualFactor
            '-v1.6
240         Call Add_Data_Record(MTLOBSETRSLTID, pusrDAS_Rec)
            'Call Process_MTLOBSETRSLT(pusrDAS_Rec)
        End With
250     If pintNum_LOBs > 1 Then
260         For i = 1 To (pintNum_LOBs - 1)
270             Call CopyMemory(pusrTmp_LOBs, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_LOBs) * (i - 1))), LenB(pusrTmp_LOBs))
280             pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrTmp_LOBs.uAcLoc.lLat))
290             pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrTmp_LOBs.uAcLoc.lLon))
300             pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(pusrTmp_LOBs.iTrueBearing))
310             pusrDAS_Rec.lStatus = pusrTmp_LOBs.bytInOutPma
                '
                '+v1.6BB
320             pusrDAS_Rec.lFlag = pusrTmp_LOBs.bytQualFactor
                '-v1.6
330             Call Add_Data_Record(MTLOBSETRSLTID, pusrDAS_Rec)
                'Call Process_MTLOBSETRSLT(pusrDAS_Rec)
           Next i
        End If
    End If                                          'SCR 14
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtlobsetrslt"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtlobsetrslt", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTLOBSETRSLT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTLOBSETRSLT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTLOBSETRSLT)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtfixrslt
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTFIXRSLT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtfixrslt(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTFIXRSLT As Mtfixrslt
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim i As Integer
    Dim pintSigID As Integer
    Dim pbytChan As Byte
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTFIXRSLT, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTFIXRSLT"
50  pusrDAS_Rec.sReport_Type = "GEO"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTFIXRSLT.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
    
80  For i = 1 To pusrMTFIXRSLT.bytNumFixes
90      With pusrMTFIXRSLT.uFixinfo(i)
100         pintSigID = agSwapBytes%(.iSigID)
110         If ((pintSigID >= LBound(gdFreq)) And (pintSigID <= UBound(gdFreq))) Then
120             pusrDAS_Rec.dFrequency = gdFreq(pintSigID)
130             pusrDAS_Rec.lEmitter_ID = glEmitter(pintSigID)
140             pusrDAS_Rec.sEmitter = gsEmitter(pintSigID)
            Else
150             pbytChan = LBound(pusrMTFIXRSLT.uFixinfo(i).uChannelData)
160             pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(.lFreq))
170             pusrDAS_Rec.lEmitter_ID = modLMBSigToLng(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
180             pusrDAS_Rec.sEmitter = modLmbSigToString(.bytRadioType, .uChannelData(pbytChan).uSignalType.bytClass)
            End If
190         pusrDAS_Rec.lStatus = .bytFixType
200         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uFixloc.lLat))
210         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uFixloc.lLon))
220         pusrDAS_Rec.lSignal_ID = pintSigID
230         pusrDAS_Rec.sSignal = Str(pintSigID)
240         If (.bytFixType = 0) Then
250             pusrDAS_Rec.sSupplemental = basCCAT.GetAlias("Operator", "OPERATOR" & agSwapBytes%(pusrMTFIXRSLT.iRequestorID), Str(agSwapBytes%(pusrMTFIXRSLT.iRequestorID)))
            End If
260         Call Add_Data_Record(MTFIXRSLTID, pusrDAS_Rec)
            'Call Process_MTFIXRSLT(pusrDAS_Rec)
        End With

    Next i
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtfixrslt"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtfixrslt", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTFIXRSLT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTFIXRSLT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTFIXRSLT)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbsigupd
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBSIGUPD message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbsigupd(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBSIGUPD As Mthbsigupd
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim plngTemp_Lat As Long
    Dim pintTemp_AID As Integer
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTHBSIGUPD, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTHBSIGUPD"
50  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBSIGUPD.uMsgHdr.iMsgFrom)
    '+v1.6TE
60  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
70  plngTemp_Lat = agSwapWords&(pusrMTHBSIGUPD.uHbSigRec.uLocation.lLat)
80  If (plngTemp_Lat = -1) Then
90      pusrDAS_Rec.sReport_Type = "SIG"
100     pusrDAS_Rec.dLatitude = 0
110     pusrDAS_Rec.dLongitude = 0
    Else
120     pusrDAS_Rec.sReport_Type = "GEO"
130     pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrMTHBSIGUPD.uHbSigRec.uLocation.lLat))
140     pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrMTHBSIGUPD.uHbSigRec.uLocation.lLon))
    End If
150 pintTemp_AID = agSwapBytes%(pusrMTHBSIGUPD.uHbSigRec.iHbSsIndex)
160 pusrDAS_Rec.lSignal_ID = pintTemp_AID
170 pintHB_ID = modGetHBID(CLng(pintTemp_AID), HBAID, pintHB_SigType, pintHB_Chan)
180 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, pintHB_Chan)
190 pusrDAS_Rec.dFrequency = CDbl(pusrDAS_Rec.lEmitter_ID)
200 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN")
210 If (pintHB_Chan <> 0) Then
220     pusrDAS_Rec.sEmitter = pusrDAS_Rec.sEmitter & pintHB_Chan
    End If
    
230 pusrDAS_Rec.sAllegiance = modAllegToString(pusrMTHBSIGUPD.uHbSigRec.bytAlleg, pusrDAS_Rec.lIFF)
240 Call Add_Data_Record(MTHBSIGUPDID, pusrDAS_Rec)
    'Call Process_MTHBSIGUPD(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbsigupd"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbsigupd", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBSIGUPD message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBSIGUPD", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBSIGUPD)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbgsrep
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBGSREP message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbgsrep(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBGSREP As Mthbgsrep
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim i As Integer
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    Dim pintNum_Ground_Sites As Integer
    Dim pusrTmp_GSRec As GroundSite
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBGSREP)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBGSREP, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBGSREP"
60  pusrDAS_Rec.sReport_Type = "GEO"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBGSREP.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE

90  pintNum_Ground_Sites = agSwapBytes%(pusrMTHBGSREP.iNumGs)
   
100 pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrMTHBGSREP.uGsrRec(0).uLocation.lLat))
110 pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrMTHBGSREP.uGsrRec(0).uLocation.lLon))
120 pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrMTHBGSREP.uGsrRec(0).bytSignal), CInt(pusrMTHBGSREP.uGsrRec(0).bytChannel))
130 pintHB_ID = modGetHBID(CLng(pusrMTHBGSREP.uGsrRec(0).bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
140 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrMTHBGSREP.uGsrRec(0).bytChannel))
150 pusrDAS_Rec.dFrequency = CDbl(pusrDAS_Rec.lSignal_ID)
160 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrMTHBGSREP.uGsrRec(0).bytChannel
   
170 pusrDAS_Rec.lStatus = pusrMTHBGSREP.uGsrRec(0).bytMethod
180 Call Add_Data_Record(MTHBGSREPID, pusrDAS_Rec)
    'Call Process_MTHBGSREP(pusrDAS_Rec)
      
190 If pintNum_Ground_Sites > 1 Then
200     For i = 1 To (pintNum_Ground_Sites - 1)
210         Call CopyMemory(pusrTmp_GSRec, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_GSRec) * (i - 1))), LenB(pusrTmp_GSRec))
220         pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrTmp_GSRec.uLocation.lLat))
230         pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrTmp_GSRec.uLocation.lLon))
240         pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrTmp_GSRec.bytSignal), CInt(pusrTmp_GSRec.bytChannel))
250         pintHB_ID = modGetHBID(CLng(pusrTmp_GSRec.bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
260         pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrTmp_GSRec.bytChannel))
270         pusrDAS_Rec.dFrequency = CDbl(pusrDAS_Rec.lSignal_ID)
280         pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrTmp_GSRec.bytChannel
290         pusrDAS_Rec.lStatus = pusrTmp_GSRec.bytMethod
300         Call Add_Data_Record(MTHBGSREPID, pusrDAS_Rec)
            'Call Process_MTHBGSREP(pusrDAS_Rec)
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbgsrep"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbgsrep", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBGSREP message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBGSREP", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBGSREP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbsemistat
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBSEMISTAT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbsemistat(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBSEMISTAT As Mthbsemistat
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBSEMISTAT)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBSEMISTAT, abytBuffer(pintStart), intMsg_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBSEMISTAT"
60  pusrDAS_Rec.sReport_Type = "GEO"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBSEMISTAT.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
   
90  Call Add_Data_Record(MTHBSEMISTATID, pusrDAS_Rec)
    'Call Process_MTHBSEMISTAT(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbsemistat"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbsemistat", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBSEMISTAT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBSEMISTAT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBSEMISTAT)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbselask
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBSELASK message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbselask(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBSELASK As Mthbselask
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTHBSELJAM, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTHBSELASK"
50  pusrDAS_Rec.sReport_Type = "EVT"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBSELASK.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
80  pusrDAS_Rec.lStatus = pusrMTHBSELASK.bytOnOff
90  pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrMTHBSELASK.bytSigType), CInt(gbytNOT_DEFINED))
100 pintHB_ID = modGetHBID(CLng(pusrMTHBSELASK.bytSigType), HBSigtype, pintHB_SigType, pintHB_Chan)
110 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN")
120 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(gbytNOT_DEFINED))
130 pusrDAS_Rec.dFrequency = CDbl(pusrDAS_Rec.lEmitter_ID)

    
140 pusrDAS_Rec.sSupplemental = pusrDAS_Rec.sEmitter & "," & pusrMTHBSELASK.bytOnOff

150 Call Add_Data_Record(MTHBSELASKID, pusrDAS_Rec)
    'Call Process_MTHBSELASK(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbselask"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbselask", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBSELASK message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBSELASK", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBSELASK)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtjamstat
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTJAMSTAT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtjamstat(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTJAMSTAT As Mtjamstat
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintSignal_ID As Integer
    Dim pintNum_Sigid As Integer
    Dim pusrTmp_Sig_Stat As SigPacket
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTJAMSTAT)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTJAMSTAT, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTJAMSTAT"
60  pusrDAS_Rec.sReport_Type = "EVT"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTJAMSTAT.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
90  pintSignal_ID = agSwapBytes%(pusrMTJAMSTAT.uSigStatus(0).iSigID)
100 pusrDAS_Rec.lSignal_ID = pintSignal_ID
110 pusrDAS_Rec.sSignal = Str(pintSignal_ID)
120 pusrDAS_Rec.lStatus = pusrMTJAMSTAT.uSigStatus(0).bytJamStatus
130 pusrDAS_Rec.sSupplemental = pintSignal_ID & "," & pusrMTJAMSTAT.uSigStatus(0).bytJamStatus & " , " & basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & pusrMTJAMSTAT.uSigStatus(0).bytJamStatus, "UNKNOWN")
    
140 If ((pintSignal_ID >= LBound(gdFreq)) And (pintSignal_ID <= UBound(gdFreq))) Then
150     pusrDAS_Rec.dFrequency = gdFreq(pintSignal_ID)
160     pusrDAS_Rec.lEmitter_ID = glEmitter(pintSignal_ID)
170     pusrDAS_Rec.sEmitter = gsEmitter(pintSignal_ID)
    End If
180 Call Add_Data_Record(MTJAMSTATID, pusrDAS_Rec)
    'Call Process_MTJAMSTAT(pusrDAS_Rec)
   
190 pintNum_Sigid = agSwapBytes%(pusrMTJAMSTAT.iNumSigID)
200 If pintNum_Sigid > 1 Then
210     For i = 1 To (pintNum_Sigid - 1)
220         Call CopyMemory(pusrTmp_Sig_Stat, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Sig_Stat) * (i - 1))), LenB(pusrTmp_Sig_Stat))
230         pintSignal_ID = agSwapBytes%(pusrTmp_Sig_Stat.iSigID)
240         pusrDAS_Rec.lEmitter_ID = glEmitter(pintSignal_ID)
250         pusrDAS_Rec.sEmitter = gsEmitter(pintSignal_ID)
260         pusrDAS_Rec.lSignal_ID = pintSignal_ID
270         pusrDAS_Rec.sSignal = Str(pintSignal_ID)
280         pusrDAS_Rec.lStatus = pusrTmp_Sig_Stat.bytJamStatus
290         pusrDAS_Rec.sSupplemental = pintSignal_ID & "," & pusrTmp_Sig_Stat.bytJamStatus & " , " & basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & pusrTmp_Sig_Stat.bytJamStatus, "UNKNOWN")
300         If ((pintSignal_ID >= LBound(gdFreq)) And (pintSignal_ID <= UBound(gdFreq))) Then
310             pusrDAS_Rec.dFrequency = gdFreq(pintSignal_ID)
320             pusrDAS_Rec.lEmitter_ID = glEmitter(pintSignal_ID)
330             pusrDAS_Rec.sEmitter = gsEmitter(pintSignal_ID)
            End If
340         Call Add_Data_Record(MTJAMSTATID, pusrDAS_Rec)
            'Call Process_MTJAMSTAT(pusrDAS_Rec)
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtjamstat"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtjamstat", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTJAMSTAT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTJAMSTAT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTJAMSTAT)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtrunmode
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTRUNMODE message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtrunmode(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTRUNMODE As Mtrunmode
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
    '
10  pintStruct_Length = LenB(pusrMTRUNMODE)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTRUNMODE, abytBuffer(pintStart), intMsg_Length)
40  pusrDAS_Rec.sReport_Type = "EVT"
    
50  pusrDAS_Rec.dReportTime = dblTime
60  pusrDAS_Rec.sMsg_Type = "MTRUNMODE"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTRUNMODE.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE

90  With pusrMTRUNMODE
100     pusrDAS_Rec.sSupplemental = "SYS@" & modRunmodeToString(agSwapBytes%(.iRunMode)) & "," & " LMB@" & modRunmodeToString(agSwapBytes%(.iLmbRunmode)) & "," & " HB@" & modRunmodeToString(agSwapBytes%(.iHbRunmode))
    End With

110 Call Add_Data_Record(MTRUNMODEID, pusrDAS_Rec)
    'modMTRUNMODE.Process_MTRUNMODE pusrDAS_Rec
    
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtrunmode"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtrunmode", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTRUNMODE message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTRUNMODE", App.HelpFile, basCCAT.IDH_TRANSLATE_MTRUNMODE)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtrunmode3_0
' AUTHOR:   Shaun Vogel
' PURPOSE:  Translate the MTRUNMODE message to DAS data structure for Blk35
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtrunmode3_0(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTRUNMODE As Mtrunmode3_0
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
    '+v1.10.11TE
    'modMTRUNMODE.Process_MTRUNMODE dblTime, intMsg_Length, abytBuffer
    '-v1.10.11TE
    '
10  pintStruct_Length = LenB(pusrMTRUNMODE)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTRUNMODE, abytBuffer(pintStart), intMsg_Length)
40  pusrDAS_Rec.sReport_Type = "EVT"
    
50  pusrDAS_Rec.dReportTime = dblTime
60  pusrDAS_Rec.sMsg_Type = "MTRUNMODE"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTRUNMODE.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE

90  With pusrMTRUNMODE
100     pusrDAS_Rec.sSupplemental = "SYS@" & modRunmodeToString(agSwapBytes%(.iRunMode)) & "," & " LMB@" & modRunmodeToString(agSwapBytes%(.iLmbRunmode)) & "," & " HB@" & modRunmodeToString(agSwapBytes%(.iHbRunmode)) & "," & " SPR@" & modRunmodeToString(agSwapBytes%(.iSprRunmode))
    End With

110 Call Add_Data_Record(MTRUNMODEID, pusrDAS_Rec)
    'Call Process_MTRUNMODE(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtrunmode3_0"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtrunmode", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTRUNMODE message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTRUNMODE", App.HelpFile, basCCAT.IDH_TRANSLATE_MTRUNMODE)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbseljam
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBSELJAM message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbseljam(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBSELJAM As Mthbseljam
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTHBSELJAM, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTHBSELJAM"
50  pusrDAS_Rec.sReport_Type = "EVT"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBSELJAM.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE

80  pusrDAS_Rec.lStatus = pusrMTHBSELJAM.bytOnOff
    
90  pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrMTHBSELJAM.bytSigType), CInt(gbytNOT_DEFINED))
100 pintHB_ID = modGetHBID(CLng(pusrMTHBSELJAM.bytSigType), HBSigtype, pintHB_SigType, pintHB_Chan)
110 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN")
120 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(gbytNOT_DEFINED))
130 pusrDAS_Rec.dFrequency = CDbl(pusrDAS_Rec.lEmitter_ID)

    
140 pusrDAS_Rec.sSupplemental = pusrDAS_Rec.sEmitter & "," & pusrMTHBSELJAM.bytOnOff

150 Call Add_Data_Record(MTHBSELJAMID, pusrDAS_Rec)
    'Call Process_MTHBSELJAM(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbseljam"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbseljam", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBSELJAM message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBSELJAM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBSELJAM)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbxmtrstat
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBXMTRSTAT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbxmtrstat(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBXMTRSTAT As Mthbxmtrstat
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Sigs As Integer
    Dim pusrTmp_Sig_Struct As SigMode
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBXMTRSTAT)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBXMTRSTAT, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBXMTRSTAT"
60  pusrDAS_Rec.sReport_Type = "EVT"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBXMTRSTAT.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
    '
90  pusrDAS_Rec.sSupplemental = modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iHboXmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iHb2XmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.uP34XmtrStatus.iBand3Xmtr)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.uP34XmtrStatus.iBand4aXmtr)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.uP34XmtrStatus.iBand4bXmtr))
        '
        ' DAP calls for Chan_Status, but that is within another embedded structure with
        ' multiple entries.
        ' Do we really need it?
    '
100 Call Add_Data_Record(MTHBXMTRSTATID, pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbxmtrstat"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbxmtrstat", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBXMTRSTAT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBXMTRSTAT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBXMTRSTAT)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbxmtrstat3_0
' AUTHOR:   Shaun Vogel
' PURPOSE:  Translate the MTHBXMTRSTAT message to DAS data structure for Blk35
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbxmtrstat3_0(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBXMTRSTAT As Mthbxmtrstat3_0
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Sigs As Integer
    Dim pusrTmp_Sig_Struct As SigMode
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBXMTRSTAT)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBXMTRSTAT, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBXMTRSTAT"
60  pusrDAS_Rec.sReport_Type = "EVT"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBXMTRSTAT.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
    '
90  pusrDAS_Rec.sSupplemental = modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iMbXmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iHb1XmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iHb2XmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iHb3XmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iSb1XmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iSpear1XmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iSpear2XmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iSpear3XmtrStatus)) & "," & _
        modXmtrstatToString(agSwapBytes%(pusrMTHBXMTRSTAT.iSpear4XmtrStatus))
        '
        ' DAP calls for Chan_Status, but that is within another embedded structure with
        ' multiple entries.
        ' Do we really need it?
    '
100 Call Add_Data_Record(MTHBXMTRSTATID, pusrDAS_Rec)
    'Call Process_MTHBXMTRSTAT(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbxmtrstat3_0"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbxmtrstat", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBXMTRSTAT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBXMTRSTAT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBXMTRSTAT)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtnavrep
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTNAVREP message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtnavrep(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTNAVREP As Mtnavrep
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    '
    '+v1.6TE
    Dim pblnUse_GPS As Boolean          ' If TRUE, uses GPS fields of NAVREP
    '-v1.6
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
    '+v1.6TE
    ' See if the user wants to use GPS values
10  pblnUse_GPS = (basCCAT.GetNumber("Miscellaneous Operations", "UseGPS", 0) = 1)
    '-v1.6
    '
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTNAVREP, abytBuffer(pintStart), intMsg_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTNAVREP"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTNAVREP.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
80  pusrDAS_Rec.sReport_Type = "TRK"
    '
    '+v1.6TE
90  If pblnUse_GPS Then
100     pusrDAS_Rec.dAltitude = CDbl(agSwapWords&(pusrMTNAVREP.uNavdata.lGpsAltitude)) / 3.28083333333333 ' Convert feet to meters
    Else
110     pusrDAS_Rec.dAltitude = agSwapWords&(pusrMTNAVREP.uNavdata.lPressureAlt)
    End If
    '-v1.6
    '
120 pusrDAS_Rec.dHeading = modBam16ToDeg(agSwapBytes%(pusrMTNAVREP.uNavdata.iHeading))
130 pusrDAS_Rec.dSpeed = agSwapBytes%(pusrMTNAVREP.uNavdata.iTrueGroundSpeed)
    '
    '+v1.6TE
140 If pblnUse_GPS Then
150     pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrMTNAVREP.uNavdata.uGpsLoc.lLat))
160     pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrMTNAVREP.uNavdata.uGpsLoc.lLon))
    Else
170     pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrMTNAVREP.uNavdata.uPosition.lLat))
180     pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrMTNAVREP.uNavdata.uPosition.lLon))
    End If
    '-v1.6
    '
    guNavPositionLat = pusrDAS_Rec.dLatitude
    guNavPositionLon = pusrDAS_Rec.dLongitude
    
190 pusrDAS_Rec.sAllegiance = modAllegToString(3, pusrDAS_Rec.lIFF)
200 pusrDAS_Rec.lStatus = pusrMTNAVREP.uNavdata.bytWow
    '
    '+v1.6TE
    ' Add roll and pitch to the supplemental data (may be useful for troubleshooting DF and ID)
210 pusrDAS_Rec.sSupplemental = pusrDAS_Rec.sSupplemental & "R:" & modBam16ToDeg(agSwapBytes%(pusrMTNAVREP.uNavdata.iRoll)) & ",P:" & modBam16ToDeg(agSwapBytes%(pusrMTNAVREP.uNavdata.iPitch))
    '-v1.6
    '
220 Call Add_Data_Record(MTNAVREPID, pusrDAS_Rec)
    'Call Process_MTNAVREP(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtnavrep"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtnavrep", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTNAVREP message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTNAVREP", App.HelpFile, basCCAT.IDH_TRANSLATE_MTNAVREP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtlhcorrelate
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTLHCORRELATE message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtlhcorrelate(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTLHCORRELATE As Mtlhcorrelate
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTLHCORRELATE, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTLHCORRELATE"
50  pusrDAS_Rec.sReport_Type = "EVT"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTLHCORRELATE.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
    'pusrMTLHCORRELATE.bytCorrType
    'pusrMTLHCORRELATE.bytObjectId
    'pusrMTLHCORRELATE.bytSubjectId
    'pusrMTLHCORRELATE.iObjectId
    'pusrMTLHCORRELATE.iSubjectId
80  Call Add_Data_Record(MTLHCORRELATEID, pusrDAS_Rec)
    'Call Process_MTLHCORRELATE(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtlhcorrelate"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtlhcorrelate", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTLHCORRELATE message", vbAbortRetryIgnore Or vbExclamation, "Error Translating MTLHCORRELATE")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtlhechmod
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTLHECHMOD message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtlhechmod(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTLHECHMOD As Mtlhechmod
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintTemp_SigID As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
    pintStart = LBound(abytBuffer)
    Call CopyMemory(pusrMTLHECHMOD, abytBuffer(pintStart), intMsg_Length)
    
10  pusrDAS_Rec.dReportTime = dblTime
20  pusrDAS_Rec.sMsg_Type = "MTLHECHMOD"
30  pusrDAS_Rec.sReport_Type = "GEO"
40  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTLHECHMOD.uMsgHdr.iMsgFrom)
    '+v1.6TE
50  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
60  pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrMTLHECHMOD.uEch.uLoc.lLat))
70  pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrMTLHECHMOD.uEch.uLoc.lLon))
80  pusrDAS_Rec.sAllegiance = modAllegToString(pusrMTLHECHMOD.uEch.bytAlleg, pusrDAS_Rec.lIFF)
    pintTemp_SigID = agSwapBytes%(pusrMTLHECHMOD.uEch.iLmbIndex)
    pusrDAS_Rec.lSignal_ID = pintTemp_SigID
    pusrDAS_Rec.lTag = agSwapBytes%(pusrMTLHECHMOD.iEchid)
    If ((pintTemp_SigID >= LBound(gdFreq)) And (pintTemp_SigID <= UBound(gdFreq))) Then
        pusrDAS_Rec.dFrequency = gdFreq(pintTemp_SigID)
        pusrDAS_Rec.lEmitter_ID = glEmitter(pintTemp_SigID)
        pusrDAS_Rec.sEmitter = gsEmitter(pintTemp_SigID)
    End If
90  pusrDAS_Rec.lStatus = pusrMTLHECHMOD.uEch.bytEchType
    '-v1.6
100 pusrDAS_Rec.lFlag = pusrMTLHECHMOD.uEch.bytLmbEchJamPri
110 Call Add_Data_Record(MTLHECHMODID, pusrDAS_Rec)
    'Call Process_MTLHECHMOD(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtlhechmod"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtlhechmod", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTLHECHMOD message", vbAbortRetryIgnore Or vbExclamation, "Error Translating MTLHECHMOD")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtlhtrackupd
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTLHTRACKUPD message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtlhtrackupd(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTLHTRACKUPD As Mtlhtrackupd
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintNum_Tracks As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTLHTRACKUPD, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTLHTRACKUPD"
50  pusrDAS_Rec.sReport_Type = "TRK"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTLHTRACKUPD.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
80  pintNum_Tracks = agSwapBytes%(pusrMTLHTRACKUPD.iNumTracks)
90  For i = 1 To pintNum_Tracks
100     pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(pusrMTLHTRACKUPD.uTrack(i).uBestloc.lLat))
110     pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(pusrMTLHTRACKUPD.uTrack(i).uBestloc.lLon))
120     With pusrMTLHTRACKUPD.uTrack(i).uTrack
130         pusrDAS_Rec.sAllegiance = modAllegToString(.bytAllegiance, pusrDAS_Rec.lIFF)
140         pusrDAS_Rec.dAltitude = agSwapBytes%(.iAltitude)
150         pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.iTrackBearing))
160         pusrDAS_Rec.dSpeed = CDbl(agSwapBytes%(.iTrackSpeed))
        End With
170     Call Add_Data_Record(MTLHTRACKUPDID, pusrDAS_Rec)
        'Call Process_MTLHTRACKCUPD(pusrDAS_Rec)
    Next i
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtlhtrackupd"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtlhtrackupd", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTLHTRACKUPD message", vbAbortRetryIgnore Or vbExclamation, "Error Translating MTLHTRACKUPD")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtlhtrackrep
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTLHTRACKREP message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtlhtrackrep(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTLHTRACKREP As Mtlhtrackrep
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Tracks As Integer
    Dim pintNum_Cont As Integer
    Dim pintRead_Pointer As Integer
    Dim i As Integer, j As Integer
    Dim pintTrack_Length As Integer
    Dim pintContrib_Length As Integer
    Dim pusrTmp_Track As HBTrackData
    Dim pusrTmp_Cont As Contrib
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer

    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintContrib_Length = LenB(pusrTmp_Cont)
20  pintTrack_Length = LenB(pusrTmp_Track) - pintContrib_Length
30  pintStruct_Length = LenB(pusrMTLHTRACKREP) - LenB(pusrTmp_Track)
40  pintStart = LBound(abytBuffer)
50  pintRead_Pointer = pintStart + pintStruct_Length
60  Call CopyMemory(pusrMTLHTRACKREP, abytBuffer(pintStart), pintStruct_Length)
    
70  pusrDAS_Rec.dReportTime = dblTime
80  pusrDAS_Rec.sMsg_Type = "MTLHTRACKREP"
90  pusrDAS_Rec.sReport_Type = "TRK"
100 pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTLHTRACKREP.uMsgHdr.iMsgFrom)
    '+v1.6TE
110 pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
120 pintNum_Tracks = agSwapBytes%(pusrMTLHTRACKREP.iMsgSize)
    
130 If (pintNum_Tracks >= 1) Then
140     For i = 1 To pintNum_Tracks
150         Call CopyMemory(pusrTmp_Track, abytBuffer(pintRead_Pointer), pintTrack_Length)
160         pintRead_Pointer = pintRead_Pointer + pintTrack_Length
170         With pusrTmp_Track
180             If (agSwapWords&(.uHbsTrackLoc.lLat) = -1) Then
190                 pusrDAS_Rec.dLatitude = 0
                Else
200                 pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uHbsTrackLoc.lLat))
                End If
210             pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uHbsTrackLoc.lLon))
220             pusrDAS_Rec.sAllegiance = modAllegToString(.bytAllegiance, pusrDAS_Rec.lIFF)
230             pusrDAS_Rec.dAltitude = agSwapBytes%(.iAltitude)
240             pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.iTrackBearing))
250             pusrDAS_Rec.dSpeed = agSwapBytes%(.iTrackSpeed)
260             pintNum_Cont = agSwapBytes%(.iNumContributors)
                ' access data like .uContributor.bytAid for contrib 0
270             If (pintNum_Cont >= 1) Then
280                 For j = 1 To pintNum_Cont
290                     Call CopyMemory(pusrTmp_Cont, abytBuffer(pintRead_Pointer), pintContrib_Length)
300                     pintRead_Pointer = pintRead_Pointer + pintContrib_Length
310                     If (.uContributor(0).bytHbsTrackMeth = 1) Then
320                         pusrDAS_Rec.sReport_Type = "VEC"
                        End If
                        
                        ' access data like pusrTmp_Cont.bytAid
                    Next j
                End If
            End With
330         Call Add_Data_Record(MTLHTRACKREPID, pusrDAS_Rec)
       Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtlhtrackrep"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtlhtrackrep", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTLHTRACKREP message", vbAbortRetryIgnore Or vbExclamation, "Error Translating MTLHTRACKREP")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtlhtrackrep3_0
' AUTHOR:   Shaun Vogel
' PURPOSE:  Translate the MTLHTRACKREP message to DAS data structure
'           Translate HB Lobs
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtlhtrackrep3_0(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTLHTRACKREP As Mtlhtrackrep
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Tracks As Integer
    Dim pintNum_Cont As Integer
    Dim pintRead_Pointer As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim pintTrack_Length As Integer
    Dim pintNum_Contrib As Integer
    Dim pintContrib_Length As Integer
    Dim pusrTmp_Track As HBTrackData
    Dim pusrTmp_Cont As CCMessageStruct.Contrib
    'added HB variables
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim plngTemp_Lat As Long
    Dim pintHB_Chan As Integer

    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintContrib_Length = LenB(pusrTmp_Cont)
20  pintTrack_Length = LenB(pusrTmp_Track) - pintContrib_Length
30  pintStruct_Length = LenB(pusrMTLHTRACKREP) - LenB(pusrTmp_Track)
40  pintStart = LBound(abytBuffer)
50  pintRead_Pointer = pintStart + pintStruct_Length
60  Call CopyMemory(pusrMTLHTRACKREP, abytBuffer(pintStart), pintStruct_Length)
    
70  pusrDAS_Rec.dReportTime = dblTime
80  pusrDAS_Rec.sMsg_Type = "MTLHTRACKREP"
100 pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTLHTRACKREP.uMsgHdr.iMsgFrom)
    '+v1.6TE
110 pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
120 pintNum_Tracks = agSwapBytes%(pusrMTLHTRACKREP.iMsgSize)
    
130 If (pintNum_Tracks >= 1) Then
140     For i = 1 To pintNum_Tracks
150         Call CopyMemory(pusrTmp_Track, abytBuffer(pintRead_Pointer), pintTrack_Length)
160         pintRead_Pointer = pintRead_Pointer + pintTrack_Length
170         With pusrTmp_Track
                If (.bytTrackClass <> 2) Then
                    pusrDAS_Rec.sReport_Type = "TRK"
180                 If (agSwapWords&(.uHbsTrackLoc.lLat) = -1) Then
190                     pusrDAS_Rec.dLatitude = 0
                    Else
200                     pusrDAS_Rec.dLatitude = modBam32ToDeg(agSwapWords&(.uHbsTrackLoc.lLat))
                    End If
210                 pusrDAS_Rec.dLongitude = modBam32ToDeg(agSwapWords&(.uHbsTrackLoc.lLon))
220                 pusrDAS_Rec.sAllegiance = modAllegToString(.bytAllegiance, pusrDAS_Rec.lIFF)
230                 pusrDAS_Rec.dAltitude = agSwapBytes%(.iAltitude)
240                 pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.iTrackBearing))
250                 pusrDAS_Rec.dSpeed = agSwapBytes%(.iTrackSpeed)
                    If .bytTrackClass = 1 Then
                        pusrDAS_Rec.sSupplemental = "HB_MOVE_PT_TRACK"
                    Else
                        'TrackClass = 3
                        pusrDAS_Rec.sSupplemental = "HB_BEARING_MRKR"
                    End If
                    pusrDAS_Rec.lTarget_ID = agSwapBytes%(.iHbsTrackId)
                    Call Add_Data_Record(MTLHTRACKREPID, pusrDAS_Rec)
                    'Call Process_MTLHTRACKREP(pusrDAS_Rec)
                    If (pintNum_Contrib > 0) Then
                        Call CopyMemory(pusrTmp_Cont, abytBuffer(pintRead_Pointer), (LenB(pusrTmp_Cont) * pintNum_Contrib))
                        pintRead_Pointer = pintRead_Pointer + (LenB(pusrTmp_Cont) * pintNum_Contrib)
                    
                    End If
                Else
                    pusrDAS_Rec.sReport_Type = "VEC"
                    pusrDAS_Rec.dLatitude = guNavPositionLat
                    pusrDAS_Rec.dLongitude = guNavPositionLon
                    pusrDAS_Rec.dBearing = modBam16ToDeg(agSwapBytes%(.iTrackBearing))
                    pusrDAS_Rec.sSupplemental = "HB_MOVE_BEARING_TRACK"     'TrackClass = 2
                        
                    pintNum_Contrib = agSwapBytes%(.iNumContributors)
                    For j = 0 To (pintNum_Contrib - 1)
                        Call CopyMemory(pusrTmp_Cont, abytBuffer(pintRead_Pointer), LenB(pusrTmp_Cont))
                        pintRead_Pointer = pintRead_Pointer + LenB(pusrTmp_Cont)
                        pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrTmp_Cont.bytSignal), CInt(pusrTmp_Cont.bytChannel))
                        pintHB_ID = modGetHBID(CLng(pusrTmp_Cont.bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
                        pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrTmp_Cont.bytChannel))
                        pusrDAS_Rec.dFrequency = CDbl(pusrDAS_Rec.lSignal_ID)
                        pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrTmp_Cont.bytChannel
                        pusrDAS_Rec.lTarget_ID = agSwapBytes%(.iHbsTrackId)
                        Call Add_Data_Record(MTLHTRACKREPID, pusrDAS_Rec)
                        'Call Process_MTLHTRACKREP(pusrDAS_Rec)
                    Next j
                End If
            End With
        Next i
    End If
    
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtlhtrackrep"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtlhtrackrep", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Debug.Print Err.Number
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTLHTRACKREP message", vbAbortRetryIgnore Or vbExclamation, "Error Translating MTLHTRACKREP")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMtsetacqsmode
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTSETACQSMODE message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtsetacqsmode(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTSETACQSMODE As Mtsetacqsmode
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTSETACQSMODE, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTSETACQSMODE"
50  pusrDAS_Rec.sReport_Type = "EVT"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTSETACQSMODE.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
80  pusrDAS_Rec.sSupplemental = Str(pusrMTSETACQSMODE.bytSubMode)
90  If (pusrDAS_Rec.sSupplemental = 3) Then
100   pusrDAS_Rec.sSupplemental = pusrDAS_Rec.sSupplemental & " SET ENV"
    End If
   
110 Call Add_Data_Record(MTSETACQSMODEID, pusrDAS_Rec)
    'Call Process_MTSETACQSMODE(pusrDAS_Rec)
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsetacqsmode"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtsetacqsmode", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSETACQSMODEID message", vbAbortRetryIgnore Or vbExclamation, "Error Translating MTSETACQSMODEID")
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'+v1.6BB
' ROUTINE:  modDasMTSSERROR
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTSSERROR message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtsserror(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTSSERROR As Mtsserror
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim i As Integer
    Dim tmpStr() As String
    '
    '
    On Error GoTo Hell
    '
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTSSERROR, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTSSERROR"
50  pusrDAS_Rec.sReport_Type = "EVT"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTSSERROR.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
80  For i = 1 To agSwapBytes%(pusrMTSSERROR.iNumberErrorReports)
        '
        '+v1.6TE
90      pusrDAS_Rec.lStatus = agSwapBytes%(pusrMTSSERROR.uErrorReport(i).iSeverity)
100     pusrDAS_Rec.lTag = agSwapBytes%(pusrMTSSERROR.uErrorReport(i).iCategory)
110     pusrDAS_Rec.lFlag = agSwapBytes%(pusrMTSSERROR.uErrorReport(i).iErrorCode)
        '-v1.6
        '
120     pusrDAS_Rec.sSupplemental = Left(StrConv(pusrMTSSERROR.uErrorReport(i).bytAmpdata, vbUnicode), 130)
        tmpStr = Split(pusrDAS_Rec.sSupplemental, "")
        pusrDAS_Rec.sSupplemental = tmpStr(0)
        If Not pusrDAS_Rec.sSupplemental = "" Then
            If pusrDAS_Rec.sOrigin = "HIGHBAND" Then
                tmpStr = Split(pusrDAS_Rec.sSupplemental, " ")
                If tmpStr(UBound(tmpStr)) = "MHz" Then
                    pusrDAS_Rec.dFrequency = tmpStr(UBound(tmpStr) - 1)
                End If
            End If
        End If
130     Call Add_Data_Record(MTSSERRORID, pusrDAS_Rec)
        'Call Process_MTSSERROR(pusrDAS_Rec)
    Next i
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMTSSERROR"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMTSSERROR", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSSERROR message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSSERROR", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSSERROR)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'+v1.6BB
' ROUTINE:  modDasMtplanarea
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTPLANAREA message to storea
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtplanarea(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMtplanarea As Mtplanarea
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim i As Integer
    '
    '
    On Error GoTo Hell
    '
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMtplanarea, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTPLANAREA"
50  pusrDAS_Rec.sReport_Type = "GEO"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMtplanarea.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
80  guPlanAreaCtr.lLat = agSwapWords&(pusrMtplanarea.uPlanLoc.lLat)
90  guPlanAreaCtr.lLon = agSwapWords&(pusrMtplanarea.uPlanLoc.lLon)
100 pusrDAS_Rec.dLatitude = modBam32ToDeg(guPlanAreaCtr.lLat)
110 pusrDAS_Rec.dLongitude = modBam32ToDeg(guPlanAreaCtr.lLon)
120 Call Add_Data_Record(MTPLANAREAID, pusrDAS_Rec)
    'Call Process_MTPLANAREA(pusrDAS_Rec)
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtplanarea"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTPLANAREA message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSSERROR", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSSERROR)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub '

'+v1.6BB
' ROUTINE:  modDasMtanareq
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTanareq message
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtanareq(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMtanareq As Mtanareq
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim i As Integer
    '
    '
    On Error GoTo Hell
    '
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMtanareq, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTANAREQ"
50  pusrDAS_Rec.sReport_Type = "EVT"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMtanareq.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    pusrDAS_Rec.lFlag = pusrDAS_Rec.lOrigin_ID
    '-v1.6TE
80  pusrDAS_Rec.dFrequency = modFreqConv(agSwapWords&(pusrMtanareq.lFreq))
120 Call Add_Data_Record(MTANAREQID, pusrDAS_Rec)
    'Call Process_MTANAREQ(pusrDAS_Rec)
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtanareq"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTANAREQ message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTANAREQ", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSSERROR)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub '

'+v1.6BB
' ROUTINE:  modDasMtdfflgs
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTDFFLGS message
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtdfflgs(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMtdfflgs As Mtdfflgs
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim i As Integer
    '
    '
    On Error GoTo Hell
    '
    '
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMtdfflgs, abytBuffer(pintStart), intMsg_Length)
    
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTDFFLGS"
50  pusrDAS_Rec.sReport_Type = "EVT"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMtdfflgs.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    pusrDAS_Rec.lFlag = pusrDAS_Rec.lOrigin_ID
    '-v1.6TE
80  pusrDAS_Rec.lStatus = pusrMtdfflgs.bytLobIntersectPma
    pusrDAS_Rec.lTag = pusrMtdfflgs.bytGoodQuality
120 Call Add_Data_Record(MTDFFLGSID, pusrDAS_Rec)
    'Call Process_MTDFFLGS(pusrDAS_Rec)

    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtdfflgs"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTDFFLGS message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTANAREQ", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSSERROR)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub '
'
'
' ROUTINE:  modDasMtrrslt
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTRRSLT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtrrslt(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTRRSLT As Mtrrslt
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Sig As Integer
    Dim i As Integer                                                                'SCR 14
    Dim pusrTimeRslt As TimeResult
    Dim pintTemp_SigID As Integer
   
    
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTRRSLT)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTRRSLT, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTRRSLT"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTRRSLT.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6

    pusrDAS_Rec.lStatus = agSwapBytes%(pusrMTRRSLT.iStatus)
    pusrDAS_Rec.sSupplemental = basCCAT.GetAlias("RRSLT", "RRSLT" & pusrDAS_Rec.lStatus, "UNKNOWN")
    pusrDAS_Rec.dLatitude = guNavPositionLat
    pusrDAS_Rec.dLongitude = guNavPositionLon

90  With pusrMTRRSLT.uTimeResult(0)
        pintTemp_SigID = agSwapBytes%(.iSignum)
        pusrDAS_Rec.lSignal_ID = pintTemp_SigID
        If ((pintTemp_SigID >= LBound(gdFreq)) And (pintTemp_SigID <= UBound(gdFreq))) Then
            pusrDAS_Rec.dFrequency = gdFreq(pintTemp_SigID)
            pusrDAS_Rec.lEmitter_ID = glEmitter(pintTemp_SigID)
            pusrDAS_Rec.sEmitter = gsEmitter(pintTemp_SigID)
        End If
        Call Add_Data_Record(MTRRSLTID, pusrDAS_Rec)
        'Call Process_MTRRSLT(pusrDAS_Rec)
   End With
   
    pintNum_Sig = agSwapBytes%(pusrMTRRSLT.iNumberOfSignals)
    If pintNum_Sig > 1 Then
        For i = 1 To (pintNum_Sig - 1)
            Call CopyMemory(pusrTimeRslt, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTimeRslt) * (i - 1))), LenB(pusrTimeRslt))
            pintTemp_SigID = agSwapBytes%(pusrTimeRslt.iSignum)
            pusrDAS_Rec.lSignal_ID = pintTemp_SigID
            If ((pintTemp_SigID >= LBound(gdFreq)) And (pintTemp_SigID <= UBound(gdFreq))) Then
                pusrDAS_Rec.dFrequency = gdFreq(pintTemp_SigID)
                pusrDAS_Rec.lEmitter_ID = glEmitter(pintTemp_SigID)
                pusrDAS_Rec.sEmitter = gsEmitter(pintTemp_SigID)
            End If
            Call Add_Data_Record(MTRRSLTID, pusrDAS_Rec)
            'Call Process_MTRRSLT(pusrDAS_Rec)
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtrrslt"
    '
    ' Process the error
    Select Case plngErr_Num
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating pusrMTRRSLT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTDFSDALARM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTDFSDALARM)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub
'
'
' ROUTINE:  modDasMthbassrep
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBASSREP message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbassrep(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBASSREP As Mthbassrep
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Ass As Integer
    Dim i As Integer
    Dim pusrTmp_Ass As AssocData
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBASSREP)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBASSREP, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBASSREP"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBASSREP.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrMTHBASSREP.uAssoc(0).bytSignal), CInt(pusrMTHBASSREP.uAssoc(0).bytChannel))
100 pintHB_ID = modGetHBID(CLng(pusrMTHBASSREP.uAssoc(0).bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
110 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrMTHBASSREP.uAssoc(0).bytChannel))
120 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrMTHBASSREP.uAssoc(0).bytChannel
130 pusrDAS_Rec.lStatus = pusrMTHBASSREP.uAssoc(0).bytAssocChan
160 Call Add_Data_Record(MTHBASSREPID, pusrDAS_Rec)
    'Call Process_MTHBASSREP(pusrDAS_Rec)
170 pintNum_Ass = agSwapBytes%(pusrMTHBASSREP.iNumAssoc)
180 If pintNum_Ass > 1 Then
190     For i = 1 To (pintNum_Ass - 1)
200         Call CopyMemory(pusrTmp_Ass, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Ass) * (i - 1))), LenB(pusrTmp_Ass))
210         pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrTmp_Ass.bytSignal), CInt(pusrTmp_Ass.bytChannel))
220         pintHB_ID = modGetHBID(CLng(pusrTmp_Ass.bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
230         pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrTmp_Ass.bytChannel))
240         pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrTmp_Ass.bytChannel
250         pusrDAS_Rec.lStatus = pusrTmp_Ass.bytAssocChan
280         Call Add_Data_Record(MTHBASSREPID, pusrDAS_Rec)
            'Call Process_MTHBASSREP(pusrDAS_Rec)
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMthbassrep"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMthbassrep", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBASSREP message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBASSREP", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBACTREP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub

'
'
' ROUTINE:  modDasMthbgsupstat
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBGSUPSTAT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbgsupstat(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBGSUPSTAT As Mthbgsupstat
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Gs As Integer
    Dim i As Integer
    Dim pusrTmp_Gs As GroundSiteUp
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBGSUPSTAT)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBGSUPSTAT, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBGSUPSTAT"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBGSUPSTAT.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrMTHBGSUPSTAT.uGsUp(0).bytSignal), CInt(pusrMTHBGSUPSTAT.uGsUp(0).bytChannel))
100 pintHB_ID = modGetHBID(CLng(pusrMTHBGSUPSTAT.uGsUp(0).bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
110 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrMTHBGSUPSTAT.uGsUp(0).bytChannel))
120 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrMTHBGSUPSTAT.uGsUp(0).bytChannel
130 pusrDAS_Rec.lStatus = agSwapBytes%(pusrMTHBGSUPSTAT.uGsUp(0).iId)
160 Call Add_Data_Record(MTHBGSUPSTATID, pusrDAS_Rec)
    'Call Process_MTHBGSUPSTAT(pusrDAS_Rec)
170 pintNum_Gs = agSwapBytes%(pusrMTHBGSUPSTAT.iNumGs)
180 If pintNum_Gs > 1 Then
190     For i = 1 To (pintNum_Gs - 1)
200         Call CopyMemory(pusrTmp_Gs, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Gs) * (i - 1))), LenB(pusrTmp_Gs))
210         pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrTmp_Gs.bytSignal), CInt(pusrTmp_Gs.bytChannel))
220         pintHB_ID = modGetHBID(CLng(pusrTmp_Gs.bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
230         pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrTmp_Gs.bytChannel))
240         pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrTmp_Gs.bytChannel
250         pusrDAS_Rec.lStatus = agSwapBytes%(pusrTmp_Gs.iId)
280         Call Add_Data_Record(MTHBGSUPSTATID, pusrDAS_Rec)
            'Call Process_MTHBGSUPSTAT(pusrDAS_Rec)
     Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMTHBGSUPSTAT"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMTHBGSUPSTAT", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBGSUPSTAT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBGSUPSTAT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBACTREP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub

'
'
' ROUTINE:  modDasMTHBGSDEL
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTHBGSDEL message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMthbgsdel(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTHBGSDEL As Mthbgsdel
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pintNum_Gs As Integer
    Dim i As Integer
    Dim pusrTmp_Gs As GroundSiteDel
    Dim pintHB_ID As Integer
    Dim pintHB_SigType As Integer
    Dim pintHB_Chan As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTHBGSDEL)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTHBGSDEL, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTHBGSDEL"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTHBGSDEL.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6
90  pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrMTHBGSDEL.uGsDel(0).bytSignal), CInt(pusrMTHBGSDEL.uGsDel(0).bytChannel))
100 pintHB_ID = modGetHBID(CLng(pusrMTHBGSDEL.uGsDel(0).bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
110 pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrMTHBGSDEL.uGsDel(0).bytChannel))
120 pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrMTHBGSDEL.uGsDel(0).bytChannel
130 pusrDAS_Rec.lStatus = agSwapBytes%(pusrMTHBGSDEL.uGsDel(0).iId)
160 Call Add_Data_Record(MTHBGSDELID, pusrDAS_Rec)
    'Call Process_MTHBGSDEL(pusrDAS_Rec)
   
170 pintNum_Gs = agSwapBytes%(pusrMTHBGSDEL.iNumGs)
180 If pintNum_Gs > 1 Then
190     For i = 1 To (pintNum_Gs - 1)
200         Call CopyMemory(pusrTmp_Gs, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_Gs) * (i - 1))), LenB(pusrTmp_Gs))
210         pusrDAS_Rec.lSignal_ID = modHBIndexToLng(CInt(pusrTmp_Gs.bytSignal), CInt(pusrTmp_Gs.bytChannel))
220         pintHB_ID = modGetHBID(CLng(pusrTmp_Gs.bytSignal), HBSigtype, pintHB_SigType, pintHB_Chan)
230         pusrDAS_Rec.lEmitter_ID = modHBIndexToLng(pintHB_SigType, CInt(pusrTmp_Gs.bytChannel))
240         pusrDAS_Rec.sEmitter = basCCAT.GetAlias("HBSIG", "HBSIG" & pintHB_ID, "UNKNOWN") & pusrTmp_Gs.bytChannel
250         pusrDAS_Rec.lStatus = agSwapBytes%(pusrTmp_Gs.iId)
280         Call Add_Data_Record(MTHBGSDELID, pusrDAS_Rec)
            'Call Process_MTHBGSDEL(pusrDAS_Rec)
        Next i
    End If
    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMTHBGSDEL"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMTHBGSDEL", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBGSDEL message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBGSDEL", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBACTREP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub

'
'
' ROUTINE:  modDasMtsysconfig
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTSYSCONFIG message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtsysconfig(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTSYSCONFIG As Mtsysconfig
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTSYSCONFIG)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTSYSCONFIG, abytBuffer(pintStart), intMsg_Length)
40  pusrDAS_Rec.sReport_Type = "EVT"
    
50  pusrDAS_Rec.dReportTime = dblTime
60  pusrDAS_Rec.sMsg_Type = "MTSYSCONFIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTSYSCONFIG.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE

90  With pusrMTSYSCONFIG
     pusrDAS_Rec.sSupplemental = "CCUA@ " & modRunmodeToString(agSwapBytes%(.iCcuMode(1)))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "CCUB@ " & modRunmodeToString(agSwapBytes%(.iCcuMode(2)))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "DWS1@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(1).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "DWS2@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(2).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "DWS3@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(3).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "DWS4@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(4).iRunMode))
  Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "DWS5@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(5).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "DWS6@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(6).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "ACQ1@ " & modRunmodeToString(agSwapBytes%(.uAcqStatus(1).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "ACQ2@ " & modRunmodeToString(agSwapBytes%(.uAcqStatus(2).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "ANA@ " & modRunmodeToString(agSwapBytes%(.uAnaStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "EXC@ " & modRunmodeToString(agSwapBytes%(.uExcStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "DF@ " & modRunmodeToString(agSwapBytes%(.uDfStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "HBS@ " & modRunmodeToString(agSwapBytes%(.uHbsStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "DPS@ " & modRunmodeToString(agSwapBytes%(.uDpStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "TNA@ " & modRunmodeToString(agSwapBytes%(.uTsStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "TECH@ " & modRunmodeToString(agSwapBytes%(.uTasStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
    End With

    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsysconfig"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSYSCONFIG message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSYSCONFIG", App.HelpFile, basCCAT.IDH_TRANSLATE_MTRUNMODE)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub

'
'
' ROUTINE:  modDasMtsysconfig3_0
' AUTHOR:   Shaun Vogel
' PURPOSE:  Translate the MTSYSCONFIG message to DAS data structure for v3.0
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtsysconfig3_0(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTSYSCONFIG As Mtsysconfig3_0
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
10  pintStruct_Length = LenB(pusrMTSYSCONFIG)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTSYSCONFIG, abytBuffer(pintStart), intMsg_Length)
40  pusrDAS_Rec.sReport_Type = "EVT"
    
50  pusrDAS_Rec.dReportTime = dblTime
60  pusrDAS_Rec.sMsg_Type = "MTSYSCONFIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTSYSCONFIG.uMsgHdr.iMsgFrom)
    '+v1.6TE
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE

90  With pusrMTSYSCONFIG
     pusrDAS_Rec.sSupplemental = "CCUA@ " & modRunmodeToString(agSwapBytes%(.iCcuMode(1)))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "CCUB@ " & modRunmodeToString(agSwapBytes%(.iCcuMode(2)))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "DWS1@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(1).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "DWS2@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(2).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "DWS3@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(3).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "DWS4@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(4).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "DWS5@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(5).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "DWS6@ " & modRunmodeToString(agSwapBytes%(.uDwStatus(6).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "ACQ1@ " & modRunmodeToString(agSwapBytes%(.uAcqStatus(1).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "ACQ2@ " & modRunmodeToString(agSwapBytes%(.uAcqStatus(2).iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "ANA@ " & modRunmodeToString(agSwapBytes%(.uAnaStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "EXC@ " & modRunmodeToString(agSwapBytes%(.uExcStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "DF@ " & modRunmodeToString(agSwapBytes%(.uDfStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "HBS@ " & modRunmodeToString(agSwapBytes%(.uHbsStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "DPS@ " & modRunmodeToString(agSwapBytes%(.uDpStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "TNA@ " & modRunmodeToString(agSwapBytes%(.uTsStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "TECH@ " & modRunmodeToString(agSwapBytes%(.uTasStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "SPEAR@ " & modRunmodeToString(agSwapBytes%(.uSprStatus.iRunMode))
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
        
    'Add SRU status
    pusrDAS_Rec.sSupplemental = "Ta Eclipse lb1 Amp@ " & modSruStatToString(.bytTaEclipseLb1Amp)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
   pusrDAS_Rec.sSupplemental = "Ta Eclipse lb2 Amp@ " & modSruStatToString(.bytTaEclipseLb2Amp)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "Ta Eclipse mb1 Amp@ " & modSruStatToString(.bytTaEclipseMb1Amp)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "Ta Eclipse mb2 Amp@ " & modSruStatToString(.bytTaEclipseMb2Amp)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "Hb Sb1 Amp@ " & modSruStatToString(.bytHbSb1Amp)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "Spear Rf 1@ " & modSruStatToString(.bytSpearRf1)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "Spear Rf 2@ " & modSruStatToString(.bytSpearRf2)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
   pusrDAS_Rec.sSupplemental = "Spear Rf 3@ " & modSruStatToString(.bytSpearRf3)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "Spear Rf 4@ " & modSruStatToString(.bytSpearRf4)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
     pusrDAS_Rec.sSupplemental = "Ta Hbsu@ " & modSruStatToString(.bytTaHbsu)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = "Ta Blanking Unit@ " & modSruStatToString(.bytTaBlankingUnit)
 Call Add_Data_Record(MTSYSCONFIGID, pusrDAS_Rec)
 'Call Process_MTSYSCONFIG(pusrDAS_Rec)
    End With


    '
    '+v1.5 TE
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsysconfig3_0"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSYSCONFIG message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSYSCONFIG", App.HelpFile, basCCAT.IDH_TRANSLATE_MTRUNMODE)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub

'
'
' ROUTINE:  modDasMtenvstat3_0
' AUTHOR:   Brad Brown
' PURPOSE:  Translate the MTENVSTAT message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDasMtenvstat(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTENVSTAT As Mtenvstat
    Dim i As Integer
    Dim pbytChan As Byte
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    Dim pintStart As Integer
    Dim pintSigID As Integer
    Dim pdblFreq As Double
    Dim plngEmitter As Long
    Dim pstrEmitter As String
    Dim sSubmode As String
    
    '
    '+v1.5 TE
    On Error GoTo Hell
    '-v1.5
    '
   
10  pintStart = LBound(abytBuffer)
20  Call CopyMemory(pusrMTENVSTAT, abytBuffer(pintStart), intMsg_Length)
     
30  pusrDAS_Rec.dReportTime = dblTime
40  pusrDAS_Rec.sMsg_Type = "MTENVSTAT"
50  pusrDAS_Rec.sReport_Type = "SIG"
60  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTENVSTAT.uMsgHdr.iMsgFrom)
    '+v1.6TE
70  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
    '-v1.6TE
    
90  With pusrMTENVSTAT

    Select Case .bytCurrentSubmode
        Case 1:
            sSubmode = "MANUAL"
        Case 2:
            sSubmode = "REANALYSIS"
        Case 3:
            sSubmode = "SET_ENV"
    End Select
    
    pusrDAS_Rec.sSupplemental = sSubmode & " segment " & .bytSegment & ", time left : " & agSwapBytes%(pusrMTENVSTAT.iTimeLeft) & " START FREQ"

    pdblFreq = modFreqConv(agSwapWords&(.lReanalysisFreq1))
    pusrDAS_Rec.dFrequency = pdblFreq
    Call Add_Data_Record(MTENVSTATID, pusrDAS_Rec)
    'Call Process_MTENVSTAT(pusrDAS_Rec)
    pusrDAS_Rec.sSupplemental = sSubmode & " segment " & .bytSegment & ", time left : " & agSwapBytes%(pusrMTENVSTAT.iTimeLeft) & " END FREQ"
    pdblFreq = modFreqConv(agSwapWords&(.lReanalysisFreq2))
    pusrDAS_Rec.dFrequency = pdblFreq
    Call Add_Data_Record(MTENVSTATID, pusrDAS_Rec)
    'Call Process_MTENVSTAT(pusrDAS_Rec)
    End With
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsigupd"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        '+v1.6TE
        'Case vbObjectError + 911:
        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtsigupd", "User prematurely terminated the translation process", App.HelpFile, lHlp
        '-v1.6
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSIGUPD message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSIGUPD", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSIGUPD)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '-v1.5
    '
End Sub

' FUNCTION: modDegToBam32
' AUTHOR:   Brad Brown
' PURPOSE:  Convert degrees to BAM32 encoded angles
' INPUT:    varDegrees = The angle in degrees
' OUTPUT:   The BAM32 encoded angle
' NOTES:
Public Function modDegToBam32(varDegrees As Variant) As Long
    Dim pdblBAM_32_Const As Double

    pdblBAM_32_Const = CDbl(2 ^ 31) / 180#
    
    modDegToBam32 = CLng(varDegrees * pdblBAM_32_Const)

End Function
' FUNCTION: modBam32ToDeg
' AUTHOR:   Brad Brown
' PURPOSE:  Convert BAM32 encoded angles to degrees
' INPUT:    lngBAM32 = The BAM32 encoded angle
' OUTPUT:   The angle in degrees
' NOTES:
Public Function modBam32ToDeg(lngBAM32 As Long) As Double
    Dim pdblBAM_32_Const As Double

    pdblBAM_32_Const = CDbl(2 ^ 31) / 180#
    
    modBam32ToDeg = CDbl(lngBAM32) / pdblBAM_32_Const

End Function
'
' FUNCTION: modBam16ToDeg
' AUTHOR:   Brad Brown
' PURPOSE:  Convert BAM16 encoded angles to degrees
' INPUT:    lngBAM16 = The BAM16 encoded angle
' OUTPUT:   The angle in degrees
' NOTES:
Public Function modBam16ToDeg(lngBAM16 As Integer) As Double
    Dim pdblBAM_16_Const As Double
    
    pdblBAM_16_Const = CDbl(2 ^ 15) / 180#
    
    modBam16ToDeg = CDbl(lngBAM16) / pdblBAM_16_Const
    '
    ' Ensure positive angles
    If (modBam16ToDeg < 0) Then
        modBam16ToDeg = CDbl(modBam16ToDeg) + 360#
    End If

End Function
'
' FUNCTION: modFreqConv
' AUTHOR:   Brad Brown
' PURPOSE:  Convert frequency data to DAS standard (MHz)
' INPUT:    lngFreq = the frequency in Hz
' OUTPUT:   The frequency in MHz
' NOTES:
Public Function modFreqConv(lngFreq As Long) As Double

    modFreqConv = modLong2Dbl(lngFreq) / 1000000#

End Function
'
' FUNCTION: modLong2Dbl
' AUTHOR:   Brad Brown
' PURPOSE:  Convert  an unsigned long (in C) to Dbl
' INPUT:    lngLong
' OUTPUT:   The unsigned value as a double
' NOTES:
Public Function modLong2Dbl(lngLong As Long) As Double

    If (lngLong < 0) Then
        modLong2Dbl = CDbl(lngLong + 4294967296#)
    Else
        modLong2Dbl = CDbl(lngLong)
    End If
    
End Function '
' FUNCTION: modHbIndexToString
' AUTHOR:   Brad Brown
' PURPOSE:  Converted the coded high band signal index to a string value
' INPUT:    lngHB_Index = the coded high band signal index
' OUTPUT:   The signal identifier
' NOTES:
Public Function modHbIndexToString(lngHB_Index As Long) As String

    modHbIndexToString = basCCAT.GetAlias("HBSIG", "HBSIG" & lngHB_Index, "UNKNOWN")

End Function
'
' FUNCTION: modHBIndexToLng
' AUTHOR:   Brad Brown
' PURPOSE:  Convert high band signal and channel data to a coded high band signal index
' INPUT:    intSignal = the signal identifier
'           intChannel = the channel number
' OUTPUT:   The signal ID
' NOTES:
Public Function modHBIndexToLng(intSignal As Integer, intChannel As Integer) As Long

    modHBIndexToLng = CLng(intSignal) * 1000
    If (intChannel <> gbytNOT_DEFINED) Then
        modHBIndexToLng = modHBIndexToLng + intChannel
    End If

End Function
'
' FUNCTION: modLmbSigToString
' AUTHOR:   Brad Brown
' PURPOSE:  Convert coded radio and class data to text
' INPUT:    bytLmbRadio = The coded radio type
'           bytLmbClass = the coded radio class
' OUTPUT:   The radio description
' NOTES:
Public Function modLmbSigToString(bytLmbRadio As Byte, bytLmbClass As Byte) As String
    '
    '+v1.5
    ' Modified this function to match the numeric value creation (radio * 1000 + class)
    'modLmbSigToString = basCCAT.GetAlias("RCV List", "RCV" & Format(bytLmbRadio, "00") & Format(bytLmbClass, "00"), "UNKNOWN")
    modLmbSigToString = basCCAT.GetAlias("RCV List", "RCV" & Filtraw2Das.modLMBSigToLng(bytLmbRadio, bytLmbClass), "UNKNOWN")
    '-v1.5
    '
    If modLmbSigToString = "UNKNOWN" Then
    '+v1.5
    ' Modified to use the INI file to extract the components.  This will make customization and
    ' the transition to the new INI file easier.
        modLmbSigToString = basCCAT.GetAlias("RCV List", "R" & bytLmbRadio, "UNK") & "-" & basCCAT.GetAlias("RCV List", "C" & bytLmbClass, "UNK")
    '    Select Case (bytLmbRadio)
    '        Case 0
    '            modLmbSigToStrig = "UNKNOWN-"
    '        Case 1
    '            modLmbSigToString = "AM-"
    '        Case 2
    '            modLmbSigToString = "FM-"
    '        Case 3
    '            modLmbSigToString = "LS-"
    '        Case 4
    '            modLmbSigToString = "US-"
    '        Case 5
    '            modLmbSigToString = "AR-"
    '        Case 6
    '            modLmbSigToString = "RF-"
    '        Case 7
    '            modLmbSigToString = "IS-"
    '        Case 8
    '            modLmbSigToString = "M1-"
    '        Case 9
    '            modLmbSigToString = "M5-"
    '        Case 10
    '            modLmbSigToString = "M3-"
    '        Case 11
    '            modLmbSigToString = "M6-"
    '        Case 12
    '            modLmbSigToString = "M2-"
    '        Case Else
    '            modLmbSigToString = "UNKNOWN-"
    '    End Select
    '
    '    Select Case (bytLmbClass)
    '        Case 0
    '            modLmbSigToString = modLmbSigToString & "UNKNOWN"
    '        Case 1
    '            modLmbSigToString = modLmbSigToString & "VOICE"
    '        Case 2
    '            modLmbSigToString = modLmbSigToString & "COMMERCIAL-TV"
    '        Case 3
    '            modLmbSigToString = modLmbSigToString & "COMMERCIAL-AUDIO"
    '        Case 4
    '            modLmbSigToString = modLmbSigToString & "TS"
    '        Case 5
    '            modLmbSigToString = modLmbSigToString & "58"
    '        Case 6
    '            modLmbSigToString = modLmbSigToString & "FJ"
    '        Case 7
    '            modLmbSigToString = modLmbSigToString & "HB"
    '        Case 8
    '            modLmbSigToString = modLmbSigToString & "SW"
    '        Case 9
    '            modLmbSigToString = modLmbSigToString & "NL"
    '        Case 10
    '            modLmbSigToString = modLmbSigToString & "MK"
    '        Case 11
    '            modLmbSigToString = modLmbSigToString & "BC"
    '        Case 12
    '            modLmbSigToString = modLmbSigToString & "KT"
    '        Case 13
    '            modLmbSigToString = modLmbSigToString & "MT"
    '        Case 14
    '            modLmbSigToString = modLmbSigToString & "BK"
    '        Case 15
    '            modLmbSigToString = modLmbSigToString & "BN"
    '        Case 16
    '            modLmbSigToString = modLmbSigToString & "BS"
    '        Case 17
    '            modLmbSigToString = modLmbSigToString & "LC"
    '        Case 18
    '            modLmbSigToString = modLmbSigToString & "PB"
    '        Case 19
    '            modLmbSigToString = modLmbSigToString & "TD"
    '        Case 20
    '            modLmbSigToString = modLmbSigToString & "TN"
    '        Case 21
    '            modLmbSigToString = modLmbSigToString & "AN"
    '        Case 22
    '            modLmbSigToString = modLmbSigToString & "NR"
    '        Case 23
    '            modLmbSigToString = modLmbSigToString & "CM"
    '        Case 24
    '            modLmbSigToString = modLmbSigToString & "WA"
    '        Case 25
    '            modLmbSigToString = modLmbSigToString & "WF"
    '        Case 26
    '            modLmbSigToString = modLmbSigToString & "WS"
    '        Case 27
    '            modLmbSigToString = modLmbSigToString & "WQ"
    '        Case 28
    '            modLmbSigToString = modLmbSigToString & "WT"
    '        Case 29
    '            modLmbSigToString = modLmbSigToString & "WK"
    '        Case 30
    '            modLmbSigToString = modLmbSigToString & "WP"
    '        Case 31
    '            modLmbSigToString = modLmbSigToString & "WB"
    '        Case Else
    '            modLmbSigToString = modLmbSigToString & "UNKNOWN"
    '    End Select
    '-v1.5
    End If
End Function
'
' FUNCTION: modLMBSigToLng
' AUTHOR:   Brad Brown
' PURPOSE:  Convert coded radio and class data to common number
' INPUT:    bytLmbRadio = coded radio type
'           bytLmbClass = coded radio class
' OUTPUT:   The combined radio and class values
' NOTES:
Public Function modLMBSigToLng(bytLmbRadio As Byte, bytLmbClass As Byte) As Long

    modLMBSigToLng = (bytLmbRadio * 1000) + bytLmbClass

End Function
'
' FUNCTION: modAllegToString
' AUTHOR:   Brad Brown
' PURPOSE:  Convert coded allegiance values to DAS standard
' INPUT:    bytAlleg = the coded allegiance from CALL
' OUTPUT:   lngIFF = the DAS standard code for the allegiance
'           The standard DAS text for the allegiance
' NOTES:
Public Function modAllegToString(bytAlleg As Byte, ByRef lngIFF As Long) As String
    '
    modAllegToString = basCCAT.GetAlias("IFF", "CALL_IFF" & bytAlleg, "UNKNOWN")
    lngIFF = basCCAT.GetNumber("IFF", "CALL_IFF" & bytAlleg, -1)
    '
    '
    If lngIFF = -1 Then
        Select Case (bytAlleg)
            Case 1
                modAllegToString = "UNKNOWN"
                lngIFF = 3
            Case 2
                modAllegToString = "RED"
                lngIFF = 2
            Case 3
                modAllegToString = "BLUE"
                lngIFF = 1
            Case Else
                modAllegToString = "UNDEF"
                lngIFF = 0
        End Select
    End If
End Function
'
' FUNCTION: modRunmodeToString
' AUTHOR:   Brad Brown
' PURPOSE:  Convert coded Run Mode values to text
' INPUT:    intRunMode = the coded run mode value
' OUTPUT:   The text description of the run mode
' NOTES:
Public Function modRunmodeToString(intRunMode As Integer) As String
    modRunmodeToString = basCCAT.GetAlias("Runmode", "RUNMODE" & intRunMode, "UNKNOWN")
    If modRunmodeToString = "UNKNOWN" Then
        Select Case (intRunMode)
            Case 1
                modRunmodeToString = "PROM"
            Case 2
                modRunmodeToString = "IDLE"
            Case 3
                modRunmodeToString = "STANDBY"
            Case 4
                modRunmodeToString = "SEARCH"
            Case 5
                modRunmodeToString = "JAM"
            Case Else
                modRunmodeToString = "UNKNOWN"
        End Select
    End If
End Function
'
' FUNCTION: modSruStatToString
' AUTHOR:   Shaun Vogel
' PURPOSE:  Convert coded SRU status to text
' INPUT:    bytSruStat = The coded SRU status
' OUTPUT:   The text version of the SRU status
' NOTES:
Public Function modSruStatToString(bytSruStat As Byte) As String
    
    'modXmtrstatToString = basCCAT.GetAlias("Xmtrstat", "XMTRSTAT" & intXmtrStat, "UNKNOWN")
    
    'If modXmtrstatToString = "UNKNOWN" Then
        Select Case (bytSruStat)
            Case 0
                modSruStatToString = "Sru Not There"
            Case 1
                modSruStatToString = "Sru Go"
            Case 2
                modSruStatToString = "Sru No Go"
            Case 3
                modSruStatToString = "Sru Go Not Tested"
            Case Else
                modSruStatToString = "Sru No Go Not Tested"
        End Select
    'End If
End Function

'
' FUNCTION: modXmtrstatToString
' AUTHOR:   Brad Brown
' PURPOSE:  Convert coded transmitter states to text
' INPUT:    intXmtrStat = The coded transmitter state
' OUTPUT:   The text version of the transmitter state
' NOTES:
Public Function modXmtrstatToString(intXmtrStat As Integer) As String
    
    modXmtrstatToString = basCCAT.GetAlias("Xmtrstat", "XMTRSTAT" & intXmtrStat, "UNKNOWN")
    
    If modXmtrstatToString = "UNKNOWN" Then
        Select Case (intXmtrStat)
            Case 0
                modXmtrstatToString = "Xmtr NA"
            Case 1
                modXmtrstatToString = "Xmtr Off"
            Case 2
                modXmtrstatToString = "Xmtr On"
            Case 4
                modXmtrstatToString = "Xmtr Low"
            Case Else
                modXmtrstatToString = "UNKNOWN"
        End Select
    End If
End Function
'
' FUNCTION: modHBFuncToString
' AUTHOR:   Brad Brown
' PURPOSE:  Convert encoded high band functions to text
' INPUT:    intHB_Func = The encoded high band function
' OUTPUT:   The function text
' NOTES:
Public Function modHBFuncToString(ByVal intHB_Func As Integer) As String

    modHBFuncToString = basCCAT.GetAlias("HBFunc", "HBFUNC" & intHB_Func, "UNKNOWN")
    If modHBFuncToString = "UNKNOWN" Then
        Select Case (intHB_Func)
            Case 0
                modHBFuncToString = "Detect"
            Case 1
                modHBFuncToString = "Search"
            Case 2
                modHBFuncToString = "Map"
            Case 3
                modHBFuncToString = "Ask"
            Case 4
                modHBFuncToString = "Jam"
            Case 5
                modHBFuncToString = "Scan"
            Case 6
                modHBFuncToString = "Search once"
            Case Else
                modHBFuncToString = "UNKNOWN"
        End Select
    End If
End Function
'
' FUNCTION: modGetHBID
' AUTHOR:   Brad Brown
' PURPOSE:  Convert high band identifiers to common format
' INPUT:    lngID = High band signal ID
'           enuType = Type of signal identifier
' OUTPUT:   intHB_SigType = The high band signal type
'           intChan = The high band channel
'           The coded high band signal index
' NOTES:
Public Function modGetHBID(lngID As Long, enuType As HBIDType, ByRef intHB_SigType As Integer, ByRef intChan As Integer) As Integer
    Dim intIndex As Integer
    Dim blnFound As Boolean
    Dim astrRange() As String

    intIndex = 1
    blnFound = False
    
    While Not blnFound
        astrRange = Split(basCCAT.GetAlias("MAP", "MAP" & intIndex, "0,9999,0,9999"), ",")
        If ((lngID >= CLng(Val(astrRange(enuType)))) And (lngID <= CLng(Val(astrRange(enuType + 1))))) Then
            blnFound = True
            intHB_SigType = CInt(Val(astrRange(2)))
            If (enuType = HBAID) Then
                intChan = CInt(lngID) - CInt(Val(astrRange(0)))
            End If
        Else
            intIndex = intIndex + 1
        End If
    Wend
    
    modGetHBID = intIndex

End Function
'
' FUNCTION: modXYToLatLon
' AUTHOR:   Brad Brown
' PURPOSE:  Convert XY Position to Lat lon
' INPUT:    X and Y Position
' OUTPUT:   Lat and Lon
' NOTES:
Public Function modXYToLatLon(uPosition As XyCopy) As LatLon
Dim vTempLat As Variant
Dim vTempLon As Variant
Dim vTemp As Variant


    'Call Cart2Geo(uPosition.dX, uPosition.dY, 0, gvPlanAreaLat, gvPlanAreaLon, 0, vTempLat, vTempLon, vTemp)
    Call libGeo.Cartesian_To_Geodetic(uPosition.dX, uPosition.dY, 0, gvPlanAreaLat, gvPlanAreaLon, 0, vTempLat, vTempLon, vTemp)
    modXYToLatLon.lLat = modDegToBam32(vTempLat)
    modXYToLatLon.lLon = modDegToBam32(vTempLon)
    
End Function '
'+v1.6TE
' FUNCTION: sGetOriginID
' AUTHOR:   Tom Elkins
' PURPOSE:  Converts coded message source ID to subsystem name
' INPUT:    lngID = the coded message source ID (the FROM field of the message header)
' OUTPUT:   The name of the subsystem
' NOTES:
Public Function sGetOriginID(lngID As Long) As String
    sGetOriginID = basCCAT.GetAlias("ORIGIN", "ORIGIN" & lngID, "UNKNOWN" & lngID)
End Function
'-v1.6
'
'+v1.6TE
' PROPERTY: ContinueProcessing
' AUTHOR:   Tom Elkins
' PURPOSE:  Allows for early termination of the lengthy translation process externally
' LET:      TRUE = Continue processing
'           FALSE = Stop processing
' OUTPUT:   The current state
' NOTES:
Public Property Let ContinueProcessing(blnState As Boolean)
    mblnContinue = blnState
End Property
'
Public Property Get ContinueProcessing() As Boolean
    ContinueProcessing = mblnContinue
End Property
'-v1.6
'
'+v1.6TE
' PROPERTY: CCOSVersion
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets/Gets the current CCOS version
' LET:      A floating point value indicating the CCOS version
' GET:      The currently set version
' NOTES:    Some message processing routines use the version to determine which parsing
'           method to use.
Public Property Let CCOSVersion(sngVersion As Single)
    msngVersion = sngVersion
End Property
'
Public Property Get CCOSVersion() As Single
    CCOSVersion = msngVersion
End Property
'-v1.6

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


'
'+1.8.10 TE
' ROUTINE:  modDASMTCNTTRG
' AUTHOR:   Tom Elkins
' PURPOSE:  Translate the MTCNTTRG message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDASMTCNTTRG(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMTCNTTRG As msgMTCNTTRG
    Dim pusrDAS_Rec As DAS_MASTER_RECORD
    
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    '
    On Error GoTo Hell
    '
10  pintStruct_Length = LenB(pusrMTCNTTRG)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMTCNTTRG, abytBuffer(pintStart), pintStruct_Length)
    
40  pusrDAS_Rec.dReportTime = dblTime
50  pusrDAS_Rec.sMsg_Type = "MTCNTTRG"
60  pusrDAS_Rec.sReport_Type = "SIG"
70  pusrDAS_Rec.lOrigin_ID = agSwapBytes%(pusrMTCNTTRG.uMsgHdr.iMsgFrom)
80  pusrDAS_Rec.sOrigin = sGetOriginID(pusrDAS_Rec.lOrigin_ID)
90  pusrDAS_Rec.lSignal_ID = agSwapBytes%(pusrMTCNTTRG.iSigID)
100 pusrDAS_Rec.lStatus = agSwapBytes%(pusrMTCNTTRG.bListType)
110 pusrDAS_Rec.lTag = pusrMTCNTTRG.bNumPackets
120 pusrDAS_Rec.lFlag = pusrMTCNTTRG.bPacketNumber
130 pusrDAS_Rec.lEmitter_ID = agSwapBytes%(pusrMTCNTTRG.iSigIDIndex)
140 pusrDAS_Rec.lCommon_ID = agSwapBytes%(pusrMTCNTTRG.iNumSigID)
160 Call Add_Data_Record(MTCNTTRGID, pusrDAS_Rec)
    'Call Process_MTCNTTRG(pusrDAS_Rec)
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMTHBGSUPSTAT"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBGSUPSTAT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBGSUPSTAT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBACTREP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '
End Sub
'-1.8.10 TE
'
'+1.8.10 TE
' ROUTINE:  modDASMTTXRFCONF
' AUTHOR:   Tom Elkins
' PURPOSE:  Translate the MTTXRFCONF message to DAS data structure
' INPUT:    dblTime = the message time stamp
'           intMsg_Length = the message length (in bytes)
'           abytBuffer() = the message contents
' OUTPUT:   None
' NOTES:
Public Sub modDASMTTXRFCONF(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
    Dim pusrMsg As msgMTTXRFCONF
    Dim pusrDAS As DAS_MASTER_RECORD
    
    Dim pintStart As Integer
    Dim pintStruct_Length As Integer
    Dim pdblStart As Double
    Dim pdblEnd As Double
    '
    On Error GoTo Hell
    '
10  pintStruct_Length = LenB(pusrMsg)
20  pintStart = LBound(abytBuffer)
30  Call CopyMemory(pusrMsg, abytBuffer(pintStart), pintStruct_Length)
40  pusrDAS.dReportTime = dblTime
50  pusrDAS.sMsg_Type = "MTTXRFCONF"
60  pusrDAS.sReport_Type = "SIG"
70  pusrDAS.lOrigin_ID = agSwapBytes%(pusrMsg.uMsgHdr.iMsgFrom)
80  pusrDAS.sOrigin = sGetOriginID(pusrDAS.lOrigin_ID)
    pusrDAS.dFrequency = modFreqConv(agSwapWords&(pusrMsg.lFreqStart))
    pusrDAS.dPRI = modFreqConv(agSwapWords&(pusrMsg.lFreqEnd))
    pusrDAS.lEmitter_ID = agSwapBytes(pusrMsg.iTx_Src)
    pusrDAS.lSignal_ID = agSwapBytes(pusrMsg.iTx_Chan)
    pusrDAS.lStatus = pusrMsg.bMsnChg
    pusrDAS.lTag = agSwapBytes(pusrMsg.iTx_Pa_ID)
    pusrDAS.lFlag = agSwapBytes(pusrMsg.iTx_Pa_Pwr_Setting)
    pusrDAS.lCommon_ID = agSwapBytes(pusrMsg.iTx_Ant_Grp)
160 Call Add_Data_Record(MTCNTTRGID, pusrDAS)
    'Call Process_MTTXRCONF(pusrDAS)
    '
    Exit Sub

Hell:
    Dim plngErr_Num As Long
    Dim plngERL As Long
    Dim pstrErr As String
    Dim pstrSource As String
    '
    plngErr_Num = Err.Number
    plngERL = Erl
    pstrErr = Err.Description
    pstrSource = Err.Source
    '
    ' Log the error
    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMTTXRFCONF"
    '
    ' Process the error
    Select Case plngErr_Num
        '
        Case Else:
            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTHBGSUPSTAT message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTHBGSUPSTAT", App.HelpFile, basCCAT.IDH_TRANSLATE_MTHBACTREP)
                Case vbAbort:
                    mblnContinue = False
                    
                Case vbRetry:
                    Resume
                    
                Case vbIgnore:
                    Resume Next
            End Select
    End Select
    '
End Sub
''
''
'' ROUTINE:  modDasMTTGTLISTUPD
'' AUTHOR:   Tom Elkins
'' PURPOSE:  Translate the MTTGTLISTUPD message to DAS data structure
'' INPUT:    dblTime = the message time stamp
''           intMsg_Length = the message length (in bytes)
''           abytBuffer() = the message contents
'' OUTPUT:   None
'' NOTES:
'Public Sub modDasMTTGTLISTUPD(dblTime As Double, intMsg_Length As Integer, abytBuffer() As Byte)
'    Dim pusrMsg As msgMTTGTLISTUPD
'    Dim pusrDAS As DAS_MASTER_RECORD
'    Dim pintStart As Integer
'    Dim pintStruct_Length As Integer
'    Dim pintNum_List As Integer
'    Dim pintLoop As Integer
'    Dim pusrTmp_List As structTgtList
'    '
'    On Error GoTo Hell
'    '
'10  pintStruct_Length = LenB(pusrMsg)
'20  pintStart = LBound(abytBuffer)
'30  Call CopyMemory(pusrMsg, abytBuffer(pintStart), pintStruct_Length)
'
'40  pusrDAS.dReportTime = dblTime
'50  pusrDAS.sMsg_Type = "MTTGTLISTUPD"
'60  pusrDAS.sReport_Type = "SIG"
'70  pusrDAS.lOrigin_ID = agSwapBytes%(pusrMsg.uMsgHdr.iMsgFrom)
'80  pusrDAS.sOrigin = sGetOriginID(pusrDAS.lOrigin_ID)
'    pusrDAS.lTarget_ID = agSwapBytes%(pusrMsg.iMsg_Flag)
'    pusrDAS.lFlag = agSwapBytes%(pusrMsg.uTgt_List(0).iCurrent_Jam)
'    pusrDAS.lIFF = agSwapBytes%(pusrMsg.uTgt_List(0).iFuture_Jam)
'    pusrDAS.lEmitter_ID = agSwapBytes%(pusrMsg.uTgt_List(0).iHB_Actual_Jam)
'    pusrDAS.lSignal_ID = agSwapBytes%(pusrMsg.uTgt_List(0).iHB_Req_Jam)
'    pusrDAS.lCommon_ID = agSwapBytes%(pusrMsg.uTgt_List(0).iHB_Target_List)
'    pusrDAS.dFrequency = CDbl(agSwapBytes%(pusrMsg.uTgt_List(0).iNum_Dup_Freqs))
'    pusrDAS.lStatus = agSwapBytes%(pusrMsg.uTgt_List(0).iStatus)
'    pusrDAS.lTag = agSwapWords&(pusrMsg.uTgt_List(0).lJam_Priority)
'    '
'140 Call Add_Data_Record(MTSIGALARMID, pusrDAS)
'
'150 pintNum_List = agSwapBytes%(pusrMsg.iNumRecs)
'160 If pintNum_List > 1 Then
'170     For pintLoop = 1 To (pintNum_List - 1)
'180         Call CopyMemory(pusrTmp_List, abytBuffer(pintStart + pintStruct_Length + (LenB(pusrTmp_List) * (pintLoop - 1))), LenB(pusrTmp_List))
'            pusrDAS.lFlag = agSwapBytes%(pusrTmp_List.iCurrent_Jam)
'            pusrDAS.lIFF = agSwapBytes%(pusrTmp_List.iFuture_Jam)
'            pusrDAS.lEmitter_ID = agSwapBytes%(pusrTmp_List.iHB_Actual_Jam)
'            pusrDAS.lSignal_ID = agSwapBytes%(pusrTmp_List.iHB_Req_Jam)
'            pusrDAS.lCommon_ID = agSwapBytes%(pusrTmp_List.iHB_Target_List)
'            pusrDAS.dFrequency = CDbl(agSwapBytes%(pusrTmp_List.iNum_Dup_Freqs))
'            pusrDAS.lStatus = agSwapBytes%(pusrTmp_List.iStatus)
'            pusrDAS.lTag = agSwapWords&(pusrTmp_List.lJam_Priority)
'            '
'240         Call Add_Data_Record(MTSIGALARMID, pusrDAS)
'        Next pintLoop
'    End If
'    '
'    '+v1.5 TE
'    Exit Sub
'
'Hell:
'    Dim plngErr_Num As Long
'    Dim plngERL As Long
'    Dim pstrErr As String
'    Dim pstrSource As String
'    '
'    plngErr_Num = Err.Number
'    plngERL = Erl
'    pstrErr = Err.Description
'    pstrSource = Err.Source
'    '
'    ' Log the error
'    basCCAT.WriteLogEntry "Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " (@ line #" & plngERL & ")", "") & " in Filtraw2Das.modDasMtsigalarm"
'    '
'    ' Process the error
'    Select Case plngErr_Num
'
'        '
'        '+v1.6TE
'        'Case vbObjectError + 911:
'        '    Err.Raise plngErr_Num, "Filtraw2DAS.modDasMtsigalarm", "User prematurely terminated the translation process", App.HelpFile, lHlp
'        '-v1.6
'        '
'        Case Else:
'            Select Case MsgBox("Error #" & plngErr_Num & " - " & pstrErr & IIf(plngERL > 0, " @ line #" & plngERL, "") & vbCr & "While translating MTSIGALARM message", vbAbortRetryIgnore Or vbExclamation Or vbMsgBoxHelpButton, "Error Translating MTSIGALARM", App.HelpFile, basCCAT.IDH_TRANSLATE_MTSIGALARM)
'                Case vbAbort:
'                    mblnContinue = False
'
'                Case vbRetry:
'                    Resume
'
'                Case vbIgnore:
'                    Resume Next
'            End Select
'    End Select
'    '-v1.5
'    '
'End Sub

'
'+v1.8.12TE
Public Function HumanTOD(uTod As TodTime) As Date
    Dim dtTime As Date
    Dim pdblSec As Double
    '
    ' Start with 1/1/<year>
    dtTime = CDate("1/1/" & Year(guCurrent.uArchive.dtArchiveDate))
    '
    ' Add the JDay (minus 1)
    dtTime = dtTime + agSwapBytes(uTod.iDayOfYear) - 1
    '
    ' Compute the number of seconds
    '   1) set the value to the high-byte
    '   2) shift the value LEFT to accomodate the remaining LONG (32 bits): 2^32 = 4,294,967,296
    '   3) add the long: this value is the number of microseconds
    '   4) divide by 1 million to convert to seconds
    pdblSec = ((CDbl(uTod.bytHighUsecs) * 4294967296#) + Abs(CDbl(agSwapWords(uTod.lUsecs)))) / 1000000#
    '
    ' Add the seconds to the days
    HumanTOD = DateAdd("s", pdblSec, dtTime)
End Function
'-v1.8.12TE
'
'
'+v1.8.12TE
Public Function LongTOD(uTod As TodTime) As Long
    Dim dtTime As Date
    Dim pdblSec As Double
    '
    ' Start with 1/1/<year>
    dtTime = CDate("1/1/" & Year(guCurrent.uArchive.dtArchiveDate))
    '
    ' Add the JDay (minus 1)
    dtTime = dtTime + agSwapBytes(uTod.iDayOfYear) - 1
    '
    ' Compute the number of seconds
    '   1) set the value to the high-byte
    '   2) shift the value LEFT to accomodate the remaining LONG (32 bits): 2^32 = 4,294,967,296
    '   3) add the long: this value is the number of microseconds
    '   4) divide by 1 million to convert to seconds
    pdblSec = ((CDbl(uTod.bytHighUsecs) * 4294967296#) + Abs(CDbl(agSwapWords(uTod.lUsecs)))) / 1000000#
    '
    ' Add the seconds to the days
    LongTOD = pdblSec
End Function
'+v1.8.12TE
Public Function SwapDoubleWords(dblNum As Double) As Double
    Dim tmpByteArray(1 To 8) As Byte
    Dim tmpByteArray2(1 To 8) As Byte
    Dim i As Integer
    '
    Call CopyMemory(tmpByteArray(1), dblNum, 8)
    For i = 1 To 8
        tmpByteArray2(9 - i) = tmpByteArray(i)
    Next i
    Call CopyMemory(SwapDoubleWords, tmpByteArray2(1), 8)
End Function
'-v1.8.12TE
'
