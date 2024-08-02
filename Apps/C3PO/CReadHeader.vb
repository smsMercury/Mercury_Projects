Imports System.Runtime.InteropServices
Imports System.IO

Public Class CReadHeader


    Private chelp As New CHelperFunctions


    Public Enum eMsgRec
        MsgTableRec = 0
        LmbSigRec
        HbSigRec
    End Enum

    Public Enum eVarTypes
        CharVarType
        ShortVarType
        LongVarType
        LongLongVarType
        FloatVarType
        DoubleVarType
        EnumVarType
        StructVarType
        TypeDefVarType
        UnionVarType
        PtrVarType
        LLVarType
        TimeVarType
        FreqVarType
        Bam16VarType
    End Enum

    Structure MsgDskEntry
        Dim lStartLoc As Integer
        Dim lSize As Integer
        Dim lCount As Integer
    End Structure

    Structure MsgEntryTable
        Dim tMsgDskEntry() As MsgDskEntry
    End Structure

    Structure SymbolType
        Dim lFldSymIdx As Integer
        Dim iBitSize As Int16
        Dim iArraySize As Int16
        Dim bUnsigned As Byte
        Dim bSubType As Byte
        Dim bPad As Int16
    End Structure

    Structure FieldType
        Dim lStrIdx As Integer
        Dim iSymbolId As Int16
        Dim iArraySize As Int16
        Dim iArrayKey As Int16
        Dim bArrayKeyType As Byte
        Dim bPad As Byte
    End Structure

    Structure DebugFileHeader
        Dim lStringSize As Integer
        Dim lNumMsgs As Integer
        Dim lNumSymRecs As Integer
        Dim lNumFldRecs As Integer
    End Structure

    Structure MtId2Type
        Dim iMsgId As Int16
        Dim iFieldId As Int16
    End Structure

    'Const ARC_DATA_FILENAME = "archive.dat"
    Const BAD_MSG_IDX = 8000
    Const BAM16_SIZE = 2
    Const FREQ_SIZE = 4
    Const MAX_SMART_ARRAYS = 65535
    Const MsgDataLimit = 4096
    Const UndefinedAryType = -1

    Private lMaxMsgId As Integer
    Private lMsgRawMode As Integer
    Private bAbort As Boolean
    Private tDebugFileHeader As DebugFileHeader
    Private sSymbolType As SymbolType
    Private sFieldType As FieldType
    Private bMsgStrings() As Byte
    Private tMtId2Type() As MtId2Type
    Private lMsgId2Idx() As Integer
    Private tSymbolType() As SymbolType
    Private tFieldType() As FieldType
    Private aStringType() As String
    Private MsgArraySizes(MAX_SMART_ARRAYS) As Integer
    Private iMultSize As Int16
    Private sFldLbl As String
    Private lCurMsgId As Integer
    Private lVarStrId As Integer
    Private iStructLevel As Int16
    Private msngVersion As Single   ' Stores the CCOS version
    Private mcurrentDb As CMasterDb

    '+vl.7BB
    'ROUTINE:	GetArraySize
    'AUTHOR:	Shaun Vogel
    'PURPOSE:	Determines the size of the array being parsed.
    'INPUT:	al - number of elements in array
    'a2 - size of array type
    'OUTPUT:	al - holds array size.
    'NOTES: 
    Public Sub GetArraySize(ByRef al As Integer, ByVal a2 As Integer)
        If al = UndefinedAryType Then
            al = a2
        ElseIf a2 <> UndefinedAryType Then
            al = al * a2
        End If
    End Sub

    '+vl.7BB
    'ROUTINE: GetString()
    'AUTHOR: Shaun(Vogel)
    'PURPOSE:	Converts a byte array into a string value.
    'INPUT:	IStrIndex - The current index into the message byte array.
    'OUTPUT:	Converted string from bMsgStrings array.
    'NOTES:   bMsgStrings is the Debug File Header read in from the "archive.dat" file.
    Public Function GetString(ByVal IStrlndex As Long) As String
        Dim lTmpCount As Long
        Dim sTmpStr As String = ""

        lTmpCount = IStrlndex
        While bMsgStrings(lTmpCount) <> 0
            sTmpStr = sTmpStr + Chr(bMsgStrings(lTmpCount))
            lTmpCount = lTmpCount + 1
        End While
        Return sTmpStr
    End Function

    '+vl.7BB
    'ROUTINE: PrcMsgData()
    'AUTHOR: Shaun(Vogel)
    'PURPOSE:	Loop through entire message and parse each field.
    'INPUT:	sStr - string, with "." notation, being parsed.
    'lFldId2 - Index into message string for field being parsed.
    'OUTPUT:	iSize - The size of the field being processed.
    'NOTES:
    Public Function PrcMsgData(ByVal sStr As String, ByVal lFldId2 As Integer) As Integer
        Dim tField As FieldType
        Dim iArraySize As Integer
        Dim iPrintSize As Integer
        Dim tSym As SymbolType
        Dim tSym2 As SymbolType
        Dim sTmpStr As String
        Dim sTmpStr2 As String
        Dim iSize As Integer
        Dim bLclAbort As Boolean
        Dim i As Integer
        Dim lTemp As Long
        Dim ldx As Long
        Dim lvalue As Long
        Dim lFldld As Long
        Dim iFrmInt As Integer

        lFldld = lFldId2
        tField = tFieldType(lFldld)
        iArraySize = tField.iArraySize
        iPrintSize = iArraySize
        iSize = 0
        bLclAbort = False
        iFrmInt = lCurMsgId
        If (tField.iSymbolId >= 0) Then
            If (tField.iSymbolId > tDebugFileHeader.lNumSymRecs) Then
                MsgBox("Found invalid symbol Id.  Record Ignored")
                bAbort = True
                PrcMsgData = 0
            Else
                bAbort = False

                tSym = tSymbolType(tField.iSymbolId)
                tSym2 = tSym
                If sStr <> "" Then
                    sFldLbl = GetString(tField.lStridx)
                    sTmpStr = sStr & "." & sFldLbl
                Else
                    sTmpStr = GetString(tField.lStridx)
                    If (Left(sTmpStr, 2) <> "mt") Then
                        Exit Function
                    End If
                    mcurrentDb.updateMsgTable(lCurMsgId, sTmpStr)
                    '•dbMsgTbl.AddNew()
                    '•dbMsgTbl!MSG_ID = lCurMsgld
                    '•dbMsgTbl!MSG_NAME = sTmpStr
                    '                'If frmWizard.IsSelected(iFrmlnt) Then dbMsgTbl!Select_Msg = True
                    '•dbMsgTbl.Update()
                End If

                GetArraySize(iArraySize, tSym2.iArraySize)
                While (tSym2.bSubType = eVarTypes.TypeDefVarType)
                    tSym2 = tSymbolType(tSym.lFldSymIdx)
                    GetArraySize(iArraySize, tSym2.iArraySize)
                End While

                If Not (iArraySize = UndefinedAryType) Then
                    iPrintSize = iArraySize
                ElseIf (tField.iArrayKey > 0) Then
                    Select Case tSym2.bSubType
                        Case eVarTypes.CharVarType
                            iSize = 1
                        Case eVarTypes.ShortVarType
                            iSize = 2
                        Case eVarTypes.LongVarType
                            iSize = 4
                        Case eVarTypes.LongLongVarType
                            iSize = 8
                        Case Else
                            iSize = 1
                    End Select
                    If (iSize > MsgDataLimit) Then
                        bLclAbort = True
                        iSize = 1
                    End If
                    lTemp = tField.iArrayKey
                    MsgArraySizes(lTemp) = iSize
                End If

                Select Case tSym2.bSubType
                    Case eVarTypes.CharVarType
                        iSize = 1
                        lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "CHAR")
                    Case eVarTypes.ShortVarType
                        iSize = 2
                        lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "SHORT")
                    Case eVarTypes.LongVarType
                        iSize = 4
                        lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "LONG")
                    Case eVarTypes.LongLongVarType
                        iSize = 8
                        lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "LONGLONG")
                    Case eVarTypes.PtrVarType
                        iSize = 4
                        lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "POINTER")
                    Case eVarTypes.FloatVarType
                        iSize = 4
                        lTemp = PrcVar(sTmpStr, iArraySize, 0, iSize, "FLOAT")
                    Case eVarTypes.DoubleVarType
                        iSize = 8
                        lTemp = PrcVar(sTmpStr, iArraySize, 0, iSize, "DOUBLE")
                    Case eVarTypes.EnumVarType
                        Select Case tSym2.iBitSize
                            Case 64
                                iSize = 8
                            Case 32
                                iSize = 4
                            Case 16
                                iSize = 2
                            Case 8
                                iSize = 1
                            Case Else
                                iSize = 4
                        End Select
                        lTemp = PrcVar(sTmpStr, iArraySize, 0, iSize, "ENUM")
                    Case eVarTypes.StructVarType, eVarTypes.LLVarType, eVarTypes.TimeVarType, eVarTypes.FreqVarType, eVarTypes.Bam16VarType
                        If (iArraySize = UndefinedAryType) Then
                            iMultSize = 0
                            iSize = PrcStruct(sTmpStr, tSym2.lFldSymIdx, tSym2.bSubType)
                        Else
                            iMultSize = iArraySize
                            'For ldx = 0 To iPrintSize - 1
                            For ldx = 0 To 0
                                sTmpStr2 = sTmpStr & "[" & ldx.ToString & "]"
                                iSize = PrcStruct(sTmpStr2, tSym2.lFldSymIdx, tSym2.bSubType)
                                If bAbort = True Then
                                    iSize = 0
                                    ldx = iPrintSize
                                ElseIf ((ldx = 0) And ((iSize * iPrintSize) > MsgDataLimit)) Then
                                    bAbort = True
                                    iSize = 0
                                    ldx = iPrintSize
                                End If
                                iMultSize = 0
                            Next ldx
                        End If
                    Case eVarTypes.UnionVarType
                        If iArraySize = UndefinedAryType Then
                            iSize = PrcUnion(sTmpStr, tSym2.lFldSymIdx)
                        Else
                            'For ldx = 0 To iPrintSize - 1 
                            For ldx = 0 To 0
                                sTmpStr2 = sTmpStr & "[" & ldx.ToString & "]"
                                iSize = PrcUnion(sTmpStr2, tSym2.lFldSymIdx)
                                If bAbort = True Then
                                    iSize = 0
                                    ldx = iPrintSize
                                ElseIf ((ldx = 0) And ((iSize * iPrintSize) > MsgDataLimit)) Then
                                    bAbort = True
                                    iSize = 0
                                    ldx = iPrintSize

                                End If
                            Next ldx
                        End If
                    Case Else
                        MsgBox("Unknown Subtype found.  Ignored!")
                End Select
                If Not (iArraySize = UndefinedAryType) Then
                    iSize = iSize * iArraySize
                End If
                If bLclAbort = True Then
                    MsgBox("Array size > limit")
                End If
                Return iSize
            End If
        Else
            MsgBox("Invalid SymbolID found.  Ignored!")
        End If
    End Function

    '+vl.7BB
    'ROUTINE: PrcUnion()
    'AUTHOR: Shaun(Vogel)
    'PURPOSE:	Processes a union struct by recursively calling PrcMsgData.
    'INPUT:	sStr - string, with "." notation, being parsed.
    'lFldId2 - Index into message string for field being parsed.
    'OUTPUT:	iSize - size of union
    'NOTES:
    Private Function PrcUnion(ByVal sStr As String, ByVal lFldId As Long) As Integer
        Dim tField As FieldType
        Dim iSum As Integer
        Dim iSize As Integer

        iSize = 0
        Do
            tField = tFieldType(lFldId)
            If tField.lStrIdx <> 0 Then
                iSum = PrcMsgData(sStr, lFldId)
                lFldId = lFldId + 1
                If (bAbort = True) Then
                    iSize = 0
                    lFldId = -1
                ElseIf (iSize <= iSum) Then
                    iSize = iSum
                End If
            Else
                lFldId = -1
            End If
        Loop Until lFldId = -1
        Return iSize
    End Function

    '+vl.7BB ROUTINE:  PrcStruct 
    'AUTHOR:   Shaun Vogel 
    'PURPOSE:  Recursively parses each structure within a message and writes it to
    'the VarStructTbl. 
    'INPUT:   sStr - string containing structure name with "." notation.
    'lFldldx - Index into the field definition structure.
    'eSubType - structure variable type. 
    'OUTPUT:   iSize - The size of the structure being processed.
    ' NOTES:   The types of structures that are processed:
    '	Baml6, LatLon, Frequency, and Time.
    Private Function PrcStruct(ByVal sStr As String, ByVal lFldldx As Long, ByVal eSubType As Byte) As Integer
        Dim iSize As Integer
        Dim tField As FieldType
        Dim iSum As Integer

        iSize = 0
        Select Case eSubType
            Case eVarTypes.Bam16VarType
                iSize = BAM16_SIZE
                'get new var_struct_table rec 
                lVarStrId = lVarStrId + 1
                mcurrentDb.updateVarStructTable(lVarStrId, lCurMsgId, sStr, iSize, "BAM16", 0, sFldLbl, 0, 0, 0, 0)

            Case eVarTypes.FreqVarType
                iSize = FREQ_SIZE
                'get new var_struct_table rec 
                lVarStrId = lVarStrId + 1
                mcurrentDb.updateVarStructTable(lVarStrId, lCurMsgId, sStr, iSize, "FREQ", 0, sFldLbl, 0, 0, 0, 0)

            Case eVarTypes.TimeVarType
                iSize = 0
                eSubType = eVarTypes.StructVarType
            Case eVarTypes.LLVarType
                iSize = 0
                eSubType = eVarTypes.StructVarType
        End Select

        If eSubType = eVarTypes.StructVarType Then
            Dim myexit As Boolean

            myexit = False
            iStructLevel = iStructLevel + 1
            tField = tFieldType(lFldldx)

            'get new var_struct_table rec 
            lVarStrId = lVarStrId + 1
            mcurrentDb.updateVarStructTable(lVarStrId, lCurMsgId, sStr, 0, "STRUCT BEGIN", 0, "", 0, iMultSize, 0, iStructLevel)

            Do
                If tField.lStrIdx <> 0 Then
                    iSum = PrcMsgData(sStr, lFldldx)
                    lFldldx = lFldldx + 1
                    tField = tFieldType(lFldldx)
                    If (bAbort = True) Then
                        iSize = 0
                        myexit = True
                    Else
                        iSize = iSize + iSum
                    End If
                Else
                    myexit = True
                End If
            Loop Until (tField.lStrIdx = 0)
            'get new var_struct_table rec lVarStrld = lVarStrld + 1
            mcurrentDb.updateVarStructTable(lVarStrId, lCurMsgId, "", 0, "STRUCT END", 0, "", 0, 0, 0, iStructLevel)
            'dbVarStrTbl.AddNew() ''save STRUCT END info
            ''update record set
            'dbVarStrTbl.Update()
            iStructLevel = iStructLevel - 1
        End If
        Return iSize
    End Function

    '+vl.7BB
    ' ROUTINE:  PrcVar
    'AUTHOR: Shaun(Vogel)
    'notation.
    'PURPOSE:  Write field with attributes to the VarStructTbl
    'INPUT:   sStr - string containing structure name with ".'*
    'iArraySize - size of array or 0 if not an array.
    'bUnsigned - True if field is unsigned else 0.
    'iSize - size of current field being processed. OUTPUT:   iTmpSize - size of field
    'NOTES:   TOC tables contain results from processing an archive

    Private Function PrcVar(ByVal sStr As String, ByVal iArraySize As Integer, ByVal bUnsigned As Byte, ByVal iSize As Integer, ByVal sVarType As String) As Integer
        Dim iTmpSize As Integer

        iTmpSize = iSize
        If Not (iArraySize = UndefinedAryType) Then
            If iArraySize > MsgDataLimit Then
                'MsgBox "Array is greater than MsgDataLimit!" 
                bAbort = True
                iTmpSize = 0
            Else
                iTmpSize = iArraySize * iSize
                'get new var_struct_table rec 
                lVarStrId = lVarStrId + 1
                mcurrentDb.updateVarStructTable(lVarStrId, lCurMsgId, sStr, iTmpSize, sVarType, 0, sFldLbl, 0, iArraySize, 0, 0)

            End If
        Else
            'get new var_struct_table rec 
            lVarStrId = lVarStrId + 1
            mcurrentDb.updateVarStructTable(lVarStrId, lCurMsgId, sStr, iTmpSize, sVarType, 0, sFldLbl, 0, 0, 0, 0)
        End If
        Return iTmpSize
    End Function

    '+vl.7BB
    'ROUTINE:	readHeader
    'AUTHOR:	Shaun Vogel
    ' PURPOSE:	Read the archive.dat file, parse the fields in each message, and store  *
    Public Sub readHeader(ByVal fname As String, ByVal currentDb As CMasterDb)
        Dim tMsgEntries As New MsgEntryTable
        Dim lTmpSize, lTmp, fTmp As Integer

        'Dim bBucket() As Byte
        ''Dim lSwapNumMsgs As Long 
        'Dim iSwapEntries As Intl6 
        'Dim btmpMsgString{) As Byte 
        'Dim lswapl As Integer 
        'Dim lswap2 As Integer 
        'Dim lswap3 As Integer 
        Dim iCurMsgld As Int16 = 0
        Dim iVarStrld As Int16 = 0
        Dim input As New IO.FileStream(fname, FileMode.Open, FileAccess.Read)
        Dim buffer() As Byte
        Dim count As Integer = 0
        Dim errMsg As String = ""

        ReDim tMsgEntries.tMsgDskEntry(eMsgRec.HbSigRec)

        mcurrentDb = currentDb
        '    On Error GoTo FileError 
        'Open fname For Binary As #1 'Get #1, , tMsgEntries
        'Read Message Entries 
        Try
            Cursor.Current = Cursors.WaitCursor
            errMsg = "reading message entries."
            count = 36 'Size of tMsgEntries.tMsgDskEntry 
            ReDim buffer(count)
            count = input.Read(buffer, 0, count)
            chelp.swapWord(buffer, 4)
            tMsgEntries.tMsgDskEntry(eMsgRec.MsgTableRec).lSize = BitConverter.ToInt32(buffer, 4)
            chelp.swapWord(buffer, 16)
            tMsgEntries.tMsgDskEntry(eMsgRec.LmbSigRec).lSize = BitConverter.ToInt32(buffer, 16)
            chelp.swapWord(buffer, 28)
            tMsgEntries.tMsgDskEntry(eMsgRec.HbSigRec).lSize = BitConverter.ToInt32(buffer, 28)
            lTmpSize = tMsgEntries.tMsgDskEntry(eMsgRec.MsgTableRec).lSize + _
            tMsgEntries.tMsgDskEntry(eMsgRec.LmbSigRec).lSize + _
            tMsgEntries.tMsgDskEntry(eMsgRec.HbSigRec).lSize
            'Skip prep stuff seek
            'ReDim bBucketd To lTmpSize)
            'Get #1, , bBucket
            errMsg = "skipping prep data."
            count = lTmpSize
            ReDim buffer(count)
            count = input.Read(buffer, 0, count)
            'Read debug file header info 
            'Get #1, , tDebugFileHeader

            errMsg = "reading debug file header info."
            count = Marshal.SizeOf(tDebugFileHeader)
            ReDim buffer(count)
            count = input.Read(buffer, 0, count)
            'Read string info
            'get stringsize
            chelp.swapWord(buffer, 0)
            tDebugFileHeader.lStringSize = BitConverter.ToInt32(buffer, 0)
            lTmpSize = tDebugFileHeader.lStringSize
            'get nummsgs
            chelp.swapWord(buffer, 4)
            tDebugFileHeader.lNumMsgs = BitConverter.ToInt32(buffer, 4)
            'get numsymrecs
            chelp.swapWord(buffer, 8)
            tDebugFileHeader.lNumSymRecs = BitConverter.ToInt32(buffer, 8)
            'get numfldrecs
            chelp.swapWord(buffer, 12)
            tDebugFileHeader.lNumFldRecs = BitConverter.ToInt32(buffer, 12)
            ReDim bMsgStrings(lTmpSize)
            'Get #1, , bMsgStrings
            errMsg = "reading message strings."
            count = lTmpSize
            count = input.Read(bMsgStrings, 0, count)
            'Read Msg Id's and Field Id index's
            ReDim buffer(tDebugFileHeader.lNumMsgs * 4) '4 bytes/record in tMtId2Type
            'Get #1, , tMtId2Type
            errMsg = "reading msg id's and field id index's."
            count = tDebugFileHeader.lNumMsgs * 4
            count = input.Read(buffer, 0, count)
            'Calculate number of valid messages in file
            lMaxMsgId = 0
            ReDim tMtId2Type(tDebugFileHeader.lNumMsgs)
            Dim pos As Integer = 0
            For lTmp = 0 To tDebugFileHeader.lNumMsgs - 1
                pos = lTmp * 4  'number to skip in buffer
                chelp.swapBytes(buffer, pos)
                tMtId2Type(lTmp).iMsgId = BitConverter.ToInt16(buffer, pos)
                chelp.swapBytes(buffer, pos + 2)
                tMtId2Type(lTmp).iFieldId = BitConverter.ToInt16(buffer, pos + 2)
                If tMtId2Type(lTmp).iMsgId > lMaxMsgId Then
                    lMaxMsgId = tMtId2Type(lTmp).iMsgId
                End If
            Next lTmp
            'Initialize msg id's
            ReDim lMsgId2Idx(lMaxMsgId)
            For lTmp = 0 To lMaxMsgId - 1
                lMsgId2Idx(lTmp) = BAD_MSG_IDX
            Next lTmp
            'Store msg id's
            For lTmp = 0 To tDebugFileHeader.lNumMsgs - 1
                lMsgId2Idx(tMtId2Type(lTmp).iMsgId) = lTmp
            Next lTmp
            'Read all Symbol Records
            lTmpSize = tDebugFileHeader.lNumSymRecs
            ReDim tSymbolType(lTmpSize)
            count = lTmpSize * Marshal.SizeOf(sSymbolType)
            'Get #1, , tSymbolType
            ReDim buffer(count)
            count = input.Read(buffer, 0, count)

            If count > 0 Then
                processSymbolType(buffer, lTmpSize)
            End If
            'Read all Field Records
            lTmpSize = tDebugFileHeader.lNumFldRecs
            ReDim tFieldType(lTmpSize)
            count = lTmpSize * Marshal.SizeOf(sFieldType)
            'Get #1, , tFieldType
            ReDim buffer(count)
            count = input.Read(buffer, 0, count)
            If count > 0 Then
                processFieldType(buffer, lTmpSize)
            End If
            input.Close()
        Catch ex As Exception
            MsgBox("Error:  " & errMsg & vbCrLf & ex.Message)
            Exit Sub
        End Try
        'Open the Database and set the record set to the var_struct_table 
        'dbVarStrTbl = guCurrent.uArchive.rsVarStruct
        currentDb.CreateVarStructTbl()
        currentDb.CreateMsgTbl()
        'Loop through Max Message Id's and store data in database
        mcurrentDb.OpenDB()
        For lTmp = 0 To tDebugFileHeader.lNumMsgs - 1
            lTmp = tMtId2Type(lTmp).iMsgId
            lCurMsgId = lTmp
            Dim i As Integer
            If lTmp = 325 Then
                i = 0
            End If
            If lTmp <= lMaxMsgId Then
                lTmp = lMsgId2Idx(lTmp)
            Else
                lTmp = BAD_MSG_IDX
            End If
            If Not (lTmp = BAD_MSG_IDX) Then
                'Process each message and store in database 
                sFldLbl = ""
                iStructLevel = 0
                iMultSize = 0
                fTmp = tMtId2Type(lTmp).iFieldId
                lTmpSize = PrcMsgData("", fTmp)
            End If
        Next lTmp
        mcurrentDb.CloseDB()
        Cursor.Current = Cursors.Default
        MsgBox("Finished Parsing HDR2 file.")
    End Sub

    Private Sub processSymbolType(ByRef buf() As Byte, ByVal numRecs As Integer)
        Dim ptr As Integer = 0

        For i As Integer = 0 To numRecs - 1
            'start at lFldSymldx 
            chelp.swapWord(buf, ptr)
            tSymbolType(i).lFldSymIdx = BitConverter.ToInt32(buf, ptr)
            ptr = ptr + 4 'skip to iBitSize
            chelp.swapBytes(buf, ptr)
            tSymbolType(i).iBitSize = BitConverter.ToInt16(buf, ptr)
            ptr = ptr + 2 'skip to iArraySize 
            chelp.swapBytes(buf, ptr)
            tSymbolType(i).iArraySize = BitConverter.ToInt16(buf, ptr)
            ptr = ptr + 2 'skip to bUnsigned 
            tSymbolType(i).bUnsigned = buf(ptr)
            ptr = ptr + 1 'skip to bSubType 
            tSymbolType(i).bSubType = buf(ptr)
            ptr = ptr + 1 'skip to bPad 
            tSymbolType(i).bPad = 0
            ptr = ptr + 2 'skip to lFldSymldx 
            If ptr > buf.Length Then Exit For
        Next
    End Sub

    Private Sub processFieldType(ByVal buf() As Byte, ByVal numRecs As Integer)
        Dim ptr As Integer = 0

        For i As Integer = 0 To numRecs - 1
            'start at IStrldx
            chelp.swapWord(buf, ptr)
            tFieldType(i).lStrIdx = BitConverter.ToInt32(buf, ptr)
            ptr = ptr + 4 'skip to iSymbolId
            chelp.swapBytes(buf, ptr)
            tFieldType(i).iSymbolId = BitConverter.ToInt16(buf, ptr)
            ptr = ptr + 2 'skip to iArraySize
            chelp.swapBytes(buf, ptr)
            tFieldType(i).iArraySize = BitConverter.ToInt16(buf, ptr)
            ptr = ptr + 2 'skip to iArrayKey
            chelp.swapBytes(buf, ptr)
            tFieldType(i).iArrayKey = BitConverter.ToInt16(buf, ptr)
            ptr = ptr + 2 'skip to bArrayKeyType
            tFieldType(i).bArrayKeyType = buf(ptr)
            ptr = ptr + 1 'skip to bPad
            tFieldType(i).bPad = 0
            ptr = ptr + 1 'skip to IStrldx
            If ptr > buf.Length Then Exit For
        Next
    End Sub
End Class

