Attribute VB_Name = "FileOps"
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' MODULE:   FileOps
' AUTHOR:   Shaun Vogel
' PURPOSE:  To create the VAR_STRUCT_TBL used to process the archive files.
' REVISION:
'   v1.7BB
'   v1.7.4SPV

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal length As Long)
'Declare Function agSwapBytes% Lib "apigid32.dll" (ByVal src%)
'Declare Function agSwapWords& Lib "apigid32.dll" (ByVal src&)

Public Enum eMsgRec
   MsgTableRec = 1
   LmbSigRec
   HbSigRec
End Enum
   
Public Enum eVarTypes
   CharVarType
   ShortVarType
   LongVarType
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

Public Enum eVarTypes2_41
   CharVarType2_41
   ShortVarType2_41
   LongVarType2_41
   LongLongVarType2_41
   FloatVarType2_41
   DoubleVarType2_41
   EnumVarType2_41
   StructVarType2_41
   TypeDefVarType2_41
   UnionVarType2_41
   PtrVarType2_41
   LLVarType2_41
   TimeVarType2_41
   FreqVarType2_41
   Bam16VarType2_41
End Enum

Type MsgDskEntry
   lStartLoc As Long
   lSize As Long
   lCount As Long
End Type
   
Type MsgEntryTable
   tMsgDskEntry(MsgTableRec To HbSigRec) As MsgDskEntry
End Type

Type SymbolType
   lFldSymIdx As Long
   iBitSize As Integer
   iArraySize As Integer
   bUnsigned As Byte
   bSubType As Byte
   bPad(1 To 2) As Byte
End Type
    
Type FieldType
   lStrIdx As Long
   iSymbolId As Integer
   iArraySize As Integer
   iArrayKey As Integer
   bArrayKeyType As Byte
   bPad As Byte
End Type

Type DebugFileHeader
   lStringSize As Long
   lNumMsgs As Long
   lNumSymRecs As Long
   lNumFldRecs As Long
End Type

Type MtId2Type
   iMsgId As Integer
   iFieldId As Integer
End Type
   
'Const ARC_DATA_FILENAME = "archive.dat"

Const BAD_MSG_IDX = 8000
Const BAM16_SIZE = 2
Const FREQ_SIZE = 4
Const MAX_SMART_ARRAYS = 65535
Const MsgDataLimit = 4096
Const UndefinedAryType = -1

Dim lMaxMsgId As Long
Dim lMsgRawMode As Long
Public bAbort As Boolean

Dim tDebugFileHeader As DebugFileHeader

Dim bMsgStrings() As Byte
Dim tMtId2Type() As MtId2Type
Dim lMsgId2Idx() As Long
Dim tSymbolType() As SymbolType
Dim tFieldType() As FieldType
Dim aStringType() As String
Dim MsgArraySizes(MAX_SMART_ARRAYS) As Long
Dim iMultSize As Integer

Dim sFldLbl As String
Dim lCurMsgId As Long
Dim lVarStrId As Long
Dim iStructLevel As Integer
Dim wsCCAT As Workspace
Dim dbCCAT As Database
Dim dbMsgTbl As Recordset
Dim dbVarStrTbl As Recordset

Private msngVersion As Single           ' Stores the CCOS version

'
'+v1.7.4SPV
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
'-v1.7.4


'
'+v1.7BB
' ROUTINE:  GetArraySize
' AUTHOR:   Shaun Vogel
' PURPOSE:  Determines the size of the array being parsed.
' INPUT:    a1 - number of elements in array
'           a2 - size of array type
' OUTPUT:   a1 - holds array size.
'
' NOTES:

Public Sub GetArraySize(a1 As Integer, a2 As Integer)

   If a1 = UndefinedAryType Then
      a1 = a2
   ElseIf a2 <> UndefinedAryType Then
      a1 = a1 * a2
   End If
   
End Sub

'
'+v1.7BB
' ROUTINE:  GetString
' AUTHOR:   Shaun Vogel
' PURPOSE:  Converts a byte array into a string value.
' INPUT:    lStrIndex - The current index into the message byte array.
' OUTPUT:   Converted string from bMsgStrings array.
'
' NOTES:    bMsgStrings is the Debug File Header read in from the "archive.dat" file.

Public Function GetString(lStrIndex As Long) As String
   Dim lTmpCount As Long
   Dim sTmpStr As String
   
   lTmpCount = lStrIndex
   
   While bMsgStrings(lTmpCount) <> 0
      sTmpStr = sTmpStr + Chr(bMsgStrings(lTmpCount))
      lTmpCount = lTmpCount + 1
   Wend
   
   GetString = sTmpStr
End Function

'
'+v1.7BB
' ROUTINE:  PrcMsgData
' AUTHOR:   Shaun Vogel
' PURPOSE:  Loop through entire message and parse each field.
' INPUT:    sStr - string, with "." notation, being parsed.
'           lFldId2 - Index into message string for field being parsed.
' OUTPUT:   iSize - The size of the field being processed.
'
' NOTES:

Public Function PrcMsgData(sStr As String, ByVal lFldId2 As Long) As Integer
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
   Dim lValue As Long
   Dim lFldId As Long
   Dim iFrmInt As Integer
   
   lFldId = lFldId2
      
   tField = tFieldType(lFldId)
   iArraySize = agSwapBytes%(tField.iArraySize)
   iPrintSize = iArraySize
   iSize = 0
   bLclAbort = False
   iFrmInt = lCurMsgId
   
   If (agSwapBytes(tField.iSymbolId) >= 0) Then
       If (agSwapBytes%(tField.iSymbolId) > agSwapWords&(tDebugFileHeader.lNumSymRecs)) Then
          MsgBox ("Found invalid symbol Id.  Record Ignored")
          bAbort = True
          PrcMsgData = 0
       Else
          bAbort = False
          tSym = tSymbolType(agSwapBytes%(tField.iSymbolId) + 1)
          tSym2 = tSym
          
          If sStr <> "" Then
             sFldLbl = GetString(agSwapWords&(tField.lStrIdx) + 1)
             sTmpStr = sStr & "." & sFldLbl
          Else
             sTmpStr = GetString(agSwapWords&(tField.lStrIdx) + 1)
             If (Left(sTmpStr, 2) <> "mt") Then
                Exit Function
             End If
             
             dbMsgTbl.AddNew
             dbMsgTbl!Msg_id = lCurMsgId
             dbMsgTbl!Msg_Name = sTmpStr
             If frmWizard.IsSelected(iFrmInt) Then dbMsgTbl!Select_Msg = True
             dbMsgTbl.Update
    
          End If
          
          GetArraySize iArraySize, agSwapBytes%(tSym2.iArraySize)
          
          While (tSym2.bSubType = TypeDefVarType)
             tSym2 = tSymbolType(agSwapWords&(tSym.lFldSymIdx) + 1)
             GetArraySize iArraySize, agSwapBytes%(tSym2.iArraySize)
          Wend
             
          If Not (iArraySize = UndefinedAryType) Then
                   iPrintSize = iArraySize
          ElseIf (agSwapBytes%(tField.iArrayKey) > 0) Then
             Select Case tSym2.bSubType
                Case CharVarType:
                   iSize = 1
                Case ShortVarType:
                   iSize = 2
                Case LongVarType:
                   iSize = 4
                Case Else:
                   iSize = 1
             End Select
             If (iSize > MsgDataLimit) Then
                bLclAbort = True
                iSize = 1
             End If
             lTemp = agSwapBytes%(tField.iArrayKey)
             MsgArraySizes(lTemp) = iSize
          End If
                
          Select Case tSym2.bSubType
             Case CharVarType:
                iSize = 1
                lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "CHAR")
             Case ShortVarType:
                iSize = 2
                lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "SHORT")
             Case LongVarType:
                iSize = 4
                lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "LONG")
             Case PtrVarType:
                iSize = 4
                lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "POINTER")
             Case FloatVarType:
                iSize = 4
                lTemp = PrcVar(sTmpStr, iArraySize, 0, iSize, "FLOAT")
             Case DoubleVarType:
                iSize = 8
                lTemp = PrcVar(sTmpStr, iArraySize, 0, iSize, "DOUBLE")
             Case EnumVarType:
                Select Case agSwapBytes%(tSym2.iBitSize)
                   Case 32:
                      iSize = 4
                   Case 16:
                      iSize = 2
                   Case 8:
                      iSize = 1
                   Case Else:
                      iSize = 4
                End Select
                lTemp = PrcVar(sTmpStr, iArraySize, 0, iSize, "ENUM")
             Case StructVarType, LLVarType, TimeVarType, FreqVarType, Bam16VarType:
                If (iArraySize = UndefinedAryType) Then
                   iMultSize = 0
                   iSize = PrcStruct(sTmpStr, agSwapWords&(tSym2.lFldSymIdx) + 1, tSym2.bSubType)
                Else
                   iMultSize = iArraySize
                   'For ldx = 0 To iPrintSize - 1
                   For ldx = 0 To 0
                      sTmpStr2 = sTmpStr & "[" & Str(ldx) & "]"
                      iSize = PrcStruct(sTmpStr2, agSwapWords&(tSym2.lFldSymIdx) + 1, tSym2.bSubType)
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
             Case UnionVarType:
                If iArraySize = UndefinedAryType Then
                   iSize = PrcUnion(sTmpStr, agSwapWords&(tSym2.lFldSymIdx) + 1)
                Else
                   'For ldx = 0 To iPrintSize - 1
                   For ldx = 0 To 0
                      sTmpStr2 = sTmpStr & "[" & Str(ldx) & "]"
                      iSize = PrcUnion(sTmpStr2, agSwapWords&(tSym2.lFldSymIdx) + 1)
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
             Case Else:
                MsgBox ("Unknown Subtype found.  Ignored!")
          End Select
          If Not (iArraySize = UndefinedAryType) Then
             iSize = iSize * iArraySize
          End If
          If bLclAbort = True Then
             MsgBox ("Array size > limit")
          End If
          PrcMsgData = iSize
       End If
    Else
        MsgBox ("Invalid SymbolID found.  Ignored!")
    End If
End Function

'
'+v1.7.4SPV
' ROUTINE:  PrcMsgData2_41
' AUTHOR:   Shaun Vogel
' PURPOSE:  Loop through entire message and parse each field.
' INPUT:    sStr - string, with "." notation, being parsed.
'           lFldId2 - Index into message string for field being parsed.
' OUTPUT:   iSize - The size of the field being processed.
'
' NOTES:    This routine parses fields based on enumVarType2_41

Public Function PrcMsgData2_41(sStr As String, ByVal lFldId2 As Long) As Integer
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
   Dim lValue As Long
   Dim lFldId As Long
   Dim iFrmInt As Integer
   
   lFldId = lFldId2
      
   tField = tFieldType(lFldId)
   iArraySize = agSwapBytes%(tField.iArraySize)
   iPrintSize = iArraySize
   iSize = 0
   bLclAbort = False
   iFrmInt = lCurMsgId
   
   If (agSwapBytes(tField.iSymbolId) >= 0) Then
       If (agSwapBytes%(tField.iSymbolId) > agSwapWords&(tDebugFileHeader.lNumSymRecs)) Then
          MsgBox ("Found invalid symbol Id.  Record Ignored")
          bAbort = True
          PrcMsgData2_41 = 0
       Else
          bAbort = False
          tSym = tSymbolType(agSwapBytes%(tField.iSymbolId) + 1)
          tSym2 = tSym
          
          If sStr <> "" Then
             sFldLbl = GetString(agSwapWords&(tField.lStrIdx) + 1)
             sTmpStr = sStr & "." & sFldLbl
          Else
             sTmpStr = GetString(agSwapWords&(tField.lStrIdx) + 1)
             If (Left(sTmpStr, 2) <> "mt") Then
                Exit Function
             End If
             
             dbMsgTbl.AddNew
             dbMsgTbl!Msg_id = lCurMsgId
             dbMsgTbl!Msg_Name = sTmpStr
             If frmWizard.IsSelected(iFrmInt) Then dbMsgTbl!Select_Msg = True
             dbMsgTbl.Update
    
          End If
          
          GetArraySize iArraySize, agSwapBytes%(tSym2.iArraySize)
          
          While (tSym2.bSubType = TypeDefVarType2_41)
             tSym2 = tSymbolType(agSwapWords&(tSym.lFldSymIdx) + 1)
             GetArraySize iArraySize, agSwapBytes%(tSym2.iArraySize)
          Wend
             
          If Not (iArraySize = UndefinedAryType) Then
                   iPrintSize = iArraySize
          ElseIf (agSwapBytes%(tField.iArrayKey) > 0) Then
             Select Case tSym2.bSubType
                Case CharVarType2_41:
                   iSize = 1
                Case ShortVarType2_41:
                   iSize = 2
                Case LongVarType2_41:
                   iSize = 4
                Case LongLongVarType2_41
                   iSize = 8
                Case Else:
                   iSize = 1
             End Select
             If (iSize > MsgDataLimit) Then
                bLclAbort = True
                iSize = 1
             End If
             lTemp = agSwapBytes%(tField.iArrayKey)
             MsgArraySizes(lTemp) = iSize
          End If
                
          Select Case tSym2.bSubType
             Case CharVarType2_41:
                iSize = 1
                lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "CHAR")
             Case ShortVarType2_41:
                iSize = 2
                lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "SHORT")
             Case LongVarType2_41:
                iSize = 4
                lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "LONG")
             Case LongLongVarType2_41:
                iSize = 8
                lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "LONGLONG")
             Case PtrVarType2_41:
                iSize = 4
                lTemp = PrcVar(sTmpStr, iPrintSize, tSym2.bUnsigned, iSize, "POINTER")
             Case FloatVarType2_41:
                iSize = 4
                lTemp = PrcVar(sTmpStr, iArraySize, 0, iSize, "FLOAT")
             Case DoubleVarType2_41:
                iSize = 8
                lTemp = PrcVar(sTmpStr, iArraySize, 0, iSize, "DOUBLE")
             Case EnumVarType2_41:
                Select Case agSwapBytes%(tSym2.iBitSize)
                   Case 32:
                      iSize = 4
                   Case 16:
                      iSize = 2
                   Case 8:
                      iSize = 1
                   Case Else:
                      iSize = 4
                End Select
                lTemp = PrcVar(sTmpStr, iArraySize, 0, iSize, "ENUM")
             Case StructVarType2_41, LLVarType2_41, TimeVarType2_41, FreqVarType2_41, Bam16VarType2_41:
                If (iArraySize = UndefinedAryType) Then
                   iMultSize = 0
                   iSize = PrcStruct2_41(sTmpStr, agSwapWords&(tSym2.lFldSymIdx) + 1, tSym2.bSubType)
                Else
                   iMultSize = iArraySize
                   'For ldx = 0 To iPrintSize - 1
                   For ldx = 0 To 0
                      sTmpStr2 = sTmpStr & "[" & Str(ldx) & "]"
                      iSize = PrcStruct2_41(sTmpStr2, agSwapWords&(tSym2.lFldSymIdx) + 1, tSym2.bSubType)
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
             Case UnionVarType2_41:
                If iArraySize = UndefinedAryType Then
                   iSize = PrcUnion(sTmpStr, agSwapWords&(tSym2.lFldSymIdx) + 1)
                Else
                   'For ldx = 0 To iPrintSize - 1
                   For ldx = 0 To 0
                      sTmpStr2 = sTmpStr & "[" & Str(ldx) & "]"
                      iSize = PrcUnion(sTmpStr2, agSwapWords&(tSym2.lFldSymIdx) + 1)
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
             Case Else:
                MsgBox ("Unknown Subtype found.  Ignored!")
          End Select
          If Not (iArraySize = UndefinedAryType) Then
             iSize = iSize * iArraySize
          End If
          If bLclAbort = True Then
             MsgBox ("Array size > limit")
          End If
          PrcMsgData2_41 = iSize
       End If
    End If
End Function

'
'+v1.7BB
' ROUTINE:  PrcUnion
' AUTHOR:   Shaun Vogel
' PURPOSE:  Processes a union struct by recursively calling PrcMsgData.
' INPUT:    sStr - string, with "." notation, being parsed.
'           lFldId2 - Index into message string for field being parsed.
' OUTPUT:   iSize - size of union
'
' NOTES:

Static Function PrcUnion(sStr As String, ByVal lFldId As Long) As Integer
   Dim i As Integer
   Dim tField As FieldType
   Dim iSum As Integer
   Dim iSize As Integer
    
   iSize = 0
   'i = iFldId
    
   Do
      tField = tFieldType(lFldId)
      If agSwapWords&(tField.lStrIdx) <> 0 Then
         '+v1.7.4SPV
         If msngVersion >= 2.41 Then
            iSum = PrcMsgData2_41(sStr, lFldId)
         Else
            iSum = PrcMsgData(sStr, lFldId)
         End If
         '-v1.7.4

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
   PrcUnion = iSize
End Function

'
'+v1.7BB
' ROUTINE:  PrcStruct
' AUTHOR:   Shaun Vogel
' PURPOSE:  Recursively parses each structure within a message and writes it to
'           the VarStructTbl.
' INPUT:    sStr - string containing structure name with "." notation.
'           lFldIdx - Index into the field definition structure.
'           eSubType - structure variable type.
' OUTPUT:   iSize - The size of the structure being processed.
'
' NOTES:    The types of structures that are processed:
'           Bam16, LatLon, Frequency, and Time.

Static Function PrcStruct(sStr As String, ByVal lFldIdx As Long, eSubType As Byte) As Integer
   Dim iSize As Integer
   Dim iArraySize As Integer
   Dim tField As FieldType
   Dim i As Integer
   Dim iSum As Integer
   
   iSize = 0
   
   If (MsgRawMode = True) Then
      eSubType = StructVarType
   End If
   
   Select Case eSubType
      Case Bam16VarType:
         iSize = BAM16_SIZE
         'get new var_struct_table rec
         lVarStrId = lVarStrId + 1
         dbVarStrTbl.AddNew
   
         'save msg_id, field_name, field_size, data_type, variable_len
         dbVarStrTbl!varStructID = lVarStrId
         dbVarStrTbl!msgid = lCurMsgId
         dbVarStrTbl!fieldname = sStr              'string with . notation
         dbVarStrTbl!FieldSize = iSize
         dbVarStrTbl!fieldlabel = sFldLbl
         dbVarStrTbl!ConvType = 0
         dbVarStrTbl!DasField = 0
         dbVarStrTbl!MultiRecPtr = 0
         dbVarStrTbl!DataType = "BAM16"
         
         'update record set
         'incRecPtr
         dbVarStrTbl.Update
      Case FreqVarType:
         iSize = FREQ_SIZE
         'get new var_struct_table rec
         lVarStrId = lVarStrId + 1
         dbVarStrTbl.AddNew
   
         'save msg_id, field_name, field_size, data_type, variable_len
         dbVarStrTbl!varStructID = lVarStrId
         dbVarStrTbl!msgid = lCurMsgId
         dbVarStrTbl!fieldname = sStr              'string with . notation
         dbVarStrTbl!FieldSize = iSize
         dbVarStrTbl!fieldlabel = sFldLbl
         dbVarStrTbl!DataType = "FREQ"
         dbVarStrTbl!ConvType = 0
         dbVarStrTbl!DasField = 0
         dbVarStrTbl!MultiRecPtr = 0
         
         'update record set
         'incRecPtr
         dbVarStrTbl.Update
      Case TimeVarType:
         iSize = 0
         eSubType = StructVarType
      Case LLVarType:
         iSize = 0
         eSubType = StructVarType
   End Select
   
   If eSubType = StructVarType Then
      Dim myexit As Boolean
      myexit = False
      iStructLevel = iStructLevel + 1
      tField = tFieldType(lFldIdx)
     
      'get new var_struct_table rec
      lVarStrId = lVarStrId + 1
      dbVarStrTbl.AddNew
     
      'save STRUCT info
      dbVarStrTbl!varStructID = lVarStrId
      dbVarStrTbl!msgid = lCurMsgId
      dbVarStrTbl!fieldname = sStr              'string with . notation
      'dbVarStrTbl!FieldSize = iTmpSize
      dbVarStrTbl!DataType = "STRUCT BEGIN"
      dbVarStrTbl!StructLevel = iStructLevel
      If (iMultSize <> 0) Then
         dbVarStrTbl!MultiEntry = iMultSize
      End If
      dbVarStrTbl!ConvType = 0
      dbVarStrTbl!DasField = 0
      dbVarStrTbl!MultiRecPtr = 0
      
      'update record set
      'incRecPtr
      dbVarStrTbl.Update
   
      Do
         If agSwapWords&(tField.lStrIdx) <> 0 Then
            iSum = PrcMsgData(sStr, lFldIdx)

            lFldIdx = lFldIdx + 1
            tField = tFieldType(lFldIdx)
            If (bAbort = True) Then
               iSize = 0
               myexit = True
            Else
               iSize = iSize + iSum
            End If
         Else
            myexit = True
         End If
      Loop Until (agSwapWords&(tField.lStrIdx) = 0)
      
      'get new var_struct_table rec
      lVarStrId = lVarStrId + 1
      dbVarStrTbl.AddNew
     
      'save STRUCT END info
      dbVarStrTbl!varStructID = lVarStrId
      dbVarStrTbl!msgid = lCurMsgId
      dbVarStrTbl!DataType = "STRUCT END"
      dbVarStrTbl!StructLevel = iStructLevel
      dbVarStrTbl!ConvType = 0
      dbVarStrTbl!DasField = 0
      dbVarStrTbl!MultiRecPtr = 0
      
      'update record set
      dbVarStrTbl.Update
      iStructLevel = iStructLevel - 1
   End If
   PrcStruct = iSize
End Function

'
'+v1.7.4
' ROUTINE:  PrcStruct2_41
' AUTHOR:   Shaun Vogel
' PURPOSE:  Recursively parses each structure within a message and writes it to
'           the VarStructTbl.
' INPUT:    sStr - string containing structure name with "." notation.
'           lFldIdx - Index into the field definition structure.
'           eSubType - structure variable type.
' OUTPUT:   iSize - The size of the structure being processed.
'
' NOTES:    The types of structures that are processed:
'           Bam16, LatLon, Frequency, and Time.
'           These types are defined in enumVarTypes2_41

Static Function PrcStruct2_41(sStr As String, ByVal lFldIdx As Long, eSubType As Byte) As Integer
   Dim iSize As Integer
   Dim iArraySize As Integer
   Dim tField As FieldType
   Dim i As Integer
   Dim iSum As Integer
   
   iSize = 0
   
   If (MsgRawMode = True) Then
      eSubType = StructVarType2_41
   End If
   
   Select Case eSubType
      Case Bam16VarType2_41:
         iSize = BAM16_SIZE
         'get new var_struct_table rec
         lVarStrId = lVarStrId + 1
         dbVarStrTbl.AddNew
   
         'save msg_id, field_name, field_size, data_type, variable_len
         dbVarStrTbl!varStructID = lVarStrId
         dbVarStrTbl!msgid = lCurMsgId
         dbVarStrTbl!fieldname = sStr              'string with . notation
         dbVarStrTbl!FieldSize = iSize
         dbVarStrTbl!fieldlabel = sFldLbl
         dbVarStrTbl!ConvType = 0
         dbVarStrTbl!DasField = 0
         dbVarStrTbl!MultiRecPtr = 0
         dbVarStrTbl!DataType = "BAM16"
         
         'update record set
         'incRecPtr
         dbVarStrTbl.Update
      Case FreqVarType2_41:
         iSize = FREQ_SIZE
         'get new var_struct_table rec
         lVarStrId = lVarStrId + 1
         dbVarStrTbl.AddNew
   
         'save msg_id, field_name, field_size, data_type, variable_len
         dbVarStrTbl!varStructID = lVarStrId
         dbVarStrTbl!msgid = lCurMsgId
         dbVarStrTbl!fieldname = sStr              'string with . notation
         dbVarStrTbl!FieldSize = iSize
         dbVarStrTbl!fieldlabel = sFldLbl
         dbVarStrTbl!DataType = "FREQ"
         dbVarStrTbl!ConvType = 0
         dbVarStrTbl!DasField = 0
         dbVarStrTbl!MultiRecPtr = 0
         
         'update record set
         'incRecPtr
         dbVarStrTbl.Update
      Case TimeVarType2_41:
         iSize = 0
         eSubType = StructVarType2_41
      Case LLVarType2_41:
         iSize = 0
         eSubType = StructVarType2_41
   End Select
   
   If eSubType = StructVarType2_41 Then
      Dim myexit As Boolean
      myexit = False
      iStructLevel = iStructLevel + 1
      tField = tFieldType(lFldIdx)
     
      'get new var_struct_table rec
      lVarStrId = lVarStrId + 1
      dbVarStrTbl.AddNew
     
      'save STRUCT info
      dbVarStrTbl!varStructID = lVarStrId
      dbVarStrTbl!msgid = lCurMsgId
      dbVarStrTbl!fieldname = sStr              'string with . notation
      'dbVarStrTbl!FieldSize = iTmpSize
      dbVarStrTbl!DataType = "STRUCT BEGIN"
      dbVarStrTbl!StructLevel = iStructLevel
      If (iMultSize <> 0) Then
         dbVarStrTbl!MultiEntry = iMultSize
      End If
      dbVarStrTbl!ConvType = 0
      dbVarStrTbl!DasField = 0
      dbVarStrTbl!MultiRecPtr = 0
      
      'update record set
      'incRecPtr
      dbVarStrTbl.Update
   
      Do
         If agSwapWords&(tField.lStrIdx) <> 0 Then
            iSum = PrcMsgData2_41(sStr, lFldIdx)
            lFldIdx = lFldIdx + 1
            tField = tFieldType(lFldIdx)
            If (bAbort = True) Then
               iSize = 0
               myexit = True
            Else
               iSize = iSize + iSum
            End If
         Else
            myexit = True
         End If
      Loop Until (agSwapWords&(tField.lStrIdx) = 0)
      
      'get new var_struct_table rec
      lVarStrId = lVarStrId + 1
      dbVarStrTbl.AddNew
     
      'save STRUCT END info
      dbVarStrTbl!varStructID = lVarStrId
      dbVarStrTbl!msgid = lCurMsgId
      dbVarStrTbl!DataType = "STRUCT END"
      dbVarStrTbl!StructLevel = iStructLevel
      dbVarStrTbl!ConvType = 0
      dbVarStrTbl!DasField = 0
      dbVarStrTbl!MultiRecPtr = 0
      
      'update record set
      dbVarStrTbl.Update
      iStructLevel = iStructLevel - 1
   End If
   PrcStruct2_41 = iSize
End Function

'
'+v1.7BB
' ROUTINE:  PrcVar
' AUTHOR:   Shaun Vogel
' PURPOSE:  Write field with attributes to the VarStructTbl
' INPUT:    sStr - string containing structure name with "." notation.
'           iArraySize - size of array or 0 if not an array.
'           bUnsigned - True if field is unsigned else 0.
'           iSize - size of current field being processed.
' OUTPUT:   iTmpSize - size of field
'
' NOTES:    TOC tables contain results from processing an archive.

Static Function PrcVar(sStr As String, iArraySize As Integer, bUnsigned As Byte, iSize As Integer, sVarType As String) As Integer
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
         dbVarStrTbl.AddNew
   
         'save msg_id, field_name, field_size, data_type, variable_len
         dbVarStrTbl!varStructID = lVarStrId
         dbVarStrTbl!msgid = lCurMsgId
         dbVarStrTbl!fieldname = sStr              'string with . notation
         dbVarStrTbl!FieldSize = iTmpSize
         dbVarStrTbl!fieldlabel = sFldLbl
         dbVarStrTbl!DataType = sVarType
         dbVarStrTbl!MultiEntry = iArraySize
         dbVarStrTbl!ConvType = 0
         dbVarStrTbl!DasField = 0
         dbVarStrTbl!MultiRecPtr = 0
         
         'update record set
         'incRecPtr
         dbVarStrTbl.Update
      End If
   Else
      'get new var_struct_table rec
      lVarStrId = lVarStrId + 1
      dbVarStrTbl.AddNew

      'save msg_id, field_name, field_size, data_type, variable_len
      dbVarStrTbl!varStructID = lVarStrId
      dbVarStrTbl!msgid = lCurMsgId
      dbVarStrTbl!fieldname = sStr              'string with . notation
      dbVarStrTbl!FieldSize = iTmpSize
      dbVarStrTbl!fieldlabel = sFldLbl
      dbVarStrTbl!DataType = sVarType
      dbVarStrTbl!ConvType = 0
      dbVarStrTbl!DasField = 0
      dbVarStrTbl!MultiRecPtr = 0
      
      'update record set
      'incRecPtr
      dbVarStrTbl.Update
   End If
   PrcVar = iTmpSize
End Function

'
'+v1.7BB
' ROUTINE:  readHeader
' AUTHOR:   Shaun Vogel
' PURPOSE:  Read the archive.dat file, parse the fields in each message, and store the
'           message structure into the VarStructTbl.
' INPUT:    fname - Name of file to parse
' OUTPUT:   none
' NOTES:    Stores message structures in the the database as the VarStructTbl.

Sub readHeader(fname As String)
   Dim tMsgEntries As MsgEntryTable
   Dim lTmpSize, lTmp, fTmp As Long
   Dim bBucket() As Byte
   Dim lSwapNumMsgs As Long
   Dim iSwapEntries As Integer
   Dim btmpMsgString() As Byte
   Dim lswap1 As Long
   Dim lswap2 As Long
   Dim lswap3 As Long
   
   iCurMsgId = 0
   iVarStrId = 0
   
   On Error GoTo FileError
   Open fname For Binary As #1
   
   Get #1, , tMsgEntries
   
   
   lTmpSize = agSwapWords&(tMsgEntries.tMsgDskEntry(MsgTableRec).lSize) + _
              agSwapWords&(tMsgEntries.tMsgDskEntry(LmbSigRec).lSize) + _
              agSwapWords&(tMsgEntries.tMsgDskEntry(HbSigRec).lSize)
              
   'Skip prep stuff  seek
   ReDim bBucket(1 To lTmpSize)
   Get #1, , bBucket
   
   'Read debug file header info
   Get #1, , tDebugFileHeader
   
   'Read string info
   lTmpSize = agSwapWords&(tDebugFileHeader.lStringSize)
   
   ReDim bMsgStrings(1 To lTmpSize)
   Get #1, , bMsgStrings
   
   'Read Msg Id's and Field Id index's
   lSwapNumMsgs = agSwapWords&(tDebugFileHeader.lNumMsgs)
   'iTmpSize = tDebugFileHeader.lNumMsgs
   ReDim tMtId2Type(1 To lSwapNumMsgs)
   Get #1, , tMtId2Type
   
   'Calculate number of valid messages if file
   lMaxMsgId = 0
   For lTmp = 1 To lSwapNumMsgs
      If agSwapBytes%(tMtId2Type(lTmp).iMsgId) > lMaxMsgId Then
         lMaxMsgId = agSwapBytes%(tMtId2Type(lTmp).iMsgId)
      End If
   Next lTmp
   
   'Initialize msg id's
   ReDim lMsgId2Idx(1 To lMaxMsgId)
   For lTmp = 1 To lMaxMsgId
      lMsgId2Idx(lTmp) = BAD_MSG_IDX
   Next lTmp
   
   'Store msg id's
   For lTmp = 1 To lSwapNumMsgs
      lMsgId2Idx(agSwapBytes%(tMtId2Type(lTmp).iMsgId)) = lTmp
   Next lTmp
      
   'Read all Symbol Records
   lTmpSize = agSwapWords&(tDebugFileHeader.lNumSymRecs)
   ReDim tSymbolType(1 To lTmpSize)
   Get #1, , tSymbolType
   
   'Read all Field Records
   lTmpSize = agSwapWords&(tDebugFileHeader.lNumFldRecs)
   ReDim tFieldType(1 To lTmpSize)
   Get #1, , tFieldType
   Close #1
   
   'Open the Database and set the record set to the var_struct_table
   Set dbMsgTbl = guCurrent.uArchive.rsMessage
   Set dbVarStrTbl = guCurrent.uArchive.rsVarStruct
   
   If Not dbVarStrTbl Is Nothing Then
        'Loop through Max Message Id's and store data in database
        For lTmp = 1 To lSwapNumMsgs
           lTmp = agSwapBytes%(tMtId2Type(lTmp).iMsgId)
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
              fTmp = agSwapBytes%(tMtId2Type(lTmp).iFieldId) + 1
              '+v1.7.4SPV
              If msngVersion >= 2.41 Then
                 lTmpSize = PrcMsgData2_41("", fTmp)
              Else
                 lTmpSize = PrcMsgData("", fTmp)
              End If
              '-v1.7.4
           End If
        Next lTmp
        
        dbVarStrTbl.Close
        
   End If
   Set dbVarStrTbl = Nothing
   Exit Sub
    
FileError:
   Debug.Print Err.Description
   Debug.Print Err.Number
   MsgBox "Error opening header file " & fname
    
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
