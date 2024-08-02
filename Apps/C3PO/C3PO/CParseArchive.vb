Imports System.IO
Imports system.text

Public Class CParseArchive
    Private mSwapBytes As New CSwapBytes
    Public Structure Summary_Record
        Dim iMsgCount As Long
        Dim dTimeFirst As Double
        Dim dTimeLast As Double
    End Structure 'Summary_Record

    Public Structure Toc_Record
        Dim iMsgId As Integer
        Dim lMsgCount As Long
        Dim dTimeStamp As Double
        Dim iStartByte As Long
        Dim iMsgSize As Integer
    End Structure 'Toc_Record

    Public Enum File_Action
        Binary_Read = 0
        Binary_Write = 1
        Text_Read = 2
        Text_Write = 3
    End Enum
    Public Structure Arc_Hdr                         '   4 bytes Fixed
        Dim lTimestamp As Long                      '   4 bytes   0 -   3
    End Structure 'Arc_Hdr
    '
    ' Message Header
    Public Structure Msg_Hdr                         '   8 bytes Fixed
        Dim iMsgLength As Integer                   '   2 bytes   0 -   1
        Dim iMsgId As Integer                       '   2 bytes   2 -   3
        Dim iMsgTo As Integer                       '   2 bytes   4 -   5
        Dim iMsgFrom As Integer                     '   2 bytes   6 -   7
    End Structure 'Msg_Hdr

    ' Archive header

    Public giRawInputFile As Integer
    Public giRawOutputFile As Integer
    Public Const sArchiveExt = ".cca"
    Public Const sFilteredExt = ".flt"


    Public iJdate As Integer
    Public dTimeZero As Double

    Public Const iMsgSizeMax = 4096   'temporary size allocation for all CC messages
    Public Const iMsgIdMax = 4096   'temporary size allocation for all CC messages
    Public Const iTocStep = 100     ' this will be redimensioned later
    Public iTocCount As Long     ' Count of number of messages in TOC
    Public iTocStart As Long     ' Starting location in filtered file
    Public iTocHiWater As Long   ' Point that we need to add more entries
    Public gaSummary(iMsgIdMax) As Summary_Record


  
    '
    '+v1.6TE
    ' FUNCTION: blnProcessRawFile
    ' AUTHOR:   Tom Elkins
    ' PURPOSE:  Copy of Proc_Raw_File -- returns boolean status
    ' INPUT:    none
    ' OUTPUT:   TRUE = Processing complete
    '           FALSE = Processing failed
    ' NOTES:
    Public Function blnProcessRawFile(ByVal strRawFile As String) As Boolean
        '        Dim tmpArcHdr As Arc_Hdr
        '        Dim tmpMsgHdr As Msg_Hdr
        '        Dim tmpByteCount As Integer
        '        Dim tmpMsgData() As Byte
        '        Dim lSwapTimestamp As Long
        '        Dim iSwapMsgId As Integer
        '        Dim iSwapMsgSize As Integer
        '        Const lBlocksize As Long = 32768
        '        Dim lBytesRemain As Long
        '        Dim iMinSize As Integer
        '        Dim dTimeStamp As Double
        '        Dim uTOCMsg As Toc_Record
        '        Dim inputfile As New FileStream(strRawFile, FileMode.Open, FileAccess.Read)
        '        '
        '        '

        '        On Error GoTo Hell
        '        blnProcessRawFile = False
        '        '
        '        lBytesRemain = lBlocksize
        '        iMinSize = (LenB(tmpArcHdr) + LenB(tmpMsgHdr))
        '        iTocStart = 1
        '        iTocCount = 0
        '        While Not EOF(giRawInputFile)
        '            'code sample
        '            Dim count As Integer = 1024
        '            Dim buffer(count - 1) As Byte
        '            count = inputfile.Read(buffer, 0, count)
        '            'code sample
        '            If (lBytesRemain < iMinSize) Then
        '                If (lBytesRemain > 0) Then
        '                    ReDim tmpMsgData(lBytesRemain)
        '                    getDataBytes(inputfile, lBytesRemain, tmpMsgData)        ' need to read it even if we don't write it
        '                    iTocStart = iTocStart + lBytesRemain
        '                End If
        '                lBytesRemain = lBlocksize
        '                '
        '                ' Take advantage of this interruption to update the form
        '                '
        '            End If

        '            getDataBytes(inputfile, lBytesRemain, tmpMsgData)
        '            getDataBytes(inputfile, lBytesRemain, tmpMsgData)


        '            'getDataBytes(giRawInputFile, , tmpArcHdr)
        '            'Get giRawInputFile, , tmpMsgHdr
        '            iTocStart = iTocStart + LenB(tmpArcHdr)

        '            '
        '            On Error Resume Next
        '            '
        '            lBytesRemain = lBytesRemain - iMinSize
        '            lSwapTimestamp = SwapBytes(tmpArcHdr.lTimestamp)
        '            iSwapMsgId = SwapBytes(tmpMsgHdr.iMsgId, 0, 2)
        '            iSwapMsgSize = SwapBytes(tmpMsgHdr, 0, 2)

        '            If (iSwapMsgId = 359) Then 'MTARCFILLERID
        '                If (lBytesRemain > 0) Then
        '                    ReDim tmpMsgData(0 To lBytesRemain - 1)
        '                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
        '                    '               iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
        '                End If
        '                iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
        '                lBytesRemain = lBlocksize
        '            ElseIf ((iSwapMsgId < 1) Or (iSwapMsgId > iMsgIdMax)) Then
        '                If (lBytesRemain > 0) Then
        '                    ReDim tmpMsgData(0 To lBytesRemain - 1)
        '                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
        '                    '               iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
        '                End If
        '                iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
        '                lBytesRemain = lBlocksize
        '            ElseIf ((lSwapTimestamp <= 0) Or (lSwapTimestamp > 864000.0#)) Then
        '                If (lBytesRemain > 0) Then
        '                    ReDim tmpMsgData(0 To lBytesRemain - 1)
        '                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
        '                    '               iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
        '                End If
        '                iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
        '                lBytesRemain = lBlocksize
        '            ElseIf ((iSwapMsgSize <= 8) Or (iSwapMsgSize > iMsgSizeMax)) Then
        '                If (lBytesRemain > 0) Then
        '                    ReDim tmpMsgData(0 To lBytesRemain - 1)
        '                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
        '                    '               iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
        '                End If
        '                iTocStart = iTocStart + lBytesRemain + LenB(tmpMsgHdr)
        '                lBytesRemain = lBlocksize
        '            Else
        '                tmpByteCount = iSwapMsgSize - 8
        '                If (tmpByteCount > 0) Then
        '                    'Dimension array to data size
        '                    ReDim tmpMsgData(0 To tmpByteCount - 1)
        '                Get giRawInputFile, , tmpMsgData()        ' need to read it even if we don't write it
        '                    lBytesRemain = lBytesRemain - tmpByteCount

        '                    dTimeStamp = lSwapTimestamp / 10.0#
        '                    iTocCount = iTocCount + 1               'v1.7B
        '                    uTOCMsg.dTimeStamp = dTimeStamp         'v1.7B
        '                    uTOCMsg.lMsgCount = iTocCount
        '                    uTOCMsg.iMsgId = iSwapMsgId             'v1.7B
        '                    uTOCMsg.iMsgSize = iSwapMsgSize         'v1.7B
        '                    uTOCMsg.iStartByte = iTocStart          'v1.7B
        '                    iTocStart = iTocStart + iSwapMsgSize    'v1.7B

        '                    Encoding.UTF8.GetByteCount(

        '                    Call basTOC.Add_TOC_Record(uTOCMsg)
        '                    ' is it a message that we want
        '                    'If (gaSummary(iSwapMsgId).iMsgCount >= 0) Then
        '                    '    Put(giRawOutputFile, , tmpArcHdr)
        '                    '    Put(giRawOutputFile, , tmpMsgHdr)
        '                    '    Put(giRawOutputFile, , tmpMsgData())
        '                    '    ' is this the first occurance of this message
        '                    '    If (gaSummary(iSwapMsgId).iMsgCount = 0) Then
        '                    '        gaSummary(iSwapMsgId).dTimeFirst = dTimeStamp
        '                    '    End If
        '                    '    ' set last occurance time
        '                    '    gaSummary(iSwapMsgId).dTimeLast = dTimeStamp
        '                    '    'Increment message count
        '                    '    gaSummary(iSwapMsgId).iMsgCount = gaSummary(iSwapMsgId).iMsgCount + 1

        '                    'End If ' message that we want
        '                End If ' byte count > 0
        '            End If 'multiple check
        '        End While
        '        blnProcessRawFile = True
        '        '
        '        Exit Function
        '        '
        '        '
        'Hell:
        '        '
        '        '
    End Function

    Private Sub getDataBytes(ByVal fsSource As FileStream, ByVal byteLength As Integer, ByRef bytes() As Byte)

        Try
            ' Read the source file into a byte array.
            Dim numBytesToRead As Integer = CType(byteLength, Integer)
            Dim numBytesRead As Integer = 0

            While (numBytesToRead > 0)
                ' Read may return anything from 0 to numBytesToRead.
                Dim n As Integer = fsSource.Read(bytes, numBytesRead, _
                    numBytesToRead)
                ' Break when the end of the file is reached.
                If (n = 0) Then
                    Exit While
                End If
                numBytesRead = (numBytesRead + n)
                numBytesToRead = (numBytesToRead - n)

            End While

        Catch ioEx As FileNotFoundException
            Console.WriteLine(ioEx.Message)
        End Try
    End Sub

    Private Sub SwapBytes(ByRef srcData() As Byte, ByRef iDataPtr As Integer, ByVal numBytes As Integer)

        If numBytes = 2 Then
            mSwapBytes.swapBytes(srcData, iDataPtr)
        ElseIf numBytes = 4 Then
            mSwapBytes.swapWord(srcData, iDataPtr)
        ElseIf numBytes = 8 Then
            mSwapBytes.swap8(srcData, iDataPtr)
        End If
    End Sub
    Public Function LenB(ByVal ObjStr As String) As Integer

        'Note that ObjStr.Length will fail if ObjStr was set to Nothing    

        If Len(ObjStr) = 0 Then Return 0

        Return System.Text.Encoding.Unicode.GetByteCount(ObjStr)

    End Function



    Public Function LenB(ByVal Obj As Object) As Integer

        If Obj Is Nothing Then Return 0

        Try 'Structure
            Return Len(Obj)

        Catch 'Leave blank for catch-all

            Try 'Type-def objects
                Return System.Runtime.InteropServices.Marshal.SizeOf(Obj)

            Catch 'Leave blank for catch-all

                Return -1 'Allow user to check for <0 as error
            End Try

        End Try

    End Function


End Class
