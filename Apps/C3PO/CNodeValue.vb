Imports System.Xml

Public Class CNodeValue

    Private mName As String = ""
    Private mRootNode As Xml.XmlNode
    Private mBinary As Boolean = True
    Private mData As String = ""
    Private mEvaluate As String = ""
    Private mSize As String = "1"
    Private mReaduntil As String = ""
    Private mStopafter As String = ""
    Private mAdvance As Boolean = True
    Private mSwapped As Boolean = False
    Private mMask As String = ""
    Private mType As String = ""
    Private mSigned As Boolean = False
    Private mScale As String = ""
    Private mOffset As String = ""
    Private mPreserve As Boolean = False
    Private mOutput As Boolean = True

    Private mSwapBytes As New CSwapBytes


    Public Sub New(ByVal xnode As Xml.XmlNode)
        Dim attList As Xml.XmlAttributeCollection
        Dim att As Xml.XmlAttribute

        mRootNode = xnode
        attList = xnode.Attributes
        Try
            For Each att In attList
                Select Case att.Name.ToLower
                    Case "binary"
                        mBinary = att.Value

                    Case "data"
                        mData = att.Value
                        'parse string

                    Case "evaluate"
                        mEvaluate = att.Value
                        'parse string

                    Case "size"
                        mSize = att.Value
                        'parse string
                        mSize = EvaluateStr(att.Value)

                    Case "readuntil"
                        mReaduntil = EvaluateStr(att.Value)
                        'parse string

                    Case "stopafter"
                        mStopafter = EvaluateStr(att.Value)

                    Case "advance"
                        mAdvance = EvaluateStr(att.Value)

                    Case "swapped"
                        mSwapped = att.Value

                    Case "mask"
                        mMask = att.Value

                    Case "type"
                        mType = att.Value

                    Case "signed"
                        mSigned = att.Value

                    Case "scale"
                        mScale = EvaluateStr(att.Value)

                    Case "offset"
                        mOffset = EvaluateStr(att.Value)

                    Case "preserve"
                        mPreserve = att.Value

                    Case "output"
                        mOutput = att.Value

                    Case "name"
                        mName = att.Value

                End Select
            Next
        Catch ex As Exception

        End Try
    End Sub

    Public Sub parseValue(ByRef srcData() As Byte, ByRef iDataPtr As Integer)

        ParseNode(srcData, iDataPtr, Me.mRootNode)
        MGlobals.DoneOutput()

    End Sub

    Private Sub ParseNode(ByRef srcData() As Byte, ByRef iDataPtr As Integer, ByVal xNode As XmlNode)
        Dim val As String = ""

        If mBinary Then
            If mSwapped Then
                'swap data
                Try
                    SwapBytes(srcData, iDataPtr, mSize)
                Catch ex As Exception
                    'unable to swap bytes
                End Try
            End If
            If mMask <> "" Then

            End If
            Select Case mType
                Case "raw"
                    For i As Integer = 0 To mSize - 1
                        val = val & srcData(iDataPtr + i)
                    Next
                Case "string"
                    val = BitConverter.ToString(srcData, iDataPtr, mSize)

                Case "hex"
                    val = HandleHex(srcData, iDataPtr)

                Case "bcd"

                Case "integer"
                    val = HandleInt(srcData, iDataPtr)

                Case "float"
                    If mSize = "4" Then
                        val = BitConverter.ToSingle(srcData, iDataPtr)
                    ElseIf mSize = "8" Then
                        val = BitConverter.ToDouble(srcData, iDataPtr)
                    Else
                        MsgBox("Unable to convert float. Size should be 4 or 8 bytes.")
                    End If

                Case "boolean"
                    Dim bool As Integer
                    If mSize = "1" Then
                        bool = srcData(iDataPtr)
                    ElseIf mSize = "2" Then
                        bool = BitConverter.ToInt16(srcData, iDataPtr)
                    Else
                        MsgBox("Unable to convert boolean. Size should be 1 or 2 bytes.")
                        Exit Select
                    End If
                    If bool = 0 Then
                        val = "false"
                    Else
                        val = "true"
                    End If

            End Select
            iDataPtr += mSize
        End If
        'place name/value in ary list
        gValueList.Replace(mName, val)
        'add to output doc
        If mOutput Then
            MGlobals.AddOutput(mName, val)
        End If
    End Sub

    Private Function HandleInt(ByRef srcData() As Byte, ByRef idataptr As Integer) As String
        Dim val As Int64 = 0

        'determine size of int
        If mSize = "2" Then
            If mSigned Then
                val = BitConverter.ToInt16(srcData, idataptr)
            Else
                val = BitConverter.ToUInt16(srcData, idataptr)
            End If
        ElseIf mSize = "4" Then
            If mSigned Then
                val = BitConverter.ToInt32(srcData, idataptr)
            Else
                val = BitConverter.ToUInt32(srcData, idataptr)
            End If
        ElseIf mSize = "8" Then
            If mSigned Then
                val = BitConverter.ToInt64(srcData, idataptr)
            Else
                val = BitConverter.ToUInt64(srcData, idataptr)
            End If
        Else
            MsgBox("Invalid integer size: " & mSize)
            Return ""
        End If

        If mScale <> "" Then

        End If
        If mOffset <> "" Then

        End If
        Return val.ToString
    End Function

    Private Function HandleHex(ByRef srcData() As Byte, ByRef iDataPtr As Integer) As String

        'determine size of int
        If mSize = "2" Then
            Dim val As Int16
            If mSigned Then
                Val = BitConverter.ToInt16(srcData, iDataPtr)
            Else
                Val = BitConverter.ToUInt16(srcData, iDataPtr)
            End If
            Return Hex(val)
        ElseIf mSize = "4" Then
            Dim val As Int32
            If mSigned Then
                Val = BitConverter.ToInt32(srcData, iDataPtr)
            Else
                Val = BitConverter.ToUInt32(srcData, iDataPtr)
            End If
            Return Hex(val)
        ElseIf mSize = "8" Then
            Dim val As Int64
            If mSigned Then
                Val = BitConverter.ToInt64(srcData, iDataPtr)
            Else
                Val = BitConverter.ToUInt64(srcData, iDataPtr)
            End If
            Return Hex(Val)
        Else
            MsgBox("Unable to convert Hex value.  Invalid hex size: " & mSize)
        End If
        Return ""
    End Function

    Private Sub SwapBytes(ByRef srcData() As Byte, ByRef iDataPtr As Integer, ByVal numBytes As Integer)

        If numBytes = 2 Then
            mSwapBytes.swapBytes(srcData, iDataPtr)
        ElseIf numBytes = 4 Then
            mSwapBytes.swapWord(srcData, iDataPtr)
        ElseIf numBytes = 8 Then
            mSwapBytes.swap8(srcData, iDataPtr)
        End If
    End Sub

    Private Function EvaluateStr(ByVal valStr As String) As String

        If valStr.StartsWith("{") And valStr.EndsWith("}") Then
            Dim tstr As String = ""
            tstr = valStr.Remove(0, 1)
            tstr = tstr.Remove(valStr.Length - 1, 1)
            If MGlobals.gValueList.Contains(tstr) Then
                Return MGlobals.gValueList(tstr)
            End If
        End If
        Return valStr
    End Function
End Class
