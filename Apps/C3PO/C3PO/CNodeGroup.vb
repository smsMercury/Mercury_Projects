Imports System.Xml


Public Class CNodeGroup

    Private mName As String = ""
    Private mRepeat As String = ""
    Private mForeach As String = ""
    Private mRootNode As Xml.XmlNode

    Public Sub New(ByVal xnode As Xml.XmlNode)
        Dim attList As Xml.XmlAttributeCollection
        Dim att As Xml.XmlAttribute

        mRootNode = xnode
        attList = xnode.Attributes
        For Each att In attList
            Select Case att.Name.ToLower
                Case "name"
                    aName = att.Value

                Case "repeat"
                    aRepeat = att.Value

                Case "foreach"
                    aForeach = att.Value

            End Select
        Next
    End Sub

    Public Sub parseGroup(ByRef srcData() As Byte, ByRef iDataPtr As Integer)
        Dim lvar As Integer = 0

        'output name if there
        If aName <> "" Then
            MGlobals.AddOutput(aName)
        End If
        'see if group is repeated
        If aRepeat <> "" Then
            'get repeat value
            Try
                lvar = aRepeat
            Catch ex As Exception
                'repeat value is not an integer so lets see if it is in the valueList
                Dim val As String = evaluateStr(aRepeat)
                If MGlobals.gValueList.Contains(val) Then
                    Try
                        lvar = MGlobals.gValueList(val)
                    Catch ex1 As Exception
                        MsgBox("Error retrieving value: " & aRepeat & vbCrLf & "Possible error in script." & vbCrLf & _
                                ex1.Message, MsgBoxStyle.OkOnly)
                        Exit Sub
                    End Try
                End If
            End Try
            If lvar > 0 Then
                For i As Integer = 1 To lvar
                    MGlobals.LoopNode(i)
                    '
                    'parse group children
                    ParseNode(srcData, iDataPtr, Me.mRootNode)
                    MGlobals.DoneOutput()
                Next
            ElseIf lvar < 0 Then
                Dim i As Integer = 1
                While iDataPtr < srcData.Length - 1
                    MGlobals.LoopNode(i)
                    '
                    'parse group children
                    ParseNode(srcData, iDataPtr, Me.mRootNode)
                    MGlobals.DoneOutput()
                    i += 1
                End While
            End If
        Else 'parse group children
            ParseNode(srcData, iDataPtr, Me.mRootNode)
        End If
        MGlobals.DoneOutput()

    End Sub

    Private Sub ParseNode(ByRef srcData() As Byte, ByRef iDataPtr As Integer, ByVal xNode As XmlNode)

        For Each cnode As XmlNode In xNode.ChildNodes
            Select Case cnode.Name.ToLower
                Case "group"
                    Dim groupNode As New CNodeGroup(cnode)
                    groupNode.parseGroup(srcData, iDataPtr)

                Case "reposition"

                Case "value"
                    Dim valNode As New CNodeValue(cnode)
                    valNode.parseValue(srcData, iDataPtr)

                Case "choose"

                Case "when"

                Case "message"

                Case "use"

                Case "bookmark"

            End Select
        Next
    End Sub

    Public Property aName() As String
        Get
            Return mName
        End Get
        Set(ByVal value As String)
            mName = value
        End Set
    End Property


    Public Property aRepeat() As String
        Get
            Return mRepeat
        End Get
        Set(ByVal value As String)
            mRepeat = value
        End Set
    End Property


    Public Property aForeach() As String
        Get
            Return mForeach
        End Get
        Set(ByVal value As String)
            mForeach = value
        End Set
    End Property

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
