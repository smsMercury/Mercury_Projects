Imports system.xml

Public Class CNodeTemplate

    Private mName As String = ""
    Private mMain As Boolean = False
    Private mBuffer As Integer = 1024
    Private mRootNode As Xml.XmlNode

    Public Sub New(ByVal xnode As Xml.XmlNode)
        Dim attList As Xml.XmlAttributeCollection

        mRootNode = xnode
        attList = xnode.Attributes
        Try
            Name = attList("name").Value
        Catch ex As Exception
            MsgBox("Template name is required in script.")
            Exit Sub
        End Try
        Try
            Main = attList("main").Value
        Catch ex As Exception

        End Try
        Try
            Buffer = attList("buffer").Value
        Catch ex As Exception
            Buffer = 1024
        End Try
    End Sub

    Public Sub parseTemplate(ByRef srcData() As Byte, ByRef iDataPtr As Integer)

        '
        'parse template node
        ParseNode(srcData, iDataPtr, Me.mRootNode)
    End Sub

    Private Sub ParseNode(ByRef srcData() As Byte, ByRef iDataPtr As Integer, ByVal xNode As XmlNode)

        For Each cnode As XmlNode In xNode.ChildNodes
            Select Case cnode.Name.ToLower
                Case "group"
                    Dim groupNode As New CNodeGroup(cnode)
                    groupNode.parseGroup(srcData, iDataPtr)

                Case "reposition"

                Case "value"

                Case "choose"

                Case "when"

                Case "message"

                Case "use"

                Case "bookmark"

            End Select
        Next
    End Sub

    Public Property Main() As Boolean
        Get
            Return mMain
        End Get
        Set(ByVal value As Boolean)
            mMain = value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return mName
        End Get
        Set(ByVal value As String)
            mName = value
        End Set
    End Property

    Public Property Buffer() As Integer
        Get
            Return mBuffer
        End Get
        Set(ByVal value As Integer)
            mBuffer = value
        End Set
    End Property

End Class
