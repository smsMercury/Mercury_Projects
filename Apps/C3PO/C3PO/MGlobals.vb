Imports System.Xml


Module MGlobals

    Private mMsgInfo As CMsgInfo
    Private mCurrXmlNode As XmlNode

    Public mCurrTOCNode As XmlNode
    Public mCurrTOCFile As String
    Public gXmlOutputDoc As XmlDocument = Nothing
    Public gXmlTOCDoc As XmlDocument = Nothing
    Public gValueList As New InterfaceDictionary
    Public gXmlOutput As New CXmlOutput

#Region "Public Methods"

    Public Property MessageInfo() As CMsgInfo
        Get
            Return mMsgInfo
        End Get
        Set(ByVal value As CMsgInfo)
            mMsgInfo = value
        End Set
    End Property

    Public Property XmlOutputDoc() As XmlDocument
        Get
            Return gXmlOutputDoc
        End Get
        Set(ByVal value As XmlDocument)
            gXmlOutputDoc = value
            If gXmlOutputDoc Is Nothing Then
                CreateNewDocument()
            End If
        End Set
    End Property

    Public Property CurrXmlNode() As XmlNode
        Get
            Return mCurrXmlNode
        End Get
        Set(ByVal value As XmlNode)
            mCurrXmlNode = value
        End Set
    End Property

    Public Sub CreateNewDocument()

        gXmlOutputDoc = New XmlDocument()
        gXmlOutputDoc.LoadXml("<?xml version='1.0' ?>" & _
                            "<root></root>")
        mCurrXmlNode = gXmlOutputDoc.DocumentElement
    End Sub

    Public Sub AddOutput(ByVal name As String, Optional ByVal value As String = "")
        Dim elem As XmlElement = gXmlOutputDoc.CreateElement(name)

        elem.InnerText = value
        mCurrXmlNode = mCurrXmlNode.AppendChild(elem)
    End Sub

    Public Sub DoneOutput()
        mCurrXmlNode = mCurrXmlNode.ParentNode
    End Sub

    Public Sub LoopNode(ByVal loopIndex As Integer)
        Dim elem As XmlElement = gXmlOutputDoc.CreateElement("array")
        Dim attr As XmlAttribute = gXmlOutputDoc.CreateAttribute("loop")

        attr.Value = loopIndex.ToString
        elem.Attributes.Append(attr)
        mCurrXmlNode = mCurrXmlNode.AppendChild(elem)
    End Sub
#End Region

End Module
