Imports System.xml
Imports C3PO.MGlobals

Public Class CParserEngine

    Private mParserXml As XmlDocument
    Private mStartPtr As Integer = 0
    Private mTemplates As New ArrayList

#Region "Public Properties"


#End Region

#Region "Public Methods"

    Public Function Parse(ByVal data() As Byte, ByRef iDataPtr As Integer, ByVal parserXml As XmlDocument, ByRef xmlOutput As XmlDocument) As Integer
        Dim srcData(data.Length) As Byte
        Dim mainTemp As Boolean = False

        'make copy of data array
        Array.Copy(data, srcData, data.Length)

        mParserXml = parserXml
        MGlobals.XmlOutputDoc = xmlOutput
        MGlobals.gValueList.Clear()
        mStartPtr = iDataPtr
        '
        'preprocess meta and templates
        parseRoot(parserXml.DocumentElement)
        '
        'start parsing using main template
        If mTemplates.Count = 1 Then
            Dim ctemp As CNodeTemplate

            ctemp = mTemplates(0)
            ctemp.parseTemplate(srcData, iDataPtr)
        Else
            For Each ctemp As CNodeTemplate In mTemplates
                mainTemp = ctemp.Main
                If mainTemp Then
                    ctemp.parseTemplate(srcData, iDataPtr)
                End If
            Next
        End If

        Return iDataPtr - mStartPtr
    End Function
#End Region

#Region "Private Methods"

    Private Sub parseRoot(ByVal xnode As XmlNode)

        For Each cnode As XmlNode In xnode.ChildNodes
            Select Case cnode.Name.ToLower
                Case "metadata"

                Case "template"
                    Dim ctemp As New CNodeTemplate(cnode)
                    mTemplates.Add(ctemp)

            End Select
        Next
    End Sub

#End Region

End Class
