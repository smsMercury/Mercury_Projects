
Imports System.Xml
Imports System.IO

Public Class CCreateXml

    Private mCurrDb As CMasterDb
    Private mXmlPath As String = ""
    Private mFieldNames As New ArrayList

#Region "Private Methods"

    Private Sub createScriptTemplate(ByRef xdoc As XmlDocument, ByVal msg As String, ByVal msgid As Integer)

        xdoc.LoadXml("<parser version='2.0'>" & _
                        "<metadata>" & _
                            "<msgid>" & msgid & "</msgid>" & _
                            "<msgname>" & msg & "</msgname>" & _
                            "<date>" & Date.Now & "</date>" & _
                        "</metadata>" & _
                        "<template name='" & msg & "' main='true'>" & _
                        "</template>" & _
                    "</parser>")

        'Create an XML declaration.
        Dim xmldecl As XmlDeclaration
        xmldecl = xdoc.CreateXmlDeclaration("1.0", Nothing, Nothing)
        'Add the new node to the document. 
        Dim root As XmlElement = xdoc.DocumentElement
        xdoc.InsertBefore(xmldecl, root)
    End Sub

    Private Sub createXml(ByVal xdoc As XmlDocument, ByVal ds As DataSet)
        Dim xnodelst As XmlNodeList
        Dim tempNode As XmlNode
        Dim row As DataRow

        'get template node to add to
        xnodelst = xdoc.GetElementsByTagName("template")
        tempNode = xnodelst(0)
        For Each row In ds.Tables(0).Rows
            addNode(xdoc, tempNode, row)
        Next
    End Sub

    Private Sub addNode(ByRef xdoc As XmlDocument, ByRef tNode As XmlNode, ByVal row As DataRow)

        Select Case row("DataType")
            Case "STRUCT BEGIN"
                Dim xelem As XmlElement
                Dim att As XmlAttribute
                Dim lc As Integer = 0
                Dim sName As String = row("FieldName")
                sName = sName.Replace(".", "_")
                sName = sName.Replace("[0]", "")
                lc = row("MultiEntry")
                xelem = xdoc.CreateElement("group")
                att = xdoc.CreateAttribute("name")
                att.Value = sName
                xelem.Attributes.Append(att)

                If lc = 1 Then
                    'variable length struct size
                    Dim dlg As New DlgXmlStructVar(sName, Me.mFieldNames)

                    If dlg.ShowDialog = DialogResult.OK Then
                        att = xdoc.CreateAttribute("repeat")
                        att.Value = "{" & dlg.getFieldName & "}"
                        xelem.Attributes.Append(att)
                    End If
                ElseIf lc > 1 Then
                    'fixed struct size
                    att = xdoc.CreateAttribute("repeat")
                    att.Value = lc.ToString
                    xelem.Attributes.Append(att)
                End If
                tNode.AppendChild(xelem)
                tNode = xelem
            Case "STRUCT END"
                tNode = tNode.ParentNode
            Case "CHAR"
                Dim xelem As XmlElement = Nothing
                Dim att As XmlAttribute
                Dim lc As Integer = 0
                Dim sName As String = row("FieldLabel")
                lc = row("MultiEntry")
                If lc <= 1 Then
                    xelem = xdoc.CreateElement("value")
                    att = xdoc.CreateAttribute("name")
                    att.Value = sName
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("size")
                    att.Value = row("FieldSize").ToString
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("type")
                    att.Value = "raw"
                    xelem.Attributes.Append(att)
                    tNode.AppendChild(xelem)
                ElseIf (sName.Contains("pad")) Then
                    xelem = xdoc.CreateElement("value")
                    att = xdoc.CreateAttribute("name")
                    att.Value = sName
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("size")
                    att.Value = row("FieldSize").ToString
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("type")
                    att.Value = "raw"
                    xelem.Attributes.Append(att)
                    tNode.AppendChild(xelem)
                Else
                    Dim dlg As New DlgXmlCharHandler(row)
                    If dlg.ShowDialog = DialogResult.OK Then
                        If dlg.rbArray.Checked Then 'treat as array of chars 
                            For i As Integer = 1 To lc
                                xelem = xdoc.CreateElement("value")
                                att = xdoc.CreateAttribute("name")
                                att.Value = sName & "_" & i.ToString
                                xelem.Attributes.Append(att)
                                att = xdoc.CreateAttribute("size")
                                att.Value = "1"
                                xelem.Attributes.Append(att)
                                att = xdoc.CreateAttribute("type")
                                att.Value = "string"
                                xelem.Attributes.Append(att)
                                tNode.AppendChild(xelem)
                            Next
                        Else 'treat as string
                            xelem = xdoc.CreateElement("value")
                            att = xdoc.CreateAttribute("name")
                            att.Value = sName
                            xelem.Attributes.Append(att)
                            att = xdoc.CreateAttribute("size")
                            att.Value = row("FieldSize")
                            xelem.Attributes.Append(att)
                            att = xdoc.CreateAttribute("type")
                            att.Value = "string"
                            xelem.Attributes.Append(att)
                            tNode.AppendChild(xelem)
                        End If
                    End If
                End If
            Case "SHORT"
                Dim xelem As XmlElement = Nothing
                Dim att As XmlAttribute
                Dim lc As Integer = 0
                Dim sName As String = row("FieldLabel")
                'add field name to list for possible use as a variable length array     * counter
                Me.mFieldNames.Add(sName)
                lc = row("MultiEntry")
                If lc = 0 Then
                    xelem = xdoc.CreateElement("value")
                    att = xdoc.CreateAttribute("name")
                    att.Value = sName
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("size")
                    att.Value = row("FieldSize").ToString
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("type")
                    att.Value = "integer"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("signed")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("swapped")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    tNode.AppendChild(xelem)
                Else
                    For i As Integer = 1 To lc
                        xelem = xdoc.CreateElement("value")
                        att = xdoc.CreateAttribute("name")
                        att.Value = sName & "_" & i.ToString
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("size")
                        att.Value = "2"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("type")
                        att.Value = "integer"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("signed")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("swapped")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        tNode.AppendChild(xelem)
                    Next
                End If

            Case "LONG"
                Dim xelem As XmlElement = Nothing
                Dim att As XmlAttribute
                Dim lc As Integer = 0
                Dim sName As String = row("FieldLabel")
                'add field name to list for possible use as a variable length array
                Me.mFieldNames.Add(sName)
                lc = row("MultiEntry")
                If lc = 0 Then
                    xelem = xdoc.CreateElement("value")
                    att = xdoc.CreateAttribute("name")
                    att.Value = sName
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("size")
                    att.Value = row("FieldSize").ToString
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("type")
                    att.Value = "integer"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("signed")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("swapped")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    tNode.AppendChild(xelem)
                Else
                    For i As Integer = 1 To lc
                        xelem = xdoc.CreateElement("value")
                        att = xdoc.CreateAttribute("name")
                        att.Value = sName & "_" & i.ToString
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("size")
                        att.Value = "4"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("type")
                        att.Value = "integer"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("signed")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("swapped")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        tNode.AppendChild(xelem)
                    Next
                End If

            Case "FLOAT"
                Dim xelem As XmlElement = Nothing
                Dim att As XmlAttribute
                Dim lc As Integer = 0
                Dim sName As String = row("FieldLabel")
                lc = row("MultiEntry")
                If lc = 0 Then
                    xelem = xdoc.CreateElement("value")
                    att = xdoc.CreateAttribute("name")
                    att.Value = sName
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("size")
                    att.Value = row("FieldSize").ToString
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("type")
                    att.Value = "float"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("swapped")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    tNode.AppendChild(xelem)
                Else
                    For i As Integer = 1 To lc
                        xelem = xdoc.CreateElement("value")
                        att = xdoc.CreateAttribute("name")
                        att.Value = sName & "_" & i.ToString
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("size")
                        att.Value = "4"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("type")
                        att.Value = "float"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("swapped")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        tNode.AppendChild(xelem)
                    Next
                End If

            Case "DOUBLE"
                Dim xelem As XmlElement = Nothing
                Dim att As XmlAttribute
                Dim lc As Integer = 0
                Dim sName As String = row("FieldLabel")

                'add field name to list for possible use as a variable length array
                Me.mFieldNames.Add(sName)
                lc = row("MultiEntry")
                If lc = 0 Then
                    xelem = xdoc.CreateElement("value")
                    att = xdoc.CreateAttribute("name")
                    att.Value = sName
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("size")
                    att.Value = row("FieldSize").ToString
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("type")
                    att.Value = "integer"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("signed")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("swapped")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    tNode.AppendChild(xelem)
                Else
                    For i As Integer = 1 To lc
                        xelem = xdoc.CreateElement("value")
                        att = xdoc.CreateAttribute("name")
                        att.Value = sName & "_" & i.ToString
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("size")
                        att.Value = "8"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("type")
                        att.Value = "integer"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("signed")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("swapped")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        tNode.AppendChild(xelem)
                    Next
                End If

            Case "BAM16"
                Dim xelem As XmlElement = Nothing
                Dim att As XmlAttribute
                Dim lc As Integer = 0
                Dim sName As String = row("FieldLabel")
                lc = row("MultiEntry")
                If lc = 0 Then
                    xelem = xdoc.CreateElement("value")
                    att = xdoc.CreateAttribute("name")
                    att.Value = sName
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("size")
                    att.Value = row("FieldSize").ToString
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("type")
                    att.Value = "integer"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("signed")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("swapped")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    tNode.AppendChild(xelem)
                Else
                    For i As Integer = 1 To lc
                        xelem = xdoc.CreateElement("value")
                        att = xdoc.CreateAttribute("name")
                        att.Value = sName & "_" & i.ToString
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("size")
                        att.Value = "2"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("type")
                        att.Value = "integer"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("signed")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("swapped")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        tNode.AppendChild(xelem)
                    Next
                End If
            Case "ENUM"
                Dim xelem As XmlElement = Nothing
                Dim att As XmlAttribute
                Dim lc As Integer = 0
                Dim sName As String = row("FieldLabel")

                lc = row("MultiEntry")
                If lc = 0 Then
                    xelem = xdoc.CreateElement("value")
                    att = xdoc.CreateAttribute("name")
                    att.Value = sName
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("size")
                    att.Value = row("FieldSize").ToString
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("type")
                    att.Value = "integer"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("signed")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("swapped")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    tNode.AppendChild(xelem)
                Else
                    For i As Integer = 1 To lc
                        xelem = xdoc.CreateElement("value")
                        att = xdoc.CreateAttribute("name")
                        att.Value = sName & "_" & i.ToString
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("size")
                        att.Value = "4"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("type")
                        att.Value = "integer"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("signed")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("swapped")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        tNode.AppendChild(xelem)
                    Next
                End If

            Case "FREQ"
                Dim xelem As XmlElement = Nothing
                Dim att As XmlAttribute
                Dim lc As Integer = 0
                Dim sName As String = row("FieldLabel")
                lc = row("MultiEntry")
                If lc = 0 Then
                    xelem = xdoc.CreateElement("value")
                    att = xdoc.CreateAttribute("name")
                    att.Value = sName
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("size")
                    att.Value = row("FieldSize").ToString
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("type")
                    att.Value = "integer"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("signed")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    att = xdoc.CreateAttribute("swapped")
                    att.Value = "true"
                    xelem.Attributes.Append(att)
                    tNode.AppendChild(xelem)
                Else
                    For i As Integer = 1 To lc
                        xelem = xdoc.CreateElement("value")
                        att = xdoc.CreateAttribute("name")
                        att.Value = sName & "_" & i.ToString
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("size")
                        att.Value = "4"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("type")
                        att.Value = "integer"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("signed")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        att = xdoc.CreateAttribute("swapped")
                        att.Value = "true"
                        xelem.Attributes.Append(att)
                        tNode.AppendChild(xelem)
                    Next
                End If
        End Select
    End Sub

#End Region

#Region "Public Methods"

    Public Sub CreateXml(ByVal msg As String)
        Dim xdoc As New XmlDocument
        Dim ds As DataSet
        Dim vsDs As DataSet
        Dim rw As DataRow
        Dim msgid As Integer

        Try
            'get msg id from msgidtbl
            ds = mCurrDb.GetMsgid(msg)
            If ds.Tables(0).Rows.Count > 0 Then
                rw = ds.Tables(0).Rows(0)
                msgid = rw(0)
            Else
                MsgBox("Unable to get msgid for:  " & msg)
                Exit Sub
            End If
            'get ds from VarStruct table
            vsDs = mCurrDb.GetMsgVarStruct(msgid)
            If vsDs Is Nothing Or vsDs.Tables(0).Rows.Count = 0 Then
                MsgBox("Unable to get varstruct for:  " & msg)
                Exit Sub
            End If
            '
            'create xml script template
            createScriptTemplate(xdoc, msg, msgid)
            'add xml to doc 
            createXml(xdoc, vsDs)
            'save xml to file
            xdoc.Save(mXmlPath & "\" & msg & ".xml")
        Catch ex As Exception
            MsgBox("Unable to create Xml for following Message:  " & msg & vbCrLf & ex.Message)
        End Try
    End Sub

    Public WriteOnly Property CurrentDb() As CMasterDb
        Set(ByVal value As CMasterDb)
            Me.mCurrDb = value
        End Set
    End Property

    Public WriteOnly Property XmlPath() As String
        Set(ByVal value As String)
            mXmlPath = value
        End Set
    End Property

#End Region

End Class

