Imports System.xml
Imports SharpPcap

Public Class frmMain

    Private mOutputProps As New COutputProps
    Private mMsgObjs As New ArrayList
    Private mOutputDoc As XmlDocument
    Private mTOCDoc As XmlDocument
    Private mTOCTable As String = ""
    Private mCurrentDb As CMasterDb

    Private mRfosPort = "7577"
    Private mNcctPort = "6002"
    Private mLink16Port = "7000"
    Private mBlnArchive As Boolean = True

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        readPropertiesFile()
        '
        'Load forms
        Me.tbXmlOutputpath.Text = Me.mOutputProps.XmlPath

    End Sub

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click
        Dim dlg As New dlgProperties(Me.mOutputProps)

        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            loadScripts(mOutputProps.XmlPath)
        End If
    End Sub

    Private Sub OpenDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenDatabaseToolStripMenuItem.Click
        Dim dlg As New OpenFileDialog

        dlg.InitialDirectory = "c:\Mercury"
        dlg.Filter = "Access Database (*.mdb)|*.mdb"
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            Me.mOutputProps.DbPath = dlg.FileName
            Me.tbCurrentDb.Text = dlg.FileName
        Else
            Exit Sub
        End If
        Try
            mCurrentDb = New CMasterDb(mOutputProps.DbPath)
            LoadMessageList()
            loadTOCpage()
        Catch ex As Exception
            MsgBox("Unable to load Database." & vbCrLf & _
                   "Non-compliant Database.  Try creating a new one.")
            Exit Sub
        End Try

        If mOutputProps.XmlPath <> "" Then
            loadScripts(mOutputProps.XmlPath)
        End If
    End Sub

    Private Sub CreateNewDatabaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateNewDatabaseToolStripMenuItem.Click
        Dim dlg As New dlgNewDb

        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            mCurrentDb = dlg.MasterDb
            Me.tbCurrentDb.Text = dlg.DbPath
            Me.mOutputProps.DbPath = dlg.DbPath
            Me.mOutputProps.UpdateProps()
            LoadMessageList()
            If mOutputProps.XmlPath <> "" Then
                loadScripts(mOutputProps.XmlPath)
            End If
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        Me.Close()
    End Sub

    Private Sub readPropertiesFile()

        Me.mOutputProps.ReadProps()

    End Sub

    Private Sub loadTOCpage()

        If mCurrentDb IsNot Nothing Then
            lbTOCs.DataSource = mCurrentDb.GetTOCtables
        End If
    End Sub

    Private Sub loadScripts(ByVal path As String)
        Dim scrpts() As String
        Dim scoll As New ArrayList

        Try
            If path <> "" Then
                Me.clbSelectedMsgs.Items.Clear()
                scrpts = IO.Directory.GetFiles(path)
                For Each scr As String In scrpts
                    Me.clbSelectedMsgs.Items.Add(IO.Path.GetFileNameWithoutExtension(scr))
                Next
            End If
        Catch ex As Exception
            MsgBox("Unable to open:  " & path)
        End Try
    End Sub

    Private Sub LoadMessageList()
        Dim ds As DataSet = Nothing
        Dim dt As DataTable
        Dim rw As DataRow
        Dim lvi As ListViewItem
        Dim xAl As New ArrayList

        If mCurrentDb Is Nothing Then Exit Sub
        Try
            'get list of existing xml scripts
            Dim fls() As String = IO.Directory.GetFiles(mOutputProps.XmlPath)
            For Each fl As String In fls
                xAl.Add(IO.Path.GetFileNameWithoutExtension(fl))
            Next
        Catch ex As Exception
            'MsgBox("Error reading Xml files.  Looking in:  " &mXmlPath) 
        End Try
        Try
            'clear items from list
            Me.lvMessages.Clear()
            'readmsgid table and display msgs
            ds = mCurrentDb.GetMsgIds()
            dt = ds.Tables(0)
            For Each rw In dt.Rows
                lvi = New ListViewItem
                lvi.Text = rw(1)
                If xAl.Contains(lvi.Text) Then
                    lvi.ForeColor = Color.Lime
                End If
                Me.lvMessages.Items.Add(lvi)
            Next
        Catch ex As Exception
            MsgBox("Unable to load MsgldTblfrom database.")
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        End Try
    End Sub

    Private Sub btnProcData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcData.Click
        '
        'validate user input
        If Not checkInputs() Then Exit Sub
        '
        'create toc table
        If Not Me.mTOCTable = "" Then
            mCurrentDb.CreateTOCtable(mTOCTable)
        End If

        If (mOutputProps.MsgList IsNot Nothing) Then
            createMsgObjects()
        End If
        initOutputDoc(mOutputProps.MsgList, mOutputProps.SourceList, mOutputProps.FilterPortString)

        'parse pcaps
        parsePcaps()

        'clear variables
        Me.mMsgObjs.Clear()
        mTOCTable = ""
    End Sub

    Private Function checkInputs() As Boolean
        Dim msgList As New ArrayList
        Dim srcList As New ArrayList
        Dim prsEng As New CParserEngine

        '
        'get scripts
        For Each msg As Object In Me.clbSelectedMsgs.CheckedItems
            msgList.Add(mOutputProps.XmlPath & "\" & msg.ToString & ".xml")
        Next
        If msgList.Count = 0 Then
            MsgBox("You don't have any Messages selected to parse.")
            mOutputProps.MsgList = Nothing
            Return False
        Else
            msgList.Sort()
            mOutputProps.MsgList = msgList
        End If
        '
        'get output type/path
        mOutputProps.OutputFile = Me.tbOutputPath.Text
        If Not Me.rbNoMsg.Checked Then
            mTOCTable = IO.Path.GetFileNameWithoutExtension(mOutputProps.OutputFile) & "_TOC"
        End If
        '
        'get data source/s
        If Me.lbSources.Items.Count = 0 Then
            MsgBox("You don't have any Data Sources selected.")
            mOutputProps.SourceList = Nothing
            Return False
        Else
            srcList = New ArrayList
            For Each str As String In Me.lbSources.Items
                srcList.Add(str)
            Next
            mOutputProps.SourceList = srcList
        End If
        '
        'get filters
        If Me.rbRfos.Checked Then
            mOutputProps.FilterPortString = Me.mRfosPort
            If Not IO.File.Exists(mOutputProps.CmuRfosHdrScript) Then
                MsgBox("CmuRfosHdrScript file not found. Must select file in Properties menu.")
                Return False
            End If
        ElseIf Me.rbNcct.Checked Then
            mOutputProps.FilterPortString = Me.mNcctPort
            If Not IO.File.Exists(mOutputProps.NcctHdrScript) Then
                MsgBox("NcctHdrScript file not found. Must select file in Properties menu.")
                Return False
            End If
        Else
            mOutputProps.FilterPortString = Me.mLink16Port
        End If
        Return True
    End Function

    Private Sub createMsgObjects()
        Dim xmldoc As Xml.XmlDocument
        Dim elemList As XmlNodeList

        For Each msg As String In mOutputProps.MsgList
            Dim cmsg As New CMsgInfo

            cmsg.Script = msg
            cmsg.Port = mOutputProps.FilterPortString
            Try
                xmldoc = New XmlDocument()
                xmldoc.Load(msg)
                elemList = xmldoc.GetElementsByTagName("msgname")
                cmsg.MsgName = elemList(0).InnerXml
                elemList = xmldoc.GetElementsByTagName("msgid")
                cmsg.MsgId = elemList(0).InnerXml
            Catch ex As Exception
                MsgBox("Unable to parse msg script:  " & msg.ToString, MsgBoxStyle.OkOnly)
            End Try
            mMsgObjs.Add(cmsg)
        Next
    End Sub

#Region "Parse Pcaps"

    Private Sub parsePcaps()
        Dim device As PcapDevice = Nothing
        Dim pkt As SharpPcap.Packets.Packet
        Dim tocWrite As Integer = 0
        Dim pb1 As Double = 0
        Dim pb2 As Double = 0

        Cursor = Cursors.WaitCursor
        'open db
        mCurrentDb.OpenDB()
        'parse each pcap
        For Each src As String In mOutputProps.SourceList
            'Get an offline pcap device
            Try
                Dim inforeader As System.IO.FileInfo
                inforeader = My.Computer.FileSystem.GetFileInfo(src)
                mCurrTOCFile = IO.Path.GetFileNameWithoutExtension(src)
                Me.lblFile.Text = "File Progress : " & mcurrtocfile
                Me.lblFile.Refresh()

                Me.ProgressBar1.Value = 0
                device = New OfflinePcapDevice(src)
                device.Open()
                pkt = device.GetNextPacket
                While pkt IsNot Nothing
                    'pb2 += pkt.Bytes.Length
                    processPacket(pkt)
                    pkt = device.GetNextPacket
                    'pb1 = pb2 / inforeader.Length
                    'If pb1 > (0.001 * (ProgressBar1.Value + 1)) Then ProgressBar1.PerformStep()
                End While
                'pb2 = 0
                'Close device
                device.Close()
                'Save output file
                mOutputDoc.Save(mOutputProps.OutputFile)
                'Save TOC
                If (Me.rbNoMsg.Checked <> True) Then
                    'mTOCDoc.Save(IO.Path.ChangeExtension(mOutputProps.OutputFile, "TOC.xml"))
                    SaveTOC()
                End If
            Catch ex As Exception
                MsgBox("Unable to open pcap file." & vbCrLf & ex.Message)
            End Try
        Next
        'merge toc and msgid tables
        mCurrentDb.MergeMsgIds(mTOCTable)
        'update TOC page
        Me.loadTOCpage()
        'close db
        mCurrentDb.CloseDB()
        Cursor = Cursors.Default
        MsgBox("Done processing files.")
        ProgressBar1.Value = 0
    End Sub

    Private Sub SaveTOC()
        Dim rnode As XmlNode
        Dim tsStr As String = ""
        Dim tsFileName As String = ""
        Dim rsID As String = ""
        Dim rsLen As String = ""
        Dim rsTo As String = ""
        Dim rsFrom As String = ""
        Dim rsMsgName As String = ""

        'Parse nodes and update TOC
        rnode = mTOCDoc.DocumentElement
        For Each chld As XmlNode In rnode.ChildNodes
            tsStr = ""
            tsFileName = ""
            For Each att As XmlAttribute In chld.Attributes
                Select Case att.Name.ToLower
                    Case "timestamp"
                        tsStr = att.Value
                    Case "filename"
                        tsFileName = att.Value
                End Select
            Next
            'Get rfos_header info
            rsID = ""
            rsLen = ""
            rsTo = ""
            rsFrom = ""
            For Each subNode As XmlNode In chld.ChildNodes
                Select Case subNode.Name.ToLower
                    Case "rfos_id"
                        rsID = subNode.InnerText
                    Case "length"
                        rsLen = subNode.InnerText
                    Case "to"
                        rsTo = subNode.InnerText
                    Case "from"
                        rsFrom = subNode.InnerText
                End Select
            Next
            Me.mCurrentDb.updateTOCTable(mTOCTable, tsStr, tsFileName, rsTo, rsFrom, rsID, rsLen, rsMsgName, "RFOS")
        Next
    End Sub

    Private Sub processPacket(ByVal pkt As SharpPcap.Packets.Packet)
        Dim tcpPkt As Packets.TCPPacket
        Dim srcp As Integer
        Dim destp As Integer
        Dim data() As Byte

        Try
            If pkt.Data.Length > 0 Then
                If TypeOf (pkt) Is SharpPcap.Packets.TCPPacket Then
                    tcpPkt = CType(pkt, Packets.TCPPacket)
                    srcp = tcpPkt.SourcePort
                    destp = tcpPkt.DestinationPort
                    If (srcp = mNcctPort) Or (destp = mNcctPort) Then
                        'ReDim data(tcpPkt.Data.Length)
                        'Array.Copy(tcpPkt.Data, data, tcpPkt.Data.Length)
                        'CNcct.ProcessPkt(tcpPkt, data, fname, tocWrite)
                    ElseIf (srcp = mRfosPort) Or (destp = mRfosPort) Then
                        ReDim data(tcpPkt.Data.Length)
                        Array.Copy(tcpPkt.Data, data, tcpPkt.Data.Length)
                        If processCmuRfosHdr(tcpPkt, data) Then
                            If (processMessage(data)) Then
                                If (Me.rbAllMsg.Checked = True) Then
                                    updateTocNode(tcpPkt.Timeval.Date)
                                End If
                                'save output
                                If MGlobals.gXmlOutputDoc IsNot Nothing Then
                                    saveMessage(tcpPkt.Timeval.Date)
                                    If (Me.rbSelectMsg.Checked = True) Then
                                        updateTocNode(tcpPkt.Timeval.Date)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("Error processing packet." & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Function processCmuRfosHdr(ByVal tcpPkt As Packets.TCPPacket, ByVal data() As Byte) As Boolean

        'check min size, 32 bytes (cmu + rfos hdr)
        If tcpPkt.Data.Length > 31 Then
            'parse header (cmu/rfos)
            Dim hdr As New CParserEngine
            Dim toDoc As XmlDocument = Nothing
            Dim cmuDoc As New XmlDocument
            Dim numb As Integer = 0
            Dim iPtr As Integer = 0

            cmuDoc.Load(mOutputProps.CmuRfosHdrScript)
            numb = hdr.Parse(data, iPtr, cmuDoc, toDoc)
        Else
            Return False
        End If
        Return True
    End Function

    Private Function processMessage(ByVal data() As Byte) As Boolean
        Dim msg As New CParserEngine
        Dim toDoc As XmlDocument = Nothing
        Dim xnode As XmlNode
        Dim nodeList As XmlNodeList
        Dim xnode2 As XmlNode
        Dim nodeList2 As XmlNodeList
        Dim mType As String = ""
        Dim iPtr As Integer = 24 'skip cmu header

        Try
            processMessage = False
            nodeList = MGlobals.gXmlOutputDoc.GetElementsByTagName("msg_type")
            If nodeList.Count > 0 Then
                xnode = nodeList(0)
                mType = xnode.InnerText
                '
                'only process messages that have a msg_type of 1,2,21
                Select Case mType
                    Case 1, 2, 21
                        processMessage = True
                        If (Me.rbNoMsg.Checked <> True) Then
                            nodeList2 = MGlobals.gXmlOutputDoc.GetElementsByTagName("rfos_header")
                            xnode2 = nodeList2(0)
                            mCurrTOCNode = MGlobals.gXmlTOCDoc.ImportNode(xnode2, True)
                        End If
                        Dim msgDoc As XmlDocument = messageParser()
                        If msgDoc IsNot Nothing Then
                            msg.Parse(data, iPtr, msgDoc, toDoc)
                        Else
                            MGlobals.gXmlOutputDoc = Nothing
                        End If
                    Case Else
                        MGlobals.gXmlOutputDoc = Nothing
                End Select
            End If
        Catch ex As Exception
            MGlobals.gXmlOutputDoc = Nothing
        End Try
        Return processMessage
    End Function

#End Region


    Private Function messageParser() As XmlDocument
        Dim xnode As XmlNode
        Dim nodeList As XmlNodeList
        Dim mId As String = ""
        Dim msgDoc As XmlDocument = Nothing

        nodeList = MGlobals.gXmlOutputdoc.GetElementsByTagName("rfos_id")
        If nodeList.Count > 0 Then
            xnode = nodeList(0)
            mId = xnode.InnerText
            For Each msg As CMsgInfo In Me.mMsgObjs
                If mId = msg.MsgId Then
                    Try
                        msgDoc = New XmlDocument
                        'load parser script
                        msgDoc.Load(msg.Script)
                        'save msginfo in globals to use later for output purposes
                        MGlobals.messageinfo = msg
                        Exit For
                    Catch ex As Exception
                        Return Nothing
                    End Try
                End If
            Next
        Else
            Return Nothing
        End If
        Return msgDoc
    End Function

    Private Sub btnSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectAll.Click

        For i As Integer = 0 To clbSelectedMsgs.Items.Count - 1
            Me.clbSelectedMsgs.SetItemChecked(i, True)
        Next
    End Sub

    Private Sub btnClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAll.Click

        For i As Integer = 0 To clbSelectedMsgs.Items.Count - 1
            Me.clbSelectedMsgs.SetItemChecked(i, False)
        Next
    End Sub

#Region "Select Data Source"

    Private Sub btnBrowseSources_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseSources.Click
        Dim dlg As New OpenFileDialog

        dlg.Multiselect = True
        dlg.Filter = "Pcap Files   (*.*)|*.*"

        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            Me.lbSources.Items.Clear()
            For Each fil As String In dlg.FileNames
                Me.lbSources.Items.Add(fil)
            Next
            mBlnArchive = False
        End If
    End Sub

#End Region

#Region "Output"

    Private Sub saveMessage(ByVal msgTime As Date)
        Dim rnode As XmlNode = mOutputDoc.DocumentElement
        Dim onode As XmlNode = MGlobals.gXmlOutputDoc.DocumentElement
        Dim newNode As XmlNode = mOutputDoc.ImportNode(onode.FirstChild, True)
        Dim newAttr As XmlAttribute = mOutputDoc.CreateAttribute("timestamp")

        newAttr.Value = msgTime.ToString
        newNode.Attributes.Append(newAttr)
        rnode.AppendChild(newNode)
    End Sub

    Private Sub updateTocNode(ByVal msgTime As Date)
        Dim rnode As XmlNode = mTOCDoc.DocumentElement
        Dim newNode As XmlNode = mTOCDoc.ImportNode(mCurrTOCNode, True)
        Dim newAttr As XmlAttribute = mTOCDoc.CreateAttribute("timestamp")
        Dim newAttr2 As XmlAttribute = mTOCDoc.CreateAttribute("filename")

        newAttr.Value = msgTime.ToString
        newNode.Attributes.Append(newAttr)
        newAttr2.Value = mCurrTOCFile
        newNode.Attributes.Append(newAttr2)
        rnode.AppendChild(newNode)
    End Sub
    'Private Sub initOutputDoc()

    '    mOutputDoc = New XmlDocument
    '    mOutputDoc.LoadXml("<?xml version='1.0' ?>" & _
    '                        "<root></root>")
    '    mTOCDoc = New XmlDocument
    '    mTOCDoc.LoadXml("<?xml version='1.0' ?>" & _
    '                        "<root></root>")
    '    gXmlTOCDoc = New XmlDocument
    '    gXmlTOCDoc.LoadXml("<?xml version='1.0' ?>" & _
    '                        "<root></root>")

    'End Sub
    Private Sub initOutputDoc(ByVal msgList As ArrayList, ByVal srcList As ArrayList, ByVal port As String)


        mOutputDoc = New XmlDocument
        mOutputDoc.LoadXml("<?xml version='1.0' ?>" & _
                            "<root>" & _
                            "    <Metadata>" & _
                            "    </Metadata>" & _
                            "</root>")
        If (Me.rbNoMsg.Checked <> True) Then
            mTOCDoc = New XmlDocument
            mTOCDoc.LoadXml("<?xml version='1.0' ?>" & _
                                "<root></root>")
            gXmlTOCDoc = New XmlDocument
            gXmlTOCDoc.LoadXml("<?xml version='1.0' ?>" & _
                                "<root></root>")
        End If

        '
        'Get Metadata Node
        Dim mnode As XmlNode = mOutputDoc.DocumentElement.ChildNodes(0)
        Dim sToday As String = Format(Now, "yyyyMMdd_HHmmss")
        Dim nnode As XmlNode
        Dim inode As XmlNode
        '
        'Add ProcessingDate
        nnode = mOutputDoc.CreateNode(XmlNodeType.Element, "", "ProcessingDate", "")
        nnode.InnerText = sToday
        mnode.AppendChild(nnode)
        '
        'Add Port
        nnode = mOutputDoc.CreateNode(XmlNodeType.Element, "", "PortFilter", "")
        nnode.InnerText = port
        mnode.AppendChild(nnode)
        '
        'Add Selected Msgs
        If (mOutputProps.MsgList IsNot Nothing) Then
            nnode = mOutputDoc.CreateNode(XmlNodeType.Element, "", "SelectedMessages", "")
            mnode.AppendChild(nnode)
            For Each str As String In msgList
                inode = mOutputDoc.CreateNode(XmlNodeType.Element, "", "Message", "")
                Dim att As Xml.XmlAttribute = mOutputDoc.CreateAttribute("name")
                att.Value = str
                inode.Attributes.Append(att)
                nnode.AppendChild(inode)
            Next
        Else
            nnode = mOutputDoc.CreateNode(XmlNodeType.Element, "", "SelectedMessages", "")
            nnode.InnerText = "None"
            mnode.AppendChild(nnode)
        End If
        '
        'Add Source List
        nnode = mOutputDoc.CreateNode(XmlNodeType.Element, "", "SourceList", "")
        mnode.AppendChild(nnode)
        For Each str As String In srcList
            inode = mOutputDoc.CreateNode(XmlNodeType.Element, "", "File", "")
            Dim att As Xml.XmlAttribute = mOutputDoc.CreateAttribute("name")
            att.Value = str
            inode.Attributes.Append(att)
            nnode.AppendChild(inode)
        Next
    End Sub

    Private Sub bBrowseOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bBrowseOutput.Click

        Dim dlg As New OpenFileDialog
        Dim fil As String = ""

        dlg.InitialDirectory = Environment.SpecialFolder.MyComputer
        dlg.Filter = "Select Xml Output file (*.xml)|*.xml"
        dlg.CheckFileExists = False
        dlg.AddExtension = True

        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            fil = dlg.FileName
            If IO.File.Exists(fil) Then
                If MsgBox("Do you want to Overwrite Xml file?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If
            Me.tbOutputPath.Text = fil
        End If

    End Sub

#End Region

#Region "Filters Tab"

#End Region

#Region "Generate ICDs"

    Private Sub btnBrowsePrsrsPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowsePrsrsPath.Click
        Dim dlg As New FolderBrowserDialog

        dlg.RootFolder = Environment.SpecialFolder.MyComputer
        dlg.ShowNewFolderButton = True
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            Me.mOutputProps.XmlPath = dlg.SelectedPath
            Me.tbXmlOutputpath.Text = Me.mOutputProps.XmlPath
            Me.mOutputProps.UpdateProps()
        End If
    End Sub

    Private Sub btnProcessMsgs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProcessMsgs.Click
        Dim icol As ListView.SelectedListViewItemCollection
        Dim msgAry As New ArrayList
        Dim crtXml As New CCreateXml

        Cursor = Cursors.WaitCursor
        'retrieve selected items from listview
        icol = Me.lvMessages.SelectedItems
        For Each itm As ListViewItem In icol
            msgAry.Add(itm.Text)
        Next
        crtXml.CurrentDb = mCurrentDb
        crtXml.XmlPath = Me.mOutputProps.XmlPath

        For Each msg As String In msgAry
            crtXml.createXml(msg)
        Next
        Me.Cursor = Cursors.Default
        Me.LoadMessageList()
    End Sub

#End Region

#Region "Edit ICDs"

    Private Sub btnLoadParser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoadParser.Click
        Dim dlg As New OpenFileDialog

        dlg.InitialDirectory = Me.mOutputProps.XmlPath
        dlg.Filter = "Xml Scripts (*.xml)|*.xml"
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            Me.tbLoadParser.Text = dlg.FileName
            Me.RichTextBox1.LoadFile(dlg.FileName, RichTextBoxStreamType.PlainText)
        End If

    End Sub

    Private Sub btnSaveParser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveParser.Click

        If (tbLoadParser.Text <> "") Then
            If MsgBox("Are you sure you want to overwrite file?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                Me.RichTextBox1.SaveFile(tbLoadParser.Text, RichTextBoxStreamType.PlainText)
            End If
        End If

    End Sub

#End Region


    Private Sub lbTOCs_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbTOCs.SelectedIndexChanged
        Dim ds As DataSet = Nothing
        Dim tstr As String = lbTOCs.SelectedItem

        mCurrentDb.GetTocDataset(ds, tstr)
        ds.Tables(0).Columns.Remove("PcapFile")
        ds.Tables(0).Columns.Remove("MsgType")
        dgTOC.DataSource = ds.Tables(0)

    End Sub
End Class
