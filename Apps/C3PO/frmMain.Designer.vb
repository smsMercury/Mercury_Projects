<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.OpenDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CreateNewDatabaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.PropertiesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.LoadConfigurationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ATNToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.CHATToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.cmnuMessages = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.SelectAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ClearAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.LoadProfileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.FilenameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.tpProcessCap = New System.Windows.Forms.TabPage
        Me.tbCurrentDb = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.rbLink16 = New System.Windows.Forms.RadioButton
        Me.rbNcct = New System.Windows.Forms.RadioButton
        Me.rbRfos = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btnSelectAll = New System.Windows.Forms.Button
        Me.btnClearAll = New System.Windows.Forms.Button
        Me.clbSelectedMsgs = New System.Windows.Forms.CheckedListBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.rbNoMsg = New System.Windows.Forms.RadioButton
        Me.rbAllMsg = New System.Windows.Forms.RadioButton
        Me.rbSelectMsg = New System.Windows.Forms.RadioButton
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.btnBrowseSources = New System.Windows.Forms.Button
        Me.lbSources = New System.Windows.Forms.ListBox
        Me.lblFile = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.bBrowseOutput = New System.Windows.Forms.Button
        Me.tbOutputPath = New System.Windows.Forms.TextBox
        Me.lOutputType = New System.Windows.Forms.Label
        Me.btnProcData = New System.Windows.Forms.Button
        Me.tpGenICD = New System.Windows.Forms.TabPage
        Me.btnProcessMsgs = New System.Windows.Forms.Button
        Me.btnBrowsePrsrsPath = New System.Windows.Forms.Button
        Me.tbXmlOutputpath = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.lvMessages = New System.Windows.Forms.ListView
        Me.tpEditICD = New System.Windows.Forms.TabPage
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.btnSaveParser = New System.Windows.Forms.Button
        Me.btnLoadParser = New System.Windows.Forms.Button
        Me.tbLoadParser = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Label8 = New System.Windows.Forms.Label
        Me.ListBox2 = New System.Windows.Forms.ListBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lbTOCs = New System.Windows.Forms.ListBox
        Me.dgTOC = New System.Windows.Forms.DataGridView
        Me.MenuStrip1.SuspendLayout()
        Me.cmnuMessages.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.tpProcessCap.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.tpGenICD.SuspendLayout()
        Me.tpEditICD.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.dgTOC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.PropertiesToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(5, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(814, 28)
        Me.MenuStrip1.TabIndex = 6
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenDatabaseToolStripMenuItem, Me.CreateNewDatabaseToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(44, 24)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'OpenDatabaseToolStripMenuItem
        '
        Me.OpenDatabaseToolStripMenuItem.Name = "OpenDatabaseToolStripMenuItem"
        Me.OpenDatabaseToolStripMenuItem.Size = New System.Drawing.Size(222, 24)
        Me.OpenDatabaseToolStripMenuItem.Text = "Open Database"
        '
        'CreateNewDatabaseToolStripMenuItem
        '
        Me.CreateNewDatabaseToolStripMenuItem.Name = "CreateNewDatabaseToolStripMenuItem"
        Me.CreateNewDatabaseToolStripMenuItem.Size = New System.Drawing.Size(222, 24)
        Me.CreateNewDatabaseToolStripMenuItem.Text = "Create New Database"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(222, 24)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'PropertiesToolStripMenuItem
        '
        Me.PropertiesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.EditToolStripMenuItem, Me.LoadConfigurationToolStripMenuItem})
        Me.PropertiesToolStripMenuItem.Name = "PropertiesToolStripMenuItem"
        Me.PropertiesToolStripMenuItem.Size = New System.Drawing.Size(88, 24)
        Me.PropertiesToolStripMenuItem.Text = "Properties"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(206, 24)
        Me.EditToolStripMenuItem.Text = "Edit"
        '
        'LoadConfigurationToolStripMenuItem
        '
        Me.LoadConfigurationToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ATNToolStripMenuItem, Me.CHATToolStripMenuItem})
        Me.LoadConfigurationToolStripMenuItem.Name = "LoadConfigurationToolStripMenuItem"
        Me.LoadConfigurationToolStripMenuItem.Size = New System.Drawing.Size(206, 24)
        Me.LoadConfigurationToolStripMenuItem.Text = "Load Configuration"
        '
        'ATNToolStripMenuItem
        '
        Me.ATNToolStripMenuItem.Name = "ATNToolStripMenuItem"
        Me.ATNToolStripMenuItem.Size = New System.Drawing.Size(116, 24)
        Me.ATNToolStripMenuItem.Text = "ATN"
        '
        'CHATToolStripMenuItem
        '
        Me.CHATToolStripMenuItem.Name = "CHATToolStripMenuItem"
        Me.CHATToolStripMenuItem.Size = New System.Drawing.Size(116, 24)
        Me.CHATToolStripMenuItem.Text = "CHAT"
        '
        'cmnuMessages
        '
        Me.cmnuMessages.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectAllToolStripMenuItem, Me.ClearAllToolStripMenuItem, Me.LoadProfileToolStripMenuItem})
        Me.cmnuMessages.Name = "cmnuMessages"
        Me.cmnuMessages.Size = New System.Drawing.Size(159, 76)
        '
        'SelectAllToolStripMenuItem
        '
        Me.SelectAllToolStripMenuItem.Name = "SelectAllToolStripMenuItem"
        Me.SelectAllToolStripMenuItem.Size = New System.Drawing.Size(158, 24)
        Me.SelectAllToolStripMenuItem.Text = "Select All"
        '
        'ClearAllToolStripMenuItem
        '
        Me.ClearAllToolStripMenuItem.Name = "ClearAllToolStripMenuItem"
        Me.ClearAllToolStripMenuItem.Size = New System.Drawing.Size(158, 24)
        Me.ClearAllToolStripMenuItem.Text = "Clear All"
        '
        'LoadProfileToolStripMenuItem
        '
        Me.LoadProfileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FilenameToolStripMenuItem})
        Me.LoadProfileToolStripMenuItem.Name = "LoadProfileToolStripMenuItem"
        Me.LoadProfileToolStripMenuItem.Size = New System.Drawing.Size(158, 24)
        Me.LoadProfileToolStripMenuItem.Text = "Load Profile"
        '
        'FilenameToolStripMenuItem
        '
        Me.FilenameToolStripMenuItem.Name = "FilenameToolStripMenuItem"
        Me.FilenameToolStripMenuItem.Size = New System.Drawing.Size(138, 24)
        Me.FilenameToolStripMenuItem.Text = "Filename"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tpProcessCap)
        Me.TabControl1.Controls.Add(Me.tpGenICD)
        Me.TabControl1.Controls.Add(Me.tpEditICD)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 28)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(814, 775)
        Me.TabControl1.TabIndex = 7
        '
        'tpProcessCap
        '
        Me.tpProcessCap.Controls.Add(Me.tbCurrentDb)
        Me.tpProcessCap.Controls.Add(Me.Label1)
        Me.tpProcessCap.Controls.Add(Me.GroupBox4)
        Me.tpProcessCap.Controls.Add(Me.GroupBox2)
        Me.tpProcessCap.Controls.Add(Me.ProgressBar1)
        Me.tpProcessCap.Controls.Add(Me.GroupBox5)
        Me.tpProcessCap.Controls.Add(Me.GroupBox3)
        Me.tpProcessCap.Controls.Add(Me.lblFile)
        Me.tpProcessCap.Controls.Add(Me.GroupBox1)
        Me.tpProcessCap.Controls.Add(Me.btnProcData)
        Me.tpProcessCap.Location = New System.Drawing.Point(4, 25)
        Me.tpProcessCap.Name = "tpProcessCap"
        Me.tpProcessCap.Padding = New System.Windows.Forms.Padding(3)
        Me.tpProcessCap.Size = New System.Drawing.Size(806, 746)
        Me.tpProcessCap.TabIndex = 0
        Me.tpProcessCap.Text = "Process Captures"
        Me.tpProcessCap.UseVisualStyleBackColor = True
        '
        'tbCurrentDb
        '
        Me.tbCurrentDb.Location = New System.Drawing.Point(166, 12)
        Me.tbCurrentDb.Name = "tbCurrentDb"
        Me.tbCurrentDb.ReadOnly = True
        Me.tbCurrentDb.Size = New System.Drawing.Size(592, 22)
        Me.tbCurrentDb.TabIndex = 24
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(124, 17)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Current Database:"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.rbLink16)
        Me.GroupBox4.Controls.Add(Me.rbNcct)
        Me.GroupBox4.Controls.Add(Me.rbRfos)
        Me.GroupBox4.Location = New System.Drawing.Point(410, 540)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox4.Size = New System.Drawing.Size(363, 128)
        Me.GroupBox4.TabIndex = 19
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Filter by Port"
        '
        'rbLink16
        '
        Me.rbLink16.AutoSize = True
        Me.rbLink16.Location = New System.Drawing.Point(13, 76)
        Me.rbLink16.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rbLink16.Name = "rbLink16"
        Me.rbLink16.Size = New System.Drawing.Size(184, 21)
        Me.rbLink16.TabIndex = 3
        Me.rbLink16.Text = "7000 - Link16 Messages"
        Me.rbLink16.UseVisualStyleBackColor = True
        '
        'rbNcct
        '
        Me.rbNcct.AutoSize = True
        Me.rbNcct.Location = New System.Drawing.Point(13, 52)
        Me.rbNcct.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rbNcct.Name = "rbNcct"
        Me.rbNcct.Size = New System.Drawing.Size(170, 21)
        Me.rbNcct.TabIndex = 2
        Me.rbNcct.Text = "6002 - Ncct Messages"
        Me.rbNcct.UseVisualStyleBackColor = True
        '
        'rbRfos
        '
        Me.rbRfos.AutoSize = True
        Me.rbRfos.Checked = True
        Me.rbRfos.Location = New System.Drawing.Point(13, 26)
        Me.rbRfos.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.rbRfos.Name = "rbRfos"
        Me.rbRfos.Size = New System.Drawing.Size(171, 21)
        Me.rbRfos.TabIndex = 1
        Me.rbRfos.TabStop = True
        Me.rbRfos.Text = "7577 - Rfos Messages"
        Me.rbRfos.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnSelectAll)
        Me.GroupBox2.Controls.Add(Me.btnClearAll)
        Me.GroupBox2.Controls.Add(Me.clbSelectedMsgs)
        Me.GroupBox2.Location = New System.Drawing.Point(22, 45)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(751, 255)
        Me.GroupBox2.TabIndex = 16
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Select Messages"
        '
        'btnSelectAll
        '
        Me.btnSelectAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSelectAll.Location = New System.Drawing.Point(13, 221)
        Me.btnSelectAll.Name = "btnSelectAll"
        Me.btnSelectAll.Size = New System.Drawing.Size(98, 24)
        Me.btnSelectAll.TabIndex = 3
        Me.btnSelectAll.Text = "Select All"
        Me.btnSelectAll.UseVisualStyleBackColor = True
        '
        'btnClearAll
        '
        Me.btnClearAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnClearAll.Location = New System.Drawing.Point(135, 221)
        Me.btnClearAll.Name = "btnClearAll"
        Me.btnClearAll.Size = New System.Drawing.Size(98, 24)
        Me.btnClearAll.TabIndex = 2
        Me.btnClearAll.Text = "Clear All"
        Me.btnClearAll.UseVisualStyleBackColor = True
        '
        'clbSelectedMsgs
        '
        Me.clbSelectedMsgs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.clbSelectedMsgs.CheckOnClick = True
        Me.clbSelectedMsgs.FormattingEnabled = True
        Me.clbSelectedMsgs.HorizontalScrollbar = True
        Me.clbSelectedMsgs.Location = New System.Drawing.Point(13, 21)
        Me.clbSelectedMsgs.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.clbSelectedMsgs.MultiColumn = True
        Me.clbSelectedMsgs.Name = "clbSelectedMsgs"
        Me.clbSelectedMsgs.Size = New System.Drawing.Size(723, 191)
        Me.clbSelectedMsgs.Sorted = True
        Me.clbSelectedMsgs.TabIndex = 1
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(22, 683)
        Me.ProgressBar1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ProgressBar1.Maximum = 1000
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(575, 27)
        Me.ProgressBar1.TabIndex = 20
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.rbNoMsg)
        Me.GroupBox5.Controls.Add(Me.rbAllMsg)
        Me.GroupBox5.Controls.Add(Me.rbSelectMsg)
        Me.GroupBox5.Location = New System.Drawing.Point(410, 431)
        Me.GroupBox5.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox5.Size = New System.Drawing.Size(363, 103)
        Me.GroupBox5.TabIndex = 22
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Table of Contents"
        '
        'rbNoMsg
        '
        Me.rbNoMsg.AutoSize = True
        Me.rbNoMsg.Location = New System.Drawing.Point(13, 74)
        Me.rbNoMsg.Margin = New System.Windows.Forms.Padding(4)
        Me.rbNoMsg.Name = "rbNoMsg"
        Me.rbNoMsg.Size = New System.Drawing.Size(148, 21)
        Me.rbNoMsg.TabIndex = 2
        Me.rbNoMsg.TabStop = True
        Me.rbNoMsg.Text = "No TOC Messages"
        Me.rbNoMsg.UseVisualStyleBackColor = True
        '
        'rbAllMsg
        '
        Me.rbAllMsg.AutoSize = True
        Me.rbAllMsg.Checked = True
        Me.rbAllMsg.Location = New System.Drawing.Point(13, 23)
        Me.rbAllMsg.Margin = New System.Windows.Forms.Padding(4)
        Me.rbAllMsg.Name = "rbAllMsg"
        Me.rbAllMsg.Size = New System.Drawing.Size(112, 21)
        Me.rbAllMsg.TabIndex = 1
        Me.rbAllMsg.TabStop = True
        Me.rbAllMsg.Text = "All Messages"
        Me.rbAllMsg.UseVisualStyleBackColor = True
        '
        'rbSelectMsg
        '
        Me.rbSelectMsg.AutoSize = True
        Me.rbSelectMsg.Location = New System.Drawing.Point(13, 47)
        Me.rbSelectMsg.Margin = New System.Windows.Forms.Padding(4)
        Me.rbSelectMsg.Name = "rbSelectMsg"
        Me.rbSelectMsg.Size = New System.Drawing.Size(185, 21)
        Me.rbSelectMsg.TabIndex = 0
        Me.rbSelectMsg.Text = "Selected Messages Only"
        Me.rbSelectMsg.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.btnBrowseSources)
        Me.GroupBox3.Controls.Add(Me.lbSources)
        Me.GroupBox3.Location = New System.Drawing.Point(22, 305)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox3.Size = New System.Drawing.Size(381, 363)
        Me.GroupBox3.TabIndex = 17
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Select Data Sources"
        '
        'btnBrowseSources
        '
        Me.btnBrowseSources.Location = New System.Drawing.Point(225, 307)
        Me.btnBrowseSources.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnBrowseSources.Name = "btnBrowseSources"
        Me.btnBrowseSources.Size = New System.Drawing.Size(137, 28)
        Me.btnBrowseSources.TabIndex = 3
        Me.btnBrowseSources.Text = "Browse Sources"
        Me.btnBrowseSources.UseVisualStyleBackColor = True
        '
        'lbSources
        '
        Me.lbSources.FormattingEnabled = True
        Me.lbSources.HorizontalScrollbar = True
        Me.lbSources.ItemHeight = 16
        Me.lbSources.Location = New System.Drawing.Point(17, 21)
        Me.lbSources.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.lbSources.Name = "lbSources"
        Me.lbSources.Size = New System.Drawing.Size(345, 276)
        Me.lbSources.Sorted = True
        Me.lbSources.TabIndex = 2
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Location = New System.Drawing.Point(23, 715)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(91, 17)
        Me.lblFile.TabIndex = 21
        Me.lblFile.Text = "File Progress"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.bBrowseOutput)
        Me.GroupBox1.Controls.Add(Me.tbOutputPath)
        Me.GroupBox1.Controls.Add(Me.lOutputType)
        Me.GroupBox1.Location = New System.Drawing.Point(410, 305)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(363, 119)
        Me.GroupBox1.TabIndex = 18
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Output Type"
        '
        'bBrowseOutput
        '
        Me.bBrowseOutput.Location = New System.Drawing.Point(272, 76)
        Me.bBrowseOutput.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.bBrowseOutput.Name = "bBrowseOutput"
        Me.bBrowseOutput.Size = New System.Drawing.Size(77, 27)
        Me.bBrowseOutput.TabIndex = 6
        Me.bBrowseOutput.Text = "Browse"
        Me.bBrowseOutput.UseVisualStyleBackColor = True
        '
        'tbOutputPath
        '
        Me.tbOutputPath.Location = New System.Drawing.Point(13, 47)
        Me.tbOutputPath.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.tbOutputPath.Name = "tbOutputPath"
        Me.tbOutputPath.Size = New System.Drawing.Size(335, 22)
        Me.tbOutputPath.TabIndex = 5
        '
        'lOutputType
        '
        Me.lOutputType.AutoSize = True
        Me.lOutputType.Location = New System.Drawing.Point(127, 28)
        Me.lOutputType.Name = "lOutputType"
        Me.lOutputType.Size = New System.Drawing.Size(83, 17)
        Me.lOutputType.TabIndex = 4
        Me.lOutputType.Text = "Output path"
        '
        'btnProcData
        '
        Me.btnProcData.Location = New System.Drawing.Point(623, 678)
        Me.btnProcData.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnProcData.Name = "btnProcData"
        Me.btnProcData.Size = New System.Drawing.Size(136, 39)
        Me.btnProcData.TabIndex = 15
        Me.btnProcData.Text = "Process Data"
        Me.btnProcData.UseVisualStyleBackColor = True
        '
        'tpGenICD
        '
        Me.tpGenICD.Controls.Add(Me.btnProcessMsgs)
        Me.tpGenICD.Controls.Add(Me.btnBrowsePrsrsPath)
        Me.tpGenICD.Controls.Add(Me.tbXmlOutputpath)
        Me.tpGenICD.Controls.Add(Me.Label6)
        Me.tpGenICD.Controls.Add(Me.Label5)
        Me.tpGenICD.Controls.Add(Me.Label4)
        Me.tpGenICD.Controls.Add(Me.lvMessages)
        Me.tpGenICD.Location = New System.Drawing.Point(4, 25)
        Me.tpGenICD.Name = "tpGenICD"
        Me.tpGenICD.Padding = New System.Windows.Forms.Padding(3)
        Me.tpGenICD.Size = New System.Drawing.Size(806, 746)
        Me.tpGenICD.TabIndex = 1
        Me.tpGenICD.Text = "Generate ICDs"
        Me.tpGenICD.UseVisualStyleBackColor = True
        '
        'btnProcessMsgs
        '
        Me.btnProcessMsgs.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnProcessMsgs.Location = New System.Drawing.Point(39, 610)
        Me.btnProcessMsgs.Name = "btnProcessMsgs"
        Me.btnProcessMsgs.Size = New System.Drawing.Size(190, 27)
        Me.btnProcessMsgs.TabIndex = 13
        Me.btnProcessMsgs.Text = "Process Messages"
        Me.btnProcessMsgs.UseVisualStyleBackColor = True
        '
        'btnBrowsePrsrsPath
        '
        Me.btnBrowsePrsrsPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnBrowsePrsrsPath.Location = New System.Drawing.Point(454, 564)
        Me.btnBrowsePrsrsPath.Name = "btnBrowsePrsrsPath"
        Me.btnBrowsePrsrsPath.Size = New System.Drawing.Size(78, 25)
        Me.btnBrowsePrsrsPath.TabIndex = 12
        Me.btnBrowsePrsrsPath.Text = "Browse"
        Me.btnBrowsePrsrsPath.UseVisualStyleBackColor = True
        '
        'tbXmlOutputpath
        '
        Me.tbXmlOutputpath.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tbXmlOutputpath.Location = New System.Drawing.Point(39, 568)
        Me.tbXmlOutputpath.Name = "tbXmlOutputpath"
        Me.tbXmlOutputpath.Size = New System.Drawing.Size(392, 22)
        Me.tbXmlOutputpath.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(38, 546)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 17)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Parsers Output path"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(36, 499)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(185, 17)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Select messages to process"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(36, 30)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 17)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Messages"
        '
        'lvMessages
        '
        Me.lvMessages.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvMessages.Location = New System.Drawing.Point(39, 50)
        Me.lvMessages.Name = "lvMessages"
        Me.lvMessages.Size = New System.Drawing.Size(735, 446)
        Me.lvMessages.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lvMessages.TabIndex = 7
        Me.lvMessages.UseCompatibleStateImageBehavior = False
        Me.lvMessages.View = System.Windows.Forms.View.List
        '
        'tpEditICD
        '
        Me.tpEditICD.Controls.Add(Me.RichTextBox1)
        Me.tpEditICD.Controls.Add(Me.btnSaveParser)
        Me.tpEditICD.Controls.Add(Me.btnLoadParser)
        Me.tpEditICD.Controls.Add(Me.tbLoadParser)
        Me.tpEditICD.Controls.Add(Me.Label7)
        Me.tpEditICD.Location = New System.Drawing.Point(4, 25)
        Me.tpEditICD.Name = "tpEditICD"
        Me.tpEditICD.Size = New System.Drawing.Size(806, 746)
        Me.tpEditICD.TabIndex = 2
        Me.tpEditICD.Text = "Edit ICDs"
        Me.tpEditICD.UseVisualStyleBackColor = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.Location = New System.Drawing.Point(39, 101)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(707, 611)
        Me.RichTextBox1.TabIndex = 10
        Me.RichTextBox1.Text = ""
        '
        'btnSaveParser
        '
        Me.btnSaveParser.Location = New System.Drawing.Point(528, 47)
        Me.btnSaveParser.Name = "btnSaveParser"
        Me.btnSaveParser.Size = New System.Drawing.Size(76, 29)
        Me.btnSaveParser.TabIndex = 9
        Me.btnSaveParser.Text = "Save"
        Me.btnSaveParser.UseVisualStyleBackColor = True
        '
        'btnLoadParser
        '
        Me.btnLoadParser.Location = New System.Drawing.Point(441, 47)
        Me.btnLoadParser.Name = "btnLoadParser"
        Me.btnLoadParser.Size = New System.Drawing.Size(76, 29)
        Me.btnLoadParser.TabIndex = 8
        Me.btnLoadParser.Text = "Browse"
        Me.btnLoadParser.UseVisualStyleBackColor = True
        '
        'tbLoadParser
        '
        Me.tbLoadParser.Location = New System.Drawing.Point(39, 50)
        Me.tbLoadParser.Name = "tbLoadParser"
        Me.tbLoadParser.Size = New System.Drawing.Size(387, 22)
        Me.tbLoadParser.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(36, 30)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(86, 17)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "Load Parser"
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.dgTOC)
        Me.TabPage1.Controls.Add(Me.Label8)
        Me.TabPage1.Controls.Add(Me.ListBox2)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.lbTOCs)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(806, 746)
        Me.TabPage1.TabIndex = 3
        Me.TabPage1.Text = "TOC Page"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(12, 383)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(78, 17)
        Me.Label8.TabIndex = 5
        Me.Label8.Text = "Output Xml"
        '
        'ListBox2
        '
        Me.ListBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.ItemHeight = 16
        Me.ListBox2.Location = New System.Drawing.Point(15, 403)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(222, 308)
        Me.ListBox2.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(245, 31)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Messages"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 31)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "TOC's"
        '
        'lbTOCs
        '
        Me.lbTOCs.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lbTOCs.FormattingEnabled = True
        Me.lbTOCs.ItemHeight = 16
        Me.lbTOCs.Location = New System.Drawing.Point(15, 51)
        Me.lbTOCs.Name = "lbTOCs"
        Me.lbTOCs.Size = New System.Drawing.Size(222, 324)
        Me.lbTOCs.Sorted = True
        Me.lbTOCs.TabIndex = 1
        '
        'dgTOC
        '
        Me.dgTOC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgTOC.Location = New System.Drawing.Point(252, 53)
        Me.dgTOC.Name = "dgTOC"
        Me.dgTOC.RowTemplate.Height = 24
        Me.dgTOC.Size = New System.Drawing.Size(535, 658)
        Me.dgTOC.TabIndex = 6
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(814, 803)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Compass Call Communications Parser (C3PO)"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.cmnuMessages.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.tpProcessCap.ResumeLayout(False)
        Me.tpProcessCap.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.tpGenICD.ResumeLayout(False)
        Me.tpGenICD.PerformLayout()
        Me.tpEditICD.ResumeLayout(False)
        Me.tpEditICD.PerformLayout()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.dgTOC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PropertiesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmnuMessages As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents SelectAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClearAllToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LoadProfileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FilenameToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tpProcessCap As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents rbLink16 As System.Windows.Forms.RadioButton
    Friend WithEvents rbNcct As System.Windows.Forms.RadioButton
    Friend WithEvents rbRfos As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents clbSelectedMsgs As System.Windows.Forms.CheckedListBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents rbNoMsg As System.Windows.Forms.RadioButton
    Friend WithEvents rbAllMsg As System.Windows.Forms.RadioButton
    Friend WithEvents rbSelectMsg As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnBrowseSources As System.Windows.Forms.Button
    Friend WithEvents lbSources As System.Windows.Forms.ListBox
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents bBrowseOutput As System.Windows.Forms.Button
    Friend WithEvents tbOutputPath As System.Windows.Forms.TextBox
    Friend WithEvents lOutputType As System.Windows.Forms.Label
    Friend WithEvents btnProcData As System.Windows.Forms.Button
    Friend WithEvents tpGenICD As System.Windows.Forms.TabPage
    Friend WithEvents tpEditICD As System.Windows.Forms.TabPage
    Friend WithEvents CreateNewDatabaseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LoadConfigurationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ATNToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CHATToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnProcessMsgs As System.Windows.Forms.Button
    Friend WithEvents btnBrowsePrsrsPath As System.Windows.Forms.Button
    Friend WithEvents tbXmlOutputpath As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lvMessages As System.Windows.Forms.ListView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSelectAll As System.Windows.Forms.Button
    Friend WithEvents btnClearAll As System.Windows.Forms.Button
    Friend WithEvents tbCurrentDb As System.Windows.Forms.TextBox
    Friend WithEvents OpenDatabaseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents btnSaveParser As System.Windows.Forms.Button
    Friend WithEvents btnLoadParser As System.Windows.Forms.Button
    Friend WithEvents tbLoadParser As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents lbTOCs As System.Windows.Forms.ListBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ListBox2 As System.Windows.Forms.ListBox
    Friend WithEvents dgTOC As System.Windows.Forms.DataGridView

End Class
