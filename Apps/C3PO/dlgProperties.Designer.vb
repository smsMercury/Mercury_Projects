<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgProperties
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbXmlScripts = New System.Windows.Forms.TextBox
        Me.btnBrwseScriptsDir = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnBrowseCmuRfosHdr = New System.Windows.Forms.Button
        Me.tbCmuRfosHdr = New System.Windows.Forms.TextBox
        Me.btnBrowseNcctHdr = New System.Windows.Forms.Button
        Me.tbNcctHdr = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(313, 208)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(195, 36)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.Location = New System.Drawing.Point(4, 4)
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(4)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(89, 28)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(101, 4)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(4)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(89, 28)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(25, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(178, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Select Xml scripts directory"
        '
        'tbXmlScripts
        '
        Me.tbXmlScripts.Location = New System.Drawing.Point(28, 44)
        Me.tbXmlScripts.Name = "tbXmlScripts"
        Me.tbXmlScripts.Size = New System.Drawing.Size(388, 22)
        Me.tbXmlScripts.TabIndex = 2
        '
        'btnBrwseScriptsDir
        '
        Me.btnBrwseScriptsDir.Location = New System.Drawing.Point(437, 41)
        Me.btnBrwseScriptsDir.Name = "btnBrwseScriptsDir"
        Me.btnBrwseScriptsDir.Size = New System.Drawing.Size(66, 29)
        Me.btnBrwseScriptsDir.TabIndex = 3
        Me.btnBrwseScriptsDir.Text = "Browse"
        Me.btnBrwseScriptsDir.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(29, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(210, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Select CMU/RFOS header script"
        '
        'btnBrowseCmuRfosHdr
        '
        Me.btnBrowseCmuRfosHdr.Location = New System.Drawing.Point(437, 96)
        Me.btnBrowseCmuRfosHdr.Name = "btnBrowseCmuRfosHdr"
        Me.btnBrowseCmuRfosHdr.Size = New System.Drawing.Size(66, 30)
        Me.btnBrowseCmuRfosHdr.TabIndex = 6
        Me.btnBrowseCmuRfosHdr.Text = "Browse"
        Me.btnBrowseCmuRfosHdr.UseVisualStyleBackColor = True
        '
        'tbCmuRfosHdr
        '
        Me.tbCmuRfosHdr.Location = New System.Drawing.Point(28, 100)
        Me.tbCmuRfosHdr.Name = "tbCmuRfosHdr"
        Me.tbCmuRfosHdr.Size = New System.Drawing.Size(388, 22)
        Me.tbCmuRfosHdr.TabIndex = 5
        '
        'btnBrowseNcctHdr
        '
        Me.btnBrowseNcctHdr.Location = New System.Drawing.Point(437, 154)
        Me.btnBrowseNcctHdr.Name = "btnBrowseNcctHdr"
        Me.btnBrowseNcctHdr.Size = New System.Drawing.Size(66, 30)
        Me.btnBrowseNcctHdr.TabIndex = 9
        Me.btnBrowseNcctHdr.Text = "Browse"
        Me.btnBrowseNcctHdr.UseVisualStyleBackColor = True
        '
        'tbNcctHdr
        '
        Me.tbNcctHdr.Location = New System.Drawing.Point(28, 158)
        Me.tbNcctHdr.Name = "tbNcctHdr"
        Me.tbNcctHdr.Size = New System.Drawing.Size(388, 22)
        Me.tbNcctHdr.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(29, 138)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(175, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Select NCCT header script"
        '
        'dlgProperties
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(524, 259)
        Me.Controls.Add(Me.btnBrowseNcctHdr)
        Me.Controls.Add(Me.tbNcctHdr)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnBrowseCmuRfosHdr)
        Me.Controls.Add(Me.tbCmuRfosHdr)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnBrwseScriptsDir)
        Me.Controls.Add(Me.tbXmlScripts)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgProperties"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Properties"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbXmlScripts As System.Windows.Forms.TextBox
    Friend WithEvents btnBrwseScriptsDir As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnBrowseCmuRfosHdr As System.Windows.Forms.Button
    Friend WithEvents tbCmuRfosHdr As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowseNcctHdr As System.Windows.Forms.Button
    Friend WithEvents tbNcctHdr As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label

End Class
