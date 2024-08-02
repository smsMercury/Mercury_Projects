<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class dlgNewDb
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
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnBrowseHdr = New System.Windows.Forms.Button
        Me.tbHdr = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnBrowseDb = New System.Windows.Forms.Button
        Me.tbDatabase = New System.Windows.Forms.TextBox
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
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(306, 160)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
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
        Me.OK_Button.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(89, 28)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "Create DB"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(101, 4)
        Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(89, 28)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(29, 79)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(99, 17)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Open HDR file"
        '
        'btnBrowseHdr
        '
        Me.btnBrowseHdr.Location = New System.Drawing.Point(430, 99)
        Me.btnBrowseHdr.Name = "btnBrowseHdr"
        Me.btnBrowseHdr.Size = New System.Drawing.Size(70, 29)
        Me.btnBrowseHdr.TabIndex = 13
        Me.btnBrowseHdr.Text = "Browse"
        Me.btnBrowseHdr.UseVisualStyleBackColor = True
        '
        'tbHdr
        '
        Me.tbHdr.Location = New System.Drawing.Point(28, 102)
        Me.tbHdr.Name = "tbHdr"
        Me.tbHdr.Size = New System.Drawing.Size(381, 22)
        Me.tbHdr.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(29, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(143, 17)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Select New Database"
        '
        'btnBrowseDb
        '
        Me.btnBrowseDb.Location = New System.Drawing.Point(430, 42)
        Me.btnBrowseDb.Name = "btnBrowseDb"
        Me.btnBrowseDb.Size = New System.Drawing.Size(70, 29)
        Me.btnBrowseDb.TabIndex = 10
        Me.btnBrowseDb.Text = "Browse"
        Me.btnBrowseDb.UseVisualStyleBackColor = True
        '
        'tbDatabase
        '
        Me.tbDatabase.Location = New System.Drawing.Point(28, 45)
        Me.tbDatabase.Name = "tbDatabase"
        Me.tbDatabase.Size = New System.Drawing.Size(381, 22)
        Me.tbDatabase.TabIndex = 9
        '
        'dlgNewDb
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(517, 211)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnBrowseHdr)
        Me.Controls.Add(Me.tbHdr)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnBrowseDb)
        Me.Controls.Add(Me.tbDatabase)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dlgNewDb"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Create Database"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents btnBrowseHdr As System.Windows.Forms.Button
    Friend WithEvents tbHdr As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnBrowseDb As System.Windows.Forms.Button
    Friend WithEvents tbDatabase As System.Windows.Forms.TextBox

End Class
