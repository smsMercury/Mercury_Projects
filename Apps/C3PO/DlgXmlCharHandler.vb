Imports System.Windows.Forms

Public Class DlgXmlCharHandler

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DlgXmlCharHandler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New(ByVal row As DataRow)
        InitializeComponent()

        Me.tbFieldName.Text = row("FieldName")
        Me.tbFieldSize.Text = row("FieldSize")
        Me.tbMultiEntry.Text = row("MultiEntry")
    End Sub
End Class
