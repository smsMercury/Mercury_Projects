Imports System.Windows.Forms

Public Class dlgProperties

    Private mProps As COutputProps

    Public Sub New(ByRef props As COutputProps)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mProps = props
    End Sub

    Private Sub dlgProperties_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.tbCmuRfosHdr.Text = mProps.CmuRfosHdrScript
        Me.tbNcctHdr.Text = mProps.NcctHdrScript
        Me.tbXmlScripts.Text = mProps.XmlPath

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        mProps.CmuRfosHdrScript = Me.tbCmuRfosHdr.Text
        mProps.NcctHdrScript = Me.tbNcctHdr.Text
        mProps.XmlPath = Me.tbXmlScripts.Text

        mProps.UpdateProps()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnBrwse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrwseScriptsDir.Click
        Dim dlg As New FolderBrowserDialog

        dlg.SelectedPath = "c:\Mercury"
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            'save path
            mProps.XmlPath = dlg.SelectedPath
            Me.tbXmlScripts.Text = mProps.XmlPath
        End If
    End Sub

    Private Sub btnBrowseCmuRfosHdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseCmuRfosHdr.Click
        Dim dlg As New OpenFileDialog

        dlg.InitialDirectory = "c:\Mercury"
        dlg.Filter = "Cmu/Rfos header script (*.xml)|*.xml"
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            mProps.CmuRfosHdrScript = dlg.FileName
            Me.tbCmuRfosHdr.Text = mProps.CmuRfosHdrScript
        End If
    End Sub

    Private Sub btnBrowseNcctHdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseNcctHdr.Click
        Dim dlg As New OpenFileDialog

        dlg.InitialDirectory = "c:\Mercury"
        dlg.Filter = "Ncct header script (*.xml)|*.xml"
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            mProps.NcctHdrScript = dlg.FileName
            Me.tbNcctHdr.Text = mProps.NcctHdrScript
        End If
    End Sub


End Class
