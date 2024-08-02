Imports System.Windows.Forms

Public Class DlgXmlStructVar

    Private mFieldName As String = ""

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        mFieldName = Me.lbFieldNames.SelectedItem.ToString

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click

        mFieldName = ""
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub DlgXmlStructVar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New(ByVal sName As String, ByVal FieldNames As ArrayList)

        InitializeComponent()

        Me.tbStruct.Text = sName
        Me.lbFieldNames.DataSource = FieldNames
    End Sub

    Public ReadOnly Property getFieldName() As String
        Get
            Return mFieldName
        End Get
    End Property
End Class
