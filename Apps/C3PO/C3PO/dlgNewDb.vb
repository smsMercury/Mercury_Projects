Imports System.Windows.Forms

Public Class dlgNewDb
    Private mDbName As String = ""
    Private CurrentDb As CMasterDb

    Private mInitialDir As String = "c:\mercury"
    Private mPropFileName As String = "\VarStructProperties.txt"

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim readHdr As New CReadHeader

        If CurrentDb Is Nothing Then
            MsgBox("You must open a Database first.")
            Exit Sub
        End If
        If (mDbName <> "") And (tbHdr.Text <> "") Then
            readHdr.readHeader(tbHdr.Text, CurrentDb)
            'LoadMessageList()
        End If

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnBrowseDb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseDb.Click
        Dim dlg As New SaveFileDialog

        dlg.InitialDirectory = mInitialDir
        dlg.Filter = "Access Database (*.mdb)|*.mdb"
        dlg.AddExtension = True
        Try
            If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
                mDbName = dlg.FileName
                If System.IO.File.Exists(mDbName) Then
                    System.IO.File.Delete(mDbName)
                End If
                CurrentDb = New CMasterDb(mDbName)
                CurrentDb.CreateDB(mDbName)
                Me.tbDatabase.Text = mDbName
            Else
                mDbName = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnBrowseHdr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowseHdr.Click
        Dim dlg As New OpenFileDialog

        dlg.InitialDirectory = mInitialDir
        dlg.Filter = "Header File (*.hdr2)|*.hdr2"
        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
            Me.tbHdr.Text = dlg.FileName
        End If
    End Sub

#Region "Public Members"

    Public ReadOnly Property MasterDb() As CMasterDb
        Get
            Return Me.CurrentDb
        End Get
    End Property

    Public ReadOnly Property DbPath() As String
        Get
            Return mDbName
        End Get
    End Property
#End Region
End Class
