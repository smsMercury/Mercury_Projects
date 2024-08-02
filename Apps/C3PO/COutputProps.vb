Public Class COutputProps

    Private mPropsName As String = "\pcapEngineProperties.txt"
    Private mAppPath As String = IO.Directory.GetCurrentDirectory

    Private mXmlPath As String = ""
    Private mDbPath As String = ""
    Private mPropsPath As String = mAppPath & mPropsName
    Private mOutputPath As String = ""
    Private mOutputFile As String = ""
    Private mMsgList As New ArrayList
    Private mSourceList As New ArrayList
    Private mFilterPortStr As String = ""
    Private mFilterTypeStr As String = ""
    Private mCmuRfosHdrScript As String = ""
    Private mNcctHdrScript As String = ""

    Public Sub UpdateProps()
        If mPropsPath <> "" Then
            Dim wtr As New IO.StreamWriter(mPropsPath)
            wtr.WriteLine(mXmlPath)
            wtr.WriteLine(mCmuRfosHdrScript)
            wtr.WriteLine(mNcctHdrScript)
            'wtr.WriteLine(mDbPath)
            wtr.Close()
        End If
    End Sub

    Public Sub ReadProps()
        Dim path As String

        Try
            If IO.File.Exists(mAppPath & mPropsName) Then
                Dim rdr As New IO.StreamReader(mAppPath & mPropsName)
                'read xmlpath
                path = rdr.ReadLine()
                If path <> "" Then
                    mXmlPath = path
                Else
                    mXmlPath = ""
                End If
                'read cmurfoshdr
                path = rdr.ReadLine()
                If path <> "" Then
                    mCmuRfosHdrScript = path
                Else
                    mCmuRfosHdrScript = ""
                End If
                'read nccthdr
                path = rdr.ReadLine()
                If path <> "" Then
                    mNcctHdrScript = path
                Else
                    mNcctHdrScript = ""
                End If
                'read dbpath
                'path = rdr.ReadLine()
                'If path <> "" Then
                '    mDbPath = path
                'Else
                '    mDbPath = ""
                'End If

                rdr.Close()
            Else
                MsgBox("Please edit properties and select the path where Xml Scripts reside.")
            End If
        Catch ex As Exception
            'rdr.Close()
        End Try
    End Sub

    Public Property XmlPath() As String
        Get
            Return mXmlPath
        End Get
        Set(ByVal value As String)
            mXmlPath = value
        End Set
    End Property

    Public Property OutputPath() As String
        Get
            Return mOutputPath
        End Get
        Set(ByVal value As String)
            mOutputPath = value
        End Set
    End Property

    Public Property OutputFile() As String
        Get
            Return mOutputFile
        End Get
        Set(ByVal value As String)
            mOutputFile = value
        End Set
    End Property

    Public Property MsgList() As ArrayList
        Get
            Return mMsgList
        End Get
        Set(ByVal value As ArrayList)
            mMsgList = value
        End Set
    End Property

    Public Property SourceList() As ArrayList
        Get
            Return mSourceList
        End Get
        Set(ByVal value As ArrayList)
            mSourceList = value
        End Set
    End Property

    Public Property FilterPortString() As String
        Get
            Return mFilterPortStr
        End Get
        Set(ByVal value As String)
            mFilterPortStr = value
        End Set
    End Property

    Public Property FilterTypeString() As String
        Get
            Return mFilterTypeStr
        End Get
        Set(ByVal value As String)
            mFilterTypeStr = value
        End Set
    End Property

    Public Property CmuRfosHdrScript() As String
        Get
            Return mCmuRfosHdrScript
        End Get
        Set(ByVal value As String)
            mCmuRfosHdrScript = value
        End Set
    End Property

    Public Property NcctHdrScript() As String
        Get
            Return mNcctHdrScript
        End Get
        Set(ByVal value As String)
            mNcctHdrScript = value
        End Set
    End Property

    Public Property DbPath() As String
        Get
            Return mDbPath
        End Get
        Set(ByVal value As String)
            mDbPath = value
        End Set
    End Property
End Class
