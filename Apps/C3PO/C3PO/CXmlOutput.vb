Imports System.xml

Public Class CXmlOutput

    Private mOutputPath As String = ""
    Private mOutputFile As String = ""

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

End Class
