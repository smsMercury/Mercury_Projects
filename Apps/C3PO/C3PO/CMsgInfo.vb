Public Class CMsgInfo

    'Public Enum MsgTypes
    '    rfos
    '    ncct
    '    link16
    'End Enum

    'Private mMsgType As MsgTypes
    Private mMsgName As String = ""
    Private mMsgId As String = ""
    Private mScript As String = ""
    Private mPort As String = ""

    'Public Property MsgType() As MsgTypes
    '    Get
    '        Return mMsgType
    '    End Get
    '    Set(ByVal value As MsgTypes)
    '        mMsgType = value
    '    End Set
    'End Property

    Public Property MsgName() As String
        Get
            Return mMsgName
        End Get
        Set(ByVal value As String)
            mMsgName = value
        End Set
    End Property

    Public Property MsgId() As String
        Get
            Return mMsgId
        End Get
        Set(ByVal value As String)
            mMsgId = value
        End Set
    End Property

    Public Property Script() As String
        Get
            Return mScript
        End Get
        Set(ByVal value As String)
            mScript = value
        End Set
    End Property

    Public Property Port() As String
        Get
            Return mPort
        End Get
        Set(ByVal value As String)
            mPort = value
        End Set
    End Property

End Class
