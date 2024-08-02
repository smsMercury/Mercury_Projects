Imports System
Imports System.Data
Imports System.Data.OleDb


Public Class CAccessDatabase

#Region "Variable Declaration"
    '=========================================================================================
    ' Variable Declarations...
    '=========================================================================================
    Dim m_ConnectionString As String
    Dim m_QueryString As String
    Dim m_Connection As OleDbConnection
    Dim m_DataSet As DataSet
    Dim m_OleDbCommand As OleDbCommand
    Dim m_OutputParameterName As String
#End Region

#Region "Initialize and Finalize"
    '=========================================================================================
    ' Initialize and Finalize
    '=======================================================================
    Public Sub New()
        ' create new instance
    End Sub
    'New()
    '=======================================================================
    Public Sub New(ByVal ConnectionString As String)
        ' create new instance and set connection string value
        m_ConnectionString = ConnectionString
    End Sub
    'New(ByVal ConnectionString As String)
    '=======================================================================
    Public Sub New(ByVal DBConnection As OleDbConnection)
        ' create new instance and set connection to the passed connection
        m_Connection = DBConnection
    End Sub
    'Finalize()
    '=======================================================================
    Protected Overrides Sub Finalize()
        'cleanup
        Dispose()
        MyBase.Finalize()
    End Sub
    'Dispose()
    '=======================================================================
    Public Sub Dispose()
        'cleanup
        On Error Resume Next
        If Not m_OleDbCommand Is Nothing Then m_OleDbCommand.Dispose()
        If Not m_DataSet Is Nothing Then m_DataSet.Dispose()
        If Not m_Connection Is Nothing Then
            If m_Connection.State <> ConnectionState.Closed Then CloseDBConnection()
            m_Connection.Dispose()
        End If
    End Sub
    'Dispose()
    '=======================================================================
#End Region

#Region "Properties"
    '=========================================================================================
    ' Properties
    '=======================================================================
    Public Property ConnectionString() As String
        ' get and set the connection string to use
        Get
            Return m_ConnectionString
        End Get
        Set(ByVal Value As String)
            m_ConnectionString = Value
        End Set
    End Property      'ConnectionString() As String
    '=======================================================================
    Public Property QueryString() As String
        ' get and set the query string to use
        Get
            Return m_QueryString
        End Get
        Set(ByVal Value As String)
            m_QueryString = Value
        End Set
    End Property      'QueryString() As String
    '=======================================================================
    Public Property DataSet() As DataSet
        ' get and set the dataset to use
        Get
            Return m_DataSet
        End Get
        Set(ByVal Value As DataSet)
            m_DataSet = Value
        End Set
    End Property      'DataSet() As DataSet
    '=======================================================================
#End Region

#Region "Helper Functions and Routines"
    '=========================================================================================
    ' Helper Functions and Routines
    '=======================================================================
    Public Sub OpenDBConnection()
        ' connection doesn't exist so create and open it
        If m_Connection Is Nothing Then
            Try
                m_Connection = New OleDbConnection(m_ConnectionString)
                m_Connection.Open()
            Catch e As Exception
                ' error happened...
                RaiseEvent ErrorRaised(e.Source, e.Message)
            End Try
        Else
            ' connection is currently open so close and reopen...
            If m_Connection.State <> ConnectionState.Closed Then
                Try
                    m_Connection.Close()
                    m_Connection.Open()
                Catch e As Exception
                    'error happened...
                    RaiseEvent ErrorRaised(e.Source, e.Message)
                End Try
            Else
                ' open a connection
                Try
                    m_Connection.Open()
                Catch e As Exception
                    'error happened...
                    RaiseEvent ErrorRaised(e.Source, e.Message)
                End Try
            End If
        End If
    End Sub  'OpenDBConnection()
    '=======================================================================
    Public Sub CloseDBConnection()
        Try
            m_Connection.Close()
        Catch e As Exception
            ' error happened
            RaiseEvent ErrorRaised(e.Source, e.Message)
        End Try
    End Sub   'CloseDBConnection()
    '=======================================================================
#End Region

#Region "Events"
    '=========================================================================================
    ' Events
    '=======================================================================
    Public Event ErrorRaised(ByVal ErrorSource As String, ByVal ErrorMessage As String)
    '=======================================================================
#End Region

#Region "Data Access"
    '=======================================================================
    ' Data Access
    '=======================================================================

    Public Sub ExecuteNonQuery()
        ' execute a non query (INSERT, UPDATE or DELETE ) using the query string value stored in local property
        Dim MyCommand As New OleDbCommand(m_QueryString, m_Connection)
        Try
            MyCommand.CommandType = CommandType.Text
            OpenDBConnection()
            MyCommand.ExecuteNonQuery()
        Catch e As Exception
            'error happened...
            RaiseEvent ErrorRaised(e.Source, e.Message)
        End Try
        MyCommand.Dispose()
    End Sub   'ExecuteNonQuery()

    '=======================================================================
    Public Sub ExecuteNonQuery(ByVal QueryString As String)
        ' execute a non query (INSERT, UPDATE or DELETE ) using the query string passed in
        Dim MyCommand As New OleDbCommand(QueryString, m_Connection)
        Try
            MyCommand.CommandType = CommandType.Text
            MyCommand.ExecuteNonQuery()
        Catch e As Exception
            'error happened...
            RaiseEvent ErrorRaised(e.Source, e.Message)
        End Try
        MyCommand.Dispose()
    End Sub   'ExecuteNonQuery(ByVal QueryString As String)
    '=======================================================================
    Public Function GetDataSet() As DataSet
        ' returns a dataset value using query string stored in local property
        Dim DataAdapter As New OleDbDataAdapter(m_QueryString, m_Connection)
        Dim DataSet As New DataSet()
        If ValidQuery(QueryString) = True Then
            Try
                DataAdapter.Fill(DataSet)
                m_DataSet = DataSet
                GetDataSet = DataSet
            Catch e As Exception
                'error happened...
                RaiseEvent ErrorRaised(e.Source, e.Message)
            End Try
        End If
        DataAdapter.Dispose()
        DataSet.Dispose()
    End Function      'GetDataSet() As DataSet
    '=======================================================================
    Public Function GetDataSet(ByVal QueryString As String) As DataSet
        ' returns a dataset value using passed in query string
        Dim DataAdapter As New OleDbDataAdapter(QueryString, m_Connection)
        Dim DataSet As New DataSet()
        If ValidQuery(QueryString) = True Then
            Try
                DataAdapter.Fill(DataSet)
                m_DataSet = DataSet
                GetDataSet = DataSet
            Catch e As Exception
                'error happened...
                RaiseEvent ErrorRaised(e.Source, e.Message)
            End Try
        End If
        DataAdapter.Dispose()
        DataSet.Dispose()
    End Function      'GetDataSet(ByVal QueryString As String) As DataSet
    '=======================================================================
    Private Function GetDataReader() As OleDbDataReader
        ' returns a datareader using query string stored in local property
        Dim MyCommand As New OleDbCommand(m_QueryString, m_Connection)
        If ValidQuery(QueryString) = True Then
            Try
                GetDataReader = MyCommand.ExecuteReader()
            Catch e As Exception
                'error happened...
                RaiseEvent ErrorRaised(e.Source, e.Message)
            End Try
        End If
        MyCommand.Dispose()
    End Function      'GetDataReader() As OleDbDataReader
    '=======================================================================
    Private Function GetDataReader(ByVal QueryString As String) As OleDbDataReader
        ' returns a datareader using passed in query string
        Dim MyCommand As New OleDbCommand(QueryString, m_Connection)
        If ValidQuery(QueryString) = True Then
            Try
                GetDataReader = MyCommand.ExecuteReader()
            Catch e As Exception
                'error happened...
                RaiseEvent ErrorRaised(e.Source, e.Message)
            End Try
        End If
        MyCommand.Dispose()
    End Function      'GetDataReader(ByVal QueryString As String) As OleDbDataReader
    '=======================================================================
    Public Function ExecuteScalarQuery() As Object
        ' returns an value from database using query string stored in local property
        Dim MyCommand As New OleDbCommand(m_QueryString, m_Connection)
        If ValidQuery(QueryString) = True Then
            Try
                MyCommand.CommandType = CommandType.Text
                ExecuteScalarQuery = MyCommand.ExecuteScalar()
            Catch e As Exception
                'error happened...
                RaiseEvent ErrorRaised(e.Source, e.Message)
            End Try
        End If
        MyCommand.Dispose()
    End Function    'ExecuteScalarQuery() As Object
    '=======================================================================
    Public Function ExecuteScalarQuery(ByVal QueryString As String) As Object
        ' returns an value from database using passed in query string
        Dim MyCommand As New OleDbCommand(QueryString, m_Connection)
        If ValidQuery(QueryString) = True Then
            Try
                MyCommand.CommandType = CommandType.Text
                ExecuteScalarQuery = MyCommand.ExecuteScalar()
            Catch e As Exception
                'error happened...
                RaiseEvent ErrorRaised(e.Source, e.Message)
            End Try
        End If
        MyCommand.Dispose()
    End Function    'ExecuteScalarQuery(ByVal QueryString As String) As Object
    '=======================================================================
    Public Function GetRecordCount() As Integer
        ' returns a record count using query string stored in local property
        Dim QueryString As String = m_QueryString
        If ValidQuery(QueryString) = True Then
            Dim DataAdapter As OleDbDataAdapter = New OleDbDataAdapter(QueryString, m_Connection)
            Dim DataSet As New DataSet()
            Try
                DataAdapter.Fill(DataSet)
                GetRecordCount = DataSet.Tables(0).Rows.Count
            Catch e As Exception
                'error happened...
                RaiseEvent ErrorRaised(e.Source, e.Message)
            End Try
            DataSet.Dispose()
            DataAdapter.Dispose()
        End If
    End Function      'GetRecordCount() As Integer
    '=======================================================================
    Public Function GetRecordCount(ByVal QueryString As String) As Integer
        ' returns a record count using passed in query string
        If ValidQuery(QueryString) = True Then
            Dim DataAdapter As OleDbDataAdapter = New OleDbDataAdapter(QueryString, m_Connection)
            Dim DataSet As New DataSet()
            Try
                DataAdapter.Fill(DataSet)
                GetRecordCount = DataSet.Tables(0).Rows.Count
            Catch e As Exception
                'error happened...
                RaiseEvent ErrorRaised(e.Source, e.Message)
            End Try
            DataSet.Dispose()
            DataAdapter.Dispose()
        End If
    End Function      'GetRecordCount(ByVal QueryString As String) As Integer
    '=======================================================================
    Public Function ExecuteStoredProcedure(Optional ByRef ReturnValue As Object = Nothing) As Boolean
        ' executes a stored procedure and returns a boolean value (success or failure)...
        ' if ReturnValue is not nothing then sets the return value = the stored procedure OUTPUT parameter
        Try
            With m_OleDbCommand
                ExecuteStoredProcedure = CBool(.ExecuteNonQuery())
                ' if expecting a return set the return variable
                If Not ReturnValue Is Nothing Then ReturnValue = .Parameters(m_OutputParameterName).Value
            End With
        Catch e As Exception
            'error happened...
            RaiseEvent ErrorRaised(e.Source, e.Message)
        End Try
    End Function    'ExecuteStoredProcedure(Optional ByRef ReturnValue As Object = Nothing) As Boolean
    '=======================================================================
    Public Sub BuildNewOleDbCommand(ByVal StoredProcedureName As String)
        'creates or initializes our command object to use with stored procedures
        Try
            ' reset our variables or create if needed
            m_OutputParameterName = ""
            If m_OleDbCommand Is Nothing Then
                m_OleDbCommand = New OleDbCommand()
            Else
                ' if exists then try to cancel if it is doing something
                m_OleDbCommand.Cancel()
            End If
            ' set our OleDbcommand values
            With m_OleDbCommand
                .CommandText = StoredProcedureName
                .CommandType = CommandType.StoredProcedure
                .Connection = m_Connection
            End With
        Catch e As Exception
            'error happened...
            RaiseEvent ErrorRaised(e.Source, e.Message)
        End Try
    End Sub   'BuildNewOleDbCommand(ByVal StoredProcedureName As String)
    '=================================================================
    Public Sub AddParameterToOleDbCommand(ByVal DataType As SqlDbType, ByVal ParameterName As String, ByVal ParameterValue As Object, ByVal ParameterDirection As ParameterDirection)
        ' adds a parameter to our command object to use with stored procedure
        Dim MyParameter As OleDbParameter
        ' correct parameter name if needed
        If Mid(Trim(ParameterName), 1, 1) = "@" Then
            MyParameter = New OleDbParameter(ParameterName, DataType)
        Else
            MyParameter = New OleDbParameter("@" & ParameterName, DataType)
        End If
        ' set parameter value
        MyParameter.Value = ParameterValue
        ' set parameter direction...input, output etc.
        MyParameter.Direction = ParameterDirection
        ' if we will be returning a value then store the output parameters name for use when executing
        If ParameterDirection = ParameterDirection.Output Then m_OutputParameterName = MyParameter.ParameterName
        ' add the parameter to our OleDb Command
        m_OleDbCommand.Parameters.Add(MyParameter)
    End Sub   'AddParameterToOleDbCommand(ByVal DataType As SqlDbType, ByVal ParameterName As String, ByVal ParameterValue As Object, ByVal ParameterDirection As ParameterDirection)
    '=================================================================
    Private Function ValidQuery(ByVal QueryString As String) As Boolean
        ' determine if query is allowed to be run
        If InStr(UCase(QueryString), "INSERT ") = 0 And InStr(UCase(QueryString), "UPDATE ") = 0 And InStr(UCase(QueryString), "DELETE ") = 0 Then
            ValidQuery = True
        Else
            ValidQuery = False
            RaiseEvent ErrorRaised("Query String", "Insert, Updates and Deletes queries are not allowed")
        End If
    End Function      'ValidQuery(ByVal QueryString As String) As Boolean
    '=================================================================
#End Region

#Region "Table Creation"

    Public Sub DeleteTable(ByVal tblName As String)
        Dim Cat As ADOX.Catalog
        Dim cn As ADODB.Connection

        Cat = New ADOX.Catalog
        cn = New ADODB.Connection
        cn.Open(m_ConnectionString)
        Cat.ActiveConnection = cn
        Cat.Tables.Delete(tblName)
        Cat = Nothing
        cn.Close()
    End Sub

    Public Sub CreateTable(ByVal tblName As String, ByVal ds As DataSet)
        Dim Cat As ADOX.Catalog
        Dim cn As ADODB.Connection

        Cat = New ADOX.Catalog
        cn = New ADODB.Connection
        cn.Open(m_ConnectionString)
        Cat.ActiveConnection = cn

        Dim objTable As New ADOX.Table
        Dim col As DataColumn
        Dim str As String

        'Create the table
        objTable.Name = tblName
        For Each col In ds.Tables(tblName).Columns
            str = col.ColumnName.ToString
            Select Case col.DataType.Name
                Case "Int32", "Int16", "Byte", "UShort", "UInteger", "UInt32", "UInt16", "Double"
                    objTable.Columns.Append(str, ADOX.DataTypeEnum.adVarWChar, 40)
                    objTable.Columns.Item(str).Attributes = ADOX.ColumnAttributesEnum.adColNullable
                Case "DateTime"
                    objTable.Columns.Append(str, ADOX.DataTypeEnum.adVarWChar, 40)
                    objTable.Columns.Item(str).Attributes = ADOX.ColumnAttributesEnum.adColNullable
                Case "Boolean"
                    objTable.Columns.Append(str, ADOX.DataTypeEnum.adBoolean)
                    objTable.Columns.Item(str).Attributes = ADOX.ColumnAttributesEnum.adColNullable
                Case Else 'String
                    objTable.Columns.Append(str, ADOX.DataTypeEnum.adLongVarWChar)
                    objTable.Columns.Item(str).Attributes = ADOX.ColumnAttributesEnum.adColNullable
            End Select
        Next

        'Append the newly created table to the Tables Collection
        Try
            Cat.Tables.Append(objTable)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        ' clean up objects
        objTable = Nothing
        Cat = Nothing
        cn.Close()
    End Sub

    Public Function tableExists(ByVal sTable As String) As Boolean
        Dim oCat As ADOX.Catalog
        Dim cn As ADODB.Connection
        Dim oTable As ADOX.Table
        Dim bFoundTable As Boolean

        oCat = New ADOX.Catalog
        cn = New ADODB.Connection
        cn.Open(m_ConnectionString)
        oCat.ActiveConnection = cn
        bFoundTable = False
        For Each oTable In oCat.Tables
            If UCase(oTable.Name) = UCase(sTable) Then
                bFoundTable = True
                Exit For
            End If
        Next
        cn.Close()
        Return bFoundTable
    End Function

    Public Function TOCtables() As ArrayList
        Dim oCat As ADOX.Catalog
        Dim cn As ADODB.Connection
        Dim oTable As ADOX.Table
        TOCtables = New ArrayList

        oCat = New ADOX.Catalog
        cn = New ADODB.Connection
        cn.Open(m_ConnectionString)
        oCat.ActiveConnection = cn
        For Each oTable In oCat.Tables
            If oTable.Name.Contains("_TOC") Then
                TOCtables.Add(oTable.Name)
            End If
        Next
        cn.Close()
        Return TOCtables
    End Function

#End Region
End Class

