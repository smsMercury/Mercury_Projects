Imports System
Imports System.IO
Imports adox

Public Class CMasterDb

    Private WithEvents mAdoDatabase As CAccessDatabase

    'Master Db Tables
    'Private mMasterDbTOCTbl As String = "TOC"
    Private mMsgIdTbl As String = "MsgIds"
    Private mVarStructTbl As String = "VarStruct"


#Region "Public Methods"

    Public Sub New()

    End Sub

    'this method instantiates a new CAccessDatabase class
    Public Sub New(ByVal dbName As String)

        mAdoDatabase = New CAccessDatabase("Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                            "Data Source=" & dbName & ";" & _
                                            "User ID=Admin;" & _
                                            "Password=")
    End Sub

    Public Sub Dispose()
        mAdoDatabase.Dispose()
    End Sub

    Public Sub CreateDB(ByVal dbName As String)
        Dim cat As Catalog = New Catalog()

        Try
            cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & dbName & ";" & _
                       "Jet OLEDB:Engine Type=5")
        Catch ex As Exception
            MsgBox("Unable to create Database." & vbCrLf & ex.Message, MsgBoxStyle.OkOnly)
        End Try
    End Sub

    'this method creates a new Master Database
    Public Sub CreateTOCtable(ByVal tableName As String)
        Dim ds As New DataSet()
        Dim dtTOC As New DataTable(tableName)

        If mAdoDatabase.tableExists(tableName) Then
            OpenDB()
            mAdoDatabase.ExecuteNonQuery("DROP Table " & tableName)
            CloseDB()
        End If

        'Define TOC table
        dtTOC.Columns.Add("MsgTime", GetType(String))
        dtTOC.Columns.Add("PcapFile", GetType(String))
        dtTOC.Columns.Add("MsgTo", GetType(String))
        dtTOC.Columns.Add("MsgFrom", GetType(String))
        dtTOC.Columns.Add("MsgID", GetType(Integer))
        dtTOC.Columns.Add("MsgSize", GetType(String))
        dtTOC.Columns.Add("MsgName", GetType(String))
        dtTOC.Columns.Add("MsgType", GetType(String))

        ds.Tables.Add(dtTOC)

        'Add toc table to db
        mAdoDatabase.CreateTable(tableName, ds)
    End Sub

    'this method creates a new VarStruct table
    Public Sub CreateVarStructTbl()

        If Not mAdoDatabase.tableExists(Me.mVarStructTbl) Then
            Dim ds As New DataSet()
            Dim dtVS As New DataTable(Me.mVarStructTbl)

            'Define TOC table
            dtVS.Columns.Add("VarStructID", GetType(Int16))
            dtVS.Columns.Add("MsgId", GetType(Int16))
            dtVS.Columns.Add("FieldName", GetType(String))
            dtVS.Columns.Add("FieldSize", GetType(Int16))
            dtVS.Columns.Add("DataType", GetType(String))
            dtVS.Columns.Add("ConvType", GetType(Int16))
            dtVS.Columns.Add("FieldLabel", GetType(String))
            dtVS.Columns.Add("DASField", GetType(Int16))
            dtVS.Columns.Add("MultiEntry", GetType(Int16))
            dtVS.Columns.Add("MultiRecPtr", GetType(Int16))
            dtVS.Columns.Add("StructLevel", GetType(Int16))

            ds.Tables.Add(dtVS)

            'Add toc table to db
            mAdoDatabase.CreateTable(mVarStructTbl, ds)
        End If
    End Sub

    'this method creates a new Msg ID table
    Public Sub CreateMsgTbl()

        If Not mAdoDatabase.tableExists(Me.mMsgIdTbl) Then
            Dim ds As New DataSet
            Dim dtMsg As New DataTable(Me.mMsgIdTbl)

            'define msg tbl
            dtMsg.Columns.Add("MsgID", GetType(String))
            dtMsg.Columns.Add("MsgName", GetType(String))

            ds.Tables.Add(dtMsg)

            'add msg table to db
            mAdoDatabase.CreateTable(mMsgIdTbl, ds)
        End If
    End Sub

    Public Function CreateTables(ByVal ds As DataSet) As Boolean
        Dim tbl As DataTable

        For Each tbl In ds.Tables
            Try
                mAdoDatabase.CreateTable(tbl.TableName, ds)
            Catch ex As Exception
                If MsgBox("Unable to create Db table for " & tbl.TableName & vbCrLf & ex.Message & vbCrLf & _
                           "Do you want to overwrite table?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
                    mAdoDatabase.DeleteTable(tbl.TableName)
                    mAdoDatabase.CreateTable(tbl.TableName, ds)
                    Return False
                End If
                Return True
            End Try
        Next
        Return False
    End Function

    Public Sub GetDataset(ByRef ds As DataSet, ByVal tblName As String)

        mAdoDatabase.OpenDBConnection()
        ds = mAdoDatabase.GetDataSet("select * from " & tblName)
        mAdoDatabase.CloseDBConnection()
    End Sub

    Public Sub GetTocDataset(ByRef ds As DataSet, ByVal tblName As String)

        mAdoDatabase.OpenDBConnection()
        'ds = mAdoDatabase.GetDataSet("select MsgTime, MsgTo, MsgFrom, MsgLen, MsgId, MsgName from " & tblName)
        ds = mAdoDatabase.GetDataSet("select * from " & tblName)
        mAdoDatabase.CloseDBConnection()
    End Sub

    Public Function GetMsgIds() As DataSet
        Dim ds As DataSet = Nothing

        GetDataset(ds, Me.mMsgIdTbl)
        Return ds
    End Function

    Public Function GetMsgId(ByVal msg As String) As DataSet
        Dim ds As DataSet = Nothing

        mAdoDatabase.OpenDBConnection()
        ds = mAdoDatabase.GetDataSet("select * from " & Me.mMsgIdTbl & " where MsgName = '" & msg & "'")
        mAdoDatabase.CloseDBConnection()
        Return ds
    End Function

    Public Function GetMsgVarStruct(ByVal msgid As Integer) As DataSet
        Dim ds As DataSet = Nothing

        mAdoDatabase.OpenDBConnection()
        ds = mAdoDatabase.GetDataSet("select * from " & Me.mVarStructTbl & " where MsgId = " & msgid & " order by VarStructID")
        mAdoDatabase.CloseDBConnection()
        Return ds
    End Function

    Public Sub updateTOCTable(ByVal tableName As String, ByVal timeStamp As String, ByVal pcapFile As String, ByVal msgTo As String, ByVal msgFrom As String, ByVal msgID As String, ByVal msgSize As String, ByVal msgName As String, ByVal msgType As String)
        Dim str As String = "Insert INTO " & tableName & " (MsgTime, PcapFile, MsgTo, MsgFrom, MsgID, MsgSize, MsgName, MsgType) " & _
                                   "VALUES ('" & timeStamp & "', '" & pcapFile & "', '" & msgTo & "', '" & msgFrom & "', '" & msgID & "', '" & msgSize & "', '" & msgName & "', '" & msgType & "')"

        mAdoDatabase.ExecuteNonQuery(str)
    End Sub

    Public Sub updateVarStructTable(ByVal id As Int16, ByVal msgId As Int16, ByVal fName As String, ByVal fSize As Int16, ByVal dataType As String, ByVal convType As Int16, ByVal fLabel As String, ByVal dasField As Int16, ByVal multiEntry As Int16, ByVal multiRecPtr As Int16, ByVal structLevel As Int16)
        Dim str As String = "Insert INTO " & Me.mVarStructTbl & " (VarStructID, MsgId, FieldName, FieldSize, DataType, ConvType, FieldLabel, DASField, MultiEntry, MultiRecPtr, StructLevel) " & _
                            "VALUES ('" & id & "', '" & msgId & "', '" & fName & "', '" & fSize & "', '" & dataType & "', '" & convType & "', '" & fLabel & "', '" & dasField & "', '" & multiEntry & "', '" & multiRecPtr & "', '" & structLevel & "')"

        mAdoDatabase.ExecuteNonQuery(str)
    End Sub

    Public Sub updateMsgTable(ByVal id As Integer, ByVal name As String)
        Dim str As String = "Insert INTO " & Me.mMsgIdTbl & " (MsgID, MsgName) " & _
                            "VALUES ('" & id & "', '" & name & "')"

        mAdoDatabase.ExecuteNonQuery(str)
    End Sub

    Public Sub UpdateTable(ByVal QryStr As String)

        mAdoDatabase.ExecuteNonQuery(QryStr)
    End Sub

    Public Sub MergeMsgIds(ByVal tocTable As String)
        Dim sql1 As String = "Insert INTO distinctTOC Select Distinct * from " & tocTable
        Dim sql2 As String = "Insert INTO tempTOC (MsgTime, PcapFile, MsgTo, MsgFrom, MsgID, MsgSize, MsgName, MsgType) SELECT " & _
                                   "distinctTOC.MsgTime, " & _
                                   "distinctTOC.PcapFile, " & _
                                   "distinctTOC.MsgTo, " & _
                                   "distinctTOC.MsgFrom, " & _
                                   "distinctTOC.MsgID, " & _
                                   "distinctTOC.MsgSize, " & _
                                   mMsgIdTbl & ".MsgName, " & _
                                   "distinctTOC.MsgType " & _
                                   "from distinctTOC, " & mMsgIdTbl & " WHERE distinctTOC.MsgID=" & mMsgIdTbl & ".MsgID"

        If tocTable = "" Then Exit Sub
        Try
            CreateTOCtable("distinctTOC")
            CreateTOCtable("tempTOC")
            OpenDB()
            mAdoDatabase.ExecuteNonQuery(sql1)
            mAdoDatabase.ExecuteNonQuery(sql2)
            mAdoDatabase.ExecuteNonQuery("DELETE * From " & tocTable)
            mAdoDatabase.ExecuteNonQuery("Insert INTO " & tocTable & " Select * from tempTOC")
            mAdoDatabase.ExecuteNonQuery("DROP Table distinctTOC, tempTOC")
            CloseDB()
        Catch ex As Exception
            MsgBox("Error updating TOC table." & vbCrLf & ex.Message)
        End Try
    End Sub

    Public Function GetTOCtables() As ArrayList

        Return mAdoDatabase.TOCtables
    End Function

    Public Sub OpenDB()
        mAdoDatabase.OpenDBConnection()
    End Sub

    Public Sub CloseDB()
        mAdoDatabase.CloseDBConnection()
    End Sub

#End Region

#Region "Private Methods"

    Private Sub mAdoDatabase_ErrorRaised(ByVal ErrorSource As String, ByVal ErrorMessage As String) Handles mAdoDatabase.ErrorRaised

        MsgBox("Database Error.  Error Source:  " & ErrorSource & vbCrLf & _
               "                 Error Message: " & ErrorMessage)
    End Sub

#End Region

End Class
