Attribute VB_Name = "basDatabase"
' COPYRIGHT (C) 1999-2001, Mercury Solutions, Inc.
' MODULE:   basDatabase
' AUTHOR:   Tom Elkins
' PURPOSE:  Maintain database handling routines
' REVISION:
'   v1.3.0  TAE Replaced token calls with INI calls
'   v1.4.0  TAE Removed periodic error in requerying the database
'               Moved stored queries to another level in the tree
'           TAE Tied some message boxes to help topics and added help button
'   v1.5.0  TAE Changed the criteria used to determine if query nodes should be displayed
'           TAE Added a custom error message to terminate translation early
'           TAE Changed data structures to use date values instead of doubles
'           TAE Changed database schema to use date fields instead of text
'           TAE Changed date/time storage to use date/time directly instead of JDay/time
'           TAE Changed data record time storage to use full date and time
'           TAE Format date/time values to be compatible with old database schema
'           TAE Changed export time conversion to handle date/time values
'           TAE Changed the binding of the data grid to an Access 2000 database
'           TAE Added a command to move the data record pointer to the beginning
'               of the record set before exporting - this solves a rare bug that appeared when the
'               customer clicked on a record then exported to a CSV file - only the records after
'               the one the user clicked on were exported.
'           TAE Added the database version to the database info structure
'           TAE Added a routine to execute SQL action statements
'           TAE Modified database nodes to display a different icon depending on version
'           TAE Added a routine to re-map values in the database to new INI values
'           TAE Added a property to store whether the user is to be prompted or not
'           TAE Added a routine to upgrade old databases to the new schema
'           TAE Modified table name search to account for archive rename
'   v1.6.0  TAE Modified the way tables are referenced, so custom-named Archives can be used
'           TAE Modified code that accessed the old Archive options form to use the new Archive Wizard
'           TAE Modified the delete archive routine to force the Grid view to use a different data source
'               so it does not lock out the delete operation.
'           TAE Replaced the Create Summary Table routine with a function that returns a boolean status
'           TAE Replaced the Create Data Table routine with a function that returns a boolean status
'   v1.6.1  TAE Added verbose logging calls
'
Option Explicit
'
' Constants
Global Const TBL_INFO = "Info"
Global Const TBL_ARCHIVES = "Archives"
'+v1.5
Global Const TBL_SUMMARY = "_Summary"
Global Const TBL_DATA = "_Data"
'-v1.5
'+v1.7BB
Global Const TBL_PROC_DATA = "_ProcData"
Global Const TBL_VAR_STRUCT = "_VarStruct"
Global Const TBL_TOC = "_TOC"
Global Const TBL_MESSAGE = "_Message"
Global Const SEP_TOC_MSG = "^"
'-v1.7BB
Global Const SEP_ARCHIVE = "@"
Global Const SEP_MESSAGE = "#"
Global Const SEP_QUERY = "%" 'v1.5
Global Const TWIPS_PER_CHARACTER = 120      ' According to Microsoft documentation
'
' Error codes
Global Const DATABASE_ALREADY_EXISTS = 3204
Global Const TABLE_ALREADY_EXISTS = 3010
Global Const NO_SUCH_TABLE = 3265
Global Const DATABASE_READ_ONLY = 3051
Global Const NO_ERROR = 0
'+v1.5
'+v1.6TE
'Global Const USER_TERMINATED = 911 ' A custom error code to see if the user terminated the translation process early
'-v1.6
Global Const CURRENT_DB_VERSION = 4#
'-v1.5
'
' SQL Structure
Public Type SQL_INFO
    sFields As String       ' Field list
    sTable As String        ' Data table
    sFilter As String       ' Filter
    sOrder As String        ' Sort list
    sQuery As String        ' Complete query
End Type
'
' Message information
Public Type MESSAGE_INFO
    sMessage As String      ' Message name
    iId As Integer          ' Message ID
    lCount As Long          ' Number of messages
    dFirst As Double        ' First time
    dLast As Double         ' Last time
End Type
'
' Archive information
Public Type ARCHIVE_INFO
    dOffset_Time As Double  ' Julian day offset
    lNum_Messages As Long   ' Number of messages
    lNum_Bytes As Long      ' Number of bytes
    lFile_Size As Long      ' Size of the file
    '+v1.5
    ' Change the data structure to use Date values instead of doubles
    'dStart_Time As Double   ' Earliest time
    'dEnd_Time As Double     ' Latest time
    dtStart_Time As Date    ' Earliest time
    dtEnd_Time As Date      ' Latest time
    dtArchiveDate As Date   ' Start date for archive data
    '-v1.5
    rsSummary As Recordset  ' Pointer to the summary table
    rsData As Recordset     ' Pointer to the data table
    rsTOC As Recordset      ' Pointer to the Table of Contents
    rsVarStruct As Recordset 'Pointer to the Variable structure table
    rsProcData As Recordset      ' Pointer to the Processed data table
    rsMessage As Recordset      ' Pointer to the Message table
End Type
'
' Database info
Public Type DB_READY_INFO
    DB As Database              ' Pointer to the current database
    sName As String             ' Database file name
    iArchive As Integer         ' Current archive
    sArchive As String          ' Current archive name
    sMessage As String          ' Message name
    iMessage As Integer         ' Message ID
    uArchive As ARCHIVE_INFO    ' Archive information
    uSQL As SQL_INFO            ' SQL Query information
    fVersion As Single          ' Database version
End Type
'
' Global variables
Public guCurrent As DB_READY_INFO   ' Global database information
Public guPrevious As DB_READY_INFO  ' global database copy for paste functions
Private pbInteractive As Boolean    'v1.5 private state property
Private psLastQuery As String
Private psCurrentQuery As String
Private psNewQuery As String
'
' ROUTINE:  Set_Columns
' AUTHOR:   Tom Elkins
' PURPOSE:  Set up column headers in the ListView based on the specified table's structure
' INPUT:    "sDB_Name" is the name of the database
'           "sTable" is the name of a table within the database
Public Sub Set_Columns(sDB_Name As String, sTable As String)
    Dim dbCurrent As Database   ' Current database
    Dim fldCurrent As Field     ' Current Field
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Set_Columns (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sDB_Name & ", " & sTable
    End If
    '-v1.6.1
    '
    ' Trap errors
    On Error GoTo ERR_HANDLER
    '
    ' Remove any existing columns
    frmMain.lvListView.ColumnHeaders.Clear
    '
    ' Open the specified database
    Set dbCurrent = OpenDatabase(sDB_Name)
    '
    ' Check for the existence of the table
    If basDatabase.bTable_Exists(dbCurrent, sTable) Then
        '
        ' Use table-level addressing
        With dbCurrent.TableDefs(sTable)
            '
            ' Loop through all of the fields
            For Each fldCurrent In .Fields
                '
                ' Ignore fields that are autoincrementing or
                ' added by the system
                If ((fldCurrent.Attributes And dbAutoIncrField) Or (fldCurrent.Attributes And dbSystemField)) = 0 Then
                    '
                    ' Add a new column for the current field
                    frmMain.lvListView.ColumnHeaders.Add , fldCurrent.Name, fldCurrent.Name
                End If
            Next fldCurrent
        End With
    End If
    '
    ' Close the database
    dbCurrent.Close
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Set_Columns (End)"
    '-v1.6.1
    '
    Exit Sub
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basDatabase.Set_Columns (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "while reading database " & dbCurrent.Name, vbOKOnly, "Error Processing Database"
End Sub
'
' ROUTINE:  Display_Session_Databases
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the list view and display the contents of the Info table
'           of all connected databases.
' INPUT:    "nodSession" is the root "Session" node of the tree view.
' OUTPUT:   None
' NOTES:
Public Sub Display_Session_Databases(nodSession As Node)
    Dim nodDB As Node           ' Database node
    Dim dbCurrent As Database   ' Current database
    Dim rsTable As Recordset    ' Current table
    Dim liDB As ListItem        ' Current list item
    Dim iField As Integer       ' Current field
    Dim sToken As String
    Dim lTokenLen As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Session_Databases (Start)"
    '-v1.6.1
    '
    ' Remove all items from the list view
    frmMain.grdData.Visible = False
    frmMain.lvListView.Visible = True
    frmMain.lvListView.ListItems.Clear
    '
    ' Remove all columns from the list view
    frmMain.lvListView.ColumnHeaders.Clear
    '
    ' Check for the existence of children database nodes
    If nodSession.Children > 0 Then
        '
        ' Look at the first child node
        Set nodDB = nodSession.Child
        '
        ' Set the list view columns based on the "Info" table structure
        ' The name of the database is stored in the "Key" property of the database node
        basDatabase.Set_Columns nodDB.Key, "Info"
        '
        ' Add the classification column
        frmMain.lvListView.ColumnHeaders.Add , "Classification", "Classification"
        '
        '+v1.5
        ' Add database version
        frmMain.lvListView.ColumnHeaders.Add , "Version", "Version"
        '-v1.5
        '
        ' Process each database node
        'While nodDB <= nodDB.LastSibling
        While Not nodDB Is Nothing
            '
            ' Open the database
            ' The name of the database is stored in the Key property
            Set dbCurrent = OpenDatabase(nodDB.Key)
            '
            ' Open the Info table
            Set rsTable = dbCurrent.OpenRecordset("Info")
            '
            ' Add a new list item for the database with the following properties
            '   Index:      None, automatically assigned
            '   Key:        Database file name (same as Node key)
            '   Text:       Database name (same as Node text)
            '   Large Icon: Database icon
            '   Small Icon: Closed database icon
            Set liDB = frmMain.lvListView.ListItems.Add(, nodDB.Key, nodDB.Text, "DB", "DB_CLOSED")
            '
            ' Set the item type
            liDB.Tag = gsDATABASE
            '
            ' Populate the columns with the field values from the Info table
            ' We do not want to explicitly reference the fields by name, because
            ' the number of fields may change from version to version, and the
            ' field names may change as well; therefore, we access the fields collection
            ' of the table.  Since some of the fields may not be included as columns
            ' (due to system-added fields or index fields) we must use the current
            ' columns added earlier to get the appropriate fields from the table.
            ' The columns are numbered from 1, so the valid range is 1 to Count;
            ' however, we skip column 1, because it is the Text value for the Item
            ' created above.  The valid range is now 2 to Count.  Since we added
            ' the Classification column after all of the fields, we have to discount
            ' it because it does not exist in the table and would cause an error if
            ' we tried to access it.  So, the valid range is now 2 to Count-1
            '+v1.5
            ' Now we have added a database version column, which also does not exist
            ' in the table; therefore, we need to reduce the upper limit by 2
            'For iField = 2 To frmMain.lvListView.ColumnHeaders.Count - 1
            For iField = 2 To frmMain.lvListView.ColumnHeaders.Count - 2
            '-v1.5
                '
                ' Copy the field value to the subitem.
                ' Subitems are numbered from 0, so we have to subtract 1 from the
                ' field counter to align the fields
                If IsNull(rsTable.Fields(frmMain.lvListView.ColumnHeaders(iField).Key).Value) Then
                    liDB.SubItems(iField - 1) = ""
                Else
                    liDB.SubItems(iField - 1) = rsTable.Fields(frmMain.lvListView.ColumnHeaders(iField).Key).Value
                End If
            Next iField
            '
            ' Display the classification level
            '+v1.5
            'liDB.SubItems(frmMain.lvListView.ColumnHeaders.Count - 1) = frmSecurity.strGetAlias("Classification Text", "SECURITY_TXT" & dbCurrent.Properties("Security").Value, "UNKNOWN")
            liDB.SubItems(frmMain.lvListView.ColumnHeaders.Count - 2) = frmSecurity.strGetAlias("Classification Text", "SECURITY_TXT" & dbCurrent.Properties("Security").Value, "UNKNOWN")
            '
            ' Display the database version
            liDB.SubItems(frmMain.lvListView.ColumnHeaders.Count - 1) = dbCurrent.Version
            '-v1.5
            '
            ' Close the Info table
            rsTable.Close
            '
            ' Close the database
            dbCurrent.Close
            '
            ' Move to the next sibling database node
            ' You must use "Set" here; otherwise only the name changes,
            ' the other properties stay the same. and you end up with
            ' all of the Nodes having the name of the next node.
            Set nodDB = nodDB.Next
        Wend
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Session_Databases (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Add_Database_Node
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds a database node to the Tree View
' INPUT:    None
' OUTPUT:   None
' NOTES:    Adds the database specified by guCurrent.DB
Public Sub Add_Database_Node()
    Dim rsTable As Recordset    ' Pointer to the Info table record
    Dim nodDB As Node           ' Database node
    '
    ' Log the event
    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Database_Node: " & guCurrent.sName & " (Start)"
    '
    ' Check if the node already exists
    If Not frmMain.blnNodeExists(guCurrent.sName) Then
        '
        ' Open the Info table
        Set rsTable = guCurrent.DB.OpenRecordset(TBL_INFO)
        '
        ' Create the database node with the following properties
        '   Relative:       Session node
        '   Relationship:   Child
        '   Key:            Name of the database file
        '   Text:           Name stored in the Info table
        '   Default Icon:   Closed database
        '   Selected Icon:  Open database
        Set nodDB = frmMain.tvTreeView.Nodes.Add(gsSESSION, tvwChild, guCurrent.sName, "<NULL>", "DB_CLOSED", "DB_OPEN")
        nodDB.Sorted = True
        '
        ' Set the node name to the name of the table if it exists
        If Not IsNull(rsTable!Name) Then nodDB.Text = rsTable!Name
        '
        ' Set the Node type
        nodDB.Tag = gsDATABASE
        '
        ' Close the Info table
        rsTable.Close
    End If
    '
    '+v1.5
    ' Set the icon based on the DB version
    Set nodDB = frmMain.tvTreeView.Nodes(guCurrent.sName)
    nodDB.Image = "DB" & Int(guCurrent.DB.Version) & "_CLOSED"
    nodDB.SelectedImage = "DB" & Int(guCurrent.DB.Version) & "_OPEN"
    Set nodDB = Nothing
    '-v1.5
    '
    ' Open the Archives table
    Set rsTable = guCurrent.DB.OpenRecordset(TBL_ARCHIVES)
    '
    ' Loop through all of the archive records
    While Not rsTable.EOF
        '
        ' Add archive records as children nodes
        basDatabase.Add_Archive_Node rsTable
        '
        ' Move to the next record
        rsTable.MoveNext
    Wend
    '
    ' Close the Archives table
    rsTable.Close
    '
    '+v1.6.1TE
    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Database_Node (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: bValid_DB
' AUTHOR:   Tom Elkins
' PURPOSE:  Validates the compatibility of the selected database
' INPUT:    "dbCurrent" is the database object being checked
' OUTPUT:   "True" if the database is compatible
'           "False" if the database is not compatible
' NOTES:    When a database is created with CCAT, a property is added to the
'           database named "CCAT" of type Boolean, with the value "True".  If
'           That property exists in the specified database, then it is considered
'           valid.
Public Function bValid_DB(dbCurrent As Database) As Boolean
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : basDatabase.bValid_DB (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & dbCurrent.Name
    End If
    '-v1.6.1
    '
    ' Suppress error reporting
    On Error Resume Next
    '
    ' Check for the existence of the unique property
    bValid_DB = CBool(dbCurrent.Properties("CCAT").Value)
    '
    ' Check for an error
    If Err Then bValid_DB = False
    '
    ' Check the validity
    If Not bValid_DB Then
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : basDatabase.bValid_DB: Invalid database: " & dbCurrent.Name
        '
        ' Inform the user
        MsgBox "Selected Database is not a CCAT database." & vbCr & "Please select another file.", vbOKOnly Or vbMsgBoxHelpButton, "Invalid Database", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, basCCAT.IDH_GUI_FILE_OPEN
    End If
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.bValid_DB (End)"
    '-v1.6.1
    '
End Function
'
' ROUTINE:  Add_Archive_Node
' AUTHOR:   Tom Elkins
' PURPOSE:  Add archive records as nodes in the TreeView
' INPUT:    "rsArchive" is the record of the current archive from the Archives table
' OUTPUT:   None
' NOTES:    The current database (stored in guCurrent.DB) is used
'           The key value of an archive node is D@I, where
'               D is the key of the parent database node (the database file name)
'               @ is a delimiting string
'               I is the archive ID
Public Sub Add_Archive_Node(rsArchive As Recordset)
    Dim rsMessage As Recordset  ' Pointer to records in the Message table
    Dim rsSummary As Recordset  ' Pointer to records in the Summary table
    Dim nodArchive As Node      ' New archive node
    Dim blnTOC As Boolean
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Archive_node (Start)"
    '-v1.6.1
    '
    '+v1.5
    On Error GoTo Hell
    '-v1.5
    '
    ' Check for the existence of the node
    If Not frmMain.blnNodeExists(guCurrent.sName & SEP_ARCHIVE & rsArchive!ID) Then
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : basDatabase.Add_Archive_Node: " & guCurrent.sName & SEP_ARCHIVE & rsArchive!ID
        '
        ' Create a new node with the following attributes:
        '   Relative:       specified database node
        '   Relationship:   Child
        '   Key:            See NOTES above
        '   Text:           Name of the archive
        Set nodArchive = frmMain.tvTreeView.Nodes.Add(guCurrent.sName, tvwChild, _
            guCurrent.sName & SEP_ARCHIVE & rsArchive!ID, rsArchive!Name)
            nodArchive.Sorted = True
            '
        ' Store the node type
        nodArchive.Tag = gsARCHIVE
        '
        ' Specify the icons to use based on the media type
        Select Case rsArchive!Media
            '
            ' Hard disk
            Case "HD":
                nodArchive.Image = "HD_CLOSED"
                nodArchive.SelectedImage = "HD_OPEN"
            '
            ' Tape
            Case "TAPE":
                nodArchive.Image = "TAPE_CLOSED"
                nodArchive.SelectedImage = "TAPE_OPEN"
            '
            ' CD or DVD
            Case "CD", "DVD":
                nodArchive.Image = "CD_CLOSED"
                nodArchive.SelectedImage = "CD_OPEN"
        End Select
    Else
        Set nodArchive = frmMain.tvTreeView.Nodes(guCurrent.sName & SEP_ARCHIVE & rsArchive!ID)
    End If
    '
    ' Check for the existence of a summary table for this archive
    '+v1.5 TE
    If bTable_Exists(guCurrent.DB, rsArchive!Name & TBL_MESSAGE) Then
        Set rsMessage = guCurrent.DB.OpenRecordset(rsArchive!Name & TBL_MESSAGE, dbOpenDynaset)
        blnTOC = True
    Else
        blnTOC = False
    End If
    
    If bTable_Exists(guCurrent.DB, rsArchive!Name & TBL_SUMMARY) Then
    '-v1.5
        '
        ' Open the summary table
        '+v1.5 TE
        'Set rsMessage = guCurrent.DB.OpenRecordset("Archive" & rsArchive!ID & "_Summary")
        Set rsSummary = guCurrent.DB.OpenRecordset(rsArchive!Name & TBL_SUMMARY)
        '-v1.5
        '
        ' If there is any data, add the query nodes
        '+v1.5
        ' Old version checked for the number of messages in the archive table; however,
        ' if there was an error while processing or the process was terminated early, the
        ' message count was 0, and the queries were not added.
        'If rsArchive!Messages > 0 Then basDatabase.Add_Query_Nodes nodArchive.Key
        '
        ' Now, check the record count for the message summary table. If there is even a partial
        ' translation, there will be records here, so we can add the query nodes.
        If rsSummary.RecordCount > 0 Then basDatabase.Add_Query_Nodes nodArchive.Key
        '-v1.5
        '
        ' Loop through the message records
        While Not rsSummary.EOF

            '
            ' Add a message node to the Tree View
            '+v1.7BB
            ' Add a TOC message node to the Tree View
            If (blnTOC = True) Then
                basTOC.Add_TOC_Node nodArchive.Key, rsSummary
            End If
            
            'bb2004
            'rsMessage.FindFirst ("Msg_ID = " & rsSummary!Msg_id)
            '
            ' Look for a match
            'If ((rsMessage.NoMatch = False) And (rsMessage!Select_Msg = True)) Then
            If (frmWizard.IsSelected(rsSummary!MSG_ID)) Then
                basDatabase.Add_Message_Node nodArchive.Key, rsSummary
            End If
            '-v1.7BB
            ' Move to the next message
            rsSummary.MoveNext
        Wend
        '
        ' Close the summary table and message table
        rsSummary.Close
        If (blnTOC = True) Then
            rsMessage.Close
        End If
    End If
'
'+v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Archive_Node (End)"
    '-v1.6.1
    '
    Exit Sub

Hell:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basDatabase.Add_Archive_Node (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    basCCAT.WriteLogEntry "Error #" & Err.Number & " - " & Err.Description
    Debug.Print "Error #" & Err.Number & " - " & Err.Description
End Sub
'
' ROUTINE:  Add_Query_Nodes
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds token file queries as nodes in the TreeView
' INPUT:    "sArchive_Node" is the currently selected archive node
' OUTPUT:   None
' NOTES:    The current database (stored in guCurrent.DB) is used
'           The key value of a query node is D@I%Q, where
'               D is the key of the parent database node (the database file name)
'               @ is a delimiting string
'               I is the archive ID
'               % is a delimiting string
'               Q is the query number
Public Sub Add_Query_Nodes(sArchive_Node As String)
    Dim nodQuery As Node        ' New query node
    Dim iQuery As Integer       ' Current query
    Dim iNum_Queries As Integer ' Total number of queries
    Dim sQuery_Token As String  ' Token prefix for queries
    Dim lTokenLength As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Query_Nodes (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sArchive_Node
    End If
    '-v1.6.1
    '
    ' Add a query folder
    On Error GoTo 0
    If Not frmMain.blnNodeExists(sArchive_Node & "_Queries") Then
        '
        On Error Resume Next
        Set nodQuery = frmMain.tvTreeView.Nodes.Add(sArchive_Node, tvwChild, sArchive_Node & "_Queries", "Stored Queries", "ClosedBook", "OpenBook")
    End If
    '
    ' Get the number of queries and the token prefix
    iNum_Queries = basCCAT.GetNumber("Queries", "MAX_QUERIES", 0)
    '
    ' Loop through all of the queries
    For iQuery = 1 To iNum_Queries
        '
        ' See if the query node already exists
        If Not frmMain.blnNodeExists(sArchive_Node & SEP_QUERY & iQuery) Then
            '
            ' See if the query is available
            If basCCAT.GetAlias("Queries", "QUERY_TITLE" & iQuery, "NULL") <> "NULL" Then
                '
                ' Add the query to the Tree
                Set nodQuery = frmMain.tvTreeView.Nodes.Add(sArchive_Node & "_Queries", tvwChild, sArchive_Node & SEP_QUERY & iQuery, basCCAT.GetAlias("Queries", "QUERY_TITLE" & iQuery, "UNKNOWN_QUERY" & iQuery))
                '
                ' Modify some of the query attributes
                nodQuery.Tag = gsQUERY
                nodQuery.Image = gsQUERY
            End If
        End If
    Next iQuery
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Query_Nodes (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Add_Message_Node
' AUTHOR:   Tom Elkins
' PURPOSE:  Add a node in the TreeView to correspond to the specified Message record
' INPUT:    "sArchive" is the archive node's key
'           "rsMsg" is the current Message record
' OUTPUT:   None
' NOTES:    The Key value for a message node is A#M, where
'               A is the Archive node's key
'               # is a separator string
'               M is the message ID
Public Sub Add_Message_Node(sArchive As String, rsMsg As Recordset)
    Dim nodMsg As Node          ' New message node
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Message_Node (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sArchive
    End If
    '-v1.6.1
    '
    ' Check for the existence of the archive node
    If frmMain.blnNodeExists(sArchive) Then
        '
        ' Check for the existence of the message node
        If Not frmMain.blnNodeExists(sArchive & SEP_MESSAGE & rsMsg!MSG_ID) Then
            '
            ' Log the event
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : basDatabase.Add_Message_Node (" & sArchive & SEP_MESSAGE & rsMsg!MSG_ID & ")"
            '
            ' Create a new node with the following properties
            '   Relative:       specified archive node
            '   Relationship:   Child
            '   Key:            See NOTES above
            '   Text:           Message name
            Set nodMsg = frmMain.tvTreeView.Nodes.Add(sArchive, tvwChild, sArchive & SEP_MESSAGE & rsMsg!MSG_ID, rsMsg!Message)
            nodMsg.Sorted = True
            '
            ' Set the icons
            nodMsg.Image = "MSG_CLOSED"
            nodMsg.SelectedImage = "MSG_OPEN"
            '
            ' Save the node type
            nodMsg.Tag = gsMESSAGE
        End If
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Message_Node (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Display_Database_Archives
' AUTHOR:   Tom Elkins
' PURPOSE:  Displays the archive records in the ListView
' INPUT:    None, the current database in guCurrent.DB is used
' OUTPUT:   None
' NOTES:
Public Sub Display_Database_Archives()
    Dim rsTable As Recordset    ' Pointer to the Archive records
    Dim liArc As ListItem       ' New archive list item
    Dim iField As Integer       ' Field counter
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Database_Archives (Start)"
    '-v1.6.1
    '
    '
    ' Remove any items in the ListView
    frmMain.grdData.Visible = False
    frmMain.lvListView.Visible = True
    frmMain.lvListView.ListItems.Clear
    '
    ' Set up columns based on the structure of the Archives table
    basDatabase.Set_Columns guCurrent.sName, TBL_ARCHIVES
    '
    ' Position the progress bar
    frmMain.barLoad.Left = frmMain.sbStatusBar.Panels(1).Left
    frmMain.barLoad.Width = frmMain.sbStatusBar.Panels(1).Width
    frmMain.barLoad.Top = frmMain.sbStatusBar.Top
    frmMain.barLoad.Height = frmMain.sbStatusBar.Height
    '
    ' Set the pointer to the first record in the Archives table
    Set rsTable = guCurrent.DB.OpenRecordset(TBL_ARCHIVES)
    '
    ' Set the limits of the progress bar
    frmMain.barLoad.Min = 0
    frmMain.barLoad.Max = rsTable.RecordCount + 1
    frmMain.barLoad.Value = 0
    frmMain.barLoad.Visible = frmMain.sbStatusBar.Visible
    '
    ' Loop through the records
    While Not rsTable.EOF
        '
        ' Create a new list item with the following properties:
        '   Index:      none -- ListView automatically creates it
        '   Key:        D@I, where D is the parent database name,
        '                          @ is a separator string, and
        '                          I is the archive ID
        '   Name:       The name of the archive
        '   Large Icon: Default to the Hard Disk icon
        '   Small Icon: Default to the Hard Disk icon
        Set liArc = frmMain.lvListView.ListItems.Add(, guCurrent.sName & SEP_ARCHIVE & rsTable!ID, _
                "<NULL>", "HD", "HD_CLOSED")
        '
        ' Use the Archive name as the item name
        If Not IsNull(rsTable!Name) Then liArc.Text = rsTable!Name
        '
        ' Store the item type
        liArc.Tag = gsARCHIVE
        '
        ' Change icons for different media
        Select Case rsTable!Media
            Case "TAPE":
                liArc.Icon = "TAPE"
                liArc.SmallIcon = "TAPE_CLOSED"
            Case "HD":
                liArc.Icon = "HD"
                liArc.SmallIcon = "HD_CLOSED"
            Case "CD", "DVD":
                liArc.Icon = "CD"
                liArc.SmallIcon = "CD_CLOSED"
        End Select
        '
        ' Populate the subitems of the new archive item.  Since the column header names
        ' are the field names, use those to extract the data.
        For iField = 2 To frmMain.lvListView.ColumnHeaders.Count
            '
            ' Populate the column with record values
            If Not IsNull(rsTable.Fields(frmMain.lvListView.ColumnHeaders(iField).Text).Value) Then
                liArc.SubItems(iField - 1) = rsTable.Fields(frmMain.lvListView.ColumnHeaders(iField).Text).Value
            Else
                liArc.SubItems(iField - 1) = ""
            End If
        Next iField
        '
        ' Update the progress bar
        frmMain.barLoad.Value = frmMain.barLoad.Value + 1
        '
        ' Move to the next record
        rsTable.MoveNext
    Wend
    '
    ' Close the table
    rsTable.Close
    '
    ' Hide the progress bar
    frmMain.barLoad.Visible = False
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Database_Archives (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Display_Archive_Messages
' AUTHOR:   Tom Elkins
' PURPOSE:  Display an archive's messages in the ListView
' INPUT:    "sArchive" is the key of the selected archive node from the TreeView
' OUTPUT:   None
' NOTES:    The database is the current database set in guCurrent.DB
Public Sub Display_Archive_Messages(sArchive As String)
    Dim iArchive As Integer     ' ID of the selected Archive's record
    Dim rsMessage As Recordset  ' Archive summary records
    Dim liMsg As ListItem       ' Message list item
    Dim iField As Integer       ' Field index
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Archive_Messages (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sArchive
    End If
    '-v1.6.1
    '
    ' Remove any items from the ListView
    frmMain.grdData.Visible = False
    frmMain.lvListView.Visible = True
    frmMain.lvListView.ListItems.Clear
    '
    ' Position the progress bar
    frmMain.barLoad.Left = frmMain.sbStatusBar.Panels(1).Left
    frmMain.barLoad.Width = frmMain.sbStatusBar.Panels(1).Width
    frmMain.barLoad.Top = frmMain.sbStatusBar.Top
    frmMain.barLoad.Height = frmMain.sbStatusBar.Height
    '
    ' Extract the indexing information from the Node's Key.
    ' The Archive Node Key is "D@I", where "D" is the parent database node's index,
    ' and "I" is the selected archive's record ID.
    iArchive = CInt(Val(Mid(sArchive, InStr(1, sArchive, SEP_ARCHIVE) + 1)))
    '
    ' Check for the existence of the summary table
    '+v1.6TE
    'If bTable_Exists(guCurrent.DB, "Archive" & iArchive & TBL_SUMMARY) Then
    If bTable_Exists(guCurrent.DB, guCurrent.sArchive & TBL_SUMMARY) Then
    '-v1.6
        '
        ' Configure the ListView for the Archive Summary table
        ' The database name is in the Key property of the Database node.
        ' The database node is the parent to the archive node.
        ' Summary tables are named "Archive#_Summary", where # is the record ID for the archive
        '+v1.6TE
        'basDatabase.Set_Columns guCurrent.sName, "Archive" & iArchive & TBL_SUMMARY
        basDatabase.Set_Columns guCurrent.sName, guCurrent.sArchive & TBL_SUMMARY
        '-v1.6
        '
        ' Open the Summary table for the selected archive.
        '+v1.6TE
        'Set rsMessage = guCurrent.DB.OpenRecordset("Archive" & iArchive & TBL_SUMMARY)
        Set rsMessage = guCurrent.DB.OpenRecordset(guCurrent.sArchive & TBL_SUMMARY)
        '-v1.6
        '
        ' Set the limits of the progress bar
        If rsMessage.RecordCount > 0 Then
            frmMain.barLoad.Min = 0
            frmMain.barLoad.Max = rsMessage.RecordCount
            frmMain.barLoad.Value = 0
            frmMain.barLoad.Visible = frmMain.sbStatusBar.Visible
        End If
        '
        ' Loop through the messages
        Do While Not rsMessage.EOF
            '
            ' Add message node
            Set liMsg = frmMain.lvListView.ListItems.Add(, sArchive & SEP_MESSAGE & rsMessage!MSG_ID, rsMessage!Message, "MSG", "MSG_CLOSED")
            '
            ' Add message type identifier
            liMsg.Tag = gsMESSAGE
            '
            ' Add the details to the subitems.
            For iField = 2 To frmMain.lvListView.ColumnHeaders.Count
                '
                ' Populate the column with record values
                If IsNull(rsMessage.Fields(frmMain.lvListView.ColumnHeaders(iField).Text).Value) Then
                    liMsg.SubItems(iField - 1) = ""
                Else
                    liMsg.SubItems(iField - 1) = rsMessage.Fields(frmMain.lvListView.ColumnHeaders(iField).Text).Value
                End If
            Next iField
            '
            ' Format the times to human-readable text
            If IsNumeric(rsMessage!First) Then _
                liMsg.SubItems(3) = basCCAT.sTSecs_To_Human_Time(rsMessage!First)
            If IsNumeric(rsMessage!Last) Then _
                liMsg.SubItems(4) = basCCAT.sTSecs_To_Human_Time(rsMessage!Last)
            '
            ' Update the progress bar -- HACK
            If frmMain.barLoad.Value < frmMain.barLoad.Max Then frmMain.barLoad.Value = frmMain.barLoad.Value + 1
            '
            ' Move to the next message
            rsMessage.MoveNext
        Loop
        '
        ' Close the table
        rsMessage.Close
        '
        ' Hide the progress bar
        frmMain.barLoad.Visible = False
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Archive_Messages (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Display_Message_Details
' AUTHOR:   Tom Elkins
' PURPOSE:  Display a message's data in the ListView
' INPUT:    "nodMsg" is the selected message node from the TreeView
' OUTPUT:   None
' NOTES:
Public Sub Display_Message_Details(nodMsg As Node)
    Dim sToken As String
    Dim lTokenLen As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Message_Details (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & nodMsg.Text & " [" & nodMsg.Key & "]"
    End If
    '-v1.6.1
    '
    ' Assign the database to the data control
    frmMain.Data1.DatabaseName = guCurrent.sName
    '
    ' Position the data grid to fit the list view space
    frmMain.grdData.Left = frmMain.lvListView.Left
    frmMain.grdData.Top = frmMain.lvListView.Top
    frmMain.grdData.Width = frmMain.lvListView.Width
    frmMain.grdData.Height = frmMain.lvListView.Height
    '
    ' Hide the list view and show the grid
    frmMain.lvListView.Visible = False
    frmMain.grdData.Visible = True
    '
    ' Create the table name, which is ARCHIVE<Archive ID>_DATA
    '+v1.6TE
    'guCurrent.uSQL.sTable = "Archive" & guCurrent.iArchive & TBL_DATA
    guCurrent.uSQL.sTable = guCurrent.sArchive & TBL_DATA
    '-v1.6
    '
    ' Create a default SQL query
    ' Get a custom field list from the token file
    guCurrent.uSQL.sFields = basCCAT.GetAlias("Message Fields", "MSG_FIELDS" & guCurrent.iMessage, "*")
    '
    ' Set the filter to extract records for the selected message
    guCurrent.uSQL.sFilter = "Msg_Type = '" & guCurrent.sMessage & "'"
    '
    ' Remove the sort order
    guCurrent.uSQL.sOrder = ""
    '
    ' If the data table exists, execute the query
    '+v1.5
    'If bTable_Exists(guCurrent.DB, guCurrent.uSQL.sTable) Then basDatabase.Requery_Data
    If bTable_Exists(guCurrent.DB, guCurrent.uSQL.sTable) Then basDatabase.QueryData basDatabase.sCreate_SQL
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Message_Details (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Add_to_Cinnabar_Tables
' AUTHOR:   Shaun Vogel
' PURPOSE:  Add support tables to database to make them cinnabar compliant
' INPUT:    database name
' OUTPUT:   None
' NOTES:
Public Sub Add_to_Cinnabar_Tables(dbCurrent As Database, entry As String)
    Dim mrsData As Recordset

    Set mrsData = dbCurrent.OpenRecordset("Sources")
    '
    '
    With mrsData
        .AddNew
        .Fields("Source") = entry
        .Fields("Origin") = "CCAT"
        .Fields("Added") = DateTime.Time
        .Update
    End With

End Sub
'
' ROUTINE:  Add_Cinnabar_Tables
' AUTHOR:   Shaun Vogel
' PURPOSE:  Add support tables to database to make them cinnabar compliant
' INPUT:    database name
' OUTPUT:   None
' NOTES:
Public Sub Add_Cinnabar_Tables(dbCurrent As Database)
    Dim ptblNew As TableDef

    If bTable_Exists(dbCurrent, "Sources") Then Exit Sub
    '
    ' Create the Sources table
    Set ptblNew = New TableDef
    With ptblNew
        .Name = "Sources"
        .Fields.Append .CreateField("SrcID", dbLong)
        .Fields.Append .CreateField("Source", dbText, 255)
        .Fields.Append .CreateField("Origin", dbMemo)
        .Fields.Append .CreateField("Added", dbDate)
        .Fields.Append .CreateField("Comment", dbMemo)
        .Fields.Append .CreateField("Catalog", dbMemo)
        
        .Fields("Origin").AllowZeroLength = True
        .Fields("Added").AllowZeroLength = True
        .Fields("Comment").AllowZeroLength = True
        .Fields("Catalog").AllowZeroLength = True
        .Fields("SrcID").Attributes = dbAutoIncrField
    End With
    dbCurrent.TableDefs.Append ptblNew
    
    'Create the Diary table
    Set ptblNew = New TableDef
    With ptblNew
        .Name = "Diary"
        .Fields.Append .CreateField("EntryTime", dbDate)
        .Fields.Append .CreateField("EntrySrc", dbText, 255)
        .Fields.Append .CreateField("Entry", dbMemo)
        .Fields.Append .CreateField("Origin", dbMemo)
        .Fields.Append .CreateField("Source", dbText, 255)
        
        .Fields("EntryTime").AllowZeroLength = True
        .Fields("EntrySrc").AllowZeroLength = True
        .Fields("Entry").AllowZeroLength = True
        .Fields("Origin").AllowZeroLength = True
        .Fields("Source").AllowZeroLength = True
    End With
    dbCurrent.TableDefs.Append ptblNew
    
    'Create the Filters table
    Set ptblNew = New TableDef
    With ptblNew
        .Name = "Filters"
        .Fields.Append .CreateField("FilterName", dbText, 255)
        .Fields.Append .CreateField("Source", dbText, 255)
        .Fields.Append .CreateField("Query", dbMemo)
        
        .Fields("FilterName").AllowZeroLength = True
        .Fields("Source").AllowZeroLength = True
        .Fields("Query").AllowZeroLength = True
    End With
    dbCurrent.TableDefs.Append ptblNew

End Sub
'
' FUNCTION: bCreate_New_Database
' AUTHOR:   Tom Elkins
' PURPOSE:  Create a new database file and structures
' INPUT:    "sDB_Name" is the name for the new database
' OUTPUT:   "TRUE" if the new database is created
'           "FALSE" if the new database is not created
' NOTES:
Public Function bCreate_New_Database(sDB_Name As String) As Boolean
    Dim dbNew As Database   ' New database
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.bCreate_New_Database: " & sDB_Name & " (Start)"
    '
    ' Suppress error reporting
    On Error Resume Next
    '
    ' Create the new database file
    Set dbNew = DBEngine.Workspaces(0).CreateDatabase(sDB_Name, dbLangGeneral)
    '
    ' Check for errors
    Select Case Err.Number
        '
        ' Database already exists
        Case DATABASE_ALREADY_EXISTS:
            '
            ' Log the error
            basCCAT.WriteLogEntry "INFO     : basDatabase.bCreate_New_Database (Database already exists)"
            '
            ' Resume error reporting
            On Error GoTo 0
            '
            ' Check for a previously marked database file
            If Dir(sDB_Name & ".del") <> "" Then Kill sDB_Name & ".del"
            '
            ' if it is the currently opened database, close it and remove it from the list
            If Not guCurrent.DB Is Nothing Then
                If sDB_Name = guCurrent.DB.Name Then
                    guCurrent.DB.Close
                    basCCAT.Remove_Database sDB_Name
                End If
            End If
            '
            ' Log the event
            basCCAT.WriteLogEntry "INFO     : basDatabase.bCreate_New_Database (Renaming existing database to " & sDB_Name & ".del)"
            '
            ' Rename the old database temporarily.
            Name sDB_Name As sDB_Name & ".del"
            '
            ' Try again
            bCreate_New_Database = bCreate_New_Database(sDB_Name)
        '
        ' No error
        Case NO_ERROR:
            '
            ' Restore error reporting
            On Error GoTo 0
            '
            ' Create a property that is unique to translator-generated databases
            ' This property will be used to determine if a selected database is valid.
            dbNew.Properties.Append dbNew.CreateProperty("CCAT", dbBoolean, True)
            '
            ' Create a classification property with the default value as "UNCLASSIFIED"
            dbNew.Properties.Append dbNew.CreateProperty("Security", dbInteger, frmSecurity.lngGetNumber("Classification bit masks", "BIT_UNCLASSIFIED", 0))
            '
            ' Add cinnabar tables
            Call Add_Cinnabar_Tables(dbNew)
            '
            ' Create the Info table
            If basDatabase.bCreate_Info_Table(dbNew) Then
                '
                ' Create the Archives table
                basDatabase.Create_Archive_Table dbNew
                '
                ' Close the database
                dbNew.Close
                '
                ' Check for a previous database
                If Dir(sDB_Name & ".del") <> "" Then
                    '
                    ' Log the event
                    basCCAT.WriteLogEntry "INFO     : basDatabase.bCreate_New_Database (Deleting old database)"
                    '
                    ' Delete the old database.  Confirmation was verified at the file box
                    Kill sDB_Name & ".del"
                End If
                '
                ' Set the return value
                bCreate_New_Database = True
            '
            ' Info table could not be created, or the user cancelled the operation
            Else
                '
                ' Close the database
                dbNew.Close
                '
                ' Log the event
                basCCAT.WriteLogEntry "INFO     : basDatabase.bCreate_New_Database (Deleting new database)"
                '
                ' Delete the database
                Kill sDB_Name
                '
                ' Check for a previous database
                If Dir(sDB_Name & ".del") <> "" Then
                    '
                    ' Log the event
                    basCCAT.WriteLogEntry "INFO     : basDatabase.bCreate_New_Database (Restoring old database)"
                    '
                    ' Rename the file
                    Name sDB_Name & ".del" As sDB_Name
                End If
                '
                ' Set the return value
                bCreate_New_Database = False
            End If
        '
        ' Unexpected Error
        Case Else:
            '
            ' Log the event
            basCCAT.WriteLogEntry "ERROR    : basDatabase.bCreate_New_Database (Error #" & Err.Number & " - " & Err.Description & ")"
            '
            ' Inform the user
            MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "while creating database.", vbOKOnly Or vbCritical, "Error Creating Database"
            '
            ' Return false
            bCreate_New_Database = False
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.bCreate_New_Database (End)"
    '-v1.6.1
    '
End Function
'
' ROUTINE:  Open_Existing_Database
' AUTHOR:   Tom Elkins
' PURPOSE:  Validates the specified database and adds it to the Tree View
' INPUT:    "sDB_Name" is the name of the database being added
' OUTPUT:   None
' NOTES:
Public Sub Open_Existing_Database(sDB_Name As String)
    Dim iAttr As Integer    ' File attributes
    '
    ' Log the event
    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Open_Existing_Database: " & sDB_Name & " (Start)"
    '
    ' Check to see if the node already exists
    If Not frmMain.blnNodeExists(sDB_Name) Then
        '
        ' Suppress error reporting
        On Error Resume Next
        '
        ' Open the specified database file
        Set guCurrent.DB = OpenDatabase(sDB_Name)
        '
        ' Check for errors
        Select Case Err.Number
            '
            ' Error 3051 -- Table already open, or no permissions
            Case DATABASE_READ_ONLY:
                '
                ' Log the event
                basCCAT.WriteLogEntry "INFO     : basDatabase.Open_Existing_Database (Database is Read-Only.  Changing attributes)"
                '
                ' Restore error reporting
                On Error GoTo 0
                '
                ' Get the file's attributes
                iAttr = GetAttr(sDB_Name)
                '
                ' Check for the read-only attribute
                If iAttr And vbReadOnly Then
                    '
                    ' Ask the user if he wants the property changed
                    If MsgBox("The selected database is Read-Only." & vbCr & sDB_Name & vbCr & "Do you wish to make the file editable?", vbYesNo Or vbQuestion, "Error Opening Database") = vbYes Then
                        '
                        ' Disable the Read-only property for the file
                        SetAttr sDB_Name, iAttr Xor vbReadOnly
                        '
                        ' Try again
                        Open_Existing_Database (sDB_Name)
                    End If
                '
                ' Some other problem
                Else
                    '
                    ' Log the event
                    basCCAT.WriteLogEntry "INFO     : basDatabase.Open_Existing_Database (database already opened)"
                    '
                    ' Inform the user
                    MsgBox "Database is already open", vbOKOnly Or vbInformation, "Error Opening Database"
                End If
            '
            ' No error
            Case NO_ERROR:
                '
                ' Save the name
                guCurrent.sName = sDB_Name
                '
                '+v1.5
                ' Store the database version
                guCurrent.fVersion = guCurrent.DB.Version
                '-v1.5
                '
                ' Check for validity
                If bValid_DB(guCurrent.DB) Then
                    '
                    ' Update the classification
                    ' The classification for the database is stored in a property
                    ' added when the database was created.
                    frmSecurity.SetClassification guCurrent.DB.Properties("Security").Value, gsDATABASE
                    '
                    ' Check for the existence of the Info table and create if necessary
                    If Not bTable_Exists(guCurrent.DB, "Info") Then basDatabase.bCreate_Info_Table (guCurrent.DB)
                    '
                    ' Check for the existence of the Archives table and create if necessary
                    If Not bTable_Exists(guCurrent.DB, "Archives") Then basDatabase.Create_Archive_Table guCurrent.DB
                    '
                    ' Add a node to the tree view
                    basDatabase.Add_Database_Node
                '
                ' Invalid database
                Else
                    '
                    ' Close the database
                    guCurrent.DB.Close
                End If
            '
            ' Unexpected error
            Case Else:
                '
                ' Log the event
                basCCAT.WriteLogEntry "ERROR    : basDatabase.Open_Existing_Database (Error #" & Err.Number & " - " & Err.Description & " while opening " & sDB_Name & ")"
                '
                ' Inform the user
                MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "while opening existing database", vbOKOnly Or vbInformation, "Error Opening Database"
        End Select
    Else
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : basDatabase.Open_Existing_Database (database is already in the session)"
        '
        ' Inform the user the database is already being used
        MsgBox "Selected database is already in the session", vbOKOnly, "Duplicate database"
        '
        ' Close the database
        guCurrent.DB.Close
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Open_Existing_Database (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: bCreate_Info_Table
' AUTHOR:   Tom Elkins
' PURPOSE:  Creates a default Info table in the specified database, and lets the
'           user edit the information
' INPUT:    "dbCurrent" is the currently selected database
' OUTPUT:   "TRUE" if the table was created and populated
'           "FALSE" if the table was not created or not populated
' NOTES:
Public Function bCreate_Info_Table(dbCurrent As Database) As Boolean
    Dim tblInfo As TableDef     ' New Info table
    Dim rsInfo As Recordset     ' New record
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.Create_Info_Table: Database = " & dbCurrent.Name & " (Start)"
    '
    ' Suppress error reporting
    On Error Resume Next
    '
    ' Create the table
    Set tblInfo = dbCurrent.CreateTableDef(TBL_INFO)
    '
    ' Check for errors
    Select Case Err.Number
        '
        ' No errors
        Case NO_ERROR:
            '
            ' Resume error reporting
            On Error GoTo 0
            '
            ' Add the fields
            tblInfo.Fields.Append tblInfo.CreateField("Name", dbText, 50)
            '+v1.5
            ' Change database schema to use dates instead of text
            'tblInfo.Fields.Append tblInfo.CreateField("Start", dbText, 20)
            'tblInfo.Fields.Append tblInfo.CreateField("End", dbText, 20)
            tblInfo.Fields.Append tblInfo.CreateField("Start", dbDate)
            tblInfo.Fields.Append tblInfo.CreateField("End", dbDate)
            '-v1.5
            tblInfo.Fields.Append tblInfo.CreateField("Description", dbText, 255)
            '
            ' Set the field attributes to allow empty strings
            tblInfo.Fields("Name").AllowZeroLength = True
            tblInfo.Fields("Start").AllowZeroLength = True
            tblInfo.Fields("End").AllowZeroLength = True
            tblInfo.Fields("Description").AllowZeroLength = True
            '
            ' Add the table to the database
            dbCurrent.TableDefs.Append tblInfo
            '
            ' Create a new, blank record
            Set rsInfo = dbCurrent.OpenRecordset(TBL_INFO)
            rsInfo.AddNew
            rsInfo.Update
            rsInfo.Close
            '
            ' Allow the user to edit the contents
            '+v1.5
            If basDatabase.Interactive Then
                bCreate_Info_Table = frmDBInfo.blnEditInfoTable(dbCurrent)
            Else
                bCreate_Info_Table = True
            End If
            '-v1.5
        '
        ' Unexpected error
        Case Else:
            '
            ' Log the event
            basCCAT.WriteLogEntry "ERROR    : basDatabase.bCreate_Info_Table (Error #" & Err.Number & " - " & Err.Description & " while creating Info table)"
            '
            ' Inform the user
            MsgBox "ERROR #" & Err.Number & " - " & Err.Description & vbCr & "While creating Info table", vbOKOnly Or vbCritical, "Error Creating Table"
            '
            ' set the function to false
            bCreate_Info_Table = False
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.bCreate_Info_Table (End)"
    '-v1.6.1
    '
End Function
'
' ROUTINE:  Create_Archive_Table
' AUTHOR:   Tom Elkins
' PURPOSE:  Creates the default, blank archive table in the specified database
' INPUT:    "dbCurrent" is the currently selected database
' OUTPUT:   None
' NOTES:
Public Sub Create_Archive_Table(dbCurrent As Database)
    Dim tblArchive As TableDef  ' New Archive table
    '
    ' Log the event
    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Create_Archive_Table: Database = " & dbCurrent.Name & " (Start)"
    '
    ' Create the table
    Set tblArchive = dbCurrent.CreateTableDef(TBL_ARCHIVES)
    '
    ' Use table-level addressing
    With tblArchive
        '
        ' Add the fields
        .Fields.Append .CreateField("ID", dbLong)
        .Fields.Append .CreateField("Name", dbText, 50)
        '+v1.5
        ' Change database schema to use dates instead of text
        '.Fields.Append .CreateField("Date", dbText, 20)
        '.Fields.Append .CreateField("Start", dbText, 20)
        '.Fields.Append .CreateField("End", dbText, 20)
        .Fields.Append .CreateField("Date", dbDate)
        .Fields.Append .CreateField("Start", dbDate)
        .Fields.Append .CreateField("End", dbDate)
        '-v1.5
        .Fields.Append .CreateField("Archive", dbText, 255)
        .Fields.Append .CreateField("Media", dbText, 5)
        '+v1.5
        '.Fields.Append .CreateField("Processed", dbText, 20)
        .Fields.Append .CreateField("Processed", dbText, 40)
        '-v1.5
        .Fields.Append .CreateField("Analysis_File", dbText, 100)
        .Fields.Append .CreateField("Messages", dbLong)
        .Fields.Append .CreateField("Bytes", dbLong)
        .Fields.Append .CreateField("Mission", dbLong)
        .Fields.Append .CreateField("Aircraft", dbLong)
        '
        ' Set the ID field to be autoincrementing
        .Fields("ID").Attributes = dbAutoIncrField
        '
        ' Set the field attribute to allow null strings
        .Fields("Name").AllowZeroLength = True
        .Fields("Start").AllowZeroLength = True
        .Fields("End").AllowZeroLength = True
        .Fields("Archive").AllowZeroLength = True
        .Fields("Media").AllowZeroLength = True
        .Fields("Processed").AllowZeroLength = True
        .Fields("Analysis_File").AllowZeroLength = True
        '
        ' Set default values for some of the fields
        .Fields("Media").DefaultValue = "HD"
        .Fields("Processed").DefaultValue = "Never"
        .Fields("Messages").DefaultValue = 0
        .Fields("Bytes").DefaultValue = 0
        .Fields("Mission").DefaultValue = 0
        .Fields("Aircraft").DefaultValue = 0
    End With
    '
    ' Add table to the database
    dbCurrent.TableDefs.Append tblArchive
    '
    '+v1.8.12 TE
    On Error Resume Next
    dbCurrent.TableDefs.Delete "MTHBDYNRSP"
    dbCurrent.TableDefs.Delete "MTRUNMODE"
    dbCurrent.TableDefs.Delete "MTSIGALARM"
    dbCurrent.TableDefs.Delete "MTSIGUPD"
    dbCurrent.TableDefs.Delete "MTTGTLISTUPD"
    dbCurrent.TableDefs.Delete "MTHBACTREP"
    On Error GoTo 0
    '-v1.8.12 TE
    '
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Create_Archive_Table (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Create_Summary_Table
' AUTHOR:   Tom Elkins
' PURPOSE:  Creates the default, blank archive summary table in the specified database
' INPUT:    "dbCurrent" is the currently selected database
'           "iArchive" is the current archive ID
' OUTPUT:   None
' NOTES:    Summary tables contain results from processing an archive.
'               Message is the name of a message in the archive
'               MSG_ID is the numeric identifier of the message type
'               Count is the number of messages in the archive of the current type
'               First is the time of the first occurance of this message in the archive
'               Last is the time of the last occurance of this message in the archive
Public Function bCreate_Summary_Table(dbCurrent As Database, iArchive As Integer) As Boolean
    Dim tblSummary As TableDef  ' New summary table
    '
    ' Trap errors
    On Error GoTo Table_Error
    '
    ' Set the default return value
    bCreate_Summary_Table = False
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.Create_Summary_Table: Archive = " & iArchive & " (Start)"
    '
    ' Create the table
    Set tblSummary = dbCurrent.CreateTableDef("Archive" & iArchive & TBL_SUMMARY)
    '
    ' Use table-level addressing
    With tblSummary
        '
        ' Add the fields
        .Fields.Append .CreateField("Message", dbText, 20)
        .Fields.Append .CreateField("MSG_ID", dbLong)
        .Fields.Append .CreateField("Count", dbLong)
        '+v1.5
        ' Changed database schema to use dates instead of text
        '.Fields.Append .CreateField("First", dbText, 20)
        '.Fields.Append .CreateField("Last", dbText, 20)
        .Fields.Append .CreateField("First", dbDate)
        .Fields.Append .CreateField("Last", dbDate)
        '-v1.5
        .Fields.Append .CreateField("Description", dbText, 255)
        .Fields.Append .CreateField("Signal", dbBoolean)
        .Fields.Append .CreateField("LOB", dbBoolean)
        .Fields.Append .CreateField("Fix", dbBoolean)
        .Fields.Append .CreateField("Track", dbBoolean)
        '
        ' Set the field attribute to allow null strings
        .Fields("Message").AllowZeroLength = True
        .Fields("First").AllowZeroLength = True
        .Fields("Last").AllowZeroLength = True
        .Fields("Description").AllowZeroLength = True
    End With
    '
    ' Add table to database
    dbCurrent.TableDefs.Append tblSummary
    '
    ' Set the return value
    bCreate_Summary_Table = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.bCreate_Summary_Table (End)"
    '-v1.6.1
    '
    ' Leave
    Exit Function
'
' Error handler
Table_Error:
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    ' Set the return value to false
    bCreate_Summary_Table = False
    '
    ' Log the error
    basCCAT.WriteLogEntry "ERROR    : basDatabase.bCreate_Summary_Table (Error #" & Err.Number & " - " & Err.Description & ")"
    '
    ' Inform the user
    MsgBox "Error #" & Err.Number & vbCr & Err.Description, vbOKOnly, "Error Creating Summary Table"
End Function
'
' FUNCTION: bCreate_Data_Table
' AUTHOR:   Tom Elkins
' PURPOSE:  Creates the default, blank message data table in the specified database
' INPUT:    "dbCurrent" is the currently selected database
'           "iArchive" is the current archive ID
'           "sMsg" is the current message name
' OUTPUT:   True if successful, False if not
' NOTES:    Data tables contain actual values from a message.  The fields are
'           from the DAS master list
'           Data table names are "<Message><Archive ID>_Data"
Public Function bCreate_Data_Table(dbCurrent As Database, iArchive As Integer) As Boolean
    Dim tblData As TableDef     ' New data table
    '
    ' Trap errors
    On Error GoTo Data_Table_Error
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.bCreate_Data_Table (Start)"
    basCCAT.WriteLogEntry "ARGUMENTS: " & dbCurrent.Name & ", " & iArchive
    '
    ' Create the table
    Set tblData = dbCurrent.CreateTableDef("Archive" & iArchive & TBL_DATA)
    '
    ' Use table-level addressing
    With tblData
        '
        ' Add the fields
        .Fields.Append .CreateField("ReportTime", dbDate)
        .Fields.Append .CreateField("Msg_Type", dbText, 30)
        .Fields.Append .CreateField("Rpt_Type", dbText, 10)
        .Fields.Append .CreateField("Origin", dbText, 30)
        .Fields.Append .CreateField("Origin_ID", dbLong)
        .Fields.Append .CreateField("Target_ID", dbLong)
        .Fields.Append .CreateField("Latitude", dbDouble)
        .Fields.Append .CreateField("Longitude", dbDouble)
        .Fields.Append .CreateField("Altitude", dbDouble)
        .Fields.Append .CreateField("Heading", dbDouble)
        .Fields.Append .CreateField("Speed", dbDouble)
        .Fields.Append .CreateField("Parent", dbText, 50)
        .Fields.Append .CreateField("Parent_ID", dbLong)
        .Fields.Append .CreateField("Allegiance", dbText, 20)
        .Fields.Append .CreateField("IFF", dbLong)
        .Fields.Append .CreateField("Emitter", dbText, 80)
        .Fields.Append .CreateField("Emitter_ID", dbLong)
        .Fields.Append .CreateField("Signal", dbText, 50)
        .Fields.Append .CreateField("Signal_ID", dbLong)
        .Fields.Append .CreateField("Frequency", dbDouble)
        .Fields.Append .CreateField("PRI", dbDouble)
        .Fields.Append .CreateField("Status", dbLong)
        .Fields.Append .CreateField("Tag", dbLong)
        .Fields.Append .CreateField("Flag", dbLong)
        .Fields.Append .CreateField("Common_ID", dbLong)
        .Fields.Append .CreateField("Range", dbDouble)
        .Fields.Append .CreateField("Bearing", dbDouble)
        .Fields.Append .CreateField("Elevation", dbDouble)
        .Fields.Append .CreateField("XX", dbDouble)
        .Fields.Append .CreateField("XY", dbDouble)
        .Fields.Append .CreateField("YY", dbDouble)
        .Fields.Append .CreateField("Other_Data", dbText)
        '
        ' Set the field attribute to allow null strings
        .Fields("Msg_Type").AllowZeroLength = True
        .Fields("Rpt_Type").AllowZeroLength = True
        .Fields("Origin").AllowZeroLength = True
        .Fields("Parent").AllowZeroLength = True
        .Fields("Allegiance").AllowZeroLength = True
        .Fields("Emitter").AllowZeroLength = True
        .Fields("Signal").AllowZeroLength = True
        .Fields("Other_Data").AllowZeroLength = True
    End With
    '
    ' Add table to the database
    dbCurrent.TableDefs.Append tblData
    '
    ' Resume error reporting
    On Error GoTo 0
    '
    ' Set the data table
    guCurrent.uSQL.sTable = "Archive" & iArchive & TBL_DATA
    '
    ' Success
    bCreate_Data_Table = True
    '
    ' Log the event
    basCCAT.WriteLogEntry CStr(bCreate_Data_Table)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.bCreate_Data_Table (End)"
    '-v1.6.1
    '
    ' Leave
    Exit Function
'
' Error handler
Data_Table_Error:
    '
    ' Report failure
    bCreate_Data_Table = False
    '
    ' Log the event
    '
    '+v1.6.1TE
    basCCAT.WriteLogEntry "ERROR    : basDatabase.bCreate_Data_Table (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    ' Inform the user
    MsgBox "ERROR #" & Err.Number & vbCr & Err.Description, vbOKOnly Or vbCritical, "Error Creating Data Table"
    '
    ' Restore error reporting
    On Error GoTo 0
End Function
'
' FUNCTION: bAdd_Archive_Record
' AUTHOR:   Tom Elkins
' PURPOSE:  Adds an archive record to the database
' INPUT:    None
' OUTPUT:   True if successful, False if not
' NOTES:
Public Function bAdd_Archive_Record() As Boolean
    Dim rsArchives As Recordset     ' Pointer to the archive table
    '
    ' Trap errors
    On Error GoTo Archive_Error
    '
    ' Set the default return value
    bAdd_Archive_Record = False
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.bAdd_Archive_Record (Start)"
    '
    ' Open the archives table
    Set rsArchives = guCurrent.DB.OpenRecordset(TBL_ARCHIVES)
    '
    ' Add a new record
    rsArchives.AddNew
    '
    ' Record the automatically generated archive ID
    guCurrent.iArchive = rsArchives!ID
    basCCAT.WriteLogEntry "New archive ID = " & guCurrent.iArchive
    '
    ' Populate the record with default values
    rsArchives!Name = "Archive" & guCurrent.iArchive
    rsArchives!Date = Date
    '+v1.5
    ' Store date/time directly instead of converting to JDAY/Time format
    'rsArchives!Start = DatePart("y", Date) & ":" & Format(Time, "hh:nn:ss")
    'rsArchives!End = DatePart("y", Date) & ":00:00:00"
    rsArchives!Start = guCurrent.uArchive.dtArchiveDate
    rsArchives!End = guCurrent.uArchive.dtArchiveDate
    '-v1.5
    rsArchives!Archive = ""
    rsArchives!Media = "HD"
    rsArchives!Processed = "Never"
    rsArchives!Analysis_File = ""
    rsArchives!Messages = 0
    rsArchives!Bytes = 0
    rsArchives!Mission = 0
    rsArchives!Aircraft = 0
    '
    ' Add the record to the table
    rsArchives.Update
    rsArchives.Close
    '
    ' Create the archive summary table
    If basDatabase.bCreate_Summary_Table(guCurrent.DB, guCurrent.iArchive) Then
        '
        ' Create the Archive data table
        bAdd_Archive_Record = basDatabase.bCreate_Data_Table(guCurrent.DB, guCurrent.iArchive)
    End If
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.bAdd_Archive_Record (End)"
    '-v1.6.1
    '
    ' Leave
    Exit Function
'
' Error handler
Archive_Error:
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    ' Report failure
    bAdd_Archive_Record = False
    '
    ' Log the event
    basCCAT.WriteLogEntry "ERROR    : basDatabase.bAdd_Archive_Record (Error #" & Err.Number & " - " & Err.Description & ")"
    '
    ' Inform the user
    MsgBox "Error #" & Err.Number & vbCr & Err.Description, vbOKOnly, "Error Adding Archive"
End Function
'
' ROUTINE:  Add_Data_Record
' AUTHOR:   Tom Elkins
' PURPOSE:  Records data values extracted from an archive
' INPUT:    "iMsg_ID" is the ID of the message extracted
'           "uSig" is the structure containing the extracted data values
' OUTPUT:   None
' NOTES:    Even though this routine really belongs with frmArchive, the VB compiler
'           would not allow passing a user-defined structure to a form.
Public Sub Add_Data_Record(iMsg_ID As Integer, uSig As DAS_MASTER_RECORD)
    '+v1.5
    ' Use date/time directly instead of TSecs
    'Dim dMsg_Time As Double     ' Extracted message time with JDay added
    Dim dtMsg_Time As Date      ' Extracted message date and time
    '-v1.5
    Dim sToken As String
    Dim lTokenLen As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Add_Data_Record(" & iMsg_ID & ")"
    '-v1.6.1
    '
    ' Trap errors
    On Error GoTo ERR_HANDLER
    '
    ' Update the current info structure
    guCurrent.iMessage = iMsg_ID
    '
    '+v1.5
    '' Update the time with the day of the year
    'dMsg_Time = uSig.dReportTime + guCurrent.uArchive.dOffset_Time
    ' Add the message time to the start date. COMPASS CALL reports time as seconds from
    ' midnight of the first mission day.
    If (uSig.dReportTime > 86400) Then
    Debug.Print
    End If
    
    dtMsg_Time = DateAdd("s", uSig.dReportTime, guCurrent.uArchive.dtArchiveDate)
    '-v1.5
    '
    ' +v1.7BB
    ' Moved much code

    ' -v1.7BB
    
    guCurrent.uArchive.lNum_Messages = guCurrent.uArchive.lNum_Messages + 1
    guCurrent.uArchive.lNum_Bytes = Loc(Filtraw2Das.giFiltInputFile)
    '
    '+v1.6
    'frmArchive.barProgress.Value = guCurrent.uArchive.lNum_Bytes
    frmWizard.barProgress.Value = guCurrent.uArchive.lNum_Bytes
    '-v1.6
    '
    ' Add a new data record
    guCurrent.uArchive.rsData.AddNew
    '
    ' Fill in the record values
    '
    '+v1.5 - Use full date/time
    'guCurrent.uArchive.rsData!ReportTime = CDate(uSig.dReportTime / 86400#)
    guCurrent.uArchive.rsData!ReportTime = dtMsg_Time
    '-v1.5
    guCurrent.uArchive.rsData!Msg_Type = uSig.sMsg_Type
    guCurrent.uArchive.rsData!Rpt_Type = uSig.sReport_Type
    guCurrent.uArchive.rsData!Origin = uSig.sOrigin
    guCurrent.uArchive.rsData!Origin_ID = uSig.lOrigin_ID
    guCurrent.uArchive.rsData!Target_ID = uSig.lTarget_ID
    guCurrent.uArchive.rsData!Latitude = uSig.dLatitude
    guCurrent.uArchive.rsData!Longitude = uSig.dLongitude
    guCurrent.uArchive.rsData!Altitude = uSig.dAltitude
    guCurrent.uArchive.rsData!Heading = uSig.dHeading
    guCurrent.uArchive.rsData!Speed = uSig.dSpeed
    guCurrent.uArchive.rsData!Parent = uSig.sParent
    guCurrent.uArchive.rsData!Parent_ID = uSig.lParent_ID
    guCurrent.uArchive.rsData!Allegiance = uSig.sAllegiance
    guCurrent.uArchive.rsData!IFF = uSig.lIFF
    guCurrent.uArchive.rsData!Emitter = uSig.sEmitter
    guCurrent.uArchive.rsData!Emitter_ID = uSig.lEmitter_ID
    guCurrent.uArchive.rsData!Signal = uSig.sSignal
    guCurrent.uArchive.rsData!Signal_ID = uSig.lSignal_ID
    guCurrent.uArchive.rsData!Frequency = uSig.dFrequency
    guCurrent.uArchive.rsData!PRI = uSig.dPRI
    guCurrent.uArchive.rsData!Status = uSig.lStatus
    guCurrent.uArchive.rsData!Tag = uSig.lTag
    guCurrent.uArchive.rsData!Flag = uSig.lFlag
    guCurrent.uArchive.rsData!Common_ID = uSig.lCommon_ID
    guCurrent.uArchive.rsData!Range = uSig.dRange
    guCurrent.uArchive.rsData!Bearing = uSig.dBearing
    guCurrent.uArchive.rsData!Elevation = uSig.dElevation
    guCurrent.uArchive.rsData!XX = uSig.dXX
    guCurrent.uArchive.rsData!XY = uSig.dXY
    guCurrent.uArchive.rsData!YY = uSig.dYY
    guCurrent.uArchive.rsData!Other_Data = uSig.sSupplemental
    '
    ' Add the new record to the table
    guCurrent.uArchive.rsData.Update
    '
    ' Update the screen periodically
    If guCurrent.uArchive.lNum_Messages Mod guGUI.lInterval = 0 Then
        '
        '+v1.6
        ''+v1.5
        '' Use direct date/time values
        '' Shorten date/time text to be compatible with old format size
        ''frmArchive.lblStart.Caption = basCCAT.sTSecs_To_Human_Time(guCurrent.uArchive.dStart_Time)
        ''frmArchive.lblEnd.Caption = basCCAT.sTSecs_To_Human_Time(guCurrent.uArchive.dEnd_Time)
        'frmArchive.lblStart.Caption = Format(guCurrent.uArchive.dtStart_Time, "mm/dd/yyyy hh:nn:ss")
        'frmArchive.lblEnd.Caption = Format(guCurrent.uArchive.dtEnd_Time, "mm/dd/yyyy hh:nn:ss")
        ''-v1.5
        'frmArchive.lblNumBytes = CDbl(Loc(Filtraw2Das.giFiltInputFile))
        'frmArchive.lblNumMsg = guCurrent.uArchive.lNum_Messages
        'frmArchive.lblProcessInfo.Caption = "Translating Filtered File - " & Int(guCurrent.uArchive.lNum_Bytes / (guCurrent.uArchive.lFile_Size / 100)) & "% Complete"
        frmWizard.lblPctDone = "Translating Filtered File - " & Int(guCurrent.uArchive.lNum_Bytes / (guCurrent.uArchive.lFile_Size / 100)) & "% Complete"
        DoEvents
        'If Not gbProcessing Then Err.Raise vbObjectError + 911, "CCAT:Add_Data_Record", "User terminated the translation process"
        '-v1.6
    End If
    Exit Sub
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basDatabase.Add_Data_Record (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
    Err.Raise Err.Number, "CCAT:Add_Data_Record", Err.Description
End Sub
'
' ROUTINE:  Delete_Archive
' AUTHOR:   Tom Elkins
' PURPOSE:  Delete an archive from a database
' INPUT:    "sArchive_Key" is the key for the corresponding archive node
' OUTPUT:   None
' NOTES:    The database is the one pointed to by guCurrent.DB
Public Sub Delete_Archive(sArchive_Key As String)
    Dim iArchive_ID As Integer
    Dim rsTable As Recordset
    '
    '+v1.6.1TE
    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Delete_Archive (Start)"
    basCCAT.WriteLogEntry "ARGUMENTS: " & sArchive_Key
    '-v1.6.1
    '
    '
    ' See if the archive node exists
    If frmMain.blnNodeExists(sArchive_Key) Then
        '
        ' See if there are items in the list view
        If frmMain.lvListView.ListItems.Count > 0 Then
            '
            ' Try to find the item
            If Not frmMain.lvListView.FindItem(frmMain.tvTreeView.Nodes(sArchive_Key).Text) Is Nothing Then
                '
                ' Make sure it is an archive item
                If frmMain.lvListView.ListItems(sArchive_Key).Tag = gsARCHIVE Then
                    '
                    ' Log the event
                    basCCAT.WriteLogEntry "INFO     : basDatabase.Delete_Archive (Removing archive item from list view)"
                    '
                    ' Remove the item
                    frmMain.lvListView.ListItems.Remove sArchive_Key
                End If
            End If
        End If
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : basDatabase.Delete_Archive (Removing archive node from tree view)"
        '
        ' Remove the node from the tree view
        frmMain.tvTreeView.Nodes.Remove sArchive_Key
    End If
    '
    ' Find the archive record in the Archives table
    '+v1.6TE
    'Set rsTable = guCurrent.DB.OpenRecordset("SELECT * FROM " & TBL_ARCHIVES & " WHERE ID = " & guCurrent.iArchive)
    Set rsTable = guCurrent.DB.OpenRecordset("SELECT * FROM " & TBL_ARCHIVES & " WHERE Name = '" & guCurrent.sArchive & "'")
    '-v1.6
    '
    ' Delete it if it exists
    If Not rsTable Is Nothing Then
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : basDatabase.Delete_Archive (Deleting archive record '" & guCurrent.sArchive & "' from table " & TBL_ARCHIVES & ")"
        '
        ' Delete the archive record
        rsTable.Delete
        rsTable.Close
    Else
        '
        ' Log the event
        basCCAT.WriteLogEntry "ERROR    : basDatabase.Delete_Archive (Cannot find archive #" & guCurrent.iArchive & " in the table)"
        MsgBox "Error -- Cannot find archive #" & guCurrent.iArchive & " in the table", , "Error Deleting Archive"
    End If
    '
    ' See if the data table exists
    '+v1.6TE
    'If bTable_Exists(guCurrent.DB, "Archive" & guCurrent.iArchive & TBL_DATA) Then
    If bTable_Exists(guCurrent.DB, guCurrent.sArchive & TBL_DATA) Then
    '-v1.6
        '
        ' Give the data grid something to look at so it doesn't bomb
'        frmMain.Data1.RecordSource = "Info"
'        frmMain.Data1.Refresh
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : basDatabase.Delete_Archive (Deleting data table)"
        '
        ' Delete the data table
        '+v1.6TE
        'guCurrent.DB.TableDefs.Delete "Archive" & guCurrent.iArchive & TBL_DATA
        '
        ' See if the table being deleted is the active table for the data grid
        If InStr(1, frmMain.Data1.RecordSource, guCurrent.sArchive & TBL_DATA) > 0 Then
            '
            ' Change the data grid table to something innocuous
            frmMain.Data1.RecordSource = TBL_INFO
            frmMain.Data1.Refresh
            '
            ' Change the display
            basDatabase.Display_Database_Archives
        End If
        '
        ' Delete the table
        guCurrent.DB.TableDefs.Delete guCurrent.sArchive & TBL_DATA
        '-v1.6
    End If
    '
    ' See if the summary table exists
    '+v1.6TE
    'If bTable_Exists(guCurrent.DB, "Archive" & guCurrent.iArchive & TBL_SUMMARY) Then
    If bTable_Exists(guCurrent.DB, guCurrent.sArchive & TBL_SUMMARY) Then
    '-v1.6
        '
        ' Log the event
        'basCCAT.WriteLogEntry "DATABASE: Delete_Archive: Deleting Summary table Archive" & guCurrent.iArchive & TBL_SUMMARY
        basCCAT.WriteLogEntry "INFO     : basDatabase.Delete_Archive (Deleting Summary table '" & guCurrent.sArchive & TBL_SUMMARY & "')"
        '
        ' Delete the summary table
        'guCurrent.DB.TableDefs.Delete "Archive" & guCurrent.iArchive & TBL_SUMMARY
        guCurrent.DB.TableDefs.Delete guCurrent.sArchive & TBL_SUMMARY
        '
    End If
    ' See if the summary table exists
    '+v1.7BB
    If bTable_Exists(guCurrent.DB, guCurrent.sArchive & TBL_PROC_DATA) Then

        basCCAT.WriteLogEntry "INFO     : basDatabase.Delete_Archive (Deleting Summary table '" & guCurrent.sArchive & TBL_PROC_DATA & "')"
        '
        ' Delete the summary table
        guCurrent.DB.TableDefs.Delete guCurrent.sArchive & TBL_PROC_DATA
        '
    End If
    '-v1.7BB
    ' See if the summary table exists
    '+v1.7BB
    If bTable_Exists(guCurrent.DB, guCurrent.sArchive & TBL_VAR_STRUCT) Then

        basCCAT.WriteLogEntry "INFO     : basDatabase.Delete_Archive (Deleting Summary table '" & guCurrent.sArchive & TBL_VAR_STRUCT & "')"
        '
        ' Delete the summary table
        guCurrent.DB.TableDefs.Delete guCurrent.sArchive & TBL_VAR_STRUCT
        '
    End If
    '-v1.7BB
    ' See if the summary table exists
    '+v1.7BB
    If bTable_Exists(guCurrent.DB, guCurrent.sArchive & TBL_TOC) Then

        basCCAT.WriteLogEntry "INFO     : basDatabase.Delete_Archive (Deleting Summary table '" & guCurrent.sArchive & TBL_TOC & "')"
        '
        ' Delete the summary table
        guCurrent.DB.TableDefs.Delete guCurrent.sArchive & TBL_TOC
        '
    End If
    '
    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Delete_Archive (End)"
    '-v1.7BB
    ' See if the summary table exists
    '+v1.7BB
    If bTable_Exists(guCurrent.DB, guCurrent.sArchive & TBL_MESSAGE) Then

        basCCAT.WriteLogEntry "INFO     : basDatabase.Delete_Archive (Deleting Summary table '" & guCurrent.sArchive & TBL_MESSAGE & "')"
        '
        ' Delete the summary table
        guCurrent.DB.TableDefs.Delete guCurrent.sArchive & TBL_MESSAGE
        '
    End If
    '
    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Delete_Archive (End)"
    '-v1.7BB
    '
End Sub
'
' FUNCTION: bTable_Exists
' AUTHOR:   Tom Elkins
' PURPOSE:  Checks for the existence of a specified table within a database
' INPUT:    "dbCurrent" is the database to check
'           "sTable" is the name of the table in question
' OUTPUT:   TRUE if the table exists
'           FALSE if the table does not exist
' NOTES:
Public Function bTable_Exists(dbCurrent As Database, sTable As String) As Boolean
    Dim tblTable As TableDef
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : basDatabase.bTable_Exists(" & dbCurrent.Name & ", " & sTable & ")"
    End If
    '-v1.6.1
    '
    '
    ' Initialize the found flag
    bTable_Exists = False
    '
    ' Loop through the tables in the database
    For Each tblTable In dbCurrent.TableDefs
        '
        ' Check the name
        If tblTable.Name = sTable Then
            '
            ' Found! Set the flag to True
            bTable_Exists = True
            '
            ' No point in continuing the search, so exit the loop
            Exit For
        End If
    Next tblTable
End Function
'
' FUNCTION: sCreate_SQL
' AUTHOR:   Tom Elkins
' PURPOSE:  Generates an SQL query based on a field list, filter clause, and sort list
' INPUT:    None
' OUTPUT:   The completed SQL query
' NOTES:    Uses the text stored in the guCurrent structure.  As elements of the
'           query are modified, the structure is updated.
Public Function sCreate_SQL() As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.sCreate_SQL (Start)"
    '-v1.6.1
    '
    '
    ' Create the field list and table portion
    guCurrent.uSQL.sQuery = "SELECT " & guCurrent.uSQL.sFields & " FROM [" & guCurrent.uSQL.sTable & "]"
    '
    ' If there is a filter, add the WHERE clause
    If Len(guCurrent.uSQL.sFilter) > 0 Then guCurrent.uSQL.sQuery = guCurrent.uSQL.sQuery & " WHERE " & guCurrent.uSQL.sFilter
    '
    ' If there is a sort list, add the ORDER BY clause
    If Len(guCurrent.uSQL.sOrder) > 0 Then guCurrent.uSQL.sQuery = guCurrent.uSQL.sQuery & " ORDER BY " & guCurrent.uSQL.sOrder
    '
    ' Return the completed query
    sCreate_SQL = guCurrent.uSQL.sQuery
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.sCreate_SQL (End) = " & guCurrent.uSQL.sQuery
    '-v1.6.1
    '
End Function
'
' FUNCTION: lExport_Table
' AUTHOR:   Tom Elkins
' PURPOSE:  Writes the contents of the specified table to a file
' INPUT:    "iFile" is the ID for the export file
' OUTPUT:   The number of records exported
' NOTES:    The database is the one pointed to by guCurrent.DB
Public Function lExport_Table(iFile As Integer) As Long
    Dim rsRecord As Recordset           ' Pointer to a data record
    Dim uDAS_Record As DAS_MASTER_RECORD   ' DAS record structure
    Dim iArchive As Integer             ' Parent archive ID
    Dim sQuery As String                ' database query
    Dim fldField As Field
    Dim sOutput As String
    Dim sTemp As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.lExport_Table (Start)"
    '-v1.6.1
    '
    '
    ' Reset the count
    lExport_Table = 0
    '
    ' Extract the update interval from the token/ini file
    If guGUI.lInterval = 0 Then guGUI.lInterval = basCCAT.GetNumber("Miscellaneous Operations", "Update_Interval", 1250)
    '
    ' Check for the existence of the specified table
    If basDatabase.bTable_Exists(guCurrent.DB, guCurrent.uSQL.sTable) Then
        '
        ' Suppress error reporting
        On Error Resume Next
        '
        ' Position the progress bar
        frmMain.barLoad.Left = frmMain.sbStatusBar.Panels(1).Left
        frmMain.barLoad.Width = frmMain.sbStatusBar.Panels(1).Width
        frmMain.barLoad.Top = frmMain.sbStatusBar.Top
        frmMain.barLoad.Height = frmMain.sbStatusBar.Height
        '
        ' Store the old query
        sQuery = frmMain.Data1.RecordSource
        '
        ' Change the mouse to "busy"
        Screen.MousePointer = vbHourglass
        '
        ' Set the new query
        ' Cannot use recordsource directly with Access 2000 databases
        '+v1.5
        sTemp = sCreate_SQL
        'frmMain.Data1.RecordSource = sCreate_SQL
        '-v1.5
        '
        ' See if the new query is different
        If guCurrent.uSQL.sQuery <> sQuery Then
            '
            ' Clean up the display
            DoEvents
            '
            ' Refresh the data grid
            '+v1.5
            'frmMain.Data1.Refresh
            ' Cannot bind the data grid directly to Access 2000 databases
            Set rsRecord = guCurrent.DB.OpenRecordset(guCurrent.uSQL.sQuery)
            Set frmMain.Data1.Recordset = rsRecord
            '-v1.5
            '
            ' See if any records were returned
            If frmMain.Data1.Recordset.RecordCount > 0 Then
                '
                ' Look at the last record to get an accurate record count
                frmMain.sbStatusBar.Panels(1).Text = "Requerying the database..."
                frmMain.Data1.Recordset.MoveLast
                frmMain.Data1.Recordset.MoveFirst
            End If
        '+v1.5
        'Else
        '    '
        '    ' Restore the old query
        '    frmMain.Data1.RecordSource = sQuery
        '-v1.5
        End If
        '
        ' Change the mouse
        Screen.MousePointer = vbDefault
        '
        ' Check for errors
        Select Case Err.Number
            '
            ' Syntax error
            Case 3075, 3061:
                '
                ' Log the event
                basCCAT.WriteLogEntry "ERROR    : basDatabase.lExport_Table (SQL Syntax error: " & guCurrent.uSQL.sQuery & ")"
                '
                ' Inform the user
                '+v1.5
                'MsgBox "Syntax error in the specified SQL query" & vbCr & "'" & guCurrent.uSQL.sQuery & "'" & vbCr & "Please try exporting again.", vbOKOnly Or vbMsgBoxHelpButton, "SQL Syntax Error", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, IDH_FilterData
                MsgBox "Syntax error in the specified SQL query" & vbCr & "'" & guCurrent.uSQL.sQuery & "'" & vbCr & "Please try exporting again.", vbOKOnly Or vbMsgBoxHelpButton, "SQL Syntax Error", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, basCCAT.IDH_DB_FILTERING
                '-v1.5
            '
            ' No errors
            Case NO_ERROR:
                '
                ' Resume error reporting
                On Error GoTo 0
                '
                ' Check for results
                If frmMain.Data1.Recordset.RecordCount > 0 Then
                    '
                    ' Make the progress bar visible
                    frmMain.barLoad.Min = 0
                    frmMain.barLoad.Value = 0
                    frmMain.barLoad.Max = frmMain.Data1.Recordset.RecordCount
                    frmMain.barLoad.Visible = True
                    '
                    ' Change the mouse
                    Screen.MousePointer = vbHourglass
                    '
                    ' Loop through all records
                    '+v1.5
                    ' Reset the record pointer to the beginning
                    ' This cures an interesting bug -- If the user happened to click in a grid cell
                    ' before exporting, only the records after the selected grid were exported.  By
                    ' forcing the pointer to move to the first record, we eliminate that problem.
                    frmMain.Data1.Recordset.MoveFirst
                    '-v1.5
                    While Not frmMain.Data1.Recordset.EOF
                        '
                        ' Write the fields to the file
                        sOutput = ""
                        For Each fldField In frmMain.Data1.Recordset.Fields
                            Select Case fldField.Type
                                Case dbText:
                                    sTemp = Replace(Trim(fldField.Value), " ", "_")
                                    If Len(sTemp) = 0 Then
                                        If fldField.Name = "Rpt_Type" Then
                                            sTemp = basCCAT.gaDAS_Rec_Type(guExport.iRec_Type)
                                        Else
                                            sTemp = "UNKNOWN"
                                        End If
                                    End If
                                    sOutput = sOutput & Trim(sTemp) & ","
                                Case dbLong:
                                    sOutput = sOutput & Format(fldField.Value, "0") & ","
                                Case dbDouble:
                                    sOutput = sOutput & Format(fldField.Value, "0.00000") & ","
                                Case dbDate:
                                    '
                                    '+v1.5
                                    'sOutput = sOutput & Format((CDbl(fldField.Value) * 86400#) + guCurrent.uArchive.dOffset_Time, "0.000") & ","
                                    '
                                    ' Subtract the days from the value, which gives the time in the day
                                    ' Multiply the time by 86400 to get seconds
                                    ' Get the day of the year (JDay) from the value and add 1 (1 Jan = day 0)
                                    ' Multiply the JDay by 86400 to convert to seconds
                                    ' Add JDay to time to get TSecs
                                    ' Output TSecs for the DAS file format.
                                    sOutput = sOutput & Format((CDbl(fldField.Value - Int(fldField.Value)) * 86400#) + ((DatePart("y", fldField.Value) - 1) * 86400#), "0.000") & ","
                                    '-v1.5
                                    '
                            End Select
                        Next fldField
                        Print #iFile, Mid(sOutput, 1, Len(sOutput) - 1)
                        '
                        ' Update the count
                        lExport_Table = lExport_Table + 1
                        frmMain.barLoad.Value = lExport_Table
                        '
                        ' Periodically release control to the system for other operations
                        If lExport_Table Mod basCCAT.guGUI.lInterval = 0 Then
                            DoEvents
                        End If
                        '
                        ' Move to the next record
                        frmMain.Data1.Recordset.MoveNext
                    Wend
                    '
                    ' Hide the progress bar
                    frmMain.barLoad.Visible = False
                    '
                    ' Change mouse
                    Screen.MousePointer = vbDefault
                Else
                    '
                    ' Inform the user no records were found
                    MsgBox "Your query '" & sQuery & "' resulted in 0 records!", , "No Records Found"
                End If
                '
                ' Log the event
                basCCAT.WriteLogEntry "INFO     : basDatabase.lExport_Table (Exported " & lExport_Table & " records)"
            '
            ' Unknown
            Case Else:
                '
                ' Log the event
                basCCAT.WriteLogEntry "ERROR    : basDatabase.lExport_Table (Error #" & Err.Number & " - " & Err.Description & " while attempting to export table " & guCurrent.uSQL.sTable & ")"
                '
                ' Inform the user
                MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "While attempting to export table", , "Unexpected Error"
        End Select
        '
        ' Resume error reporting
        On Error GoTo 0
    Else
        '
        ' Log the event
        basCCAT.WriteLogEntry "ERROR    : basDatabase.lExport_Table (Table " & guCurrent.uSQL.sTable & " not found!)"
        '
        ' Inform the user
        MsgBox "Table '" & guCurrent.uSQL.sTable & "' not found!", , "Table not found"
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.lExport_Table (End)"
    '-v1.6.1
    '
End Function
'
' ROUTINE:  Requery_Data
' AUTHOR:   Tom Elkins
' PURPOSE:  Populates the data grid
' INPUT:    None
' OUTPUT:   None
' NOTES:    The database is the one pointed to by guCurrent.DB
Public Sub Requery_Data()
    Dim sWhere As String        ' WHERE clause
    Dim sSort As String         ' ORDER BY clause
    Dim sOld_Query As String    ' Old query
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Requery_Data (Start)"
    '-v1.6.1
    '
    '+v1.5
    Dim rsQuery As Recordset    ' Required to use Intrinsic Data control with Access 2000
    '-v1.5
    '
    ' Save old query
    sOld_Query = frmMain.Data1.RecordSource
    '
    ' Change the mouse pointer to "busy"
    Screen.MousePointer = vbHourglass
    '
    ' Suppress errors
    On Error Resume Next
    '
    ' Create the new query
    '+v1.5
    'frmMain.Data1.RecordSource = basDatabase.sCreate_SQL & ";"
    ' Using the old method will cause a "Unknown Database Format" error
    ' Must set recordset object to the query, then assign the recordset object to the grid
    Set rsQuery = guCurrent.DB.OpenRecordset(basDatabase.sCreate_SQL)
    Set frmMain.Data1.Recordset = rsQuery
    '
    '' Display the results
    ' Assigning the recordset precludes the use of the Refresh method.  Method will cause and error
    'frmMain.Data1.Refresh
    '-v1.5
    '
    ' Return the mouse pointer to normal
    Screen.MousePointer = vbDefault
    '
    ' Check for errors
    If Err.Number <> NO_ERROR Then
        '
        ' Inform the user
        MsgBox "Could not execute your query!" & vbCr & "Error #" & Err.Number & " - " & Err.Description & " in " & Err.Source & vbCr & vbCr & "Restoring old query", vbOKOnly Or vbExclamation Or vbMsgBoxHelpButton, "Error Executing Query", App.HelpFile, basCCAT.IDH_DB_FILTERING
        '
        ' Note it in the log
        basCCAT.WriteLogEntry "ERROR    : basDatabase.Requery_Data (Error executing query " & guCurrent.uSQL.sQuery
        basCCAT.WriteLogEntry "ERROR    : basDatabase.Requery_Date (Error #" & Err.Number & " - " & Err.Description & " in " & Err.Source & ")"
        '
        ' Resume error reporting
        On Error GoTo 0
        '
        ' restore the old query
        basDatabase.Parse_SQL sOld_Query
        '+v1.5
        'frmMain.Data1.RecordSource = sOld_Query
        'frmMain.Data1.Refresh
        '
        ' Required to use dbGrid with Access 2000 database
        Set rsQuery = guCurrent.DB.OpenRecordset(sOld_Query)
        Set frmMain.Data1.Recordset = rsQuery
        '-v1.5
    Else
        '
        ' See if any records were returned
        If frmMain.Data1.Recordset.RecordCount > 0 Then
            '
            ' Look at the last record to get an accurate record count
            frmMain.Data1.Recordset.MoveLast
            frmMain.Data1.Recordset.MoveFirst
        End If
        '
        ' Update the status bar
        frmMain.sbStatusBar.Panels(1).Text = frmMain.Data1.Recordset.RecordCount & " records matched query."
        frmMain.sbStatusBar.Panels(1).ToolTipText = guCurrent.uSQL.sQuery
    End If
    '
    ' Resume error reporting
    On Error GoTo 0
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Requery_Data (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: bReset_Archive_Processing
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets the database up to reprocess an archive
' INPUT:    None
' OUTPUT:   None
' NOTES:    The database is the one pointed to by guCurrent.DB
Public Function bReset_Archive_Processing() As Boolean
    Dim bContinue As Boolean        ' Continuation flag
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.bReset_Archive_Processing (Start)"
    '-v1.6.1
    '
    ' Set the default processing flag
    bContinue = True
    '
    ' See if the archive node exists
    If frmMain.blnNodeExists(guCurrent.sName & SEP_ARCHIVE & guCurrent.iArchive) Then
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : basDatabase.bReset_Archive_Processing (Deleting archive node " & guCurrent.sName & SEP_ARCHIVE & guCurrent.iArchive & ")"
        '
        ' Remove the node and all of its children
        frmMain.tvTreeView.Nodes.Remove guCurrent.sName & SEP_ARCHIVE & guCurrent.iArchive
    End If
    '
    ' Create the summary table name
    guCurrent.uSQL.sTable = "Archive" & guCurrent.iArchive & TBL_SUMMARY
    '
    ' Check for the existence of the summary table
    If basDatabase.bTable_Exists(guCurrent.DB, guCurrent.uSQL.sTable) Then
        '
        ' Remove all records from the summary table
        guCurrent.DB.Execute "DELETE * FROM " & guCurrent.uSQL.sTable
        '
        ' Log the number of records deleted
        basCCAT.WriteLogEntry "INFO     : basDatabase.bReset_Archive_Processing (" & guCurrent.DB.RecordsAffected & " records deleted from " & guCurrent.uSQL.sTable & ")"
    Else
        '
        ' Create the summary table
        bContinue = basDatabase.bCreate_Summary_Table(guCurrent.DB, guCurrent.iArchive)
    End If
    '
    ' Check for continuing
    If bContinue Then
        '
        ' Open the summary table
        Set guCurrent.uArchive.rsSummary = guCurrent.DB.OpenRecordset(guCurrent.uSQL.sTable, dbOpenDynaset)
        '
        ' Create the data table name
        guCurrent.uSQL.sTable = "Archive" & guCurrent.iArchive & TBL_DATA
        '
        ' Check for the existence of the data table
        If basDatabase.bTable_Exists(guCurrent.DB, guCurrent.uSQL.sTable) Then
            '
            ' Remove all records from the data table
            guCurrent.DB.Execute "DELETE * FROM " & guCurrent.uSQL.sTable
            '
            ' Log the number of records deleted
            basCCAT.WriteLogEntry "INFO     : basDatabase.bReset_Archive_Processing (" & guCurrent.DB.RecordsAffected & " records deleted from " & guCurrent.uSQL.sTable & ")"
        Else
            '
            ' Create the data table
            bContinue = basDatabase.bCreate_Data_Table(guCurrent.DB, guCurrent.iArchive)
        End If
        '
        ' Check for continuing
        If bContinue Then
            '
            ' Open the data table
            Set guCurrent.uArchive.rsData = guCurrent.DB.OpenRecordset(guCurrent.uSQL.sTable, dbOpenDynaset, dbAppendOnly)
        End If
    End If
    '
    ' Return the status
    bReset_Archive_Processing = bContinue
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.bReset_Archive_Processing (End)"
    '-v1.6.1
    '
End Function
'
' ROUTINE:  Display_Query Results
' AUTHOR:   Tom Elkins
' PURPOSE:  Display the results of the selected query
' INPUT:    "iQuery" is the ID of the selected query
' OUTPUT:   None
' NOTES:
Public Sub Display_Query_Results(iQuery As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Query_Results (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & iQuery
    End If
    '-v1.6.1
    '
    ' Set the grid to the current database
    frmMain.Data1.DatabaseName = guCurrent.sName
    '
    ' Position the grid to cover the list view
    frmMain.grdData.Left = frmMain.lvListView.Left
    frmMain.grdData.Top = frmMain.lvListView.Top
    frmMain.grdData.Width = frmMain.lvListView.Width
    frmMain.grdData.Height = frmMain.lvListView.Height
    '
    ' Hide the list view and show the grid
    frmMain.lvListView.Visible = False
    frmMain.grdData.Visible = True
    '
    ' Set the data table name
    '+v1.6TE
    'guCurrent.uSQL.sTable = "Archive" & guCurrent.iArchive & TBL_DATA
    guCurrent.uSQL.sTable = guCurrent.sArchive & TBL_DATA
    '-v1.6
    '
    ' Extract the custom field list from the token file
    guCurrent.uSQL.sFields = basCCAT.GetAlias("Queries", "QUERY_FIELDS" & iQuery, "*")
    '
    ' Extract the custom filter from the token file
    guCurrent.uSQL.sFilter = basCCAT.GetAlias("Queries", "QUERY" & iQuery, "")
    '
    ' Extract the custom sort list from the token file
    guCurrent.uSQL.sOrder = basCCAT.GetAlias("Queries", "QUERY_SORT" & iQuery, "")
    '
    ' Make sure the data table exists
    If bTable_Exists(guCurrent.DB, guCurrent.uSQL.sTable) Then
        '
        ' Execute the query
        basDatabase.Requery_Data
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Display_Query_Results (End)"
    '-v1.6.1
    '
End Sub
'
'
' ROUTINE:  Parse_SQL
' AUTHOR:   Tom Elkins
' PURPOSE:  Parses a SQL query into components that are then stored in a structure
' INPUT:    "sQuery" is the query to be parsed
' OUTPUT:   None
' NOTES:
Public Sub Parse_SQL(sQuery As String)
    Dim iField_List As Integer  ' Location of the end of the field list
    Dim iFilter As Integer      ' Location of the filter
    Dim iSort As Integer        ' Location of the sort list
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Parse_SQL (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sQuery
    End If
    '-v1.6.1
    '
    '
    ' Use structure-level addressing
    With guCurrent.uSQL
        '
        ' Update the query
        .sQuery = sQuery
        '
        ' Find the end of the field list
        iField_List = InStr(1, UCase(.sQuery), " FROM ")
        '
        ' Find the beginning of the filter
        iFilter = InStr(iField_List + 1, UCase(.sQuery), " WHERE ")
        '
        ' Find the beginning of the sort list
        iSort = InStr(iFilter + 1, UCase(.sQuery), " ORDER BY ")
        '
        ' Suppress error reporting
        On Error Resume Next
        '
        ' Extract the field list (between the SELECT and the FROM)
        .sFields = Mid(.sQuery, 8, iField_List - 8)
        '
        ' See if there is a filter
        If iFilter > 0 Then
            '
            ' See if there is a sort list
            If iSort > 0 Then
                '
                ' Extract the filter from WHERE to ORDER BY
                .sFilter = Mid(.sQuery, iFilter + 7, iSort - iFilter - 7)
            Else
                '
                ' Extract the remaining string after WHERE
                .sFilter = Mid(.sQuery, iFilter + 7)
            End If
        Else
            .sFilter = ""
        End If
        '
        ' See if there is a sort list
        If iSort > 0 Then
            '
            ' Extract the sort list (after the ORDER BY)
            .sOrder = Mid(.sQuery, iSort + 10)
        Else
            .sOrder = ""
        End If
        '
        ' Resume error reporting
        On Error GoTo 0
    End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.Parse_SQL (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' ROUTINE:  ExecuteSQLAction
' AUTHOR:   Tom Elkins
' PURPOSE:  Executes the specified SQL Action statement
' INPUT:    "sSQL" an optional SQL action statement.  If not specified, the user is prompted
' OUTPUT:   None
' NOTES:    This can be dangerous, because it opens up an avenue to the data definition statements
'           Examples:
'           To Delete from the database, "DELETE FROM <table> WHERE <conditions>"
'           To Edit values in the database, "UPDATE <table> SET <field> = <new value> [WHERE <conditions>"
Public Sub ExecuteSQLAction(Optional sSQL As String)
    Dim bLocal As Boolean
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.ExecuteSQLAction (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sSQL
    End If
    '-v1.6.1
    '
    ' Trap errors
    On Error GoTo SQLError
    '
    ' Assume the command is internal
    bLocal = False
    '
    ' Check the length of the incoming statement
    If Len(sSQL) = 0 Then
        '
        ' Prompt for entry
        sSQL = InputBox("Enter SQL command", "Modify Database " & basDatabase.guCurrent.sName, sSQL, , , App.HelpFile, basCCAT.IDH_GUI_TOOLS_SQL)
        '
        ' Specify that the user is interacting
        bLocal = True
    End If
    '
    ' See if there is a query
    If Len(sSQL) <> 0 Then
        '
        ' Log the action
        basCCAT.WriteLogEntry "INFO     : basDatabase.ExecuteSQLAction (Query = " & sSQL & ")"
        '
        ' Mark the interface busy
        frmMain.MousePointer = vbHourglass
        Screen.MousePointer = vbHourglass
        '
        ' Update the status
        frmMain.UpdateStatusText "Modifying database " & basDatabase.guCurrent.sName
        '
        ' Execute the query
        ' dbFailOnError will roll-back any changes if the query does not complete
        basDatabase.guCurrent.DB.Execute sSQL, dbFailOnError
        '
        ' Reset the mouse
        frmMain.MousePointer = vbDefault
        Screen.MousePointer = vbDefault
        '
        ' Log the number of records affected
        basCCAT.WriteLogEntry "INFO     : basDatabase.ExecuteSQLAction (" & basDatabase.guCurrent.DB.RecordsAffected & " records were modified)"
        frmMain.UpdateStatusText basDatabase.guCurrent.DB.RecordsAffected & " records were modified."
        '
        ' Inform the user about the number of records affected
        If bLocal Then MsgBox basDatabase.guCurrent.DB.RecordsAffected & " records were modified.", vbOKOnly, "SQL command executed"
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.ExecuteSQLAction (End)"
    '-v1.6.1
    '
    Exit Sub
'
' Error handler
SQLError:
    '
    ' Reset the mouser
    frmMain.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    '
    ' Log the error
    frmMain.UpdateStatusText "Error executing SQL Action.  No changes were made to the database."
    basCCAT.WriteLogEntry "ERROR    : basDatabase.ExecuteSQLAction (ERROR #" & Err.Number & " (" & Err.Description & "). No rows were affected)"
    '
    ' Inform the user of the error
    If bLocal Then MsgBox "Error #" & Err.Number & " - " & Err.Description & vbCr & "While executing SQL statement:" & vbCr & sSQL & vbCr & vbCr & "No changes made to database.", vbOKOnly Or vbMsgBoxHelpButton, "Error executing SQL", App.HelpFile, basCCAT.IDH_DB_FILTERING
    '
    ' Reset error trapping
    On Error GoTo 0
End Sub
'-v1.5
'
'+v1.5
' ROUTINE:  RemapINI
' AUTHOR:   Tom Elkins
' PURPOSE:  Re-maps certain numeric values to new text values if the INI file changed
' INPUT:    None
' OUTPUT:   None
' NOTES:    This is useful if the user changes the INI file and wants the database to be updated
'           to the new values
Public Sub RemapINI()
    Dim rsRec As Recordset  ' Records to be modified
    Dim sText As String     ' New mapped INI file value
    Dim sQuery As String    ' SQL action query to update the values
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.RemapINI (Start)"
    '-v1.6.1
    '
    ' Warn the user
    If MsgBox("This operation can take a long time, and cannot be stopped before completion." & vbCrLf & "Do you wish to start?", vbYesNo Or vbQuestion Or vbMsgBoxHelpButton, "Re-map INI file confirmation", App.HelpFile, 0) = vbYes Then
        '
        ' Mark the interface as busy
        frmMain.MousePointer = vbHourglass
        '
        ' Re-map RCV values
        '   1. Get a list of the current values for LMB only
        '   2. For each value in the list, find the new translated text
        '   3. If the new text is different than the old text, update the database
        ' Need to find a way to differentiate between LMB and HB signals
        'Set rsRec = basDatabase.guCurrent.DB.OpenRecordset("SELECT Distinct Emitter, Emitter_ID FROM Archive" & basDatabase.guCurrent.iArchive & "_Data GROUP BY Emitter_ID, Emitter")
        '
        ' Cannot do RUNMODE because the numeric code is not stored in the database
        ' Re-map Jamstat values
        '   1. Get unique Status values from the database
        '   2. For each Status value:
        '   3. Get the new text value from the INI file
        '   4. Construct a new Supplemental data field value
        '   5. Update the Other_Data field for the status value
        '   6. Return to 2
        ' Get the list of unique status values
        '+v1.6TE
        'Set rsRec = basDatabase.guCurrent.DB.OpenRecordset("SELECT Distinct Status FROM Archive" & basDatabase.guCurrent.iArchive & "_Data WHERE Msg_Type = 'MTJAMSTAT'")
        Set rsRec = basDatabase.guCurrent.DB.OpenRecordset("SELECT Distinct Status FROM [" & basDatabase.guCurrent.sArchive & basDatabase.TBL_DATA & "] WHERE Msg_Type = 'MTJAMSTAT'")
        '-v1.6
        '
        ' Loop through the list
        While Not rsRec.EOF
            '
            ' Get the new value for the current status
            sText = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & rsRec!Status, "UNKNOWN")
            '
            ' Create the action query
            '+v1.6TE
            'sQuery = "UPDATE Archive" & basDatabase.guCurrent.iArchive & "_Data SET Other_Data = [Signal_ID] & "", "" & [Status] & "", "" & " & sText & " WHERE Msg_Type = 'MTJAMSTAT' AND Status = " & rsRec!Status
            sQuery = "UPDATE [" & basDatabase.guCurrent.sArchive & basDatabase.TBL_DATA & "] SET Other_Data = [Signal_ID] & "", "" & [Status] & "", "" & '" & sText & "' WHERE Msg_Type = 'MTJAMSTAT' AND Status = " & rsRec!Status
            '-v1.6
            '
            ' Execute the action query
            basDatabase.ExecuteSQLAction sQuery
            '
            ' Move to the next status
            rsRec.MoveNext
        Wend
        '
        '+v1.6TE
        ' Re-map ORIGIN_ID values
        '   1. Get unique ORIGIN_ID values from the database
        '   2. For each ORIGIN_ID value:
        '   3.  Get the ORIGIN code from the INI file
        '   4.  Update the ORIGIN field in the database with the new value
        '   5. Return to 2
        '
        ' Get the list of unique ORIGIN_ID values
        Set rsRec = basDatabase.guCurrent.DB.OpenRecordset("SELECT Distinct Origin_ID FROM [" & basDatabase.guCurrent.sArchive & basDatabase.TBL_DATA & "]")
        '
        ' Loop through the list
        While Not rsRec.EOF
            '
            ' Get the new value for the origin code
            sText = basCCAT.GetAlias("ORIGIN", "ORIGIN" & rsRec!Origin_ID, "UNKNOWN")
            '
            ' Create the action query
            sQuery = "UPDATE [" & basDatabase.guCurrent.sArchive & basDatabase.TBL_DATA & "] SET Origin='" & sText & "' WHERE Origin_ID = " & rsRec!Origin_ID
            '
            ' Execute the action query
            basDatabase.ExecuteSQLAction sQuery
            '
            ' Move to the next Origin code
            rsRec.MoveNext
        Wend
        '-v1.6
        '
        ' Reset the mouse
        frmMain.MousePointer = vbDefault
        '
        ' Cannot do XMTRSTAT because the numeric values are not stored in the database
        ' Cannot do HBFUNC because the numeric values are not stored in the database
        ' Cannot do HBOPT because the numeric values are not stored in the database
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.RemapINI (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' PROPERTY: Interactive
' AUTHOR:   Tom Elkins
' PURPOSE:  Determines if the user is to be prompted for information or not
' STATES:   TRUE = The user is to be prompted to provide information
'           FALSE = The user is not prompted for information
' NOTES:
Public Property Let Interactive(bState As Boolean)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: basDatabase.Interactive Let (Start)"
    '-v1.6.1
    '
    pbInteractive = bState
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: basDatabase.Interactive Let (End) = " & bState
    '-v1.6.1
    '
End Property
'
Public Property Get Interactive() As Boolean
    Interactive = pbInteractive
End Property
'-v1.5
'
'+v1.5
' ROUTINE:  UpgradeDatabase
' AUTHOR:   Tom Elkins
' PURPOSE:  Converts and copies data from an old database into a new database with the new schema
' INPUT:    None
' OUTPUT:   None
' NOTES:    The old database is kept, but renamed
Public Sub UpgradeDatabase()
    Dim dbNew As Database           ' The new database created in the current version
    Dim rsOld As Recordset          ' Pointer to the old database tables
    Dim rsNew As Recordset          ' Pointer to the new database tables
    Dim rsOldData As Recordset      ' Pointer to the old database data
    Dim rsNewData As Recordset      ' Pointer to the new database data
    Dim dtArchiveDate As Date       ' The archive date stamp - all time values will be relative to this date
    Dim sTimeParts() As String      ' Array to hold the components of the old DDD:HH:MM:SS time
    Dim dOffset As Double           ' Difference (in days) between the old time values and the archive date stamp
    Dim sName As String             ' Database file name
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.UpgradeDatabase (Start)"
    '-v1.6.1
    '
    ' Trap errors
    On Error GoTo UpgradeFailed
    '
    ' prevent user from being presented forms for data entry
    basDatabase.Interactive = False
    '
    ' Prevent the user from interacting with the other controls
    frmMain.MousePointer = vbHourglass
    '
    ' Log the start of the operation
    basCCAT.WriteLogEntry "INFO     : basDatabase.UpgradeDatabase (Starting upgrade process on " & guCurrent.DB.Name & ")"
    '
    ' Save the database name
    sName = guCurrent.DB.Name
    '
    ' Update status
    frmMain.UpdateStatusText "Creating new database and schema"
    '
    ' Create a new database
    If basDatabase.bCreate_New_Database(sName & ".new") Then
        '
        ' Open the new database
        frmMain.UpdateStatusText "Opening new database"
        Set dbNew = OpenDatabase(sName & ".new")
        '
        ' Open the Info table of both databases
        frmMain.UpdateStatusText "Copying Info table contents"
        basCCAT.WriteLogEntry "INFO     : basDatabase.UpgradeDatabase (Copying Info table contents)"
        Set rsOld = guCurrent.DB.OpenRecordset("Info")
        Set rsNew = dbNew.OpenRecordset("Info")
        '
        ' Copy the old info data into the new table
        rsNew.Edit
        rsNew!Name = rsOld!Name
        rsNew!Start = rsOld!Start
        rsNew!End = rsOld!End
        rsNew!Description = rsOld!Description
        rsNew.Update
        rsNew.Close
        rsOld.Close
        '
        ' Open the Archives table of both databases
        Set rsOld = guCurrent.DB.OpenRecordset("Archives")
        Set rsNew = dbNew.OpenRecordset("Archives")
        '
        ' Loop through all of the archive entries in the old table
        While Not rsOld.EOF
            '
            ' Add and populate an archive entry
            frmMain.UpdateStatusText "Copying information for Archive" & rsOld!ID
            basCCAT.WriteLogEntry "INFO     : basDatabase.UpgradeDatabase (Copying information for Archive" & rsOld!ID & ")"
            rsNew.AddNew
            rsNew!ID = rsOld!ID
            rsNew!Name = rsOld!Name
            '
            ' Confirm the archive date stamp
            frmMain.MousePointer = vbDefault
            dtArchiveDate = Int(CDate(InputBox("Please make sure the correct archive start date is shown below." & vbCr & "If it is incorrect, enter the correct date." & vbCr & "The time value for all data records will be computed from this date.", "Confirm Archive Date", rsOld!Date, , , App.HelpFile, basCCAT.IDH_GUI_TOOLS_UPDATE)))
            rsNew!Date = dtArchiveDate
            frmMain.MousePointer = vbHourglass
            basCCAT.WriteLogEntry "INFO     : basDatabase.UpgradeDatabase (Archive Date stamp = " & dtArchiveDate & ")"
            '
            ' Get the TRUE start and stop times from the data
            frmMain.UpdateStatusText "Querying database for first and last times"
            basCCAT.WriteLogEntry "INFO     : basDatabase.UpgradeDatabase (Querying database for first and last message times)"
            Set rsOldData = basDatabase.guCurrent.DB.OpenRecordset("SELECT Min(ReportTime) As Min, Max(ReportTime) As Max FROM " & rsOld!Name & TBL_DATA)
            frmMain.UpdateStatusText "Converting old time strings to true date/time values"
            If Not rsOldData.EOF Then
                '
                ' Compute the start time from the archive date
                ' Compute the end time from the start time
                rsNew!Start = dtArchiveDate + (rsOldData!Min - Int(rsOldData!Min))
                rsNew!End = rsNew!Start + (rsOldData!Max - rsOldData!Min)
            Else
                '
                ' Parse the time strings, compute the seconds, and add to the archive date
                dOffset = basCCAT.dHumanTimeToTSecs(rsOld!Start)
                rsNew!Start = dtArchiveDate + ((dOffset / 86400#) - Int(dOffset / 86400#))
                rsNew!End = DateAdd("s", basCCAT.dHumanTimeToTSecs(rsOld!End) - dOffset, rsNew!Start)
            End If
            rsOldData.Close
            '
            ' Copy other archive info
            rsNew!Archive = rsOld!Archive
            rsNew!Media = rsOld!Media
            rsNew!Processed = rsOld!Processed
            rsNew!Analysis_File = rsOld!Analysis_File
            rsNew!Messages = rsOld!Messages
            rsNew!Bytes = rsOld!Bytes
            rsNew.Update
            '
            ' Add a summary table for this archive
            frmMain.UpdateStatusText "Creating summary table in new database"
            If basDatabase.bCreate_Summary_Table(dbNew, rsOld!ID) Then
                Set rsOldData = basDatabase.guCurrent.DB.OpenRecordset(rsOld!Name & TBL_SUMMARY)
                Set rsNewData = dbNew.OpenRecordset(rsOld!Name & TBL_SUMMARY)
                '
                ' Add and populate a message summary entry
                While Not rsOldData.EOF
                    frmMain.UpdateStatusText "Copying summary information for message " & rsOldData!Message
                    basCCAT.WriteLogEntry "INFO     : basdatabase.UpgradeDatabase (Copying summary information for message " & rsOldData!Message & ")"
                    rsNewData.AddNew
                    rsNewData!Message = rsOldData!Message
                    rsNewData!MSG_ID = rsOldData!MSG_ID
                    rsNewData!Count = rsOldData!Count
                    If IsDate(rsOldData!First) Then
                        rsNewData!First = dtArchiveDate + (CDate(rsOldData!First) - Int(CDate(rsOldData!First)))
                        rsNewData!Last = rsNewData!First + (CDate(rsOldData!Last) - CDate(rsOldData!First))
                    Else
                        rsNewData!First = dtArchiveDate + ((rsOldData!First / 86400#) - Int(rsOldData!First / 86400#))
                        rsNewData!Last = DateAdd("s", rsOldData!Last - rsOldData!First, rsNewData!First)
                    End If
                    rsNewData!Description = rsOldData!Description
                    rsNewData.Update
                    '
                    rsOldData.MoveNext
                Wend
                rsOldData.Close
                rsNewData.Close
            End If
            '
            ' Add a data table for this archive
            frmMain.UpdateStatusText "Creating data table in new database"
            If basDatabase.bCreate_Data_Table(dbNew, rsOld!ID) Then
                Set rsOldData = basDatabase.guCurrent.DB.OpenRecordset(rsOld!Name & TBL_DATA)
                rsOldData.MoveLast
                rsOldData.MoveFirst
                Set rsNewData = dbNew.OpenRecordset(rsOld!Name & TBL_DATA)
                '
                ' Compute offset
                dOffset = dtArchiveDate - Int(rsOldData!ReportTime)
                basCCAT.WriteLogEntry "INFO     : basDatabase.UpgradeDatabase (Archive time offset (new - old) = " & dOffset & " days)"
                '
                ' Set up the progress bar
                frmMain.ShowProgressBar 0, 100, 0
                basCCAT.WriteLogEntry "INFO     : basdatabase.UpgradeDatabase (Copying data table)"
                DoEvents
                '
                ' Loop through the data records
                While Not rsOldData.EOF
                    '
                    ' Add and populate a data record
                    rsNewData.AddNew
                    rsNewData!ReportTime = rsOldData!ReportTime + dOffset
                    rsNewData!Msg_Type = rsOldData!Msg_Type
                    rsNewData!Rpt_Type = rsOldData!Rpt_Type
                    rsNewData!Origin = rsOldData!Origin
                    rsNewData!Origin_ID = rsOldData!Origin_ID
                    rsNewData!Target_ID = rsOldData!Target_ID
                    rsNewData!Latitude = rsOldData!Latitude
                    rsNewData!Longitude = rsOldData!Longitude
                    rsNewData!Altitude = rsOldData!Altitude
                    rsNewData!Heading = rsOldData!Heading
                    rsNewData!Speed = rsOldData!Speed
                    rsNewData!Parent = rsOldData!Parent
                    rsNewData!Parent_ID = rsOldData!Parent_ID
                    rsNewData!Allegiance = rsOldData!Allegiance
                    rsNewData!IFF = rsOldData!IFF
                    rsNewData!Emitter = rsOldData!Emitter
                    rsNewData!Emitter_ID = rsOldData!Emitter_ID
                    rsNewData!Signal = rsOldData!Signal
                    rsNewData!Signal_ID = rsOldData!Signal_ID
                    rsNewData!Frequency = rsOldData!Frequency
                    rsNewData!PRI = rsOldData!PRI
                    rsNewData!Status = rsOldData!Status
                    rsNewData!Tag = rsOldData!Tag
                    rsNewData!Flag = rsOldData!Flag
                    rsNewData!Common_ID = rsOldData!Common_ID
                    rsNewData!Range = rsOldData!Range
                    rsNewData!Bearing = rsOldData!Bearing
                    rsNewData!Elevation = rsOldData!Elevation
                    rsNewData!XX = rsOldData!XX
                    rsNewData!XY = rsOldData!XY
                    rsNewData!YY = rsOldData!YY
                    rsNewData!Other_Data = rsOldData!Other_Data
                    rsNewData.Update
                    '
                    ' Update the progress bar
                    frmMain.barLoad.Value = rsOldData.PercentPosition
                    '
                    rsOldData.MoveNext
                Wend
                rsOldData.Close
                rsNewData.Close
                frmMain.barLoad.Visible = False
                DoEvents
            End If
            rsOld.MoveNext
        Wend
    End If
    '
    ' Report completion
    frmMain.UpdateStatusText "Upgrade complete, renaming and loading new database..."
    basDatabase.Interactive = True
    Set rsOldData = Nothing
    Set rsOld = Nothing
    Set rsNewData = Nothing
    Set rsNew = Nothing
    dbNew.Close
    Set dbNew = Nothing
    '
    ' Close the current database
    basDatabase.guCurrent.DB.Close
    '
    ' Rename the old database
    Name sName As sName & ".old"
    '
    ' Rename the new database
    Name sName & ".new" As sName
    '
    ' Open the new database
    Set basDatabase.guCurrent.DB = OpenDatabase(sName)
    '
    ' Reset the mouse pointer
    frmMain.MousePointer = vbDefault
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.UpgradeDatabase (End)"
    '-v1.6.1
    '
    Exit Sub
'
UpgradeFailed:
    frmMain.MousePointer = vbDefault
    '
    basCCAT.WriteLogEntry "ERROR    : basDatabase.UpgradeDatabase (Error #" & Err.Number & " - " & Err.Description & ")"
    '
    MsgBox "Database Upgrade Operation Failed.", vbOKOnly Or vbExclamation, "Operation Failed"
End Sub
'-v1.5
'
'+v1.5
' PROPERTY: LastQuery
' AUTHOR:   Tom Elkins
' PURPOSE:  Stores the last query executed
' LET:      A string value containing the query to save
' GET:      A string value containing the last query saved
' NOTES:
Public Property Let LastQuery(sQ As String)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: basDatabase.LastQuery (Start) = " & sQ
    '-v1.6.1
    '
    psLastQuery = sQ
End Property
'
Public Property Get LastQuery() As String
    LastQuery = psLastQuery
End Property
'-v1.5
'
'+v1.5
' PROPERTY: CurrentQuert
' AUTHOR:   Tom Elkins
' PURPOSE:  Stores the current query executed
' LET:      A string value containing the query just executed
' GET:      A string value containing the current query
' NOTES:
Public Property Let CurrentQuery(sQ As String)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: basDatabase.CurrentQuery Let (Start) = " & sQ
    '-v1.6.1
    '
    psCurrentQuery = sQ
End Property
'
Public Property Get CurrentQuery() As String
    CurrentQuery = psCurrentQuery
End Property
'-v1.5
'
'+v1.5
' PROPERTY: NewQuery
' AUTHOR:   Tom Elkins
' PURPOSE:  Stores the query being created
' LET:      A string value containing a new query to be executed
' GET:      A string value containing the latest stage of the new query
' NOTES:    Not currently in use
Public Property Let NewQuery(sQ As String)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "PROPERTY: basDatabase.NewQuery (Start) = " & sQ
    '-v1.6.1
    '
    psNewQuery = sQ
End Property
'
Public Property Get NewQuery() As String
    NewQuery = psNewQuery
End Property
'-v1.5
'
'+v1.5
' ROUTINE:  QueryData
' AUTHOR:   Tom Elkins
' PURPOSE:  Executes a query and displays the results in the data grid
' INPUT:    "sQuery" is an optional string value containing the query to be executed
' OUTPUT:   None
' NOTES:
Public Sub QueryData(Optional sQuery As String)
    Dim rsData As Recordset     ' The recordset containing the results of the query
    Dim bChanged As Boolean     ' A flag to indicate if the data grid was changed
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.QueryData (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sQuery
    End If
    '-v1.6.1
    '
    ' Reset the flag to false (no changes to data grid yet)
    bChanged = False
    '
    ' If a query was not passed in, use the last saved query
    If sQuery = "" Then sQuery = basDatabase.LastQuery
    '
    ' Trap any errors
    On Error GoTo BadQuery
    '
    ' Mark the interface as busy
    frmMain.MousePointer = vbHourglass
    '
    ' Inform the user that we are busy executing the query
    frmMain.UpdateStatusText "Querying database..."
    '
    ' Execute the query
    Set rsData = basDatabase.guCurrent.DB.OpenRecordset(sQuery, dbOpenSnapshot, dbReadOnly, dbReadOnly)
   '
    ' See if any data was passed back
    If Not rsData Is Nothing Then
        '
        ' Save the last query
        basDatabase.LastQuery = basDatabase.guCurrent.uSQL.sQuery
        '
        ' Save the current query
        basDatabase.guCurrent.uSQL.sQuery = sQuery
        '
        ' Parse the current query
        basDatabase.Parse_SQL sQuery
        '
        ' Update the data grid
        Set frmMain.Data1.Recordset = rsData
        '
        ' Mark that the data grid was changed
        bChanged = True
        '
        ' Move to the last record and back to get an accurate count
        frmMain.Data1.Recordset.MoveLast
        frmMain.Data1.Recordset.MoveFirst
        '
        ' Inform the user of the number of records
        frmMain.UpdateStatusText frmMain.Data1.Recordset.RecordCount & " record" & IIf(frmMain.Data1.Recordset.RecordCount > 1, "s match", " matches") & " the specified query."
        '
        ' Display the query above the data grid
        frmMain.lblTitle(1).Caption = sQuery
        '
        ' Free the memory for the record pointer
        Set rsData = Nothing
    End If
    '
    ' Return the mouse pointer to its default icon
    frmMain.MousePointer = vbDefault
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.QueryData (End)"
    '-v1.6.1
    '
    ' Return to the calling routine
    Exit Sub
'
' Handle any errors
BadQuery:
    Dim lNum As Long        ' Persist the error number
    Dim sErr As String      ' Persist the error description
    Dim sSrc As String      ' Persist the error source
    '
    ' Save the error values
    lNum = Err.Number
    sErr = Err.Description
    sSrc = Err.Source
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basDatabase.QueryData (Error #" & lNum & " - " & sErr & " reported in " & sSrc & ")"
    '-v1.6.1
    '
    ' Reset the mouse pointer
    frmMain.MousePointer = vbDefault
    '
    ' Free up the memory for the record pointer
    Set rsData = Nothing
    '
    ' Check the error number
    Select Case lNum
        '
        ' No current record - occurs when a query returns no records
        Case 3021:
            '
            ' Inform the user that no records were found
            frmMain.UpdateStatusText "0 records match the specified query"
            MsgBox "0 Records returned!", vbOKOnly Or vbInformation, "No records found"
        '
        ' All other errors
        Case Else:
            '
            ' Write the log entry, and inform the user there was a problem
            frmMain.UpdateStatusText "Error processing query!"
            MsgBox "Error #" & lNum & " - " & sErr & vbCrLf & vbCrLf & "Query : """ & sQuery & """", vbOKOnly, sSrc & " reported an error wile processing query", App.HelpFile, basCCAT.IDH_DB_FILTERING
            If bChanged Then basDatabase.QueryData
            'Err.Raise lNum, sSrc, sErr
    End Select
End Sub
'-v1.5
'
'+v1.5
' FUNCTION: lExportGrid
' AUTHOR:   Tom Elkins
' PURPOSE:  Exports the contents of the data grid to a file
' INPUT:    "iFile" is the file ID to write the records
' OUTPUT:   A long integer with the number of records exported
' NOTES:
Public Function lExportGrid(iFile As Integer) As Long
    Dim fldField As Field       ' Pointer to a field object
    Dim sOutput As String       ' The output record
    Dim sTemp As String         ' A temporary holding variable
    Dim lNum_Output As Long     ' The number of records output
    Dim sTxtDelimiter As String
    Dim sTxtSpace As String
    Dim sTxtBlank As String
    Dim sLngFmt As String
    Dim sDblFmt As String
    Dim sDTmFmt As String
    Dim bTSecs As Boolean
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.lExportGrid (Start)"
    '-v1.6.1
    '
    ' Trap all errors
    On Error GoTo Bad_Export
    '
    ' Mark the interface busy
    frmMain.MousePointer = vbHourglass
    '
    ' If the data grid is invalid, re-execute the query
    If frmMain.Data1.Recordset Is Nothing Then basDatabase.QueryData basDatabase.sCreate_SQL
    '
    ' Reset the output record count
    lNum_Output = 0
    '
    ' Get the update interval from the INI file. This is used for long exports to let the
    ' user know it is still working.
    If guGUI.lInterval = 0 Then guGUI.lInterval = basCCAT.GetNumber("Miscellaneous Operations", "Update_Interval", 1250)
    '
    ' Use recordset-level addressing
    With frmMain.Data1.Recordset
        '
        ' Ensure we have a complete record count and the pointer is at the top of the record set
        .MoveLast
        .MoveFirst
        '
        ' Set up the progress bar
        frmMain.ShowProgressBar 0, .RecordCount, lNum_Output
        '
        ' Get the export formats from the INI file
        sTxtDelimiter = basCCAT.GetAlias("Export", "Text_Delimiter", "")
        sTxtSpace = basCCAT.GetAlias("Export", "Text_Space", " ")
        sTxtBlank = basCCAT.GetAlias("Export", "Text_Blank", "UNKNOWN")
        sLngFmt = basCCAT.GetAlias("Export", "Long_Format", "0")
        sDblFmt = basCCAT.GetAlias("Export", "Double_Format", "0.00000")
        sDTmFmt = basCCAT.GetAlias("Export", "Time_Format", "mm/dd/yyyy hh:nn:ss.000")
        bTSecs = (basCCAT.GetNumber("Export", "TSecs", 1) = 1)
        '
        ' Loop through all of the records
        While Not .EOF
            '
            ' Clear the output string
            sOutput = ""
            '
            ' Loop through all the fields
            For Each fldField In .Fields
                '
                ' Format the data based on field type
                Select Case fldField.Type
                    '
                    ' Text fields
                    Case dbText:
                        '
                        ' Replace spaces within text fields
                        sTemp = Replace(Trim(fldField.Value), " ", sTxtSpace)
                        '
                        ' See if the text is blank
                        If Len(sTemp) = 0 Then
                            '
                            ' See if this is the Report Type field
                            If fldField.Name = "Rpt_Type" Then
                                '
                                ' Assume the record matches the file type
                                sTemp = basCCAT.gaDAS_Rec_Type(guExport.iRec_Type)
                            Else
                                '
                                ' Use the default blank value
                                sTemp = sTxtBlank
                            End If
                        End If
                        '
                        ' Add the field to the output record
                        sOutput = sOutput & sTxtDelimiter & Trim(sTemp) & sTxtDelimiter & ","
                    '
                    ' Long integers
                    Case dbLong:
                        '
                        ' Format the number and add it to the output record
                        sOutput = sOutput & Format(fldField.Value, sLngFmt) & ","
                    '
                    ' Double floats
                    Case dbDouble:
                        '
                        ' Format the number and add it to the output record
                        sOutput = sOutput & Format(fldField.Value, sDblFmt) & ","
                    '
                    ' Date format
                    Case dbDate:
                        '
                        '+v1.5
                        'sOutput = sOutput & Format((CDbl(fldField.Value) * 86400#) + guCurrent.uArchive.dOffset_Time, "0.000") & ","
                        '
                        If bTSecs Then
                            ' Subtract the days from the value, which gives the time in the day
                            ' Multiply the time by 86400 to get seconds
                            ' Get the day of the year (JDay) from the value and add 1 (1 Jan = day 0)
                            ' Multiply the JDay by 86400 to convert to seconds
                            ' Add JDay to time to get TSecs
                            ' Output TSecs for the DAS file format.
                            sOutput = sOutput & Format((CDbl(fldField.Value - Int(fldField.Value)) * 86400#) + ((DatePart("y", fldField.Value) - 1) * 86400#), "0.000") & ","
                        Else
                            sOutput = sOutput & Format(fldField.Value, sDTmFmt) & ","
                        End If
                        '-v1.5
                        '
                End Select
            Next fldField
            '
            ' Write the output string to the file (minus the trailing comma)
            Print #iFile, Mid(sOutput, 1, Len(sOutput) - 1)
            '
            ' Update the count
            lNum_Output = lNum_Output + 1
            frmMain.barLoad.Value = lNum_Output
            '
            ' Periodically release control to the system for other operations
            If lNum_Output Mod basCCAT.guGUI.lInterval = 0 Then
                DoEvents
            End If
            '
            ' Move to the next record
            frmMain.Data1.Recordset.MoveNext
        Wend
        '
        ' Hide the progress bar
        frmMain.barLoad.Visible = False
        '
        ' Report the number of records output
        frmMain.UpdateStatusText "Exported " & lNum_Output & " out of " & .RecordCount & " records."
        '
        ' Report if 0 records were output
        If lNum_Output = 0 Then MsgBox lNum_Output & " out of " & .RecordCount & " records were exported." & vbCr & vbCr & "Your query was '" & guCurrent.uSQL.sQuery & "'", vbOKOnly Or vbExclamation Or vbMsgBoxHelpButton, "No Records Found", App.HelpFile, basCCAT.IDH_DB_FILTERING
        '
        ' Log the event
        basCCAT.WriteLogEntry "INFO     : basDatabase.lExportGrid (Exported " & lNum_Output & " out of " & .RecordCount & " records)"
    End With
    '
    ' Change mouse
    frmMain.MousePointer = vbDefault
    '
    ' Pass back the number of records exported
    lExportGrid = lNum_Output
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basDatabase.lExportGrid (End)"
    '-v1.6.1
    '
    Exit Function
'
' Handle errors
Bad_Export:
    Dim lNum As Long
    Dim sErr As String
    Dim sSrc As String
    '
    '
    lNum = Err.Number
    sErr = Err.Description
    sSrc = Err.Source
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basDatabase.lExportGrid (Error #" & lNum & " - " & sErr & " from " & sSrc & ")"
    '-v1.6.1
    '
    '
    Select Case lNum
        '
        ' Unknown
        Case Else:
            '
            ' Inform the user
            MsgBox "Error #" & lNum & " - " & sErr & vbCr & "While attempting to export grid", vbOKOnly Or vbExclamation Or vbMsgBoxHelpButton, "Error reported by " & sSrc, App.HelpFile, basCCAT.IDH_DB_FILTERING
    End Select
    'Err.Raise lNum, sSrc, sErr & vbCrLf & "In basDatabase.lExportGrid"
End Function
'-v1.5
'
'+v1.6TE
' ROUTINE:  blnCreateSummaryTable
' AUTHOR:   Tom Elkins
' PURPOSE:  Creates the default, blank archive summary table in the specified database
' INPUT:    "dbCurrent" is the currently selected database
'           "strArchive" is the name of the new archive
' OUTPUT:   True if the table was created
'           False if the table was not created
' NOTES:    Summary tables contain results from processing an archive.
'               Message is the name of a message in the archive
'               MSG_ID is the numeric identifier of the message type
'               Count is the number of messages in the archive of the current type
'               First is the time of the first occurance of this message in the archive
'               Last is the time of the last occurance of this message in the archive
Public Function blnCreateSummaryTable(dbCurrent As Database, strArchive As String) As Boolean
    Dim tblSummary As TableDef  ' New summary table
    '
    ' Trap errors
    On Error GoTo Hell
    '
    ' Set the default return value
    blnCreateSummaryTable = False
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.blnCreateSummaryTable (Start)"
    basCCAT.WriteLogEntry "ARGUMENTS: DB = " & dbCurrent.Name & ", Archive = " & strArchive
    '
    ' Create the table
    Set tblSummary = dbCurrent.CreateTableDef(strArchive & TBL_SUMMARY)
    '
    ' Use table-level addressing
    With tblSummary
        '
        ' Add the fields
        .Fields.Append .CreateField("Message", dbText, 20)
        .Fields.Append .CreateField("MSG_ID", dbLong)
        .Fields.Append .CreateField("Count", dbLong)
        '+v1.5
        ' Changed database schema to use dates instead of text
        '.Fields.Append .CreateField("First", dbText, 20)
        '.Fields.Append .CreateField("Last", dbText, 20)
        .Fields.Append .CreateField("First", dbDate)
        .Fields.Append .CreateField("Last", dbDate)
        '-v1.5
        .Fields.Append .CreateField("Description", dbText, 255)
        '.Fields.Append .CreateField("Signal", dbBoolean)
        '.Fields.Append .CreateField("LOB", dbBoolean)
        '.Fields.Append .CreateField("Fix", dbBoolean)
        '.Fields.Append .CreateField("Track", dbBoolean)
        '
        ' Set the field attribute to allow null strings
        .Fields("Message").AllowZeroLength = True
        .Fields("First").AllowZeroLength = True
        .Fields("Last").AllowZeroLength = True
        .Fields("Description").AllowZeroLength = True
    End With
    '
    ' Add table to database
    dbCurrent.TableDefs.Append tblSummary
    '
    ' Set the return value
    blnCreateSummaryTable = True
    '
    '+v1.6.1TE
    basCCAT.WriteLogEntry "ROUTINE  : basDatabase.blnCreateSummaryTable (End)"
    '-v1.6.1
    '
    ' Leave
    Exit Function
'
' Error handler
Hell:
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    ' Set the return value to false
    blnCreateSummaryTable = False
    '
    ' Log the error
    basCCAT.WriteLogEntry "ERROR    : basDatabase.blnCreateSummaryTable (Error #" & Err.Number & " - " & Err.Description & ")"
    '
    ' Inform the user
    MsgBox "Error #" & Err.Number & vbCr & Err.Description, vbOKOnly, "Error Creating Summary Table"
End Function
'-v1.6
'
'+v1.6TE
' FUNCTION: blnCreateDataTable
' AUTHOR:   Tom Elkins
' PURPOSE:  Creates the default, blank message data table in the specified database
' INPUT:    "dbCurrent" is the currently selected database
'           "strArchive" is the new archive name
' OUTPUT:   True if successful, False if not
' NOTES:    Data tables contain actual values from a message.  The fields are
'           from the DAS master list
'           Data table names are "<Archive Name>_Data"
Public Function blnCreateDataTable(dbCurrent As Database, strArchive As String) As Boolean
    Dim tblData As TableDef     ' New data table
    '
    ' Trap errors
    On Error GoTo Hell
    '
    ' Log the event
    basCCAT.WriteLogEntry "FUNCTION : basDatabase.blnCreateDataTable (Start)"
    basCCAT.WriteLogEntry "ARGUMENTS: DB = " & dbCurrent.Name & ", Archive = " & strArchive
    '
    ' Create the table
    Set tblData = dbCurrent.CreateTableDef(strArchive & TBL_DATA)
    '
    ' Use table-level addressing
    With tblData
        '
        ' Add the fields
        .Fields.Append .CreateField("ReportTime", dbDate)
        .Fields.Append .CreateField("Msg_Type", dbText, 30)
        .Fields.Append .CreateField("Rpt_Type", dbText, 10)
        .Fields.Append .CreateField("Origin", dbText, 30)
        .Fields.Append .CreateField("Origin_ID", dbLong)
        .Fields.Append .CreateField("Target_ID", dbLong)
        .Fields.Append .CreateField("Latitude", dbDouble)
        .Fields.Append .CreateField("Longitude", dbDouble)
        .Fields.Append .CreateField("Altitude", dbDouble)
        .Fields.Append .CreateField("Heading", dbDouble)
        .Fields.Append .CreateField("Speed", dbDouble)
        .Fields.Append .CreateField("Parent", dbText, 50)
        .Fields.Append .CreateField("Parent_ID", dbLong)
        .Fields.Append .CreateField("Allegiance", dbText, 20)
        .Fields.Append .CreateField("IFF", dbLong)
        .Fields.Append .CreateField("Emitter", dbText, 80)
        .Fields.Append .CreateField("Emitter_ID", dbLong)
        .Fields.Append .CreateField("Signal", dbText, 50)
        .Fields.Append .CreateField("Signal_ID", dbLong)
        .Fields.Append .CreateField("Frequency", dbDouble)
        .Fields.Append .CreateField("PRI", dbDouble)
        .Fields.Append .CreateField("Status", dbLong)
        .Fields.Append .CreateField("Tag", dbLong)
        .Fields.Append .CreateField("Flag", dbLong)
        .Fields.Append .CreateField("Common_ID", dbLong)
        .Fields.Append .CreateField("Range", dbDouble)
        .Fields.Append .CreateField("Bearing", dbDouble)
        .Fields.Append .CreateField("Elevation", dbDouble)
        .Fields.Append .CreateField("XX", dbDouble)
        .Fields.Append .CreateField("XY", dbDouble)
        .Fields.Append .CreateField("YY", dbDouble)
        .Fields.Append .CreateField("Other_Data", dbText)
        '
        ' Set the field attribute to allow null strings
        .Fields("Msg_Type").AllowZeroLength = True
        .Fields("Rpt_Type").AllowZeroLength = True
        .Fields("Origin").AllowZeroLength = True
        .Fields("Parent").AllowZeroLength = True
        .Fields("Allegiance").AllowZeroLength = True
        .Fields("Emitter").AllowZeroLength = True
        .Fields("Signal").AllowZeroLength = True
        .Fields("Other_Data").AllowZeroLength = True
    End With
    '
    ' Add table to the database
    dbCurrent.TableDefs.Append tblData
    '
    ' Add table to Sources table
    Call Add_to_Cinnabar_Tables(dbCurrent, strArchive & TBL_DATA)
    '
    ' Resume error reporting
    On Error GoTo 0
    '
    ' Set the data table
    guCurrent.uSQL.sTable = strArchive & TBL_DATA
    '
    ' Success
    blnCreateDataTable = True
    '
    ' Log the event
    basCCAT.WriteLogEntry "INFO     : basDatabase.blnCreateDataTable (End - Successfully created data table for " & strArchive & ")"
    '
    ' Leave
    Exit Function
'
' Error handler
Hell:
    '
    ' Report failure
    blnCreateDataTable = False
    '
    ' Log the event
    basCCAT.WriteLogEntry "ERROR    : basDatabase.blnCreateDataTable (Error #" & Err.Number & " - " & Err.Description & ")"
    '
    ' Inform the user
    MsgBox "ERROR #" & Err.Number & vbCr & Err.Description, vbOKOnly Or vbCritical, "Error Creating Data Table"
    '
    ' Restore error reporting
    On Error GoTo 0
End Function
'-v1.6
'
'+v1.5
' ROUTINE:  AddValueFilter
' AUTHOR:   Tom Elkins
' PURPOSE:  Appends a filter to the current query
' INPUT:    "strFilter" is the filter to be added to the list
' OUTPUT:   None
' NOTES:
Public Sub AddValueFilter(strFilter As String)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basDatabase.AddValueFilter (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & strFilter
    End If
    '-v1.6.1
    '
    ' Surround the current filter with parentheses and add the new filter
    basDatabase.guCurrent.uSQL.sFilter = "(" & basDatabase.guCurrent.uSQL.sFilter & ") AND " & strFilter
    '
    ' Execute the query
    basDatabase.QueryData basDatabase.sCreate_SQL
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basDatabase.AddValueFilter (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
