Attribute VB_Name = "VSOps"
Dim rsVS As Recordset
Dim newVS As Recordset
Dim defVS As Recordset

Const VS_DEFAULT = "DefaultVS"

Public Sub CreateDefaultVS()

   Dim bTblExists As Boolean
   
   bTblExists = False
   
   On Error GoTo NoVS
   
   Set newVS = guCurrent.DB.OpenRecordset(VS_DEFAULT & "_VarStruct", dbOpenDynaset)
   If Not (newVS Is Nothing) Then
      If (MsgBox("Default VarStruct already exists.  Do you want to Overwrite it?", vbYesNo) = vbNo) Then
         Exit Sub
      End If
      bTblExists = True
   End If
               
NoVS:
   If Not bTblExists Then
      'Create Default VarStruct table in current db
      If basTOC.blnCreateVarStructTable(basDatabase.guCurrent.DB, VS_DEFAULT) Then
         'Open new CopyTo VarStruct table
         Set newVS = guCurrent.DB.OpenRecordset(VS_DEFAULT & "_VarStruct", dbOpenDynaset)
      Else
         MsgBox ("Unable to open DefaultVS Table in current DB")
         Exit Sub
      End If
   End If
   
   Screen.MousePointer = vbHourglass

   'Open Current selected VarStruct table
   Set rsVS = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_VarStruct", dbOpenDynaset)

   rsVS.MoveFirst
   If bTblExists Then
      newVS.MoveFirst
   End If
   
   Dim i As Integer
   'Loop through current VarStruct table and save values to Default VS table
   For i = 1 To rsVS.RecordCount
      If bTblExists Then
         newVS.Edit
      Else
         newVS.AddNew
      End If
      newVS!varStructID = rsVS!varStructID
      newVS!msgid = rsVS!msgid
      newVS!fieldname = rsVS!fieldname
      newVS!FieldSize = rsVS!FieldSize
      newVS!DataType = rsVS!DataType
      newVS!ConvType = rsVS!ConvType
      newVS!fieldlabel = rsVS!fieldlabel
      newVS!DasField = rsVS!DasField
      newVS!MultiEntry = rsVS!MultiEntry
      newVS!MultiRecPtr = rsVS!MultiRecPtr
      newVS!StructLevel = rsVS!StructLevel
      newVS.Update
      
      rsVS.MoveNext
      If bTblExists Then
         newVS.MoveNext
      End If
   Next i
   
   newVS.Close
   rsVS.Close
   
   Screen.MousePointer = vbDefault

End Sub

Public Sub ImportDefaultVS()

   'Open default VarStruct table
   Set defVS = guCurrent.DB.OpenRecordset(VS_DEFAULT & "_VarStruct", dbOpenDynaset)
   'Open current selected VarStruct table
   Set rsVS = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_VarStruct", dbOpenDynaset)

   If defVS Is Nothing Then
      MsgBox ("There is NO Default VarStruct table in the current DB to import.")
      Exit Sub
   End If
      
   If (MsgBox("You are about to Overwrite " & guCurrent.sArchive & " VarStruct" & _
              " Are you sure you want to continue?", vbYesNo) = vbNo) Then
       Exit Sub
   End If

   rsVS.MoveFirst
   defVS.MoveFirst
   
   Screen.MousePointer = vbHourglass
   Dim i As Integer
   'Loop through default VarStruct table and save values to current VarStruct table
   For i = 1 To defVS.RecordCount
      rsVS.Edit
      rsVS!varStructID = defVS!varStructID
      rsVS!msgid = defVS!msgid
      rsVS!fieldname = defVS!fieldname
      rsVS!FieldSize = defVS!FieldSize
      rsVS!DataType = defVS!DataType
      rsVS!ConvType = defVS!ConvType
      rsVS!fieldlabel = defVS!fieldlabel
      rsVS!DasField = defVS!DasField
      rsVS!MultiEntry = defVS!MultiEntry
      rsVS!MultiRecPtr = defVS!MultiRecPtr
      rsVS!StructLevel = defVS!StructLevel
      rsVS.Update
      
      rsVS.MoveNext
      defVS.MoveNext
   Next i
   defVS.Close
   rsVS.Close
   
   Screen.MousePointer = vbDefault

End Sub
