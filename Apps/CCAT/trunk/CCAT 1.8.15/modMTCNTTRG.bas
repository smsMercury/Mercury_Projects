Attribute VB_Name = "modMTCNTTRG"
Private MSG_NAME As String
Private Const MSG_ID = 22

Private mrsData As Recordset
'
'
Public Sub Create_MTCNTTRG()
    MSG_NAME = guCurrent.sArchive & "_MTCNTTRG"
    Dim ptblNew As TableDef
    '
    ' See if the table exists
    If bTable_Exists(guCurrent.DB, MSG_NAME) Then Exit Sub
    '
    ' Create the table
    Set ptblNew = New TableDef
    With ptblNew
        .Name = MSG_NAME
        .Fields.Append .CreateField("ReportTime", dbDate)
        .Fields.Append .CreateField("ReportType", dbText, 50)
        .Fields.Append .CreateField("Origin", dbText, 50)
        .Fields.Append .CreateField("Origin_ID", dbLong)
'        .Fields.Append .CreateField("Target_ID", dbLong)
'        .Fields.Append .CreateField("Latitude", dbDouble)
'        .Fields.Append .CreateField("Longitude", dbDouble)
'        .Fields.Append .CreateField("Altitude", dbDouble)
'        .Fields.Append .CreateField("Heading", dbDouble)
'        .Fields.Append .CreateField("Speed", dbDouble)
'        .Fields.Append .CreateField("Parent", dbText, 50)
'        .Fields.Append .CreateField("Parent_ID", dbLong)
'        .Fields.Append .CreateField("Allegiance", dbText, 50)
'        .Fields.Append .CreateField("IFF", dbLong)
'        .Fields.Append .CreateField("Emitter", dbText, 50)
        .Fields.Append .CreateField("SigIDIndex", dbLong)
'        .Fields.Append .CreateField("Signal", dbText, 50)
        .Fields.Append .CreateField("Signal_ID", dbLong)
'        .Fields.Append .CreateField("Frequency", dbDouble)
'        .Fields.Append .CreateField("PRI", dbDouble)
        .Fields.Append .CreateField("ListType", dbLong)
        .Fields.Append .CreateField("NumPackets", dbLong)
        .Fields.Append .CreateField("PacketNumber", dbLong)
        .Fields.Append .CreateField("NumSigID", dbLong)
'        .Fields.Append .CreateField("Range", dbDouble)
'        .Fields.Append .CreateField("Bearing", dbDouble)
'        .Fields.Append .CreateField("Elevation", dbDouble)
'        .Fields.Append .CreateField("XX", dbDouble)
'        .Fields.Append .CreateField("XY", dbDouble)
'        .Fields.Append .CreateField("YY", dbDouble)
'        .Fields.Append .CreateField("Other_Data", dbText)
        
        .Fields("ReportType").AllowZeroLength = True
        .Fields("Origin").AllowZeroLength = True
    End With
    guCurrent.DB.TableDefs.Append ptblNew
End Sub
'
Public Sub Process_MTCNTTRG(uSig As DAS_MASTER_RECORD)

    '
    '
    Create_MTCNTTRG
    Set mrsData = guCurrent.DB.OpenRecordset(MSG_NAME)
    '
    '
    With mrsData
        .AddNew
        .Fields("ReportTime") = DateAdd("s", uSig.dReportTime, guCurrent.uArchive.dtArchiveDate)
        .Fields("ReportType") = uSig.sReport_Type
        .Fields("Origin") = uSig.sOrigin
        .Fields("Origin_ID") = uSig.lOrigin_ID
'        .Fields("Target_ID") = uSig.lTarget_ID
'        .Fields("Latitude") = uSig.dLatitude
'        .Fields("Longitude") = uSig.dLongitude
'        .Fields("Altitude") = uSig.dAltitude
'        .Fields("Heading") = uSig.dHeading
'        .Fields("Speed") = uSig.dSpeed
'        .Fields("Parent") = uSig.sParent
'        .Fields("Parent_ID") = uSig.lParent_ID
'        .Fields("Allegiance") = uSig.sAllegiance
'        .Fields("IFF") = uSig.lIFF
'        .Fields("Emitter") = uSig.sEmitter
        .Fields("SigIDIndex") = uSig.lEmitter_ID
'        .Fields("Signal") = uSig.sSignal
        .Fields("Signal_ID") = uSig.lSignal_ID
'        .Fields("Frequency") = uSig.dFrequency
'        .Fields("PRI") = uSig.dPRI
        .Fields("ListType") = uSig.lStatus
        .Fields("NumPackets") = uSig.lTag
        .Fields("PacketNumber") = uSig.lFlag
        .Fields("NumSigID") = uSig.lCommon_ID
'        .Fields("Range") = uSig.dRange
'        .Fields("Bearing") = uSig.dBearing
'        .Fields("Elevation") = uSig.dElevation
'        .Fields("XX") = uSig.dXX
'        .Fields("XY") = uSig.dXY
'        .Fields("YY") = uSig.dYY
'        .Fields("Other_Data") = uSig.sSupplemental
        .Update
    End With
End Sub
