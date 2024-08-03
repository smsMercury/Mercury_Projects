Attribute VB_Name = "modMTANARSLT"
Private MSG_NAME As String
Private Const MSG_ID = 42

Private mrsData As Recordset
'
'
Public Sub Create_MTANARSLT()
    MSG_NAME = guCurrent.sArchive & "_MTANARSLT"
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
        .Fields.Append .CreateField("SignalPresentDf", dbLong)
        .Fields.Append .CreateField("Latitude", dbDouble)
        .Fields.Append .CreateField("Longitude", dbDouble)
'        .Fields.Append .CreateField("Altitude", dbDouble)
'        .Fields.Append .CreateField("Heading", dbDouble)
'        .Fields.Append .CreateField("Speed", dbDouble)
'        .Fields.Append .CreateField("Parent", dbText, 50)
        .Fields.Append .CreateField("PassFreq", dbLong)
'        .Fields.Append .CreateField("Allegiance", dbText, 50)
'        .Fields.Append .CreateField("IFF", dbLong)
        .Fields.Append .CreateField("Emitter", dbText, 80)
        .Fields.Append .CreateField("Emitter_ID", dbLong)
        .Fields.Append .CreateField("Signal", dbText, 50)
        .Fields.Append .CreateField("Signal_ID", dbLong)
        .Fields.Append .CreateField("Frequency", dbDouble)
'        .Fields.Append .CreateField("PRI", dbDouble)
        .Fields.Append .CreateField("SignalPresentAna", dbLong)
        .Fields.Append .CreateField("Variant", dbLong)
        .Fields.Append .CreateField("RequestorID", dbLong)
'        .Fields.Append .CreateField("Common_ID", dbLong)
'        .Fields.Append .CreateField("Range", dbDouble)
        .Fields.Append .CreateField("Bearing", dbDouble)
'        .Fields.Append .CreateField("Elevation", dbDouble)
'        .Fields.Append .CreateField("XX", dbDouble)
'        .Fields.Append .CreateField("XY", dbDouble)
'        .Fields.Append .CreateField("YY", dbDouble)
        .Fields.Append .CreateField("SigUsage", dbText)
        
        ' Set the field attribute to allow null strings
        .Fields("ReportType").AllowZeroLength = True
        .Fields("Origin").AllowZeroLength = True
        .Fields("Emitter").AllowZeroLength = True
        .Fields("Signal").AllowZeroLength = True
        .Fields("SigUsage").AllowZeroLength = True
    End With
    guCurrent.DB.TableDefs.Append ptblNew
End Sub
'
Public Sub Process_MTANARSLT(uSig As DAS_MASTER_RECORD)

    '
    '
    Create_MTANARSLT
    Set mrsData = guCurrent.DB.OpenRecordset(MSG_NAME)
    '
    '
    With mrsData
        .AddNew
        .Fields("ReportTime") = DateAdd("s", uSig.dReportTime, guCurrent.uArchive.dtArchiveDate)
        .Fields("ReportType") = uSig.sReport_Type
        .Fields("Origin") = uSig.sOrigin
        .Fields("Origin_ID") = uSig.lOrigin_ID
        .Fields("SignalPresentDf") = uSig.lTarget_ID
        .Fields("Latitude") = uSig.dLatitude
        .Fields("Longitude") = uSig.dLongitude
'        .Fields("Altitude") = uSig.dAltitude
'        .Fields("Heading") = uSig.dHeading
'        .Fields("Speed") = uSig.dSpeed
'        .Fields("Parent") = uSig.sParent
        .Fields("PassFreq") = uSig.lParent_ID
'        .Fields("Allegiance") = uSig.sAllegiance
'        .Fields("IFF") = uSig.lIFF
        .Fields("Emitter") = uSig.sEmitter
        .Fields("Emitter_ID") = uSig.lEmitter_ID
        .Fields("Signal") = uSig.sSignal
        .Fields("Signal_ID") = uSig.lSignal_ID
        .Fields("Frequency") = uSig.dFrequency
'        .Fields("PRI") = uSig.dPRI
        .Fields("SignalPresentAna") = uSig.lStatus
        .Fields("Variant") = uSig.lTag
        .Fields("RequestorID") = uSig.lFlag
'        .Fields("Common_ID") = uSig.lCommon_ID
'        .Fields("Range") = uSig.dRange
        .Fields("Bearing") = uSig.dBearing
'        .Fields("Elevation") = uSig.dElevation
'        .Fields("XX") = uSig.dXX
'        .Fields("XY") = uSig.dXY
'        .Fields("YY") = uSig.dYY
        .Fields("SigUsage") = uSig.sSupplemental
        .Update
    End With
End Sub

