Attribute VB_Name = "modMTTXRCONF"
Private MSG_NAME As String
Private Const MSG_ID = 100

Private mrsData As Recordset
'
'
Public Sub Create_MTTXRCONF()
    MSG_NAME = guCurrent.sArchive & "_MTTXRCONF"
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
        .Fields.Append .CreateField("Tx_Src", dbLong)
'        .Fields.Append .CreateField("Signal", dbText, 50)
        .Fields.Append .CreateField("Tx_Chan", dbLong)
        .Fields.Append .CreateField("Frequency", dbDouble)
        .Fields.Append .CreateField("FreqEnd", dbDouble)
        .Fields.Append .CreateField("MsnChg", dbLong)
        .Fields.Append .CreateField("Tx_Pa_ID", dbLong)
        .Fields.Append .CreateField("Tx_Pa_Pwr_Setting", dbLong)
        .Fields.Append .CreateField("Tx_Ant_Grp", dbLong)
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
Public Sub Process_MTTXRCONF(uSig As DAS_MASTER_RECORD)

    '
    '
    Create_MTTXRCONF
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
        .Fields("Tx_Src") = uSig.lEmitter_ID
'        .Fields("Signal") = uSig.sSignal
        .Fields("Tx_Chan") = uSig.lSignal_ID
        .Fields("Frequency") = uSig.dFrequency
        .Fields("FreqEnd") = uSig.dPRI
        .Fields("MsnChg") = uSig.lStatus
        .Fields("Tx_Pa_ID") = uSig.lTag
        .Fields("Tx_Pa_Pwr_Setting") = uSig.lFlag
        .Fields("Tx_Ant_Grp") = uSig.lCommon_ID
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

