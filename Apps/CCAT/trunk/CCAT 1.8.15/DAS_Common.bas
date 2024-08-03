Attribute VB_Name = "DAS_Common"
' COPYRIGHT (C) 1999-2001, Mercury Solutions, Inc.
' MODULE:   DAS_Common
' AUTHOR:   Russ Brown
' PURPOSE:  Define DAS structures
' REVISION:
'   v1.6.1  TAE Added verbose logging calls
Option Explicit

Public Type DAS_MASTER_RECORD
   dReportTime As Double
   sMsg_Type As String
   sReport_Type As String
   sOrigin As String
   lOrigin_ID As Long
   lTarget_ID As Long
   dLatitude As Double
   dLongitude As Double
   dAltitude As Double
   dHeading As Double
   dSpeed As Double
   sParent As String
   lParent_ID As Long
   sAllegiance As String
   lIFF As Long
   sEmitter As String
   lEmitter_ID As Long
   sSignal As String
   lSignal_ID As Long
   dFrequency As Double
   dPRI As Double
   lStatus As Long
   lTag As Long
   lFlag As Long
   lCommon_ID As Long
   dRange As Double
   dBearing As Double
   dElevation As Double
   dXX As Double
   dXY As Double
   dYY As Double
   sSupplemental As String
End Type

Public Type DAS_SIG_RECORD
   fdTsecs As Double
   sMsgType As String
   sReportType As String
   sOrigin As String
   lOriginID As Long
   sAllegiance As String
   lAllegianceID As Long
   sEmitter As String
   lEmitterID As Long
   sSignal As String
   lSignalID As Long
   fdFrequency As Double
   fdPRI As Double
   lStatus As Long
   lTag As Long
   lFlag As Long
   lCommonID As Long
   sExtraFields As String
End Type

Public Type DAS_MTF_RECORD
   fdTsecs As Double
   sMsgType As String
   sReportType As String
   sOrigin As String
   lOriginID As String
   lTargetID As Long
   fdLat As Double
   fdLng As Double
   fdAltitude As Double
   fdHeading As Double
   fdSpeed As Double
   sAllegiance As String
   lAllegianceID As Long
   sEmitter As String
   lEmitterID As Long
   sSignal As String
   lSignalID As Long
   fdFrequency As Double
   fdPRI As Double
   lStatus As Long
   lTag As Long
   lFlag As Long
   lCommonID As Long
   sExtraFields As String
End Type

Public Type DAS_LOB_RECORD
   fdTsecs As Double
   sMsgType As String
   sReportType As String
   sOrigin As String
   lOriginID As String
   lTargetID As Long
   fdLat As Double
   fdLng As Double
   fdAltitude As Double
   fdHeading As Double
   fdSpeed As Double
   sAllegiance As String
   lAllegianceID As Long
   sEmitter As String
   lEmitterID As Long
   sSignal As String
   lSignalID As Long
   fdFrequency As Double
   fdPRI As Double
   lStatus As Long
   lTag As Long
   lFlag As Long
   lCommonID As Long
   fdRangeToTarget As Double
   fdBearingAngle As Double
   fdElevationAngle As Double
   sExtraFields As String
End Type

Public Type DAS_STF_RECORD
   fdTsecs As Double
   sMsgType As String
   sReportType As String
   sOrigin As String
   lOriginID As String
   lTargetID As Long
   fdLat As Double
   fdLng As Double
   fdAltitude As Double
   sParent As String
   lParentID As Long
   sAllegiance As String
   lAllegianceID As Long
   sEmitter As String
   lEmitterID As Long
   sSignal As String
   lSignalID As Long
   fdFrequency As Double
   fdPRI As Double
   lStatus As Long
   lTag As Long
   lFlag As Long
   lCommonID As Long
   sExtraFields As String
End Type

Public Function Parse_MTF_Record(sRecord As String) As DAS_MTF_RECORD
    
   Dim sRecTmp As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Parse_MTF_Record (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sRecord
    End If
    '-v1.6.1
    '
'
'  Copy record to temp variable
'
   sRecTmp = sRecord
'
'  Parse fields into structure
'
   With Parse_MTF_Record
      .fdTsecs = CDbl(Val(Extract_Field(sRecTmp)))
      .sMsgType = Extract_Field(sRecTmp)
      .sReportType = Extract_Field(sRecTmp)
      .sOrigin = Extract_Field(sRecTmp)
      .lOriginID = CLng(Val(Extract_Field(sRecTmp)))
      .lTargetID = CLng(Val(Extract_Field(sRecTmp)))
      .fdLat = CDbl(Val(Extract_Field(sRecTmp)))
      .fdLng = CDbl(Val(Extract_Field(sRecTmp)))
      .fdAltitude = CDbl(Val(Extract_Field(sRecTmp)))
      .fdHeading = CDbl(Val(Extract_Field(sRecTmp)))
      .fdSpeed = CDbl(Val(Extract_Field(sRecTmp)))
      .sAllegiance = Extract_Field(sRecTmp)
      .lAllegianceID = CLng(Val(Extract_Field(sRecTmp)))
      .sEmitter = Extract_Field(sRecTmp)
      .lEmitterID = CLng(Val(Extract_Field(sRecTmp)))
      .sSignal = Extract_Field(sRecTmp)
      .lSignalID = CLng(Val(Extract_Field(sRecTmp)))
      .fdFrequency = CDbl(Val(Extract_Field(sRecTmp)))
      .fdPRI = CDbl(Val(Extract_Field(sRecTmp)))
      .lStatus = CLng(Val(Extract_Field(sRecTmp)))
      .lTag = CLng(Val(Extract_Field(sRecTmp)))
      .lFlag = CLng(Val(Extract_Field(sRecTmp)))
      .lCommonID = CLng(Val(Extract_Field(sRecTmp)))
      .sExtraFields = Extract_Field(Trim(sRecTmp))
  End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Parse_MTF_Record (End)"
    '-v1.6.1
    '

End Function

Public Function Parse_LOB_Record(sRecord As String) As DAS_LOB_RECORD
    
   Dim sRecTmp As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Parse_LOB_Record (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sRecord
    End If
    '-v1.6.1
    '
'
'  Copy record to temp variable
'
   sRecTmp = sRecord
'
'  Parse fields into structure
'
   With Parse_LOB_Record
      .fdTsecs = CDbl(Val(Extract_Field(sRecTmp)))
      .sMsgType = Extract_Field(sRecTmp)
      .sReportType = Extract_Field(sRecTmp)
      .sOrigin = Extract_Field(sRecTmp)
      .lOriginID = CLng(Val(Extract_Field(sRecTmp)))
      .lTargetID = CLng(Val(Extract_Field(sRecTmp)))
      .fdLat = CDbl(Val(Extract_Field(sRecTmp)))
      .fdLng = CDbl(Val(Extract_Field(sRecTmp)))
      .fdAltitude = CDbl(Val(Extract_Field(sRecTmp)))
      .fdHeading = CDbl(Val(Extract_Field(sRecTmp)))
      .fdSpeed = CDbl(Val(Extract_Field(sRecTmp)))
      .sAllegiance = Extract_Field(sRecTmp)
      .lAllegianceID = CLng(Val(Extract_Field(sRecTmp)))
      .sEmitter = Extract_Field(sRecTmp)
      .lEmitterID = CLng(Val(Extract_Field(sRecTmp)))
      .sSignal = Extract_Field(sRecTmp)
      .lSignalID = CLng(Val(Extract_Field(sRecTmp)))
      .fdFrequency = CDbl(Val(Extract_Field(sRecTmp)))
      .fdPRI = CDbl(Val(Extract_Field(sRecTmp)))
      .lStatus = CLng(Val(Extract_Field(sRecTmp)))
      .lTag = CLng(Val(Extract_Field(sRecTmp)))
      .lFlag = CLng(Val(Extract_Field(sRecTmp)))
      .lCommonID = CLng(Val(Extract_Field(sRecTmp)))
      .fdRangeToTarget = CDbl(Val(Extract_Field(sRecTmp)))
      .fdBearingAngle = CDbl(Val(Extract_Field(sRecTmp)))
      .fdElevationAngle = CDbl(Val(Extract_Field(sRecTmp)))
      .sExtraFields = Extract_Field(Trim(sRecTmp))
  End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Parse_LOB_Record (End)"
    '-v1.6.1
    '

End Function


Public Function Parse_SIG_Record(sRecord As String) As DAS_SIG_RECORD
    
   Dim sRecTmp As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Parse_SIG_Record (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sRecord
    End If
    '-v1.6.1
    '
'
'  Copy record to temp variable
'
   sRecTmp = sRecord
'
'  Parse fields into structure
'
   With Parse_SIG_Record
      .fdTsecs = CDbl(Val(Extract_Field(sRecTmp)))
      .sMsgType = Extract_Field(sRecTmp)
      .sReportType = Extract_Field(sRecTmp)
      .sOrigin = Extract_Field(sRecTmp)
      .lOriginID = CLng(Val(Extract_Field(sRecTmp)))
      .sAllegiance = Extract_Field(sRecTmp)
      .lAllegianceID = CLng(Val(Extract_Field(sRecTmp)))
      .sEmitter = Extract_Field(sRecTmp)
      .lEmitterID = CLng(Val(Extract_Field(sRecTmp)))
      .sSignal = Extract_Field(sRecTmp)
      .lSignalID = CLng(Val(Extract_Field(sRecTmp)))
      .fdFrequency = CDbl(Val(Extract_Field(sRecTmp)))
      .fdPRI = CDbl(Val(Extract_Field(sRecTmp)))
      .lStatus = CLng(Val(Extract_Field(sRecTmp)))
      .lTag = CLng(Val(Extract_Field(sRecTmp)))
      .lFlag = CLng(Val(Extract_Field(sRecTmp)))
      .lCommonID = CLng(Val(Extract_Field(sRecTmp)))
      .sExtraFields = Extract_Field(Trim(sRecTmp))
  End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Parse_SIG_Record (End)"
    '-v1.6.1
    '

End Function


Public Function Parse_STF_Record(sRecord As String) As DAS_STF_RECORD
    
   Dim sRecTmp As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Parse_STF_Record (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sRecord
    End If
    '-v1.6.1
    '
'
'  Copy record to temp variable
'
   sRecTmp = sRecord
'
'  Parse fields into structure
'
   With Parse_STF_Record
      .fdTsecs = CDbl(Val(Extract_Field(sRecTmp)))
      .sMsgType = Extract_Field(sRecTmp)
      .sReportType = Extract_Field(sRecTmp)
      .sOrigin = Extract_Field(sRecTmp)
      .lOriginID = CLng(Val(Extract_Field(sRecTmp)))
      .lTargetID = CLng(Val(Extract_Field(sRecTmp)))
      .fdLat = CDbl(Val(Extract_Field(sRecTmp)))
      .fdLng = CDbl(Val(Extract_Field(sRecTmp)))
      .fdAltitude = CDbl(Val(Extract_Field(sRecTmp)))
      .sParent = Extract_Field(sRecTmp)
      .lParentID = CLng(Val(Extract_Field(sRecTmp)))
      .sAllegiance = Extract_Field(sRecTmp)
      .lAllegianceID = CLng(Val(Extract_Field(sRecTmp)))
      .sEmitter = Extract_Field(sRecTmp)
      .lEmitterID = CLng(Val(Extract_Field(sRecTmp)))
      .sSignal = Extract_Field(sRecTmp)
      .lSignalID = CLng(Val(Extract_Field(sRecTmp)))
      .fdFrequency = CDbl(Val(Extract_Field(sRecTmp)))
      .fdPRI = CDbl(Val(Extract_Field(sRecTmp)))
      .lStatus = CLng(Val(Extract_Field(sRecTmp)))
      .lTag = CLng(Val(Extract_Field(sRecTmp)))
      .lFlag = CLng(Val(Extract_Field(sRecTmp)))
      .lCommonID = CLng(Val(Extract_Field(sRecTmp)))
      .sExtraFields = Extract_Field(Trim(sRecTmp))
  End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Parse_STF_Record (End)"
    '-v1.6.1
    '

End Function


Public Function Extract_Field(ByRef sRecord As String) As String
    
   Dim iDelimiter As Integer
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Extract_Field (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sRecord
    End If
    '-v1.6.1
    '
'
'  Find the next delimiter
'
   iDelimiter = InStr(1, sRecord, ",")
'
'  Copy the field contents
'
   If iDelimiter = 0 Then
      Extract_Field = Trim(sRecord)
   Else
      Extract_Field = Trim(Mid(sRecord, 1, iDelimiter - 1))
   End If
'
'  Remove the field from the record
'
   If iDelimiter > 0 Then
      sRecord = Mid(sRecord, iDelimiter + 1)
   Else
      sRecord = ""
   End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : DAS_Common.Extract_Field (End)"
    '-v1.6.1
    '
    
End Function

