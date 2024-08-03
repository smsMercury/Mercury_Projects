Attribute VB_Name = "CCMessageStruct"
' COPYRIGHT (C) 1999-2001, Mercury Solutions, Inc.
' MODULE:   CCMessageStruct
' AUTHOR:   Brad Brown
' PURPOSE:  Establish CCOS message structures
' REVISION:
'   1.6     BDB Added structures for CCOS 2.3 variants of messages
'              Added structure for MTSSERROR message
'           TAE Updated variables to reflect current state
'   1.8     SPV Added updates for Block 35
'
' Declare API calls
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal length As Long)
'Declare Function agSwapBytes% Lib "apigid32.dll" (ByVal src%)
'Declare Function agSwapWords& Lib "apigid32.dll" (ByVal src&)
'Declare Function agSwapBytes% Alias SwapBytes (ByVal src%)
'Declare Function agSwapWords& Alias SwapWords (ByVal src&)

'
' Set up Message identifiers
Public Enum MsgIds
    '
    '+v1.71
    'NUMFILTERMTIDS = 32
    NUMFILTERMTIDS = 39
    '-v1.71
    MTRANDOM1 = 1       ' these are used only for random message creation
    MTRANDOM2 = 2       '
    MTRANDOM3 = 3       '
    MTDEFPMAID = 51
    MTSIGALARMID = 2304
    MTANAALARMID = 2570
    MTHBDYNRSPID = 94
    MTSDALARMID = 2306
    MTHBACTREPID = 91
    MTDFALARMID = 3328
    MTRFSTATUSID = 2308
    MTDFSDALARMID = 3329
    MTLOBUPDID = 49
    MTANARSLTID = 42
    MTSIGUPDID = 19
    MTHBLOBUPDID = 566
    MTULDDATAID = 50
    MTLOBRSLTID = 58
    MTLOBSETRSLTID = 60
    MTFIXRSLTID = 48
    MTHBSIGUPDID = 70
    MTHBGSREPID = 95
    MTHBSEMISTATID = 96
    MTHBSELASKID = 449
    MTJAMSTATID = 35
    MTRUNMODEID = 112
    MTHBSELJAMID = 450
    MTHBXMTRSTATID = 97
    MTSETACQSMODEID = 29
    MTTIMESYNID = 308
    MTARCFILLERID = 359
    MTNAVREPID = 53
    MTLHCORRELATEID = 345
    MTLHECHMODID = 76
    MTLHTRACKUPDID = 484
    MTLHTRACKREPID = 98
    MTLHBEARINGONLYID = 223     ' v1.12
    MTSSERRORID = 113           ' v1.6
    MTPLANAREAID = 423
    MTANAREQID = 45
    MTDFFLGSID = 52
    MTRRSLTID = 2569
    MTHBASSREPID = 92   ' v1.71
    MTSYSCONFIGID = 116   ' v1.71
    MTHBGSUPSTATID = 123   ' v1.71
    MTHBGSDELID = 124   ' v1.71
    MTENVSTATID = 26 ' v1.71
    MTDPSEXTRSPID = 127
    MTCNTTRGID = 22 ' TE 1.8.10
    MTTXRFCONFID = 100 ' TE 1.8.10
    MTTGTLISTUPDID = 434 ' TE 1.8.10
End Enum
'
' High Band ID
Public Enum HBIDType
    HBAID = 0
    HBSigtype = 2
End Enum
'
' Global constants
Global Const gbytCLASSIFIED As Byte = 8
Global Const gbytNOT_DEFINED As Byte = 255
'
' Archive header
Public Type Arc_Hdr                         '   4 bytes Fixed
    lTimestamp As Long                      '   4 bytes   0 -   3
End Type 'Arc_Hdr
'
' Message Header
Public Type Msg_Hdr                         '   8 bytes Fixed
    iMsgLength As Integer                   '   2 bytes   0 -   1
    iMsgId As Integer                       '   2 bytes   2 -   3
    iMsgTo As Integer                       '   2 bytes   4 -   5
    iMsgFrom As Integer                     '   2 bytes   6 -   7
End Type 'Msg_Hdr
'
' Latitude/Longitude
Public Type LatLon                          '   8 bytes Fixed
    lLat As Long                            '   4 bytes   0 -   3
    lLon As Long                            '   4 bytes   4 -   7
End Type 'LatLon
'
' Time of Day
Public Type TodTime                         '   8 bytes Fixed
    iDayOfYear As Integer                   '   2 bytes   0 -   1
    bytSpare2 As Byte                       '   1 byte    2 -   2
    bytHighUsecs As Byte                    '   1 byte    3 -   3
    lUsecs As Long                          '   4 bytes   4 -   7
End Type 'TodTime
'
' Primary Mission Area
Public Type PMA                             '  52 bytes Fixed
    bytNumberOfVertices As Byte             '   1 byte    0 -   0
    cPad(1 To 3) As Byte                    '   3 bytes   1 -   3
    uPmaVertices(1 To 6) As LatLon          '  48 bytes   4 -  51
End Type 'PMA
'
' Signal packet
Public Type SigPacket                       '   4 bytes Fixed
    iSigID As Integer                       '   2 bytes   0 -   1
    bytJamStatus As Byte                    '   1 byte    2 -   2
    bytPad As Byte                          '   1 byte    3 -   3
End Type 'SigPacket
'
' Alarm
Public Type AlarmData                       '  16 bytes Fixed
    lCtrFreq As Long                        '   4 bytes   0 -   3
    lBandwidth As Long                      '   4 bytes   4 -   7
    lPeakFrequency As Long                  '   4 bytes   8 -  11
    iAmplitude As Integer                   '   2 bytes  12 -  13
    bytPad(1 To 2) As Byte                  '   2 bytes  14 -  15
End Type ' AlarmData
'
' Alarm packet
Public Type AlarmPacket                     '  20 bytes Fixed
    uData As AlarmData                      '  16 bytes   0 -  15
    bytRfStatus As Byte                     '   1 byte   16 -  16
    bytPad(1 To 3) As Byte                  '   3 bytes  17 -  19
End Type ' AlarmPacket
'
' Signal alarm
Public Type SignalAlarmData                 '  32 Bytes Fixed
    uTodTime As TodTime                     '   8 bytes   0 -   7
    iSignalStatus As Integer                '   2 bytes   8 -   9
    bytPad(1 To 2) As Byte                  '   2 bytes  10 -  11
    uData As AlarmPacket                    '  20 bytes  12 -  31
End Type
'
' LOB packet
Public Type LobPacket                       '  16 bytes Fixed
    uAcLoc As LatLon                        '   8 bytes   0 -   7
    iTrueBearing As Integer                 '   2 bytes   8 -   9
    bytInOutPma As Byte                     '   1 byte   10 -  10
    bytQualFactor As Byte                   '   1 byte   11 -  11
    bytOperRequested As Byte                '   1 byte   12 -  12
    bytPad(1 To 3) As Byte                  '   3 bytes  13 -  15
End Type 'LobPacket
'
' Channel status
Public Type ChanStatus                      '   8 bytes Fixed
    bytChan As Byte                         '   1 byte    0 -   0
    bytPad(1 To 3) As Byte                  '   3 bytes   1 -   3
    lFrequency As Long                      '   4 bytes   4 -   7
End Type 'ChanStatus
'
' Signal mode
Public Type SigMode                         '   Variable
    bytSigType As Byte                      '   1 byte    0 -   0
    bytMode As Byte                         '   1 byte    1 -   1
    iNumChans As Integer                    '   2 bytes   2 -   3
    uChanStatus(0) As ChanStatus            '  x8 bytes   4 - ...
End Type
'
' P34 status
Public Type P34Stat                         '   8 bytes Fixed
    iBand3Xmtr As Integer                   '   2 bytes   0 -   1
    iBand4aXmtr As Integer                  '   2 bytes   2 -   3
    iBand4bXmtr As Integer                  '   2 bytes   4 -   5
    bytPad(1 To 2) As Byte                  '   2 bytes   6 -   7
End Type 'P34Stat
'
' Activity report
Public Type ActRec                          '  20 bytes Fixed
    bytSignal As Byte                       '   1 byte    0 -   0
    bytFunction As Byte                     '   1 byte    1 -   1
    bytOption As Byte                       '   1 byte    2 -   2
    bytChannel As Byte                      '   1 byte    3 -   3
    bytPmaNonPma As Byte                    '   1 byte    4 -   4
    bytStatus As Byte                       '   1 byte    5 -   5
    bytNumTgts As Byte                      '   1 byte    6 -   6
    bytPad0 As Byte                         '   1 byte    7 -   7
    lFreq As Long                           '   4 bytes   8 -  11
    iActCnt As Integer                      '   2 bytes  12 -  13
    iDuration As Integer                    '   2 bytes  14 -  15
    bytBeam As Byte                         '   1 byte   16 -  16
    bytPad1(1 To 3) As Byte                 '   3 bytes  17 -  19
End Type 'ActRec
'
' Activity report
Public Type ActRec3_0                       '  20 bytes Fixed
    bytSignal As Byte                       '   1 byte    0 -   0
    bytFunction As Byte                     '   1 byte    1 -   1
    bytOption As Byte                       '   1 byte    2 -   2
    bytChannel As Byte                      '   1 byte    3 -   3
    bytPmaNonPma As Byte                    '   1 byte    4 -   4
    bytStatus As Byte                       '   1 byte    5 -   5
    iNumTgts As Integer                     '   2 bytes   6 -   7
    lFreq As Long                           '   4 bytes   8 -  11
    iActCnt As Integer                      '   2 bytes  12 -  13
    iDuration As Integer                    '   2 bytes  14 -  15
    bytBeam As Byte                         '   1 byte   16 -  16
    bytPad1(1 To 3) As Byte                 '   3 bytes  17 -  19
End Type 'ActRec3_0
'
'
Public Type WPSignalData                    '  12 bytes fixed
    iItem1 As Integer                       '   2 bytes   0 -   1
    iItem2 As Integer                       '   2 bytes   2 -   3
    iItem3 As Integer                       '   2 bytes   4 -   5
    bytItem4 As Byte                        '   1 byte    6 -   6
    bytItem5 As Byte                        '   1 byte    7 -   7
    lItem6 As Byte                          '   4 bytes   8 -  11
End Type 'WPSignalData
'
'
Public Type WPResultsData                   '   8 bytes fixed
    iItem1 As Integer                       '   2 bytes   0 -   1
    bytItem2 As Byte                        '   1 byte    2 -   2
    bytItem3 As Byte                        '   1 byte    3 -   3
    lItem4 As Long                          '   4 bytes   4 -   7
End Type 'WPResultsData
'
' High Band dynamic info
Public Type HbdynRec                        '   8 bytes Fixed
    bytOnList As Byte                       '   1 byte    0 -   0
    bytOption As Byte                       '   1 byte    1 -   1
    bytSource As Byte                       '   1 byte    2 -   2
    bytPmaNonPma As Byte                    '   1 byte    3 -   3
    lFrequency As Long                      '   4 bytes   4 -   7
End Type 'HbdynRec
'
' High band dynamic response
Public Type HbdynRsp                        '  12 bytes
    bytResponse As Byte                     '   1 byte    0 -   0
    bytPad(1 To 3) As Byte                  '   3 bytes   1 -   3
    uHbDynRec As HbdynRec                   '   8 bytes   4 -  11
End Type 'HbdynRsp
'
'
Public Type HbdynRspRec                     ' Variable length
    bytSignal As Byte                       '   1 byte    0 -   0
    bytFunction As Byte                     '   1 byte    1 -   1
    bytNumOfChannels As Byte                '   1 byte    2 -   2
    bytPad As Byte                          '   1 byte    3 -   3
    uHbDynRsp(0) As HbdynRsp                ' x12 bytes   4 -  ...
End Type 'HbdynRspRec

Public Type NavData                         '  56 bytes Fixed
    uTod As TodTime                         '   8 bytes   0 -   7
    uPosition As LatLon                     '   8 bytes   8 -  15
    iHeading As Integer                     '   2 bytes  16 -  17
    iPitch As Integer                       '   2 bytes  18 -  19
    iRoll As Integer                        '   2 bytes  20 -  21
    iTrueGroundSpeed As Integer             '   2 bytes  22 -  23
    iGrndTrackAngle As Integer              '   2 bytes  24 -  25
    bytPad0(1 To 2) As Byte                 '   2 bytes  26 -  27
    lPressureAlt As Long                    '   4 bytes  28 -  31
    bytWow As Byte                          '   1 byte   32 -  32
    bytPad1(1 To 3) As Byte                 '   3 bytes  33 -  35
    lGpsUtcTime As Long                     '   4 bytes  36 -  39
    uGpsLoc As LatLon                       '   8 bytes  40 -  47
    lGpsAltitude As Long                    '   4 bytes  48 -  51
    iGpsFom As Integer                      '   2 bytes  52 -  53
    bytPad2(1 To 2) As Byte                 '   2 bytes  54 -  55
End Type 'NavData

Public Type Sigclassvar                     '   4 bytes Fixed
    bytClass As Byte                        '   1 byte    0 -   0
    bytVariant As Byte                      '   1 byte    1 -   1
    bytPad0(1 To 2) As Byte                 '   2 bytes   2 -   3
End Type ' Sigclassvar

Public Type ChannelData                     '   8 bytes
    bytChannelAnalyzed As Byte              '   1 byte    0 -   0
    bytChannelActive As Byte                '   1 byte    1 -   1
    bytSigUsage As Byte                     '   1 byte    2 -   2
    bytPad0 As Byte                         '   1 byte    3 -   3
    uSignalType As Sigclassvar              '   4 bytes   4 -   7
End Type 'ChannelData
   
Public Type LobInfo                         '  48 bytes Fixed
    uLobData As LobPacket                   '  16 bytes   0 -  15
    iAPrioriLoc As Integer                  '   2 bytes  16 -  17
    iSigID(1 To 3) As Integer               '   6 bytes  18 -  23 (3 * 2)
    iNumberOfFrequencys As Integer          '   2 bytes  24 -  25
    bytPad(1 To 2) As Byte                  '   2 bytes  26 -  27
    lFrequency(1 To 5) As Long              '  20 bytes  28 -  47 (5 * 4)
End Type ' LobInfo

Public Type DeferredLobData                 '  52 bytes Fixed
    bytRetryCount As Byte                   '   1 byte    0 -   0
    bytPad(1 To 3) As Byte                  '   3 bytes   1 -   3
    uLobData As LobInfo                     '  48 bytes   4 -  51
End Type ' DeferredLobData
'
'+v1.6BB
Public Type DeferredLobData2_3              '  56 bytes Fixed
    bytRetryCount As Byte                   '   1 byte    0 -   0
    bytPad(1 To 3) As Byte                  '   3 bytes   1 -   3
    lAlarmFrequency As Long                 '   4 bytes   4 -   7 This is the culprit
    uLobData As LobInfo                     '  48 bytes   8 -  55
End Type ' DeferredLobData2_3
'-v1.6
'
Public Type ShortDurationAlarmData              ' 708 bytes Fixed
    bytSignalStatus As Byte                     '   1 byte    0 -   0
    bytNumberOfLobs As Byte                     '   1 byte    1 -   1
    bytPad(1 To 2) As Byte                      '   2 bytes   2 -   3
    uAcqTime As TodTime                         '   8 bytes   4 -  11
    uAlarm As AlarmData                         '  16 bytes  12 -  27
    uDeferredLobs(1 To 20) As DeferredLobData   ' 960 bytes  28 - 987 (20 * 48)
End Type 'ShortDurationAlarmData
'
'+v1.6BB
Public Type ShortDurationAlarmData2_3           ' 788 bytes Fixed
    bytSignalStatus As Byte                     '   1 byte    0 -   0
    bytNumberOfLobs As Byte                     '   1 byte    1 -   1
    bytPad(1 To 2) As Byte                      '   2 bytes   2 -   3
    uAcqTime As TodTime                         '   8 bytes   4 -  11
    uAlarm As AlarmData                         '  16 bytes  12 -  27
    uDeferredLobs(1 To 20) As DeferredLobData2_3 '760 bytes  28 - 787 (20 * 38)
End Type 'ShortDurationAlarmData2_3
'-v1.6
'
Public Type DfDsaSupportData                    '  12 bytes Fixed
    lBandwidth As Long                          '   4 bytes   0 -   3
    iAttenuation As Integer                     '   2 bytes   4 -   5
    iSignalStrength As Integer                  '   2 bytes   6 -   7
    bytNumberOfDfSignals As Byte                '   1 byte    8 -   8
    bytPad(1 To 3) As Byte                      '   3 bytes   9 -  11
End Type ' DfDsaSupportData

Public Type DfAlarmData                         ' 648 bytes Fixed
    uAlarm As AlarmPacket                       '  20 bytes   0 -  19
    uTimeAcquired As TodTime                    '   8 bytes  20 -  27
    uDsaSupportData As DfDsaSupportData         '  12 bytes  28 -  39
    iRequestorID As Integer                     '   2 bytes  40 -  41
    bytAlarmType As Byte                        '   1 byte   42 -  42
    bytLobRequestType As Byte                   '   1 byte   43 -  43
    bytNumberOfLobs As Byte                     '   1 byte   44 -  44
    bytPad(1 To 3) As Byte                      '   3 bytes  45 -  47
    uLobs(1 To 20) As LobInfo                   ' 600 bytes  48 - 647 (20 * 30)
End Type ' DfAlarmData

Public Type DfShortDurationResultsData          ' 616 bytes Fixed
    lAcqCenterFrequency As Long                 '   4 bytes   0 -   3
    uTimeAcquired As TodTime                    '   8 bytes   4 -  11
    bytPad1(1 To 3) As Byte                     '   3 bytes  12 -  14
    bytNumberOfLobs As Byte                     '   1 byte   15 -  15
    uLobs(1 To 20) As LobInfo                   ' 600 bytes  16 - 615 (20 * 30)
End Type 'DfShortDurationResultsData

Public Type GroundSite                          '  40 bytes Fixed
    iId As Integer                              '   2 bytes   0 -   1
    bytSignal As Byte                           '   1 byte    2 -   2
    bytChannel As Byte                          '   1 byte    3 -   3
    bytMethod As Byte                           '   1 byte    4 -   4
    bytPad0(1 To 3) As Byte                     '   3 bytes   5 -   7
    uLocation As LatLon                         '   8 bytes   8 -  15
    dCovxx As Double                            '   8 bytes  16 -  23
    dCovxy As Double                            '   8 bytes  24 -  31
    dCovyy As Double                            '   8 bytes  32 -  39
End Type 'GroundSite

Public Type SemisigtypeRec                      '   8 bytes Fixed
    ' do this later
    itemp(1 To 4) As Integer                    '   8 bytes   0 -   7 (4 * 2)
End Type ' SemisigtypeRec

Public Type SemistatRec                         '  84 bytes Fixed
    ' do this later
    itemp(1 To 42) As Integer                   '  84 bytes   0 -  83 (42 * 2)
End Type ' SemistatRec

Public Type Mthbsemiinfostruct                  ' Variable length
    uSemisigtype As SemisigtypeRec              '   8 bytes   0 -   7
    iNumRecs As Integer                         '   2 bytes   8 -   9
    bytPad0(1 To 2) As Byte                     '   2 bytes  10 -  11
    uSemistatRec(0) As SemistatRec              ' x84 bytes  12 - ...
End Type 'Mthbsemiinfostruct
    
Public Type ContData                            '  48 bytes Fixed
    'do this later, size is correct
    bytTemp(1 To 5) As Byte                     '   5 bytes   0 -   4
    bytChannel As Byte                          '   1 byte    5 -   5
    bytSignal As Byte                           '   1 byte    6 -   6
    byttemp2(1 To 41) As Byte                   '  41 bytes   7 -  47
    'itemp(1 To 24) As Integer
End Type ' ContData

Public Type HbLobRec                            ' 568 bytes Fixed
    iHbsTrackId As Integer                      '   2 bytes   0 -   1
    bytTrackConf As Byte                        '   1 byte    2 -   2
    bytPad As Byte                              '   1 byte    3 -   3
    iTrackBearing(1 To 8) As Integer            '  16 bytes   4 -  19 (8 * 2)
    uContrib(1 To 10) As ContData               ' 480 bytes  20 - 499 (10 * 48)
    uOwnShip(1 To 8) As LatLon                  '  64 bytes 500 - 563 (8 * 8)
    iBearId As Integer                          '   2 bytes 564 - 565
    iMostRecent As Integer                      '   2 bytes 566 - 567
End Type ' HbLobRec

Public Type HbLobTblEntry                       ' 572 bytes Fixed
    uHblobData As HbLobRec                      ' 568 bytes   0 - 567
    iTblNum As Integer                          '   2 bytes 568 - 569
    bytPad0(1 To 2) As Byte                     '   2 bytes 570 - 571
End Type 'HbLobTblEntry

Public Type HbSigRec                            ' 112 bytes Fixed
    iSignum As Integer                          '   2 bytes   0 -   1
    iSigID As Integer                           '   2 bytes   2 -   3
    bytSigStatus As Byte                        '   1 byte    4 -   4
    bytRespOpr As Byte                          '   1 byte    5 -   5
    bytSemStat As Byte                          '   1 byte    6 -   6
    bytLocType As Byte                          '   1 byte    7 -   7
    uLocation As LatLon                         '   8 bytes   8 -  15
    iTmtn As Integer                            '   2 bytes  16 -  17
    iTellStatus As Integer                      '   2 bytes  18 -  19
    iSource As Integer                          '   2 bytes  20 -  21
    bytAlleg As Byte                            '   1 byte   22 -  22
    bytAllegSource As Byte                      '   1 byte   23 -  23
    bytExcChan As Byte                          '   1 byte   24 -  24
    bytActChan As Byte                          '   1 byte   25 -  25
    bytRemark(1 To 11) As Byte                  '  11 bytes  26 -  36
    bytCsid(1 To 19) As Byte                    '  19 bytes  37 -  55
    lTimeAcq As Long                            '   4 bytes  56 -  59
    iLocSource As Integer                       '   2 bytes  60 -  61
    bytFloat64Pad(1 To 2) As Byte               '   2 bytes  62 -  63
    dHbMsmtXx As Double                         '   8 bytes  64 -  71
    dHbMsmtXy As Double                         '   8 bytes  72 -  79
    dHbMsmtYy As Double                         '   8 bytes  80 -  87
    iEchIndex As Integer                        '   2 bytes  88 -  89
    iHbSsIndex As Integer                       '   2 bytes  90 -  91
    iChanLink As Integer                        '   2 bytes  92 -  93
    iEchLink As Integer                         '   2 bytes  94 -  95
    iNumContribEchs As Integer                  '   2 bytes  96 -  97
    iCorrId As Integer                          '   2 bytes  98 -  99
    iMpEchid As Integer                         '   2 bytes 100 - 101
    iOtherEchId As Integer                      '   2 bytes 102 - 103
    bytRstrOverride As Byte                     '   1 byte  104 - 104
    bytTpgType As Byte                          '   1 byte  105 - 105
    bytDeconOverride As Byte                    '   1 byte  106 - 106
    bytEnjChan As Byte                          '   1 byte  107 - 107
    bytDfltLocInd As Byte                       '   1 byte  108 - 108
    bytPad0(1 To 3) As Byte                     '   3 bytes 109 - 111
    iChanId As Integer                          '   2 bytes 112 - 113
    iPaId As Integer                            '   2 bytes 114 - 115
    bytPad64(1 To 4) As Byte                    '   4 bytes 116 - 119
End Type ' HbSigRec

Public Type SigData                             ' 316 bytes Fixed
    bytSemStat As Byte                          '   1 byte    0 -   0
    bytEmitterAnalyzed As Byte                  '   1 byte    1 -   1
    iSignum As Integer                          '   2 bytes   2 -   3
    iSigID As Integer                           '   2 bytes   4 -   5
    bytNumActiveChan As Byte                    '   1 byte    6 -   6
    bytPad0 As Byte                             '   1 byte    7 -   7
    lFreq As Long                               '   4 bytes   8 -  11
    lBandwidth As Long                          '   4 bytes  12 -  15
    iModType As Integer                         '   2 bytes  16 -  17
    bytRadioType As Byte                        '   1 byte   18 -  18
    bytNumberOfChannels As Byte                 '   1 byte   19 -  19
    uChannelData(1 To 12) As ChannelData        '  96 bytes  20 - 115 (12 * 8)
    iEchid As Integer                           '   2 bytes 116 - 117
    iMasterSig As Integer                       '   2 bytes 118 - 119
    bytSigjampri As Byte                        '   1 byte  120 - 120
    bytJamMeth As Byte                          '   1 byte  121 - 121
    bytJamConf As Byte                          '   1 byte  122 - 122
    bytListenConf As Byte                       '   1 byte  123 - 123
    bytActBySig As Byte                         '   1 byte  124 - 124
    bytIncBySig As Byte                         '   1 byte  125 - 125
    bytEnjBySig As Byte                         '   1 byte  126 - 126
    bytPtctOverride As Byte                     '   1 byte  127 - 127
    iAsignum As Integer                         '   2 bytes 128 - 129
    iCorrId As Integer                          '   2 bytes 130 - 131
    iEchLink As Integer                         '   2 bytes 132 - 133
    bytLobValidated As Byte                     '   1 byte  134 - 134
    bytAcqDrngJam As Byte                       '   1 byte  135 - 135
    lTimeAcq As Long                            '   4 bytes 136 - 139
    iFreqLink As Integer                        '   2 bytes 140 - 141
    iMpEchid As Integer                         '   2 bytes 142 - 143
    iOtherEchId As Integer                      '   2 bytes 144 - 145
    iOprAsg As Integer                          '   2 bytes 146 - 147
    bytSigStatus As Byte                        '   1 byte  148 - 148
    bytMpStatus As Byte                         '   1 byte  149 - 149
    bytChanged As Byte                          '   1 byte  150 - 150
    bytValStat As Byte                          '   1 byte  151 - 151
    lFixId As Long                              '   4 bytes 152 - 155
    uFixloc As LatLon                           '   8 bytes 156 - 163
    iFixBrng As Integer                         '   2 bytes 164 - 165
    iMinorAxis As Integer                       '   2 bytes 166 - 167
    iMajorAxis As Integer                       '   2 bytes 168 - 169
    bytFixType As Byte                          '   1 byte  170 - 170
    bytOprqIndex As Byte                        '   1 byte  171 - 171
    iFoiQueue As Integer                        '   2 bytes 172 - 173
    iAltFreqLink As Integer                     '   2 bytes 174 - 175
    bytLang(1 To 3) As Byte                     '   3 bytes 176 - 178
    bytNetwork(1 To 19) As Byte                 '  19 bytes 179 - 197
    bytAud(1 To 9) As Byte                      '   9 bytes 198 - 206
    bytChannel(1 To 3) As Byte                  '   3 bytes 207 - 209
    bytWstype(1 To 17) As Byte                  '  17 bytes 210 - 226
    bytStationNum(1 To 11) As Byte              '  11 bytes 227 - 237
    bytNarr(1 To 21) As Byte                    '  21 bytes 238 - 258
    bytCollArea As Byte                         '   1 byte  259 - 259
    iSource As Integer                          '   2 bytes 260 - 261
    iCallLink As Integer                        '   2 bytes 262 - 263
    bytPickedSignal As Byte                     '   1 byte  264 - 264
    bytRespOpr As Byte                          '   1 byte  265 - 265
    iTmtn As Integer                            '   2 bytes 266 - 267
    iTellStatus As Integer                      '   2 bytes 268 - 269
    bytRemoteSource As Byte                     '   1 byte  270 - 270
    bytAllegiance As Byte                       '   1 byte  271 - 271
    bytAllegSource As Byte                      '   1 byte  272 - 272
    bytDeconOverride As Byte                    '   1 byte  273 - 273
    iNumFrps As Integer                         '   2 bytes 274 - 275
    bytFrpChan(1 To 12) As Byte                 '  12 bytes 276 - 287 (12 * 1)
    iFrpIndex(1 To 12) As Integer               '  24 bytes 288 - 311 (12 * 2)
    bytTpglType As Byte                         '   1 byte  312 - 312
    bytOpSjp As Byte                            '   1 byte  313 - 313
    iLocSource As Integer                       '   2 bytes 314 - 315
End Type ' SigData

Public Type SigData3_0                          ' 328 bytes Fixed
    bytSemStat As Byte                          '   1 byte    0 -   0
    bytEmitterAnalyzed As Byte                  '   1 byte    1 -   1
    iSignum As Integer                          '   2 bytes   2 -   3
    iSigID As Integer                           '   2 bytes   4 -   5
    bytNumActiveChan As Byte                    '   1 byte    6 -   6
    bPtd As Byte                                '   1 byte    7 -   7
    lFreq As Long                               '   4 bytes   8 -  11
    lBandwidth As Long                          '   4 bytes  12 -  15
    iModType As Integer                         '   2 bytes  16 -  17
    bytRadioType As Byte                        '   1 byte   18 -  18
    bytNumberOfChannels As Byte                 '   1 byte   19 -  19
    uChannelData(1 To 12) As ChannelData        '  96 bytes  20 - 115 (12 * 8)
    iEchid As Integer                           '   2 bytes 116 - 117
    iMasterSig As Integer                       '   2 bytes 118 - 119
    bytSigjampri As Byte                        '   1 byte  120 - 120
    bytJamMeth As Byte                          '   1 byte  121 - 121
    bytJamConf As Byte                          '   1 byte  122 - 122
    bytListenConf As Byte                       '   1 byte  123 - 123
    bytActBySig As Byte                         '   1 byte  124 - 124
    bytIncBySig As Byte                         '   1 byte  125 - 125
    bytEnjBySig As Byte                         '   1 byte  126 - 126
    bytPtctOverride As Byte                     '   1 byte  127 - 127
    iAsignum As Integer                         '   2 bytes 128 - 129
    iCorrId As Integer                          '   2 bytes 130 - 131
    iEchLink As Integer                         '   2 bytes 132 - 133
    bytLobValidated As Byte                     '   1 byte  134 - 134
    bytAcqDrngJam As Byte                       '   1 byte  135 - 135
    lTimeAcq As Long                            '   4 bytes 136 - 139
    iFreqLink As Integer                        '   2 bytes 140 - 141
    iMpEchid As Integer                         '   2 bytes 142 - 143
    iOtherEchId As Integer                      '   2 bytes 144 - 145
    iOprAsg As Integer                          '   2 bytes 146 - 147
    bytSigStatus As Byte                        '   1 byte  148 - 148
    bytMpStatus As Byte                         '   1 byte  149 - 149
    bytChanged As Byte                          '   1 byte  150 - 150
    bytValStat As Byte                          '   1 byte  151 - 151
    lFixId As Long                              '   4 bytes 152 - 155
    uFixloc As LatLon                           '   8 bytes 156 - 163
    iFixBrng As Integer                         '   2 bytes 164 - 165
    iMinorAxis As Integer                       '   2 bytes 166 - 167
    iMajorAxis As Integer                       '   2 bytes 168 - 169
    bytFixType As Byte                          '   1 byte  170 - 170
    bytOprqIndex As Byte                        '   1 byte  171 - 171
    iFoiQueue As Integer                        '   2 bytes 172 - 173
    iAltFreqLink As Integer                     '   2 bytes 174 - 175
    bytLang(1 To 3) As Byte                     '   3 bytes 176 - 178
    bytNetwork(1 To 19) As Byte                 '  19 bytes 179 - 197
    bytAud(1 To 9) As Byte                      '   9 bytes 198 - 206
    bytChannel(1 To 3) As Byte                  '   3 bytes 207 - 209
    bytWstype(1 To 17) As Byte                  '  17 bytes 210 - 226
    bytStationNum(1 To 11) As Byte              '  11 bytes 227 - 237
    bytNarr(1 To 21) As Byte                    '  21 bytes 238 - 258
    bytCollArea As Byte                         '   1 byte  259 - 259
    iSource As Integer                          '   2 bytes 260 - 261
    iCallLink As Integer                        '   2 bytes 262 - 263
    bytPickedSignal As Byte                     '   1 byte  264 - 264
    bytRespOpr As Byte                          '   1 byte  265 - 265
    iTmtn As Integer                            '   2 bytes 266 - 267
    iTellStatus As Integer                      '   2 bytes 268 - 269
    bytRemoteSource As Byte                     '   1 byte  270 - 270
    bytAllegiance As Byte                       '   1 byte  271 - 271
    bytAllegSource As Byte                      '   1 byte  272 - 272
    bytDeconOverride As Byte                    '   1 byte  273 - 273
    iNumFrps As Integer                         '   2 bytes 274 - 275
    bytFrpChan(1 To 12) As Byte                 '  12 bytes 276 - 287 (12 * 1)
    iFrpIndex(1 To 12) As Integer               '  24 bytes 288 - 311 (12 * 2)
    bytTpglType As Byte                         '   1 byte  312 - 312
    bytOpSjp As Byte                            '   1 byte  313 - 313
    iLocSource As Integer                       '   2 bytes 314 - 315
    uWPSignalData As WPSignalData               '  12 bytes 316 - 327
    iChanId As Integer                          '   2 bytes 328 - 329
    iPaId As Integer                            '   2 bytes 330 - 331
End Type ' SigData3_0

Public Type DynamicSigData                      '   4 bytes fixed
    bytUpDown As Byte                           '   1 byte    0 -   0
    bytJamStatus As Byte                        '   1 byte    1 -   1
    bytSigAmp As Byte                           '   1 byte    2 -   2
    bytPad0 As Byte                             '   1 byte    3 -   3
End Type 'DynamicSigData

Public Type SignalRec                           ' 320 bytes fixed
    uSigData As SigData                         ' 316 bytes   0 - 315
    uDynamiSigData As DynamicSigData            '   4 bytes 316 - 319
End Type ' SignalRec

Public Type SignalRec3_0                        ' 320 bytes fixed
    uSigData As SigData3_0                      ' 316 bytes   0 - 315
    uDynamiSigData As DynamicSigData            '   4 bytes 316 - 319
End Type ' SignalRec3_0

Public Type LobSet                              '  40 bytes fixed
    uLobData As LobPacket                       '  16 bytes   0 -  15
    iNumberOfSigIds As Integer                  '   2 bytes  16 -  17
    iCorrelatedSigId(1 To 10) As Integer        '  20 bytes  18 -  37 (10 * 2)
    bytPad0(1 To 2) As Byte                     '   2 bytes  38 -  39 (2 * 1)
End Type ' LobSet

Public Type VarRecHdr
    lRecSize As Long
    lIndex As Long
End Type ' VarRecHdr

Public Type DlddataInfo
    lRecSize As Long
    lStartIndex As Long
    lSize As Long
    lByteOffset As Long
End Type ' DlddataInfo

Public Type TiSysTime
    iYear As Integer
    iMonth As Integer
    iDate As Integer
    iDay As Integer
    iHour As Integer
    iMin As Integer
    iSec As Integer
    iJday As Integer
    lSysRfosSec As Long
    lStMissSec As Long
End Type ' TiSysTime

Public Type Gtsystime
    uTime As TiSysTime
    iBeginRfosYear As Integer
    iBeginRfosMonth As Integer
    iBeginRfosDate As Integer
    iBeginRfosHour As Integer
    iBeginRfosMin As Integer
    iBeginRfosSec As Integer
    iBeginRfosDay As Integer
End Type ' Gtsystime

Public Type Fxdef                               ' 136 bytes fixed
    uFixloc As LatLon                           '   8 bytes   0 -   7
    iMajaxis As Integer                         '   2 bytes   8 -   9
    iMinaxis As Integer                         '   2 bytes  10 -  11
    iTrubear As Integer                         '   2 bytes  12 -  13
    bytInPma As Byte                            '   1 byte   14 -  14
    bytPad0 As Byte                             '   1 byte   15 -  15
    iSigID As Integer                           '   2 bytes  16 -  17
    bytPad1(1 To 2) As Byte                     '   2 bytes  18 -  19 (2 * 1)
    lFreq As Long                               '   4 bytes  20 -  23
    lBandwidth As Long                          '   4 bytes  24 -  27
    bytRadioType As Byte                        '   1 byte   28 -  28
    bytNumberOfChannels As Byte                 '   1 byte   29 -  29
    bytPad2(1 To 2) As Byte                     '   2 bytes  30 -  31 (2 * 1)
    uChannelData(1 To 12) As ChannelData        '  96 bytes  32 - 127 (12 * 8)
    bytFixType As Byte                          '   1 byte  128 - 128
    bytHowAssoc As Byte                         '   1 byte  129 - 129
    bytPad3(1 To 2) As Byte                     '   2 bytes 130 - 131 (2 * 1)
    lFixId As Long                              '   4 bytes 132 - 135
End Type ' Fxdef

Public Type LhEchRec
    bytSemStat As Byte
    bytDisplay As Byte
    bytEchType As Byte
    bytLmbEchJamPri As Byte
    uLoc As LatLon                          '   8 bytes
    bytIncByEch As Byte
    bytActByEch As Byte
    bytRestricted As Byte
    bytEchLabel(1 To 9) As Byte
    iCorrIndex As Integer
    iNetIndex As Integer
    iTrackIndex As Integer
    iSource As Integer
    iTibsTellStatus As Integer
    iTibsTmtn As Integer
    bytTibsSimulated As Byte
    bytAlleg As Byte
    bytAllegSource As Byte
    bytPad0 As Byte
    iOtherEchId As Integer
    bytOtherPrimary As Byte
    bytWsMtStatus As Byte
    iWsMtIndex As Integer
    iWsMtLink As Integer
    iLmbIndex As Integer
    iHbIndex As Integer
    iNextRecLink As Integer
    iContribSigs(1 To 10) As Integer
    bytDeconOverride As Byte
    bytOpEjp As Byte
    iEchRelodtgtId As Integer
    bytPad1(1 To 2) As Byte
End Type 'LhEchRec

Public Type LhtrackData
    bytSemStat As Byte
    bytFirstTime As Byte
    bytTibsSimulated As Byte
    bytTrackType As Byte
    iTibsRaidSize As Integer
    iTibsTmtn As Integer
    iTibsTellStatus As Integer
    iTibsUnitNmbr As Integer
    uTibsTrackLoc As LatLon             '   8 bytes
    iTibsLocAccy As Integer
    iTibsPec As Integer
    uHbsTrackLoc As LatLon              '   8 bytes
    iHbsTrackId As Integer
    bytTrackClass As Byte
    bytTrackConf As Byte
    iTrackBearing As Integer
    iTrackSpeed As Integer
    iTrackAngle As Integer
    iDpsTrackId As Integer
    uDpsTrackLoc As LatLon              '   8 bytes
    iDpsAttackId As Integer
    bytDpsRefStatus As Byte
    bytAllegiance As Byte
    iAltitude As Integer
    iNumContributors As Integer
    uMsmtTime As TodTime                '   8 bytes
End Type 'LhtrackData

Public Type Trackupd
    iRecnum As Integer
    bytPad0(1 To 2) As Byte
    uTrack As LhtrackData
    uBestloc As LatLon                  '   8 bytes
    iEchid As Integer
    bytPad1(1 To 2) As Byte
End Type 'Trackupd

Public Type XyCoordinate
    dX As Long
    dY As Long
End Type 'XyCoordinate

Public Type XyCopy
    dX As Single
    dY As Single
End Type 'XyCopy

Public Type FreqEntry
    uNextEntry As Long 'pointer to next freq entry
    uPrevEntry As Long ' pointer to previous freq entry
    lStartFreq As Long
    lStopFreq As Long
    iLobTableId As Integer
End Type 'FreqEntry

Public Type LobFlags
    lLobBitMap As Long
End Type 'LobFlags

Public Type LobRecord
    lLobId As Long
    uTod As TodTime
    uPosition As XyCoordinate
    uFlags As LobFlags
    iTrueBearing As Integer
    iFixAssoc As Integer
End Type 'LobRecord

Public Type LobListRecord
    iLobListId As Integer
    iLobTableId As Integer
    iNumberOfLobs As Integer
    iNextLobListId As Integer
    uLobs(1 To 30) As LobRecord
End Type 'LobListRecord

Public Type TaskingRec
    lFrequency As Long
    lBandwidth As Long
End Type 'TaskingRec

Public Type LobTableEntry
    iLobTableId As Integer
    iDrivingSigid As Integer
    bytNumberOfSigids As Byte
    iSigIds(1 To 10) As Integer
    bytEldbCutRequested As Byte
    bytEldbFixPending As Byte
    uNextEldbCut As TodTime
    lEldbCutInterval As Long
    uTotalTasking As TaskingRec
    uLobTasking(1 To 10) As TaskingRec
    iNumberOfCutReqs As Integer
    iLobListId As Integer
    uFreqPtr As FreqEntry
    bytLastEldbFailed As Byte
    bytAttenuation As Byte
End Type 'LobTableEntry

Public Type FixEntry
    iSigID As Integer
    iFixId As Integer
    uFix As XyCoordinate
    iMajaxis As Integer
    iMinaxis As Integer
    iOrientAngle As Integer
    bytInPma As Byte
    bytPad As Byte
End Type 'FixEntry

Public Type EldbEntry
    iSigID As Integer
    iNumCuts As Integer
    iLobTableId As Integer
    bytOperFixPending As Byte
    bytCurrentFixType As Byte
    bytNumFailedFixes As Byte
    bytPad(1 To 3) As Byte
    lEldbCutInterval As Long
    iFixStartBear As Integer
    iDefaultLocation As Integer
    uFixStartTime As TodTime                '   8 bytes
    uFixEndTime As TodTime                  '   8 bytes
    uFixMinEndTime As TodTime               '   8 bytes
    uFixStartLoc As XyCoordinate
    lFixStartLobId As Long
    uFixinfo(1 To 4) As FixEntry
End Type 'EldbEntry

Public Type HbsigData
    bytTemp(1 To 16) As Byte
End Type 'HbsigData

Public Type Contrib
    iHbsTrackId As Integer
    bytAid(1 To 6) As Byte
    bytOnBoard As Byte
    bytSource As Byte
    bytChannel As Byte
    bytSignal As Byte
    bytHbsTrackMeth As Byte
    bytBeam As Byte
    bytSlsAmpDiff As Byte
    bytHbsJamEffPred As Byte
    lMopFreq As Long
    uLastUpdTime As TodTime             '   8 bytes
    iDpsSigId As Integer
    iDpsLmbChan As Integer
    uData As HbsigData
End Type 'Contrib

Public Type HBTrackData
    bytFirstTime As Byte
    bytTibsSimulated As Byte
    iTibsRaidSize As Integer
    iTibsTmtn As Integer
    iTibsTellStatus As Integer
    iTibsUnitNmbr As Integer
    iTibsLocAccy As Integer
    uTibsTrackLoc As LatLon             '   8 bytes
    iTibsPec As Integer
    iHbsTrackId As Integer
    uHbsTrackLoc As LatLon              '   8 bytes
    bytTrackClass As Byte
    bytTrackConf As Byte
    iTrackBearing As Integer
    iTrackSpeed As Integer
    iTrackAngle As Integer
    iDpsTrackId As Integer
    iBestContribIndex As Integer
    uDpsTrackLoc As LatLon              '   8 bytes
    iDpsAttackId As Integer
    bytDpsRefStatus As Byte
    bytAllegiance As Byte
    iAltitude As Integer
    bytTrackType As Byte
    bytMergedIdSource(1 To 5) As Byte
    iMergedSourceId(1 To 5) As Integer
    bytMsmtFlag As Byte
    bytMsmtSignal As Byte
    uMsmtTime As TodTime                '   8 bytes
    snglMsmtRange As Single
    snglMsmtBearing As Single
    snglMsmtXpos As Single
    snglMsmtYpos As Single
    lTidno As Long
    dMsmtCovXx As Double
    dMsmtCovXy As Double
    dMsmtCovYy As Double
    dMsmtBearVar As Double
    bytTibsAoi(1 To 6) As Byte
    iNumContributors As Integer
    uContributor(0) As Contrib
End Type 'HBTrackData

' start of message definitions

Public Type Mtdefpma                            '  64 bytes Fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uPMA As PMA                                 '  52 bytes  12 -  63
End Type ' Mtdefpma

Public Type Mtsigalarm                          ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumAlarms As Integer                       '   2 bytes  12 -  13
    bytPad(1 To 2) As Byte                      '   2 bytes  14 -  15
    uAlarm(0) As SignalAlarmData                ' x32 bytes  16 - ...
End Type 'Mtsigalarm

Public Type Mtanaalarm                          ' 290 bytes
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iSigID As Integer                           '   2 bytes  12 -  13
    iPad0 As Integer                            '   2 bytes  14 -  15
    lFrequency As Long                          '   4 bytes  16 -  19
    lBandwidth As Long                          '   4 bytes  20 -  23
    bytRadioType As Byte                        '   1 byte   24 -  24
    bytPad1 As Byte                             '   1 byte   25 -  25
    iNumberOfChannels As Integer                '   2 bytes  26 -  27
    uChannelData(1 To 12) As ChannelData        '  96 bytes  28 - 123
    bytTimeDataAvail As Byte                    '   1 byte  124 - 124
    bytPad3(1 To 3) As Byte                     '   3 bytes 125 - 127
    uSignalTime(1 To 12) As TodTime             '  96 bytes 128 - 223
    lSignalLength(1 To 12) As Long              '  48 bytes 224 - 271
    iRequestorID As Integer                     '   2 bytes 272 - 273
    bytAmplitude As Byte                        '   1 byte  274 - 274
    bytPassFreq As Byte                         '   1 byte  275 - 275
    iRunMode As Integer                         '   2 bytes 276 - 277
    bytChangeSigclassvar As Byte                '   1 byte  278 - 278
    bytPad4 As Byte                             '   1 byte  279 - 279
    uSigAnaTime As TodTime                      '   8 bytes 280 - 287
End Type 'Mtanaalaarm

Public Type Mtanaalarm3_0                       ' 290 bytes
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iSigID As Integer                           '   2 bytes  12 -  13
    iPad0 As Integer                            '   2 bytes  14 -  15
    lFrequency As Long                          '   4 bytes  16 -  19
    lBandwidth As Long                          '   4 bytes  20 -  23
    bytRadioType As Byte                        '   1 byte   24 -  24
    bOperRequest As Byte                        '   1 byte   25 -  25
    iNumberOfChannels As Integer                '   2 bytes  26 -  27
    uChannelData(1 To 12) As ChannelData        '  96 bytes  28 - 123
    bytTimeDataAvail As Byte                    '   1 byte  124 - 124
    bytPad3(1 To 3) As Byte                     '   3 bytes 125 - 127
    uSignalTime(1 To 12) As TodTime             '  96 bytes 128 - 223
    lSignalLength(1 To 12) As Long              '  48 bytes 224 - 271
    iRequestorID As Integer                     '   2 bytes 272 - 273
    bytAmplitude As Byte                        '   1 byte  274 - 274
    bytPassFreq As Byte                         '   1 byte  275 - 275
    iRunMode As Integer                         '   2 bytes 276 - 277
    bytChangeSigclassvar As Byte                '   1 byte  278 - 278
    bytPad4 As Byte                             '   1 bytes 279 - 279
    uSigAnaTime As TodTime                      '   8 bytes 280 - 287
    uWPResultsData As WPResultsData             '   8 bytes 288 - 295
End Type 'Mtanaalaarm3_0

Public Type Mthbdynrsp                          ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumOfRsps As Integer                       '   2 bytes  12 -  13
    iRequestorID As Integer                     '   2 bytes  14 -  15
    uHbDynRsp(0) As HbdynRspRec                 '   ? bytes  16 - ...
End Type 'Mthbdynrsp

Public Type Mtsdalarm                           '   Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumAlarms As Integer                       '   2 bytes  12 -  13
    bytPad(1 To 2) As Byte                      '   2 bytes  14 -  15
    uAlarm(0) As ShortDurationAlarmData         'x708 bytes  16 - ...
End Type 'Mtsdalarm
'
'+v1.6BB
Public Type Mtsdalarm2_3
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumAlarms As Integer
    bytPad(1 To 2) As Byte
    uAlarm(0) As ShortDurationAlarmData2_3
End Type 'Mtsdalarm2_3
'-v1.6

Public Type Mthbactrep
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uTime As TodTime                            '   8 bytes  12 -  19
    iNumSigs As Integer                         '   2 bytes  20 -  21
    bytPad(1 To 2) As Byte                      '   2 bytes  22 -  23
    uAct(0) As ActRec                           '
End Type 'Mthbactrep

Public Type Mthbactrep3_0
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uTime As TodTime                            '   8 bytes  12 -  19
    iNumSigs As Integer                         '   2 bytes  20 -  21
    bytPad(1 To 2) As Byte                      '   2 bytes  22 -  23
    uAct(0) As ActRec3_0                        '
End Type 'Mthbactrep3_0

Public Type Mtdfalarm
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumAlarms As Integer                       '   2 bytes  12 -  13
    bytPad(1 To 2) As Byte                      '   2 bytes  14 -  15
    uAlarm(0) As DfAlarmData                    'x648 bytes  16 - ...
End Type 'Mtdfalarm

Public Type Mtrfstatus
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uRfStatusTime As TodTime                    '   8 bytes
    bytStatus(1 To gbytCLASSIFIED) As Byte
End Type 'Mtrfstatus

Public Type Mtdfsdalarm                         ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumAlarms As Integer                       '   2 bytes  12 -  13
    bytPad(1 To 2) As Byte                      '   2 bytes  14 -  15
    uAlarm(0) As DfShortDurationResultsData     'x616 bytes  16 - ...
End Type 'Mtdfsdalarm

Public Type Mtlobupd                            ' 272 bytes Fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iSigID As Integer                           '   2 bytes  12 -  13
    bytNumLobs As Byte                          '   1 byte   14 -  14
    bytPad As Byte                              '   1 byte   15 -  15
    uLobData(1 To 16) As LobPacket              ' 256 bytes  16 - 271 (16 * 16)
End Type 'Mtlobupd

Public Type Mtanarslt                           ' 932 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iSigID As Integer                           '   2 bytes  12 -  13
    bytPad0(1 To 2) As Byte                     '   2 bytes  14 -  15 (2 * 1)
    lFreq As Long                               '   4 bytes  16 -  19
    lBandwidth As Long                          '   4 bytes  20 -  23
    bytRadioType As Byte                        '   1 byte   24 -  24
    bytPad1 As Byte                             '   1 byte   25 -  25
    iNumberOfChannels As Integer                '   2 bytes  26 -  27
    uChannelData(1 To 12) As ChannelData        '  96 bytes  28 - 123 (12 * 8)
    bytSignalPresentAna As Byte                 '   1 byte  124 - 124
    bytSignalPresentDf As Byte                  '   1 byte  125 - 125
    bytAmplitude As Byte                        '   1 byte  126 - 126
    bytPassFreq As Byte                         '   1 byte  127 - 127
    bytLobRequestType As Byte                   '   1 byte  128 - 128
    bytLobAvail As Byte                         '   1 byte  129 - 129
    bytTimeDataAvail As Byte                    '   1 byte  130 - 130
    bytNumLobs As Byte                          '   1 byte  131 - 131
    uLobData(1 To 16) As LobSet                 ' 640 bytes 132 - 771 (16 * 40)
    iRequestorID As Integer                     '   2 bytes 772 - 773
    iRunMode As Integer                         '   2 bytes 774 - 775
    uSigAnaTime As TodTime                      '   8 bytes 776 - 783
    bytChangeSigvar As Byte                     '   1 byte  784 - 784
    bytLobValidated As Byte                     '   1 byte  785 - 785
    bytSignalCreated As Byte                    '   1 byte  786 - 786
    bytPad2 As Byte                             '   1 byte  787 - 787
    uSignalTime(1 To 12) As TodTime             '  96 bytes 788 - 883 (12 * 8)
    lSignalLength(1 To 12) As Long              '  48 bytes 884 - 931 (12 * 4)
End Type 'Mtanarslt

Public Type Mtanarslt3_0                        ' 932 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iSigID As Integer                           '   2 bytes  12 -  13
    bytPad0(1 To 2) As Byte                     '   2 bytes  14 -  15 (2 * 1)
    lFreq As Long                               '   4 bytes  16 -  19
    lBandwidth As Long                          '   4 bytes  20 -  23
    bytRadioType As Byte                        '   1 byte   24 -  24
    bytPad1 As Byte                             '   1 byte   25 -  25
    iNumberOfChannels As Integer                '   2 bytes  26 -  27
    uChannelData(1 To 12) As ChannelData        '  96 bytes  28 - 123 (12 * 8)
    bytSignalPresentAna As Byte                 '   1 byte  124 - 124
    bytSignalPresentDf As Byte                  '   1 byte  125 - 125
    bytAmplitude As Byte                        '   1 byte  126 - 126
    bytPassFreq As Byte                         '   1 byte  127 - 127
    bytLobRequestType As Byte                   '   1 byte  128 - 128
    bytLobAvail As Byte                         '   1 byte  129 - 129
    bytTimeDataAvail As Byte                    '   1 byte  130 - 130
    bytNumLobs As Byte                          '   1 byte  131 - 131
    uLobData(1 To 16) As LobSet                 ' 640 bytes 132 - 771 (16 * 40)
    iRequestorID As Integer                     '   2 bytes 772 - 773
    iRunMode As Integer                         '   2 bytes 774 - 775
    uSigAnaTime As TodTime                      '   8 bytes 776 - 783
    bytChangeSigvar As Byte                     '   1 byte  784 - 784
    bytLobValidated As Byte                     '   1 byte  785 - 785
    bytSignalCreated As Byte                    '   1 byte  786 - 786
    bytPad2 As Byte                             '   1 byte  787 - 787
    uSignalTime(1 To 12) As TodTime             '  96 bytes 788 - 883 (12 * 8)
    lSignalLength(1 To 12) As Long              '  48 bytes 884 - 931 (12 * 4)
    uWPResultsData As WPResultsData             '   8 bytes 932 - 939
End Type 'Mtanarslt3_0

Public Type Mtsigupd                            ' 336 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iRequestorID As Integer                     '   2 bytes  12 -  13
    bytPado(1 To 2) As Byte                     '   2 bytes  14 -  15
    uSig As SignalRec                           ' 320 bytes  16 - 335
End Type 'Mtsigupd

Public Type Mtsigupd3_0                         ' 336 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iRequestorID As Integer                     '   2 bytes  12 -  13
    bytPado(1 To 2) As Byte                     '   2 bytes  14 -  15
    uSig As SignalRec3_0                        ' 320 bytes  16 - 335
End Type 'Mtsigupd3_0

Public Type Mthblobupd                          ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iRecsNMsg As Integer                        '   2 bytes  12 -  13
    bytDone As Byte                             '   1 byte   14 -  14
    bytPado As Byte                             '   1 byte   15 -  15
    lBurstTime(1 To 8) As Long                  '  32 bytes  16 -  47 (8 * 4)
    uHblobEntry(0) As HbLobTblEntry             'x572 bytes  48 - ...
End Type 'Mthblobupd

Public Type Mtulddata
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    lMsgSubtype As Long
    lMsgSeqNo As Long
    lType As Long
    bytMoreData As Byte
    bytPad0(1 To 3) As Byte
    uInfo As DlddataInfo
    'bytData(1) As Byte
    'bytPad1(1 To 3) As Byte
End Type 'Mtulddata

Public Type Mtlobsetrslt                        ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iSigID As Integer                           '   2 bytes  12 -  13
    iRequestorID As Integer                     '   2 bytes  14 -  15
    iNumRec As Integer                          '   2 bytes  16 -  17
    bytPad(1 To 2) As Byte                      '   2 bytes  18 -  19
    uLobs(0) As LobPacket                       ' x16 bytes  20 - ...
End Type 'Mtlobsetrslt

Public Type LobData
    uLobs As LobPacket
    iNumSigIds As Integer
    iCorrSigId As Integer
    Pad0(1 To 2) As Byte
End Type 'LobData

Public Type Mtlobrslt                           ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    lFreq As Long                               '   4 bytes  12 -  15
    lBandwidth As Long                          '   4 bytes  16 -  19
    bytLobStatus As Byte                        '   1 byte   20 -  20
    Pad0 As Byte                                '   1 bytes  21 -  21
    iRequestorID As Integer                     '   2 bytes  22 -  23
    iNumRec As Integer                          '   2 bytes  24 -  25
    pad1(1 To 2) As Byte                        '   2 bytes  26 -  27
    uLobData(0) As LobData                       ' x16 bytes  20 - ...
End Type 'Mtlobrslt

Public Type Mtfixrslt                           '1108 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes    4 -   11
    bytNumFixes As Byte                         '   1 byte    12 -   12
    bytPad0(1 To 3) As Byte                     '   3 bytes   13 -   15 (3 * 1)
    uFixinfo(1 To 8) As Fxdef                   '1088 bytes   16 - 1103 (8 * 136)
    iRequestorID As Integer                     '   2 bytes 1104 - 1105
    bytFixStatus As Byte                        '   1 byte  1106 - 1106
    bytPad1 As Byte                             '   1 byte  1107 - 1107
End Type 'Mtfixrslt

Public Type Mthbsigupd                          ' 132 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iRequestorID As Integer                     '   2 bytes  12 -  13
    iSignum As Integer                          '   2 bytes  14 -  15
    bytFloat64Pad(1 To 4) As Byte               '   4 bytes  16 -  19
    uHbSigRec As HbSigRec                       ' 112 bytes  20 - 131
End Type 'Mthbsigupd

Public Type Mthbgsrep                           ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uTime As TodTime                            '   8 bytes  12 -  19
    iNumGsBurst As Integer                      '   2 bytes  20 -  21
    iNumGs As Integer                           '   2 bytes  22 -  23
    bytFloat64Pad(1 To 4) As Byte               '   4 bytes  24 -  27
    uGsrRec(0) As GroundSite                    ' x40 bytes  28 - ...
End Type 'Mthbgsrep

Public Type Mthbsemistat
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uTime As TodTime                            '   8 bytes  12 -  19
    iNumEntries As Integer                      '   2 bytes  20 -  21
    bytPad0(1 To 2) As Byte                     '   2 bytes  22 -  23
    uHbsemiinfo(0) As Mthbsemiinfostruct        ' ??? bytes  24 - ...
End Type 'Mthbsemistat

Public Type Mthbselask                          '  14 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    bytOnOff As Byte                            '   1 byte   12 -  12
    bytSigType As Byte                          '   1 byte   13 -  13
End Type 'Mthbselask

Public Type Mtjamstat                           ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    bytNumPackets As Byte                       '   1 byte   12 -  12
    bytPacketNum As Byte                        '   1 byte   13 -  13
    iNumSigID As Integer                        '   2 bytes  14 -  15
    uSigStatus(0) As SigPacket                  '  x4 bytes  16 - ...
End Type 'Mtjamstat

Public Type Mtrunmode                           ' 20 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iRunMode As Integer                         '   2 bytes  12 -  13
    iRequestorID As Integer                     '   2 bytes  14 -  15
    iLmbRunmode As Integer                      '   2 bytes  16 -  17
    iHbRunmode As Integer                       '   2 bytes  18 -  19
End Type 'Mtrunmode

Public Type Mtrunmode3_0                        ' 20 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iRunMode As Integer                         '   2 bytes  12 -  13
    iRequestorID As Integer                     '   2 bytes  14 -  15
    iLmbRunmode As Integer                      '   2 bytes  16 -  17
    iHbRunmode As Integer                       '   2 bytes  18 -  19
    iSprRunmode As Integer                      '   2 bytes  20 -  21
    bytPad0(1 To 2) As Byte                     '   2 bytes  22 -  23
End Type 'Mtrunmode3_0

Public Type Mthbseljam                          '  14 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    bytOnOff As Byte                            '   1 byte   12 -  12
    bytSigType As Byte                          '   1 byte   13 -  13
End Type 'Mthbseljam

Public Type Mthbxmtrstat                        ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iHboXmtrStatus As Integer                   '   2 bytes  12 -  13
    iHb2XmtrStatus As Integer                   '   2 bytes  14 -  15
    uP34XmtrStatus As P34Stat                   '   8 bytes  16 -  23
    iNumSigs As Integer                         '   2 bytes  24 -  25
    bytPad(1 To 2) As Byte                      '   2 bytes  26 -  27
    uSigMode(0) As SigMode                      ' ??? bytes  28 - ...
End Type 'Mthbxmtrstat

Public Type Mthbxmtrstat3_0                     ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iMbXmtrStatus As Integer                    '   2 bytes  12 -  13
    iHb1XmtrStatus As Integer                   '   2 bytes  14 -  15
    iHb2XmtrStatus As Integer                   '   2 bytes  16 -  17
    iHb3XmtrStatus As Integer                   '   2 bytes  18 -  19
    iSb1XmtrStatus As Integer                   '   2 bytes  20 -  21
    iSpear1XmtrStatus As Integer                '   2 bytes  22 -  23
    iSpear2XmtrStatus As Integer                '   2 bytes  24 -  25
    iSpear3XmtrStatus As Integer                '   2 bytes  26 -  27
    iSpear4XmtrStatus As Integer                '   2 bytes  28 -  29
    iNumSigs As Integer                         '   2 bytes  30 -  31
    uSigMode(0) As SigMode                      ' ??? bytes  32 - ...
End Type 'Mthbxmtrstat

Public Type Mtnavrep                            '  68 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uNavdata As NavData                         '  56 bytes  12 -  67
End Type 'Mtnavrep

Public Type Mtacqsigstat
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumSigIds As Integer
    iStarting_Sigid As Integer
    bytAlarmData(0) As Byte
End Type 'Mtacqsigstat

Public Type Mtsetacqsmode
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    bytSubMode As Byte
    bytStartSegment As Byte
End Type 'Mtsetacqsmode

Public Type Mttimesyn
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uSystime As Gtsystime
End Type 'Mttimesyn

Public Type Mtlhcorrelate
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iRequestorID As Integer
    bytCorrType As Byte
    bytPad0 As Byte
    iSubjectId As Integer
    bytSubjectId As Byte
    bytPad1 As Byte
    iObjectId As Integer
    bytObjectId As Byte
    bytPad2 As Byte
End Type 'Mtlhcorrelate

Public Type Mtlhechmod
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iRequestorID As Integer
    iEchid As Integer
    iDpsId As Integer
    bytEnder As Byte
    bytPad0 As Byte
    uEch As LhEchRec
End Type 'Mtlhechmod

Public Type Mtlhtrackupd
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumTracks As Integer
    bytPad0(1 To 2) As Byte
    uTrack(1 To 60) As Trackupd
End Type 'Mtlhtrackupd

Public Type Mtlhtrackrep
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uTime As TodTime                            '   8 bytes
    bytLastFlag As Byte
    bytMsgSeqNum As Byte
    iMsgSize As Integer
    bytFloat64Pad(1 To 4) As Byte
    uTrackMsg(0) As HBTrackData
End Type 'Mtlhtrackupd
'
'+v1.6BB
Public Type SsErrorReport                       ' 192 bytes fixed
    iSeverity As Integer                        '   2 bytes   0 -   1
    iCategory As Integer                        '   2 bytes   2 -   3
    iErrorCode As Integer                       '   2 bytes   4 -   5
    bytPad0(1 To 2) As Byte                     '   2 bytes   6 -   7 (2 * 1)
    uReportTime As TodTime                      '   8 bytes   8 -  15
    bytFile(1 To 40) As Byte                    '  40 bytes  16 -  55 (40 * 1)
    lLine As Long                               '   4 bytes  56 -  59
    iAmpDataLength As Integer                   '   2 bytes  60 -  61
    bytAmpdata(1 To 130) As Byte                ' 130 bytes  62 - 191 (130 * 1)
End Type 'SsErrorReport
'-v1.6
'
'+v1.6BB
Public Type Mtsserror                           '1936 bytes fixed
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumberErrorReports As Integer              '   2 bytes  12 -  13
    bytPad0(1 To 2) As Byte                     '   2 bytes  14 -  15 (2 * 1)
    uErrorReport(1 To 10) As SsErrorReport      '1920 bytes  16 -1935 (10 * 192)
End Type 'Mtsserror
'-v1.6
'
'BB 2002
Public Type Mtplanarea
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uPlanLoc As LatLon
    bytDwsNo As Byte
    bytMakegeo As Byte
    iNorthOffset As Integer
    lEarthRad As Long
End Type 'Mtplanarea
'BB 2002

'BB 2002
Public Type Mtanareq
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    lFreq As Long
    lBandwidth As Long
    iRequestorID As Integer
    bytPad(1 To 2) As Byte
End Type 'Mtanareq
'BB 2002

'BB 2002
Public Type Mtdfflgs
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    bytLobIntersectPma As Byte
    bytGoodQuality As Byte
    bytPad(1 To 2) As Byte
End Type 'Mtdfflgs
'BB 2002

Public Type TimeResult
    iSignum As Integer
    bytChannelNumber As Byte
    bytPad1 As Byte
    uTime As TodTime
    bytValidLength As Byte
    bytPad2(1 To 3) As Byte
    lLength As Long
End Type 'uTimeResult

Public Type Mtrrslt
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iStatus As Integer
    iNumberOfSignals As Integer
    uTimeResult(0) As TimeResult
End Type 'mtrrslt

Public Type AssocData                          '   bytes Fixed
    bytSignal As Byte                          '   1 byte    2 -   2
    bytChannel As Byte                         '   1 byte    3 -   3
    bytNumAssChan As Byte                      '   1 byte    4 -   4
    bytAssocChan(0) As Byte
End Type 'AssocData


Public Type Mthbassrep                           ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uTime As TodTime                            '   8 bytes  12 -  19
    iNumAssoc As Integer                        '   2 bytes  20 -  21
    bytPad(1 To 2) As Byte
    uAssoc(0) As AssocData                     ' x40 bytes  28 - ...
End Type 'Mthbassrep

Public Type GroundSiteUp                          '   bytes Fixed
    bytSignal As Byte                          '   1 byte    2 -   2
    bytChannel As Byte                         '   1 byte    3 -   3
    iId As Integer                             '   1 byte    4 -   4
End Type 'GroundSiteUp


Public Type Mthbgsupstat                        ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    uTime As TodTime                            '   8 bytes  12 -  19
    iNumGs As Integer                        '   2 bytes  20 -  21
    bytPad(1 To 2) As Byte
    uGsUp(0) As GroundSiteUp                     ' x40 bytes  28 - ...
End Type 'Mthbgsupstat

Public Type GroundSiteDel                          '   bytes Fixed
    bytSignal As Byte                          '   1 byte    2 -   2
    bytChannel As Byte                         '   1 byte    3 -   3
    iId As Integer                             '   1 byte    4 -   4
End Type 'GroundSiteDel


Public Type Mthbgsdel                        ' Variable length
    uMsgHdr As Msg_Hdr                          '   8 bytes   4 -  11
    iNumGs As Integer                        '   2 bytes  20 -  21
    bytPad(1 To 2) As Byte
    uGsDel(0) As GroundSiteDel                     ' x40 bytes  28 - ...
End Type 'Mthbgsdel

Public Type SsStatus                           '   bytes Fixed
    iRunMode As Integer
    iLan1 As Integer
    iLan2 As Integer
    iENet As Integer
End Type 'SsStatus

Public Type LanManage                          '   bytes Fixed
    byteLan1 As Byte
    byteLan2 As Byte
    byteNet1 As Byte
    byteNet2 As Byte
End Type 'LanManage

Public Type Mtsysconfig
    uMsgHdr As Msg_Hdr
    uLanUse As LanManage
    iCcuMode(1 To 2) As Integer
    uCuStatus(1 To 2) As SsStatus
    uDwStatus(1 To 6) As SsStatus
    uAcqStatus(1 To 2) As SsStatus
    uAnaStatus As SsStatus
    uExcStatus As SsStatus
    uDfStatus As SsStatus
    uHbsStatus As SsStatus
    uDpStatus As SsStatus
    uTsStatus As SsStatus
    uTasStatus As SsStatus
End Type 'Mtsysconfig

Public Type Mtsysconfig3_0
    uMsgHdr As Msg_Hdr
    uLanUse As LanManage
    iCcuMode(1 To 2) As Integer
    uCuStatus(1 To 2) As SsStatus
    uDwStatus(1 To 6) As SsStatus
    uAcqStatus(1 To 2) As SsStatus
    uAnaStatus As SsStatus
    uExcStatus As SsStatus
    uDfStatus As SsStatus
    uHbsStatus As SsStatus
    uDpStatus As SsStatus
    uTsStatus As SsStatus
    uTasStatus As SsStatus
    uSprStatus As SsStatus
    bytTaEclipseLb1Amp As Byte
    bytTaEclipseLb2Amp As Byte
    bytTaEclipseMb1Amp As Byte
    bytTaEclipseMb2Amp As Byte
    bytHbSb1Amp As Byte
    bytSpearRf1 As Byte
    bytSpearRf2 As Byte
    bytSpearRf3 As Byte
    bytSpearRf4 As Byte
    bytTaHbsu As Byte
    bytTaBlankingUnit As Byte
    bytPad As Byte
End Type 'Mtsysconfig3_0

Public Type Mtenvstat
    uMsgHdr As Msg_Hdr
    bytCurrentSubmode As Byte
    bytSegment As Byte
    iTimeLeft As Byte
    lReanalysisFreq1 As Long
    lReanalysisFreq2 As Long
End Type 'Mtenvstat

Public Type ExtRecType
    iDisc As Integer            'discriminant (Data Structure)
    iNumRecs As Integer         'number of Data recs in this msg
    iSigFrmFlag As Integer
    bytPad0(1 To 2) As Byte
   ' bytData(0) As Byte    'Data recs which is a union of all the ExtData types listed below
    'iRecType As Integer
End Type

Public Type DpsRespData
    iSignalID As Integer        'signal id
    bytError As Byte            'type of error if raised in dsa dps processing
    bytRadioType As Byte          'radio type designation
    bytChanNum As Byte          'channel number
    bytPad0(1 To 3) As Byte
    lFreq As Long               'dps use only
    uSigType As Sigclassvar     'sig type as determined by dsa
    uUsage As Byte              'usage determined by dsa
    bytLastPacket As Byte       'processing complete for the orig request, always true for partial extractions
    bytPad1(1 To 2) As Byte
    uTime As TodTime            'time of sig processing by dsa for data
    bytSigPresent As Byte       'signal was present during processing
    bytDataStatus As Byte       'sig data status as defined above
    bytDataType As Byte         'type of data returned by dsa dps processing
    bytPad2 As Byte
    iDataLength As Integer      'length of data in bytes
    bytPad3(1 To 2) As Byte
   ' uExtData(0) As Byte            'dps data - start on 4 byte boundary
   ' bytPad4(1 To 3) As Byte
End Type

Public Type Mtdpsextrsp
    uMsgHdr As Msg_Hdr
    uRspData As DpsRespData
End Type

Global Const CLASS18_SPEC_REC    As Byte = 3    'for class18 special records
Global Const CLASS14_TRACK_REC   As Byte = 10   'for class14 track records */
Global Const CLASS22_TRACK_REC   As Byte = 8    'for class22 track records */
Global Const REG_TRACK_REC       As Byte = 6    'for other track records */
Global Const CLASS14_SPEC_WORDS  As Byte = 40   'for class14 special record */
Global Const CLASS18_SPEC_WORDS  As Byte = 3    'for class18 special record */
Global Const CLASS22_SPEC_WORDS  As Byte = 5    'for class22 special record */
Global Const REG_SPEC_WORDS      As Byte = 17   'for other special records */

Public Type ExtData_TrackRec            '42 bytes
    iRecType As Integer
    iValidDataFlag As Integer
    iTan As Integer
    iIff As Integer
    lXval As Long
    lYval As Long
    lZalt As Long
    iXdot As Integer
    iYdot As Integer
    iZdot As Integer
    iAuxData(1 To 8) As Integer
    iParityFlags As Integer
End Type

Public Type ExtData_MissleRec
    uTrackRec As ExtData_TrackRec
End Type

Public Type ExtData_GciRec              '42 bytes
    iRecType As Integer
    iValidDataFlag As Integer
    iAddress As Integer
    iCourse As Integer
    iSpeed As Integer
    lAltitude As Long
    lRange As Long
    iRangeInd As Integer
    iAzimuth As Integer
    iElevation As Integer
    iCloseRate As Integer
    iAspect As Integer
    iAuxData(1 To 6) As Integer
    iParityFlags As Integer
End Type

Public Type ExtData_SpecRec             '38 bytes
    iRecType As Integer
    iNumWords As Integer
    iData(1 To REG_SPEC_WORDS) As Integer
End Type

Public Type ExtData_SpecRecDS           '40 bytes
    iRecType As Integer
    iDataStruct As Integer
    iNumWords As Integer
    iData(1 To REG_SPEC_WORDS) As Integer
End Type

Public Type Class6_SpecRec
    iRecType As Integer
    iDataStruct As Integer
    iNumWords As Integer
    iPDisc As Integer
    iSDisc As Integer
    iBatAddr As Integer
    iTanA As Integer
    iTanB As Integer
    iDinome As Integer
    iSector As Integer
    iAlt As Integer
    iXval As Integer
    iYval As Integer
    iBearing As Integer
    iDist As Integer
End Type

Public Type Class7_ZRV_SpecRec
    iRecType As Integer
    iNumWords As Integer
    iDisc As Integer
    iSubAddr As Integer
    iSpecfld(1 To 14) As Integer
End Type

Public Type Class7_RTV_SpecRec
    iRecType As Integer
    iDataStruct As Integer
    iNumWords As Integer
    iDisc As Integer
    iSubAddr As Integer
    iSpecfld(1 To 14) As Integer
End Type

Public Type Class8var1_3SpecRec
    iRecType As Integer
    iDataStruct As Integer
    iNumWords As Integer
    iXscrn As Integer
    iYscrn As Integer
    iSymAmp1 As Integer
    iSymAmp2 As Integer
End Type

Public Type Class8var2_4SpecRec
    iRecType As Integer
    iDataStruct As Integer
    iNumWords As Integer
    iSpecFlds(1 To 5) As Integer
End Type

Public Type Class9_SpecRec
    iRecType As Integer
    iDataStruct As Integer
    iNumWords As Integer
    iData As Integer
End Type

Public Type Class12_SpecRec
    iRecType As Integer
    iNumWords As Integer
    iDataStruct As Integer
    sText As String * 65
End Type

Public Type Class13_SpecRec
    iRecType As Integer
    iNumWords As Integer
    iDataStruct As Integer
    sText As String * 144
End Type

Public Type Commands
    iTanPar As Integer
    iTan As Integer
    iCmd As Integer
End Type

Public Type Responses
    iRspPar As Integer
    iRsp As Integer
End Type

Public Type Class14_SpecRec
    iRecType As Integer
    iNumWords As Integer
    uCmds(1 To 8) As Commands
    uRsps(1 To 8) As Responses
End Type

Public Type Class16_SpecRec
    iRecType As Integer
    iNumWords As Integer
    iResponse As Integer
End Type

Public Type ExtData_Class14Rec
    uSpecialData As Class14_SpecRec
    uTrackData(1 To CLASS14_TRACK_REC) As ExtData_TrackRec
End Type

Public Type Class17_SubRec
    iAangle As Integer
    iEangle As Integer
    iLcmd As Integer
End Type

Public Type Class17_SpecRec
    iRecType As Integer
    iDataStruct As Integer
    iNumWords As Integer
    uSubRec(1 To 3) As Class17_SubRec
End Type

Public Type Class18_SpecRec
    iRecType As Integer
    iDataStruct As Integer
    iNumWords As Integer
    iData(1 To 3) As Integer
End Type

Public Type Class19_SpecRec
    iRecType As Integer
    iNumWords As Integer
    iMsgType As Integer
    iTanA As Integer
    iTanB As Integer
    iXval As Integer
    iYval As Integer
End Type
    
Public Type Class22_SpecRec
    iRecType As Integer
    iNumWords As Integer
    iSync1 As Integer
    iSync2 As Integer
    iSync3 As Integer
    iSync4 As Integer
    iAck As Integer
End Type

'Public Type ExtData_Class22Rec
    'uSpecialData As Class22SpecRec
    'iTrackData(1 To CLASS22_TRACK_REC) As ExtData_TrackRec
'End Type

Public Type ExtData_UnkRec
    iRecType As Integer
    iFrmLen As Integer
    iUnkData As Integer
End Type
'
'+v1.8.10 TE
Public Type msgMTCNTTRG
    uMsgHdr As Msg_Hdr              ' 8
    bListType As Byte               ' 1
    bNumPackets As Byte             ' 1
    bPacketNumber As Byte           ' 1
    bytPad0(0) As Byte              ' 1
    uTimeStamp As TodTime           ' 8
    iSigIDIndex As Integer          ' 2
    iNumSigID As Integer            ' 2
    iSigID As Integer               ' 2
    bytPad1(1) As Byte              ' 2
End Type
'
Public Type msgMTTXRFCONF
    uMsgHdr As Msg_Hdr              ' 8
    bMsnChg As Byte                 ' 1 Status
    bytPad0(2) As Byte              ' 3
    lFreqStart As Long              ' 4 Freq
    lFreqEnd As Long                ' 4 PRI
    iTx_Src As Integer              ' 2 EmitterID
    iTx_Chan As Integer             ' 2 SignalID
    iTx_Pa_ID As Integer            ' 2 Tag
    iTx_Pa_Pwr_Setting As Integer   ' 2 Flag
    iTx_Ant_Grp As Integer          ' 2 Common
    bytPad1(1) As Byte              ' 2
End Type
'

