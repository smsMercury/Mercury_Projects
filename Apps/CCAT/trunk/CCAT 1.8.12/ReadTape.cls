VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReadTape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' COPYRIGHT (C) 2001, Mercury Solutions, Inc.
'
' MODULE:   ReadTape
' AUTHOR:   Keith Gibby
' PURPOSE:  Tape Library to Control and Read Files From Tape
' NOTES:    API Declarations from Public Domain Code at http://www.allapi.net/
'
' HISTORY:  Version 1.0 - 20011107
'
Option Explicit

Dim hTape, hDisk As Long
Dim SecurityAttributes As SECURITY_ATTRIBUTES
Dim bAbortRead As Boolean
Dim lStatus As Long
Dim sStatusMessage As String

Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function SetTapeParameters Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, lpTapeInformation As Any) As Long
Private Declare Function PrepareTape Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, ByVal bimmediate As Long) As Long
Private Declare Function SetTapePosition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPositionMethod As Long, ByVal dwPartition As Long, ByVal dwOffsetLow As Long, ByVal dwOffsetHigh As Long, ByVal bimmediate As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTapeParameters Lib "kernel32" (ByVal hDevice As Long, ByVal dwOperation As Long, lpdwSize As Long, lpTapeInformation As Any) As Long
Private Declare Function GetTapePosition Lib "kernel32" (ByVal hDevice As Long, ByVal dwPositionType As Long, lpdwPartition As Long, lpdwOffsetLow As Long, lpdwOffsetHigh As Long) As Long
Private Declare Function GetTapeStatus Lib "kernel32" (ByVal hDevice As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const GET_TAPE_DRIVE_INFORMATION = 1
Const GET_TAPE_MEDIA_INFORMATION = 0
Const SET_TAPE_DRIVE_INFORMATION = 1
Const SET_TAPE_MEDIA_INFORMATION = 0

Const GENERIC_READ = &H80000000      '###ADD-IN
Const OPEN_EXISTING = 3              '###ADD-IN
Const FILE_SHARE_READ = &H1          '###ADD-IN
Const FILE_SHARE_WRITE = &H2         '###ADD-IN
Const GENERIC_WRITE = &H40000000     '###ADD-IN
Const CREATE_ALWAYS = 2              '###ADD-IN
Const INVALID_HANDLE_VALUE = -1
Const NO_ERROR = 0
Const ERROR_BEGINNING_OF_MEDIA = 1102
Const ERROR_DEVICE_NOT_PARTITIONED = 1107
Const ERROR_END_OF_MEDIA = 1100
Const ERROR_EOM_OVERFLOW = 1129
Const ERROR_FILEMARK_DETECTED = 1101
Const ERROR_INVALID_BLOCK_LENGTH = 1106
Const ERROR_NO_DATA_DETECTED = 1104
Const ERROR_PARTITION_FAILURE = 1105
Const ERROR_SETMARK_DETECTED = 1103

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type TAPE_GET_MEDIA_PARAMETERS
    Capacity As LARGE_INTEGER
    Remaining As LARGE_INTEGER
    BlockSize As Long
    PartitionCount As Long
    WriteProtected As Byte
    PadBytes0 As Byte
    PadBytes1 As Long                        'WARNING****Structure has padding bytes  In C++ the structure is 32 bytes in VB it is 26 bytes
    PadBytes2 As Boolean                     'See MSDN Article Q249962, PadBytes might have to be removed in future versions
    'PadBytes3 As Boolean                     'WARNING****Structure has padding bytes  In C++ the structure is 32 bytes in VB it is 26 bytes

End Type

Private Type TAPE_GET_DRIVE_PARAMETERS
    ECC As Boolean
    Compression As Boolean
    DataPadding As Boolean
    ReportSetmarks As Boolean
    DefaultBlockSize As Long
    MaximumBlockSize As Long
    MinimumBlockSize As Long
    MaximumPartitionCount As Long
    FeaturesLow As Long
    FeaturesHigh As Long
    EOTWarningZoneSize As Long
End Type
Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
' following taken from Winnt.h
'
' IOCTL_TAPE_ERASE definitions
'
Const TAPE_ERASE_SHORT = 0
Const TAPE_ERASE_LONG = 1
Private Type TAPE_ERASE
    Type As Long
    Immediate As Boolean
' TAPE_ERASE, *PTAPE_ERASE;
End Type
'
' IOCTL_TAPE_PREPARE definitions
'
Const TAPE_LOAD = 0
Const TAPE_UNLOAD = 1
Const TAPE_TENSION = 2
Const TAPE_LOCK = 3
Const TAPE_UNLOCK = 4
Const TAPE_FORMAT = 5
Private Type TAPE_PREPARE
    Operation As Long
    Immediate As Boolean
' TAPE_PREPARE, *PTAPE_PREPARE;
End Type
'
' IOCTL_TAPE_WRITE_MARKS definitions
'
Const TAPE_SETMARKS = 0
Const TAPE_FILEMARKS = 1
Const TAPE_SHORT_FILEMARKS = 2
Const TAPE_LONG_FILEMARKS = 3
Private Type TAPE_WRITE_MARKS
    Type As Long
    Count As Long
    Immediate As Boolean
' TAPE_WRITE_MARKS, *PTAPE_WRITE_MARKS;
End Type
'
' IOCTL_TAPE_GET_POSITION definitions
'

Const TAPE_ABSOLUTE_POSITION = 0
Const TAPE_LOGICAL_POSITION = 1
Const TAPE_PSEUDO_LOGICAL_POSITION = 2
Private Type TAPE_GET_POSITION
    Type As Long
    Partition As Long
    offset As Long
' TAPE_GET_POSITION, *PTAPE_GET_POSITION;
End Type
'
' IOCTL_TAPE_SET_POSITION definitions
'
Const TAPE_REWIND = 0&
Const TAPE_ABSOLUTE_BLOCK = 1&
Const TAPE_LOGICAL_BLOCK = 2&
Const TAPE_PSEUDO_LOGICAL_BLOCK = 3&
Const TAPE_SPACE_END_OF_DATA = 4&
Const TAPE_SPACE_RELATIVE_BLOCKS = 5&
Const TAPE_SPACE_FILEMARKS = 6&
Const TAPE_SPACE_SEQUENTIAL_FMKS = 7&
Const TAPE_SPACE_SETMARKS = 8&
Const TAPE_SPACE_SEQUENTIAL_SMKS = 9&
Private Type TAPE_SET_POSITION
    Method As Long
    Partition As Long
    offset As Long
    Immediate As Boolean
' TAPE_SET_POSITION, *PTAPE_SET_POSITION;
End Type
'
' IOCTL_TAPE_GET_DRIVE_PARAMS definitions
'
'
' Definitions for FeaturesLow parameter
'
Const TAPE_DRIVE_FIXED = &H1
Const TAPE_DRIVE_SELECT = &H2
Const TAPE_DRIVE_INITIATOR = &H4
Const TAPE_DRIVE_ERASE_SHORT = &H10
Const TAPE_DRIVE_ERASE_LONG = &H20
Const TAPE_DRIVE_ERASE_BOP_ONLY = &H40
Const TAPE_DRIVE_ERASE_IMMEDIATE = &H80
Const TAPE_DRIVE_TAPE_CAPACITY = &H100
Const TAPE_DRIVE_TAPE_REMAINING = &H200
Const TAPE_DRIVE_FIXED_BLOCK = &H400
Const TAPE_DRIVE_VARIABLE_BLOCK = &H800
Const TAPE_DRIVE_WRITE_PROTECT = &H1000
Const TAPE_DRIVE_EOT_WZ_SIZE = &H2000
Const TAPE_DRIVE_ECC = &H10000
Const TAPE_DRIVE_COMPRESSION = &H20000
Const TAPE_DRIVE_PADDING = &H40000
Const TAPE_DRIVE_REPORT_SMKS = &H80000
Const TAPE_DRIVE_GET_ABSOLUTE_BLK = &H100000
Const TAPE_DRIVE_GET_LOGICAL_BLK = &H200000
Const TAPE_DRIVE_SET_EOT_WZ_SIZE = &H400000
Const TAPE_DRIVE_EJECT_MEDIA = &H1000000
Const TAPE_DRIVE_CLEAN_REQUESTS = &H2000000
Const TAPE_DRIVE_SET_CMP_BOP_ONLY = &H4000000
Const TAPE_DRIVE_RESERVED_BIT = &H80000000
'
' Definitions for FeaturesHigh parameter
'
Const TAPE_DRIVE_LOAD_UNLOAD = &H80000001
Const TAPE_DRIVE_TENSION = &H80000002
Const TAPE_DRIVE_LOCK_UNLOCK = &H80000004
Const TAPE_DRIVE_REWIND_IMMEDIATE = &H80000008
Const TAPE_DRIVE_SET_BLOCK_SIZE = &H80000010
Const TAPE_DRIVE_LOAD_UNLD_IMMED = &H80000020
Const TAPE_DRIVE_TENSION_IMMED = &H80000040
Const TAPE_DRIVE_LOCK_UNLK_IMMED = &H80000080
Const TAPE_DRIVE_SET_ECC = &H80000100
Const TAPE_DRIVE_SET_COMPRESSION = &H80000200
Const TAPE_DRIVE_SET_PADDING = &H80000400
Const TAPE_DRIVE_SET_REPORT_SMKS = &H80000800
Const TAPE_DRIVE_ABSOLUTE_BLK = &H80001000
Const TAPE_DRIVE_ABS_BLK_IMMED = &H80002000
Const TAPE_DRIVE_LOGICAL_BLK = &H80004000
Const TAPE_DRIVE_LOG_BLK_IMMED = &H80008000
Const TAPE_DRIVE_END_OF_DATA = &H80010000
Const TAPE_DRIVE_RELATIVE_BLKS = &H80020000
Const TAPE_DRIVE_FILEMARKS = &H80040000
Const TAPE_DRIVE_SEQUENTIAL_FMKS = &H80080000
Const TAPE_DRIVE_SETMARKS = &H80100000
Const TAPE_DRIVE_SEQUENTIAL_SMKS = &H80200000
Const TAPE_DRIVE_REVERSE_POSITION = &H80400000
Const TAPE_DRIVE_SPACE_IMMEDIATE = &H80800000
Const TAPE_DRIVE_WRITE_SETMARKS = &H81000000
Const TAPE_DRIVE_WRITE_FILEMARKS = &H82000000
Const TAPE_DRIVE_WRITE_SHORT_FMKS = &H84000000
Const TAPE_DRIVE_WRITE_LONG_FMKS = &H88000000
Const TAPE_DRIVE_WRITE_MARK_IMMED = &H90000000
Const TAPE_DRIVE_FORMAT = &HA0000000
Const TAPE_DRIVE_FORMAT_IMMEDIATE = &HC0000000
Const TAPE_DRIVE_HIGH_FEATURES = &H80000000    'mask for high features flag
'
' IOCTL_TAPE_SET_DRIVE_PARAMETERS definitions
'
Private Type TAPE_SET_DRIVE_PARAMETERS
    ECC As Boolean
    Compression As Boolean
    DataPadding As Boolean
    ReportSetmarks As Boolean
    EOTWarningZoneSize As Boolean
' TAPE_SET_DRIVE_PARAMETERS, *PTAPE_SET_DRIVE_PARAMETERS;
End Type
'
' IOCTL_TAPE_SET_MEDIA_PARAMETERS definitions
'
Private Type TAPE_SET_MEDIA_PARAMETERS
    BlockSize As Long
' TAPE_SET_MEDIA_PARAMETERS, *PTAPE_SET_MEDIA_PARAMETERS;
End Type
'
' IOCTL_TAPE_CREATE_PARTITION definitions
'
Const TAPE_FIXED_PARTITIONS = 0&
Const TAPE_SELECT_PARTITIONS = 1&
Const TAPE_INITIATOR_PARTITIONS = 2&
Private Type TAPE_CREATE_PARTITION
    Method As Boolean
    Count As Boolean
    Size As Boolean
' TAPE_CREATE_PARTITION, *PTAPE_CREATE_PARTITION;
End Type

'Error Message Functions
Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
Const SUBLANG_DEFAULT = &H1

Dim TapeDrive As TAPE_GET_DRIVE_PARAMETERS
Dim TapeMedia As TAPE_GET_MEDIA_PARAMETERS
' PROPERTY: StatusMessage
' AUTHOR:   Keith Gibby
' PURPOSE:  Retrieve the current status message
' INPUT:    None
' OUTPUT:   A string containing the current status
' NOTES:
Property Get StatusMessage() As String
    If sStatusMessage = "" Then sStatusMessage = "Tape Library Active"
    StatusMessage = sStatusMessage
End Property
'
'
' FUNCTION: cLargeInt
' AUTHOR:   From Public Domain Source - MSDN
' PURPOSE:  Convert LARGER_INTEGER Type to value
' INPUT:    Low Part and High Part of LARGER_INTEGER
' OUTPUT:   Value
' NOTES:
'
Public Function cLargeInt(LoPart As Long, HiPart As Long) As Double

    Dim dblLo As Double
    Dim dblHi As Double
    
    If LoPart < 0 Then
        dblLo = 2 ^ 32 + LoPart
    Else
        dblLo = LoPart
    End If
    
    If HiPart < 0 Then
        dblHi = 2 ^ 32 + HiPart
    Else
        dblHi = HiPart
    End If
    
    cLargeInt = dblLo + dblHi * 2 ^ 32

End Function
'
'
' FUNCTION: ErrorMessage
' AUTHOR:   Keith Gibby
' PURPOSE:
' INPUT:    ShowMessage : TRUE - Display Message : FALSE - Do Not Display Message
'           ErrorCode   : API Status
' OUTPUT:   Error Message
' NOTES:
'
Public Function ErrorMessage(bShowMessage As Boolean, lErrorCode As Long) As String

    Dim sBuffer As String
    
    sBuffer = Space(200)   'Allocate a string buffer
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrorCode, LANG_NEUTRAL, sBuffer, 200, ByVal 0&     'Format message string
    ErrorMessage = sBuffer
    If lErrorCode = ERROR_NO_DATA_DETECTED Then bAbortRead = True
    sStatusMessage = sBuffer
    If bShowMessage Then MsgBox sStatusMessage, vbOKOnly, "Error Message"
    
End Function
'
'
' FUNCTION: ReadTapeFile
' AUTHOR:   Keith Gibby
' PURPOSE:  Read a file from tape
' INPUT:    None
' OUTPUT:   None
' NOTES:
'
Public Function ReadTapeFile(sFileOut As String) As String

    Const lMaxBytes = 32768
    
    Dim abyData(lMaxBytes) As Byte
    Dim lLastError As Long
    Dim lBytesRead As Long
    Dim lBytesWrote As Long
    Dim bDone As Boolean
    Dim fso, fsoFile, fsoStream
    Dim bBytesSaved As Boolean
    
    ReadTapeFile = ""
    
    SecurityAttributes.bInheritHandle = 0&
    SecurityAttributes.lpSecurityDescriptor = 0&
    SecurityAttributes.nLength = 0&
    
    OpenTape
    
    sStatusMessage = "Reading Tape"
    
    hDisk = CreateFile(sFileOut, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, SecurityAttributes, CREATE_ALWAYS, 0, 0&)
    
    bDone = False
    bBytesSaved = False
    While Not bDone
        Erase abyData
        lLastError = ReadFile(hTape, abyData(1), lMaxBytes, lBytesRead, ByVal 0&)
        lLastError = GetLastError
        
        If lLastError = ERROR_NO_DATA_DETECTED Then
            bAbortRead = True
            sStatusMessage = "End of data found."
            ReadTapeFile = "End of data"
            bDone = True
        End If
        
        If Not bDone Then
            If lBytesRead = 0 Then
                bDone = True
                If bBytesSaved Then
                    bBytesSaved = False
                    CloseHandle (hDisk)
                    ReadTapeFile = "No Error"
                Else
                    sStatusMessage = "No data written to file."
                    ReadTapeFile = "Error"
                    bAbortRead = True
                End If
            Else
                bBytesSaved = True
                
                '###########################################################Filter Function Start###############################################################
                '
                '   Pseudo Code
                '       Given MessageToFilter
                '       Look in abyData array for MessageToFilter occurences
                '       When MessageToFilter is found determine how many bytes are used for it
                '
                '       Modify abyData to exclude MessageToFilter and its data and change lBytesRead as required
                '       ___OR___
                '       Overwrite each byte with dummy value
                '
                '       Repeat Until MessageToFilter is not in abyData
                '###########################################################Filter Function End#################################################################
                
                lLastError = WriteFile(hDisk, abyData(1), lBytesRead, lBytesWrote, ByVal 0&)        'Write the data to disk
                
                DoEvents
                If lLastError = 0 Then
                    lLastError = GetLastError
                End If
            End If
        End If
        DoEvents
    Wend
    
    CloseHandle (hTape)
    CloseHandle (hDisk)
    
    If Not ReadTapeFile = "No Error" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(sFileOut) Then
           Set fsoFile = fso.GetFile(sFileOut)
            fsoFile.Delete
        End If
    End If

End Function 'ReadTapeFile
'
' FUNCTION: TapeControl
' AUTHOR:   Keith Gibby
' PURPOSE:  Basic Tape Control
' INPUT:    iControlCode - Action to take
'           lSpaceCount -  Moves the tape forward (or backward) to the first occurrence of lSpaceCount filemarks in the current partition
' OUTPUT:   Status of function
' NOTES:
'
Public Function TapeControl(iControlCode As Integer, lSpaceCount As Long) As Long
  
    Dim lSpace As Long
    Dim lPositionMethod As Long
    
    
    lSpace = lSpaceCount
    
    Select Case iControlCode
        Case 0
         lPositionMethod = TAPE_REWIND
        Case 1
         lPositionMethod = TAPE_SPACE_FILEMARKS
        Case 3
         lPositionMethod = TAPE_SPACE_END_OF_DATA
        Case 4
         lPositionMethod = TAPE_SPACE_RELATIVE_BLOCKS
    End Select
    
    lStatus = SetTapePosition(hTape, lPositionMethod, 1, lSpace, 0, False)
    
    If lStatus <> 0 Then
     sStatusMessage = ErrorMessage(False, lStatus)
    End If
    
    TapeControl = lStatus

End Function
'
'
' ROUTINE:  OpenTape
' AUTHOR:   Keith Gibby
' PURPOSE:  Create Tape Handle & Set tape Parameters
' INPUT:    None
' OUTPUT:   None
' NOTES:
'
Public Sub OpenTape()

    Dim TapeName As String
   
    sStatusMessage = "Opening Tape"
    
    TapeName = "\\.\Tape0"
    hTape = CreateFile(TapeName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, SecurityAttributes, OPEN_EXISTING, 0, 0&)
    lStatus = SetTapeParameters(hTape, SET_TAPE_MEDIA_INFORMATION, 0) ' variable block length!
    If (hTape = INVALID_HANDLE_VALUE) Then
     sStatusMessage = ErrorMessage(False, lStatus)
     CloseHandle (hTape)
    End If
    
    DoEvents
    
End Sub
'
'
' FUNCTION: InitializeTape
' AUTHOR:   Keith Gibby
' PURPOSE:  Check the tape for write protection status
' INPUT:    None
' OUTPUT:   Status of Write Protection - TRUE: Write Protection Enabled - FALSE: Write Protection NOT Enabled
' NOTES:
'
Public Function InitializeTape() As Boolean

Dim dblAvailable As Double
Dim dblFree As Double
Dim bWriteProtect As Boolean
Dim lVarSize
On Error GoTo Jail

    sStatusMessage = "Initializing Tape"
    InitializeTape = False
    bWriteProtect = False
    
    bAbortRead = False
    OpenTape
    lStatus = GetTapeStatus(hTape)
    If lStatus = 1112 Then
      InitializeTape = True
      MsgBox "No Media In Drive, Insert The Tape"
      CloseHandle (hTape)
      Exit Function
    End If
    
    lVarSize = Len(TapeMedia)

    lStatus = GetTapeParameters(hTape, GET_TAPE_MEDIA_INFORMATION, lVarSize, TapeMedia)
 'MsgBox cLargeInt(TapeMedia.Capacity.LowPart, TapeMedia.Capacity.HighPart) & vbNewLine & _
        cLargeInt(TapeMedia.Remaining.LowPart, TapeMedia.Remaining.HighPart) & vbNewLine & _
        TapeMedia.BlockSize & " " & TapeMedia.PartitionCount & vbNewLine & _
        TapeMedia.WriteProtected & "   " & TapeMedia.PadBytes1 & " " & TapeMedia.PadBytes2
 
    If (TapeMedia.WriteProtected = 1) Then
     sStatusMessage = "Write Protect ON"
    Else
     sStatusMessage = "Write Protect OFF"
    End If

    InitializeTape = TapeMedia.WriteProtected
    CloseHandle (hTape)
    Exit Function
    
Jail:
    MsgBox Err.Number & " " & Err.Description
    CloseHandle (hTape)
    
End Function
'
'
' ROUTINE:  EjectTape
' AUTHOR:   Keith Gibby
' PURPOSE:  Eject The tape
' INPUT:    None
' OUTPUT:   None
' NOTES:
'
Public Sub EjectTape()
On Error GoTo ErrorHandler
    OpenTape
    sStatusMessage = "Preparing to eject tape, please wait"
    lStatus = PrepareTape(hTape, TAPE_UNLOAD, False)
    If (lStatus <> 0) Then
       sStatusMessage = ErrorMessage(False, lStatus)
       If lStatus = 1110 Then    'Media Change Warning - Try to eject again
           lStatus = PrepareTape(hTape, TAPE_UNLOAD, False)
       End If
    Else
       sStatusMessage = "Tape Ejected"
    End If

    CloseHandle (hTape)

    Exit Sub
    
ErrorHandler:
    CloseHandle (hTape)
    MsgBox Err.Number & "-" & Err.Description
End Sub
'
'
' ROUTINE:  Rewind The tape
' AUTHOR:   Keith Gibby
' PURPOSE:  Rewind The tape
' INPUT:    None
' OUTPUT:   None
' NOTES:
'
Public Sub RewindTape()

    OpenTape
    sStatusMessage = "Rewinding Tape"
    
    lStatus = TapeControl(0, 0)
    
    If lStatus <> 0 Then
        sStatusMessage = "Rewind Error"
    Else
        sStatusMessage = "Tape Rewound"
    End If
    
    CloseHandle (hTape)
    
End Sub
'
'
' ROUTINE:  ScanTape
' AUTHOR:   Keith Gibby
' PURPOSE:  Dump Tape Files To Disk
' INPUT:    sBaseFileName - Base Name Of Files To Create
' OUTPUT:   Files From tape
' NOTES:    sBaseFileName is appended with digit incremented for each file found
'
Public Sub ScanTape(sBaseFileName As String)  'Scan Tape START
    Static FileSum As Long
    Dim sOutfileName As String
    
    While Not bAbortRead
        FileSum = FileSum + 1
        sOutfileName = Replace(sBaseFileName, ".", FileSum & ".", 1)
        ReadTapeFile sOutfileName
    Wend
    RewindTape
    FileSum = 0

End Sub    'Scan Tape END
'
'
' ROUTINE:  SeekEndOfFile
' AUTHOR:   Keith Gibby
' PURPOSE:  Find End Of File Mark
' INPUT:    None
' OUTPUT:   None
' NOTES:
'
Public Sub SeekEndOfFile()   'Seek EOF Mark START

    OpenTape
    
    sStatusMessage = "Seeking End Of File Mark"
    
    lStatus = TapeControl(1, 1)
    If (lStatus <> NO_ERROR) Then
       sStatusMessage = ErrorMessage(False, lStatus)
    Else
       sStatusMessage = "End Of File Mark"
    End If
    
    CloseHandle (hTape)
    
End Sub  'Seek EOF Mark END
'
'
' ROUTINE:  EndOfData
' AUTHOR:   Keith Gibby
' PURPOSE:  Find End Of Data
' INPUT:    None
' OUTPUT:   None
' NOTES:
'
Public Sub EndOfData()
    
    OpenTape
    sStatusMessage = "Seeking End Of Data"
    DoEvents
    lStatus = TapeControl(3, 0)
    
    If (lStatus <> NO_ERROR) Then
       sStatusMessage = ErrorMessage(False, lStatus)
    Else
       sStatusMessage = "End Of Data Reached"
       bAbortRead = True
    End If
    
    CloseHandle (hTape)

End Sub

'###########################Development###############################
'
'
' FUNCTION: ReadDiskTapeFile
' AUTHOR:   Keith Gibby
' PURPOSE:
' INPUT:    sFileOut Name Of File To Create
' OUTPUT:   File
' NOTES:    This function mirrors the Readtape Function except it reads from a local file
'           on the hard disk.  This will allow development of the filter function without
'           having the tape drive available.
'
'           Disk Files can be read by individual bytes; however, tapes must be read in blocks
'
'
Public Function ReadDiskTapeFile(sFileOut As String) As String
    
    Const lMaxBytes = 32768

    Const DISK_INPUT_FILE = "DiskFile_IN.ZIP"   ' Added to simulate tape file
    
    Dim abyData(lMaxBytes) As Byte
    Dim lLastError As Long
    Dim lBytesRead As Long
    Dim lBytesWrote As Long
    Dim bDone As Boolean
    Dim fso, fsoFile, fsoStream
    Dim bBytesSaved As Boolean
    
    ReadDiskTapeFile = ""
    
    SecurityAttributes.bInheritHandle = 0&
    SecurityAttributes.lpSecurityDescriptor = 0&
    SecurityAttributes.nLength = 0&

    Dim DiskTapeName As String
    Dim hDiskTape As Long
    
    sStatusMessage = "Opening DiskTape"
    DiskTapeName = App.Path & "\" & DISK_INPUT_FILE
    hDiskTape = CreateFile(DiskTapeName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, SecurityAttributes, OPEN_EXISTING, 0, 0&)
    If (hDiskTape = INVALID_HANDLE_VALUE) Then
     sStatusMessage = ErrorMessage(False, lStatus)
     CloseHandle (hDiskTape)
    End If
    
    sStatusMessage = "Reading DiskTape"
    
    hDisk = CreateFile(sFileOut, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, SecurityAttributes, CREATE_ALWAYS, 0, 0&)
    
    bDone = False
    bBytesSaved = False
    While Not bDone
        Erase abyData
        lLastError = ReadFile(hDiskTape, abyData(1), lMaxBytes, lBytesRead, ByVal 0&)
        lLastError = GetLastError
              
        If lLastError = ERROR_NO_DATA_DETECTED Then
            bAbortRead = True
            sStatusMessage = "End of data found."
            ReadDiskTapeFile = "End of data"
            bDone = True
        End If
        
        If Not bDone Then
            If lBytesRead = 0 Then
                bDone = True
                If bBytesSaved Then
                    bBytesSaved = False
                    CloseHandle (hDisk)
                    ReadDiskTapeFile = "No Error"
                Else
                    sStatusMessage = "No data written to file."
                    ReadDiskTapeFile = "Error"
                    bAbortRead = True
                End If
            Else
                bBytesSaved = True
                
                
'###########################################################Filter Function Start###############################################################
'
'   Pseudo Code
'       Given MessageToFilter
'       Look in abyData array for MessageToFilter occurences
'       When MessageToFilter is found determine how many bytes are used for it
'
'       Modify abyData to exclude MessageToFilter and its data and change lBytesRead as required
'       ___OR___
'       Overwrite each byte with dummy value
'
'       Repeat Until MessageToFilter is not in abyData
'###########################################################Filter Function End#################################################################
                
                
                lLastError = WriteFile(hDisk, abyData(1), lBytesRead, lBytesWrote, ByVal 0&)        'Write the data to disk
                sStatusMessage = "Finished Reading DiskTape"
                If lLastError = 0 Then
                    lLastError = GetLastError
                End If
            End If
        End If
        DoEvents
    Wend
    
    CloseHandle (hDiskTape)
    CloseHandle (hDisk)
    
    If Not ReadDiskTapeFile = "No Error" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(sFileOut) Then
           Set fsoFile = fso.GetFile(sFileOut)
            fsoFile.Delete
        End If
    End If

End Function

'###########################Development###############################

