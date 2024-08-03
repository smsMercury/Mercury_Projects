Attribute VB_Name = "basCCAT"
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
' MODULE:   basCCAT
' AUTHOR:   Tom Elkins
' PURPOSE:  Translator related routines and functions
' REVISION:
'   v1.3    TAE Added code for INI files
'   v1.4    TAE Added code to look for INI files in application directory
'           TAE Added code to asks user to find INI file if not in either directory
'           TAE Added code to create a default INI file if user cannot find
'           TAE Updated INI file creation contents to Dec 99 version
'           TAE Modified DAS Header record to split SQL query into multiple lines
'           TAE Added error checks in session database list
'           TAE Updated help file constants
'   v1.5    TAE Added a global flag to terminate translation process prematurely
'           TAE Upgraded help system to use the new HTML help file
'           TAE Added a function that extracts help context IDs from the INI file
'           TAE Added a routine to write to the log file
'           TAE Added a function that parses a human-readable time string and converts it to TSecs
'           TAE Modified INI creation routine to write the new INI file
'   v1.6    TAE Changed the timestamp in the WriteLogEntry routine to include the date
'           TAE Added a method to write to the Token (INI) file
'           TAE Updated the contents of the default INI file
'           TAE Added a method to populate a supplied combo box with CCOS versions read from the INI file
'           TAE Added a method to populate a supplied list view control with the list of CCOS messages from the INI file
'           TAE Updated the Event file output to match the current format
'           TAE Removed the Help ID map from the INI file and included the constants locally
'   v1.6.1  TAE Added verbose logging calls
'
Option Explicit     ' Forces variables to be declared before they can be used
'
' Declarations for the Windows API to use INI files instead of token files
Declare Function lGetINIString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal Section As String, ByVal KeyName As String, ByVal Default As String, ByVal ReturnedString As String, ByVal Size As Long, ByVal INIFileName As String) As Long
Declare Function lGetININum Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal Section As String, ByVal KeyName As String, ByVal Default As Long, ByVal INIFileName As String) As Long
Declare Function lPutINIString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Section As String, ByVal KeyName As Any, ByVal Default As Any, ByVal INIFileName As String) As Long
Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'+v1.5
Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Global Const HH_HELP_CONTEXT = &HF
Global Const HH_HELP_TOPIC = &H0
'Global Const HH_SET_WIN_TYPE = &H4
'Global Const HH_GET_WIN_TYPE = &H5
'Global Const HH_GET_WIN_HANDLE = &H6
'Global Const HH_DISPLAY_TEXT_POPUP = &HE        ' Display string resource ID or text in a pop-up window.
'Global Const HH_TP_HELP_CONTEXTMENU = &H10      ' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU.
'Global Const HH_TP_HELP_WM_HELP = &H11          ' text pop-up help
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'-v1.5
'
'+v1.6TE
' Help Topics
Global Const IDH_CCAT_EXPORT = 301
Global Const IDH_GUI_ABOUT = 100
Global Const IDH_GUI_ABOUT_CLASS = 101
Global Const IDH_GUI_ABOUT_DESC = 102
Global Const IDH_GUI_ABOUT_OK = 103
Global Const IDH_GUI_ABOUT_SYSINFO = 105
Global Const IDH_GUI_ABOUT_TITLE = 104
Global Const IDH_GUI_ABOUT_VERSION = 106
Global Const IDH_GUI_ABOUT_WARNING = 107
Global Const IDH_GUI_ARCHIVE_PROPERTIES = 250
Global Const IDH_GUI_WIZARD_MESSAGES_ALL = 232
Global Const IDH_GUI_WIZARD_BACK = 297
Global Const IDH_GUI_WIZARD_CANCEL = 296
Global Const IDH_GUI_WIZARD_INFORMATION_CLASS = 224
Global Const IDH_GUI_WIZARD_SOURCE_SOURCE = 211
Global Const IDH_GUI_WIZARD_INFORMATION_DATE = 222
Global Const IDH_GUI_WIZARD_SOURCE_FILE = 213
Global Const IDH_GUI_WIZARD_FINAL = 204
Global Const IDH_GUI_WIZARD_FINISH = 299
Global Const IDH_GUI_WIZARD_HELP = 295
Global Const IDH_GUI_WIZARD_INTRO_HIDE = 209
Global Const IDH_GUI_WIZARD_INFORMATION = 202
Global Const IDH_GUI_WIZARD_INTRO = 200
Global Const IDH_GUI_WIZARD_MESSAGES_INVERT = 234
Global Const IDH_GUI_WIZARD_MESSAGES_LIST = 231
Global Const IDH_GUI_WIZARD_MESSAGES = 203
Global Const IDH_GUI_WIZARD_INFORMATION_NAME = 221
Global Const IDH_GUI_WIZARD_NEXT = 298
Global Const IDH_GUI_WIZARD_MESSAGES_NONE = 233
Global Const IDH_GUI_WIZARD_FINAL_PROGRESS = 242
Global Const IDH_GUI_WIZARD_SOURCE = 201
Global Const IDH_GUI_WIZARD_SOURCE_TAPE = 212
Global Const IDH_GUI_WIZARD_TASKS = 241
Global Const IDH_GUI_WIZARD_INFORMATION_VERSION = 223
Global Const IDH_DAS = 300
Global Const IDH_DAS_EVENT = 305
Global Const IDH_DAS_MTF = 303
Global Const IDH_DAS_SIGNAL = 302
Global Const IDH_DAS_STF = 304
Global Const IDH_GUI_FILTER = 400
Global Const IDH_GUI_FILTER_FIELDS = 410
Global Const IDH_GUI_FILTER_FILTER = 430
Global Const IDH_GUI_FILTER_SORT = 450
Global Const IDH_GUI_FILTER_ASSISTANT = 401
Global Const IDH_GUI_FILTER_CLEAR = 491
Global Const IDH_GUI_FILTER_CLOSE = 490
Global Const IDH_GUI_FILTER_EXECUTE = 493
Global Const IDH_GUI_FILTER_FIELD_LIST = 402
Global Const IDH_GUI_FILTER_FIELDS_ACCEPT = 422
Global Const IDH_GUI_FILTER_FIELDS_DOT = 416
Global Const IDH_GUI_FILTER_FIELDS_EVENT = 413
Global Const IDH_GUI_FILTER_FIELDS_GEO = 419
Global Const IDH_GUI_FILTER_FIELDS_MTF = 414
Global Const IDH_GUI_FILTER_FIELDS_PREDEFINED = 411
Global Const IDH_GUI_FILTER_FIELDS_SELECT = 421
Global Const IDH_GUI_FILTER_FIELDS_SIGNAL = 412
Global Const IDH_GUI_FILTER_FIELDS_STF = 415
Global Const IDH_GUI_FILTER_FIELDS_TRK = 417
Global Const IDH_GUI_FILTER_FIELDS_USER = 420
Global Const IDH_GUI_FILTER_FIELDS_VEC = 418
Global Const IDH_GUI_FILTER_FILTER_LIST = 403
Global Const IDH_GUI_FILTER_FILTERS_ACCEPT = 436
Global Const IDH_GUI_FILTER_FILTERS_AND = 431
Global Const IDH_GUI_FILTER_FILTERS_FIELD = 433
Global Const IDH_GUI_FILTER_FILTERS_OPERATORS = 434
Global Const IDH_GUI_FILTER_FILTERS_OR = 432
Global Const IDH_GUI_FILTER_FILTERS_VALUE = 435
Global Const IDH_GUI_FILTER_MANUAL = 405
Global Const IDH_GUI_FILTER_SAVE = 492
Global Const IDH_GUI_FILTER_SORT_ACCEPT = 453
Global Const IDH_GUI_FILTER_SORT_BY = 451
Global Const IDH_GUI_FILTER_SORT_FIELD = 452
Global Const IDH_GUI_FILTER_SORT_LIST = 404
Global Const IDH_GUI_FILTER_SQL = 406
Global Const IDH_DB = 500
Global Const IDH_DB_ARCHIVE = 502
Global Const IDH_DB_DATA = 504
Global Const IDH_DB_FILTERING = 505
Global Const IDH_DB_INFO = 501
Global Const IDH_DB_SUMMARY = 503
Global Const IDH_GUI_DBINFO = 600
Global Const IDH_GUI_DBINFO_ACCEPT = 605
Global Const IDH_GUI_DBINFO_CANCEL = 606
Global Const IDH_GUI_DBINFO_CLASS = 604
Global Const IDH_GUI_DBINFO_DATE = 602
Global Const IDH_GUI_DBINFO_DESCRIPTION = 603
Global Const IDH_GUI_DBINFO_NAME = 601
Global Const IDH_ERR_3026 = 701
Global Const IDH_ERR_53 = 702
Global Const IDH_GUI_CLOSE = 804
Global Const IDH_GUI_CONVERT = 858
Global Const IDH_GUI_DATA_COL = 873
Global Const IDH_GUI_GRID = 811
Global Const IDH_GUI_DATA_HIDE = 872
Global Const IDH_GUI_DATA_LAST = 871
Global Const IDH_GUI_DATA_MENU = 870
Global Const IDH_GUI_DATA_VALUE_MENU = 877
Global Const IDH_GUI_DATA_ROW = 876
Global Const IDH_GUI_DATA_SORT = 874
Global Const IDH_GUI_DATA_VALUE = 875
Global Const IDH_GUI_EDIT_ADD = 831
Global Const IDH_GUI_EDIT_FILTER = 833
Global Const IDH_GUI_EDIT = 830
Global Const IDH_GUI_EDIT_PROPERTIES = 832
Global Const IDH_GUI_FILE_DELETE = 824
Global Const IDH_GUI_FILE = 820
Global Const IDH_GUI_FILE_NEW = 822
Global Const IDH_GUI_FILE_OPEN = 821
Global Const IDH_GUI_FILE_REMOVE = 825
Global Const IDH_GUI_FILE_SAVE = 823
Global Const IDH_GUI_HELP_CONTENTS = 861
Global Const IDH_GUI_HELP = 860
Global Const IDH_GUI_ICON = 801
Global Const IDH_GUI_VIEW_STATUS = 842
Global Const IDH_GUI_LISTVIEW = 806
Global Const IDH_GUI_MAIN = 800
Global Const IDH_GUI_MAXIMIZE = 803
Global Const IDH_GUI_MINIMIZE = 802
Global Const IDH_GUI_MODES = 812
Global Const IDH_GUI_RESIZE = 810
Global Const IDH_GUI_SECURITY = 813
Global Const IDH_GUI_SPLASH = 814
Global Const IDH_GUI_SPLITTER = 815
Global Const IDH_GUI_STATUS = 807
Global Const IDH_GUI_DATE = 809
Global Const IDH_GUI_SECURITY_PANEL = 808
Global Const IDH_GUI_TOOLS_DEGREE = 852
Global Const IDH_GUI_TOOLS_INI = 853
Global Const IDH_GUI_TOOLS = 850
Global Const IDH_GUI_TOOLS_REMAP = 854
Global Const IDH_GUI_TOOLS_SAVE = 855
Global Const IDH_GUI_TOOLS_SQL = 856
Global Const IDH_GUI_TOOLS_TIME = 851
Global Const IDH_GUI_TOOLS_UPDATE = 857
Global Const IDH_GUI_TREE_ARCHIVE = 883
Global Const IDH_GUI_TREE_DATABASE = 882
Global Const IDH_GUI_TREE_MESSAGE = 885
Global Const IDH_GUI_TREE_QUERY = 884
Global Const IDH_GUI_TREE_SESSION = 881
Global Const IDH_GUI_TREE = 805
Global Const IDH_GUI_VIEW_ARRANGE = 847
Global Const IDH_GUI_VIEW_DETAILS = 846
Global Const IDH_GUI_VIEW_LARGE = 843
Global Const IDH_GUI_VIEW_LIST = 845
Global Const IDH_GUI_VIEW = 840
Global Const IDH_GUI_VIEW_REFRESH = 848
Global Const IDH_GUI_VIEW_SMALL = 844
Global Const IDH_GUI_VIEW_TOOLBAR = 841
Global Const IDH_GUI_MESSAGE = 900
Global Const IDH_GUI_MESSAGE_DESCRIPTION = 903
Global Const IDH_GUI_MESSAGE_ID = 901
Global Const IDH_GUI_MESSAGE_NAME = 902
Global Const IDH_GUI_MESSAGE_OK = 904
Global Const IDH_TOKEN_OTHER = 1009
Global Const IDH_TOKEN = 1000
Global Const IDH_TOKEN_CLASS = 1001
Global Const IDH_TOKEN_EOB = 1008
Global Const IDH_TOKEN_EXPORT = 1011
Global Const IDH_TOKEN_FIELDS = 1003
Global Const IDH_TOKEN_IFF = 1006
Global Const IDH_TOKEN_MESSAGE = 1002
Global Const IDH_TOKEN_MISC = 1010
Global Const IDH_TOKEN_RCV = 1005
Global Const IDH_TOKEN_SIGNAL = 1007
Global Const IDH_TOKEN_SQL = 1004
Global Const IDH_FILTERING = 1
Global Const IDH_TRANSLATE_MSG_HDR = 1102
Global Const IDH_TRANSLATE = 1100
Global Const IDH_TRANSLATE_MTANAALARM = 1106
Global Const IDH_TRANSLATE_MTANARSLT = 1113
Global Const IDH_TRANSLATE_MTDEFPMA = 1103
Global Const IDH_TRANSLATE_MTDFALARM = 1110
Global Const IDH_TRANSLATE_MTDFSDALARM = 1111
Global Const IDH_TRANSLATE_MTFIXRSLT = 1117
Global Const IDH_TRANSLATE_MTHBACTREP = 1109
Global Const IDH_TRANSLATE_MTHBDYNRSP = 1107
Global Const IDH_TRANSLATE_MTHBGSREP = 1119
Global Const IDH_TRANSLATE_MTHBLOBUPD = 1115
Global Const IDH_TRANSLATE_MTHBSELASK = 1121
Global Const IDH_TRANSLATE_MTHBSELJAM = 1124
Global Const IDH_TRANSLATE_MTHBSEMISTAT = 1120
Global Const IDH_TRANSLATE_MTHBSIGUPD = 1118
Global Const IDH_TRANSLATE_MTHBXMTRSTAT = 1125
Global Const IDH_TRANSLATE_MTJAMSTAT = 1122
Global Const IDH_TRANSLATE_MTLOBSETRSLT = 1116
Global Const IDH_TRANSLATE_MTLOBUPD = 1112
Global Const IDH_TRANSLATE_MTNAVREP = 1104
Global Const IDH_TRANSLATE_MTRUNMODE = 1123
Global Const IDH_TRANSLATE_MTSDALARM = 1108
Global Const IDH_TRANSLATE_MTSIGALARM = 1105
Global Const IDH_TRANSLATE_MTSIGUPD = 1114
Global Const IDH_TRANSLATE_MTSSERROR = 1126
Global Const IDH_TRANSLATE_TIME_HDR = 1101
'-v1.6
'
' Constants
Global Const DAS_TOKEN_PATH = "\Token Files\"       ' Path to the token files
Global Const DAS_HELP_PATH = "\Help Files\"         ' Path to the help files
'+v1.5
'Global Const CCAT_HELP_FILE = "ccat.hlp"  ' Name of the CCAT help file
Global Const CCAT_HELP_FILE = "CCAT.chm"  ' Name of the CCAT help file
'-v1.5
Global Const CCAT_INI_FILE = "ccat.ini"   ' INI file
'
' Constants used for interface mode and node/item types
Global Const gsSESSION = "Session"
Global Const gsDATABASE = "Database"
Global Const gsARCHIVE = "Archive"
Global Const gsMESSAGE = "Message"
Global Const gsTOCMSG = "TOCMSG"
Global Const gsQUERY = "Query"
Global Const gsDATA = "Data"
'
' DAS File types
Global Const giDAS_SIG = 1      ' DAS Signal activity file
Global Const giDAS_MTF = 2      ' DAS Moving target file
Global Const giDAS_STF = 3      ' DAS Stationary target file
Global Const giDAS_EVT = 4      ' DAS Event file
Global Const giUSR_TXT = 5      ' User-defined text file
'
' DAS Record types
Global Const giREC_DOT = 0      ' Dot - detection
Global Const giREC_TRK = 1      ' Track
Global Const giREC_VEC = 2      ' Vector - Line of bearing
Global Const giREC_GEO = 3      ' Geolocation - fix
Global Const giREC_SIG = 4      ' Signal activity
Global Const giREC_EVT = 5      ' Event
'
' Export file information
Public Type EXPORT_INFO
    sFile As String             ' File name
    iFile_Type As Integer       ' DAS File type
    sFields As String           ' Field list
    iRec_Type As Integer        ' Record type
    iRec_Src As Integer         ' Source of records
    sSQL As String              ' SQL query
End Type
'
' GUI Settings and information
Public Type GUI_INFO
    sMode As String             ' Current mode
    sNode As String             ' Currently selected node
    sItem As String             ' Currently selected item
    bTreeView As Boolean        ' True if the TreeView has focus
    fMouse_X As Single          ' Mouse X position
    fMouse_Y As Single          ' Mouse Y position
    bRight_Button As Boolean    ' True if the right button was pressed
    lInterval As Long           ' Translation interrupt interval
End Type
'
' Archive properties
Public Type ARCHIVE_FILE
    sFile As String             ' Archive file name
    iType As Integer            ' Archive file type
    sDate As String             ' Archive start date
    sClass As String            ' Classification level
End Type
'
' Help topics mapping
'v1.5 Moved to CCAT.INI file
'
' Global variables
Public guExport As EXPORT_INFO          ' Keeps the current export file info
Public gaDAS_Rec_Type(6) As String      ' Array of DAS record types
Public giLog_File As Integer            ' Log file
Public guGUI As GUI_INFO                ' Current GUI info
Public guArchive As ARCHIVE_FILE        ' Current Archive information
Public gsCCAT_INI_Path As String        ' Stores the location of the INI file
'+v1.5
Public gbProcessing As Boolean          ' TRUE if currently processing an archive
'-v1.5
'
'+v1.6.1TE
Private mblnVerbose As Boolean          ' TRUE if verbose reporting is on
'-v1.6.1
'
' ROUTINE:  Initialize_Translator
' AUTHOR:   Tom Elkins
' PURPOSE:  Initialize the properties and procedures for the translator, including the
'           Token file
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Sub Initialize_Translator()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Initialize_Translator (Start)"
    '-v1.6.1
    '
    ' Look at the default path for the INI file
    gsCCAT_INI_Path = App.Path & DAS_TOKEN_PATH & CCAT_INI_FILE
    '
    ' Check for the existence of the INI file
    If Dir(gsCCAT_INI_Path) = "" Then
        '
        ' It doesn't exist, so
        ' look at the secondary path
        gsCCAT_INI_Path = App.Path & "\" & CCAT_INI_FILE
        '
        ' Check for the existence of the INI file
        If Dir(gsCCAT_INI_Path) = "" Then
            '
            ' It doesn't exist, so
            ' ask the user to find the file
            With frmMain.dlgCommonDialog
                '
                .CancelError = False
                .DialogTitle = "Find CCAT INI file"
                .FileName = ""
                .Filter = "CCAT settings file|" & CCAT_INI_FILE
                .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
                .InitDir = App.Path
                .ShowOpen
                '
                ' See if user found the file
                If .FileName <> "" Then
                    '
                    ' Copy the file to an appropriate place
                    FileCopy .FileName, gsCCAT_INI_Path
                Else
                    '
                    ' Create a default version of the file
                    Create_CCAT_Token_File
                End If
            End With
        End If
    End If
    '
    ' Update the classification level
    frmSecurity.SetClassification basCCAT.GetNumber("Classification", CCAT_INI_FILE & "_CLASS", 0), "CCAT Token File"
    '
    ' Initialize GUI structure
    guGUI.sItem = ""
    guGUI.sMode = ""
    guGUI.sNode = ""
    guGUI.lInterval = basCCAT.GetNumber("Miscellaneous Operations", "UPDATE_INTERVAL", 1250)
    guGUI.bRight_Button = False
    guGUI.bTreeView = True
    '
    ' Initialize DAS Record Type Array
    gaDAS_Rec_Type(giREC_DOT) = "DOT"
    gaDAS_Rec_Type(giREC_TRK) = "TRK"
    gaDAS_Rec_Type(giREC_VEC) = "VEC"
    gaDAS_Rec_Type(giREC_GEO) = "GEO"
    gaDAS_Rec_Type(giREC_SIG) = "SIG"
    gaDAS_Rec_Type(giREC_EVT) = "EVT"
    '
    ' Initialize the database structure
    guCurrent.iArchive = 0
    guCurrent.iMessage = 0
    guCurrent.sMessage = ""
    guCurrent.sName = ""
    guCurrent.uSQL.sFields = "*"
    guCurrent.uSQL.sFilter = ""
    guCurrent.uSQL.sOrder = ""
    guCurrent.uSQL.sTable = ""
    '
    '+v1.5
    ' Check interface features
    frmMain.mnuToolsExecuteSQL.Visible = (basCCAT.GetNumber("Miscellaneous Operations", "ADVSQL", 0) = 1)
    basDatabase.Interactive = True
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Initialize_Translator (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: sGet_Version
' AUTHOR:   Tom Elkins
' PURPOSE:  Returns a string stating the current application version
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Function sGet_Version()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basCCAT.sGet_Version (Start)"
    '-v1.6.1
    '
    ' Construct the version string
    sGet_Version = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basCCAT.sGet_Version (End)"
    '-v1.6.1
    '
End Function
'
' ROUTINE:  Create_CCAT_Token_File
' AUTHOR:   Tom Elkins
' PURPOSE:  Creates a default token file
' INPUT:    None
' OUTPUT:   None
' NOTES:    If the executable is moved to a new directory but the token files are not,
'           this routine will create a token file to use.
Public Sub Create_CCAT_Token_File()
    Dim iFile As Integer    ' File identifier
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Create_CCAT_Token_File (Start)"
    '-v1.6.1
    '
    ' Log the event
    'basCCAT.WriteLogEntry "CCAT: Create_CCAT_Token_File"
    '
    ' Assign a file identifier
    iFile = FreeFile
    '
    ' Create the token file
    Open gsCCAT_INI_Path For Output As iFile
    '
    ' Write the contents
    '
    '+v1.6TE
    Print #iFile, "; Compass Call Archive Translator - Tokens file - v" & sGet_Version
    Print #iFile, "; FILE GENERATED BY CCAT"
    '-v1.6
    Print #iFile, "; These tokens configure the operation of the translator"
    Print #iFile, "; DO NOT MODIFY unless you are confident of altering these files."
    Print #iFile, ";"
    Print #iFile, "[File]"
    Print #iFile, "; File managment tokens"
    Print #iFile, ";<filename>_FILE=<1=loaded>"
    Print #iFile, CCAT_INI_FILE & "_FILE=1"
    Print #iFile, ";"
    Print #iFile, "[Classification]"
    Print #iFile, ";<filename>_CLASS=<classification ID>"
    Print #iFile, ";   0=UNCLASSIFIED"
    Print #iFile, ";   +1=SAR"
    Print #iFile, ";   +2=SCI"
    Print #iFile, ";   +4=CONFIDENTIAL"
    Print #iFile, ";   +8=SECRET"
    Print #iFile, ";   +16=TOP SECRET"
    Print #iFile, CCAT_INI_FILE & "_CLASS=0 ;UNCLASSIFIED"
    Print #iFile, ";"
    Print #iFile, "[Message List]"
    Print #iFile, "; Message List"
    Print #iFile, ";<Message token>=<message name>"
    '
    '+v1.8.12
    Print #iFile, "CC_MESSAGES=33"
    '-v1.8.12
    Print #iFile, "CC_MSG1=MTDEFPMA"
    Print #iFile, "CC_MSG2=MTSIGALARM"
    Print #iFile, "CC_MSG3=MTANAALARM"
    Print #iFile, "CC_MSG4=MTHBDYNRSP"
    Print #iFile, "CC_MSG5=MTSDALARM"
    Print #iFile, "CC_MSG6=MTHBACTREP"
    Print #iFile, "CC_MSG7=MTDFALARM"
    Print #iFile, "CC_MSG8=MTRFSTATUS"
    Print #iFile, "CC_MSG9=MTDFSDALARM"
    Print #iFile, "CC_MSG10=MTLOBUPD"
    Print #iFile, "CC_MSG11=MTANARSLT"
    Print #iFile, "CC_MSG12=MTSIGUPD"
    Print #iFile, "CC_MSG13=MTHBLOBUPD"
    Print #iFile, "CC_MSG14=MTULDDATA"
    Print #iFile, "CC_MSG15=MTLOBSETRSLT"
    Print #iFile, "CC_MSG16=MTFIXRSLT"
    Print #iFile, "CC_MSG17=MTHBSIGUPD"
    Print #iFile, "CC_MSG18=MTHBGSREP"
    Print #iFile, "CC_MSG19=MTHBSEMISTAT"
    Print #iFile, "CC_MSG20=MTHBSELASK"
    Print #iFile, "CC_MSG21=MTJAMSTAT"
    Print #iFile, "CC_MSG22=MTRUNMODE"
    Print #iFile, "CC_MSG23=MTHBSELJAM"
    Print #iFile, "CC_MSG24=MTHBXMTRSTAT"
    Print #iFile, "CC_MSG25=MTNAVREP"
    Print #iFile, "CC_MSG26=MTLHECHMOD"
    Print #iFile, "CC_MSG27=MTLHCORRELATE"
    Print #iFile, "CC_MSG28=MTLHTRACKUPD"
    Print #iFile, "CC_MSG29=MTLHTRACKREP"
    Print #iFile, "CC_MSG30=MTSETACQSMODE"
    Print #iFile, "CC_MSG31=MTSSERROR"
    Print #iFile, "CC_MSG32=MTLOBRSLT"
    '+v1.8.12
    Print #iFile, "CC_MSG33=MTLHBEARINGONLY"
    '+v1.8.12
  
    Print #iFile, ";"
    Print #iFile, ";<messagename>ID=<message ID>"
    Print #iFile, "[Message ID]"
    Print #iFile, "MTDEFPMAID=51"
    Print #iFile, "MTSIGALARMID=2304"
    Print #iFile, "MTANAALARMID=2570"
    Print #iFile, "MTHBDYNRSPID=94"
    Print #iFile, "MTSDALARMID=2306"
    Print #iFile, "MTHBACTREPID=91"
    Print #iFile, "MTDFALARMID=3328"
    Print #iFile, "MTRFSTATUSID=2308"
    Print #iFile, "MTDFSDALARMID=3329"
    Print #iFile, "MTLOBUPDID=49"
    Print #iFile, "MTANARSLTID=42"
    Print #iFile, "MTSIGUPDID=19"
    Print #iFile, "MTHBLOBUPDID=566"
    Print #iFile, "MTULDDATAID=50"
    Print #iFile, "MTLOBSETRSLTID=60"
    Print #iFile, "MTFIXRSLTID=48"
    Print #iFile, "MTHBSIGUPDID=70"
    Print #iFile, "MTHBGSREPID=95"
    Print #iFile, "MTHBSEMISTATID=96"
    Print #iFile, "MTHBSELASKID=449"
    Print #iFile, "MTJAMSTATID=35"
    Print #iFile, "MTRUNMODEID=112"
    Print #iFile, "MTHBSELJAMID=450"
    Print #iFile, "MTHBXMTRSTATID=97"
    Print #iFile, "MTNAVREPID=53"
    Print #iFile, "MTLHECHMODID=76"
    Print #iFile, "MTLHCORRELATEID=345"
    Print #iFile, "MTLHTRACKUPDID=484"
    Print #iFile, "MTLHTRACKREPID=98"
    Print #iFile, "MTSETACQSMODEID=29"
    Print #iFile, "MTSSERRORID=113"
    Print #iFile, "MTLOBRSLTID=58"
    '-v1.8.12
    Print #iFile, "MTLHBEARINGONLYID=223"
    '-v1.8.12
    Print #iFile, ";"
    Print #iFile, "[Message Names]"
    Print #iFile, "; Block 30 Message IDs"
    Print #iFile, ";CC_MSGID<Message ID>=<Message name>"
    Print #iFile, "CC_MSGID51=MTDEFPMA"
    Print #iFile, "CC_MSGID2304=MTSIGALARM"
    Print #iFile, "CC_MSGID2570=MTANAALARM"
    Print #iFile, "CC_MSGID94=MTHBDYNRSP"
    Print #iFile, "CC_MSGID2306=MTSDALARM"
    Print #iFile, "CC_MSGID91=MTHBACTREP"
    Print #iFile, "CC_MSGID3328=MTDFALARM"
    Print #iFile, "CC_MSGID2308=MTRFSTATUS"
    Print #iFile, "CC_MSGID3329=MTDFSDALARM"
    Print #iFile, "CC_MSGID49=MTLOBUPD"
    Print #iFile, "CC_MSGID42=MTANARSLT"
    Print #iFile, "CC_MSGID19=MTSIGUPD"
    Print #iFile, "CC_MSGID566=MTHBLOBUPD"
    Print #iFile, "CC_MSGID50=MTULDDATA"
    Print #iFile, "CC_MSGID60=MTLOBSETRSLT"
    Print #iFile, "CC_MSGID48=MTFIXRSLT"
    Print #iFile, "CC_MSGID70=MTHBSIGUPD"
    Print #iFile, "CC_MSGID95=MTHBGSREP"
    Print #iFile, "CC_MSGID96=MTHBSEMISTAT"
    Print #iFile, "CC_MSGID449=MTHBSELASK"
    Print #iFile, "CC_MSGID35=MTJAMSTAT"
    Print #iFile, "CC_MSGID112=MTRUNMODE"
    Print #iFile, "CC_MSGID450=MTHBSELJAM"
    Print #iFile, "CC_MSGID97=MTHBXMTRSTAT"
    Print #iFile, "CC_MSGID53=MTNAVREP"
    Print #iFile, "CC_MSGID76=MTLHECHMOD"
    Print #iFile, "CC_MSGID345=MTLHCORRELATE"
    Print #iFile, "CC_MSGID484=MTLHTRACKUPD"
    Print #iFile, "CC_MSGID98=MTLHTRACKREP"
    Print #iFile, "CC_MSGID29=MTSETACQSMODE"
    Print #iFile, "CC_MSGID113=MTSSERROR"
    Print #iFile, "CC_MSGID58=MTLOBRSLT"
    '+v1.8.12
    Print #iFile, "CC_MSGID223=MTLHBEARINGONLY"
    '-v1.8.12
    Print #iFile, ";"
    Print #iFile, ";CC_MSG_DESC<message ID>=<message description>"
    Print #iFile, "[Message Descriptions]"
    Print #iFile, "CC_MSG_DESC51=Defines the Prime Mission Area (PMA)"
    Print #iFile, "CC_MSG_DESC2304=Reports low/mid band signals detected by the Acquisition subsystem"
    Print #iFile, "CC_MSG_DESC2570=Reports low/mid band signals detected by the Analysis subsystem"
    Print #iFile, "CC_MSG_DESC94=High band dynamic response"
    Print #iFile, "CC_MSG_DESC2306=Reports short-duration low/mid band signals detected by the Acquisition subsystem"
    Print #iFile, "CC_MSG_DESC91=High band signal activity report"
    Print #iFile, "CC_MSG_DESC3328=Reports low/mid band signals detected by the DF subsystem"
    Print #iFile, "CC_MSG_DESC2308=Reports RF activity for all low/mid band receiver bins"
    Print #iFile, "CC_MSG_DESC3329=Reports short-duration low/mid band signals detected by the DF subsystem"
    Print #iFile, "CC_MSG_DESC49=Line-of-bearing update for low/mid band emitters"
    Print #iFile, "CC_MSG_DESC42=Signal analysis results"
    Print #iFile, "CC_MSG_DESC19=Low/mid band signal record update"
    Print #iFile, "CC_MSG_DESC566=Line-of-bearing update for high band emitters"
    Print #iFile, "CC_MSG_DESC50=Uploaded line-of-bearing data for low/mid band emitters"
    Print #iFile, "CC_MSG_DESC60=Line-of-bearing to low/mid band emitter"
    Print #iFile, "CC_MSG_DESC48=Computed low/mid band emitter ground location"
    Print #iFile, "CC_MSG_DESC70=High band signal record update"
    Print #iFile, "CC_MSG_DESC95=Computed high band emitter ground location"
    Print #iFile, "CC_MSG_DESC96=High band signal semi-static record update"
    Print #iFile, "CC_MSG_DESC449=High band signal selection status"
    Print #iFile, "CC_MSG_DESC35=Low/mid band signal transmitter status"
    Print #iFile, "CC_MSG_DESC112=Operating mode"
    Print #iFile, "CC_MSG_DESC450=High band signal selection status"
    Print #iFile, "CC_MSG_DESC97=High band transmitter status"
    Print #iFile, "CC_MSG_DESC53=Navigation report"
    Print #iFile, "CC_MSG_DESC76=Echelon update"
    Print #iFile, "CC_MSG_DESC345=Correlation Line"
    Print #iFile, "CC_MSG_DESC484=Track update"
    Print #iFile, "CC_MSG_DESC98=Track report"
    Print #iFile, "CC_MSG_DESC29=Set environment message"
    Print #iFile, "CC_MSG_DESC113=Error and status messages"
    '
    '+v1.6TE
    Print #iFile, "CC_MSG_DESC58=Lob result messages"
    '-v1.6
    Print #iFile, ";"
    Print #iFile, "[Message Fields]"
    Print #iFile, "; MSG_FIELDS<Message ID>=<field list>"
    Print #iFile, "MSG_FIELDS51=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID"
    Print #iFile, "MSG_FIELDS2304=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID"
    Print #iFile, "MSG_FIELDS2570=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status AS Active, Tag AS Variant, Flag, Common_ID"
    Print #iFile, "MSG_FIELDS94=ReportTime, Msg_Type, Rpt_Type, Emitter, Emitter_ID, Signal, Signal_ID, Frequency AS JamOverrideFreq, Status AS Response, Tag AS On_List, Other_Data AS Function"
    Print #iFile, "MSG_FIELDS2306=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, Status, Bearing"
    Print #iFile, "MSG_FIELDS91=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Other_Data AS Function"
    Print #iFile, "MSG_FIELDS3328=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status AS In_Out_PMA, Tag, Flag, Common_ID, Range, Bearing, Elevation"
    Print #iFile, "MSG_FIELDS2308=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Other_Data"
    Print #iFile, "MSG_FIELDS3329=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status AS In_Out_PMA, Tag, Flag, Common_ID, Range, Bearing, Elevation"
    Print #iFile, "MSG_FIELDS49=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status AS In_Out_PMA, Tag, Flag, Common_ID, Range, Bearing, Elevation"
    Print #iFile, "MSG_FIELDS42=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status AS Active, Tag AS Variant, Flag AS RequestorID, Common_ID, Range, Bearing, Elevation"
    Print #iFile, "MSG_FIELDS19=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status AS Active, Tag AS Variant, Flag AS Operator, Common_ID"
    Print #iFile, "MSG_FIELDS566=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, Bearing, Other_Data AS Contributors"
    Print #iFile, "MSG_FIELDS50=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Range, Bearing, Elevation"
    Print #iFile, "MSG_FIELDS60=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status AS In_Out_PMA, Tag, Flag, Common_ID, Range, Bearing, Elevation"
    Print #iFile, "MSG_FIELDS48=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status AS FixType, Tag, Flag, Other_Data AS Requestor"
    Print #iFile, "MSG_FIELDS70=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Range, Bearing, Elevation"
    Print #iFile, "MSG_FIELDS95=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, XX, XY, YY"
    Print #iFile, "MSG_FIELDS96=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Range, Bearing, Elevation"
    Print #iFile, "MSG_FIELDS449=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, Other_Data"
    Print #iFile, "MSG_FIELDS35=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, Status, Other_Data"
    Print #iFile, "MSG_FIELDS112=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Other_Data"
    Print #iFile, "MSG_FIELDS450=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Status AS OnOff, Other_Data"
    Print #iFile, "MSG_FIELDS97=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Other_Data"
    Print #iFile, "MSG_FIELDS53=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID"
    Print #iFile, "MSG_FIELDS76=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Other_Data"
    Print #iFile, "MSG_FIELDS345=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Other_Data"
    Print #iFile, "MSG_FIELDS484=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID"
    Print #iFile, "MSG_FIELDS113=ReportTime, Msg_Type, Origin, Origin_ID, Other_Data"
    '
    '+v1.8.7
    Print #iFile, "MSG_FIELDS58=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Range, Bearing, Elevation"
    Print #iFile, "MSG_FIELDS98=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Range, Bearing, Elevation, Other_data as TrackClass"
    '-v1.8.7
    '+v1.8.12
    Print #iFile, "MSG_FIELDS223=ReportTime, Msg_Type, Rpt_Type, Origin, Origin_ID, Target_ID, Latitude, Longitude, Altitude, Heading, Speed, Allegiance, IFF, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, PRI, Status, Tag, Flag, Common_ID, Range, Bearing, Elevation, Other_data as TrackClass"
    '-v1.8.12


    Print #iFile, ";"
    Print #iFile, "[Queries]"
    Print #iFile, "; MAX_QUERIES=<title token>=<number of queries>"
    Print #iFile, "; <title token>;=<Name to show the user>=;"
    Print #iFile, "; QUERY_FIELDS;=<Field list>=;"
    Print #iFile, "; QUERY;=<Where clause>=;"
    Print #iFile, "; QUERY_SORT;=<Sort fields>=;"
    Print #iFile, ";"
    Print #iFile, "MAX_QUERIES=16"
    Print #iFile, ";"
    Print #iFile, "; Query 1 - Detection Information"
    Print #iFile, "QUERY_TITLE1=MOP 1 - Detection Quick-Look"
    Print #iFile, "QUERY_FIELDS1=Frequency, Count(*) AS Count, MIN(ReportTime) AS First_Time, Max(ReportTime) as Last_Time"
    Print #iFile, "QUERY1=(Msg_Type = 'MTSIGALARM' OR Msg_Type = 'MTANAALARM' OR Msg_Type = 'MTSDALARM' OR Msg_Type = 'MTDFALARM' OR Msg_Type = 'MTDFSDALARM' OR Msg_Type = 'MTHBACTREP') AND Rpt_Type = 'SIG' GROUP BY Frequency"
    Print #iFile, ";"
    Print #iFile, "; Query 2 - PMA Determination"
    Print #iFile, "QUERY_TITLE2=MOP 2 - PMA Quick-Look"
    Print #iFile, "QUERY_FIELDS2=Frequency, Status, Count(*) As Count, Min(ReportTime) As First_Time, Min(Bearing) AS Min_Brng, Max(Bearing) AS Max_Brng"
    Print #iFile, "QUERY2=(Msg_Type = 'MTSDALARM' OR Msg_Type = 'MTDFALARM' OR Msg_Type = 'MTDFSDALARM' OR Msg_Type = 'MTLOBUPD' OR Msg_Type = 'MTANARSLT') AND Rpt_Type = 'VEC' GROUP BY Frequency, Status"
    Print #iFile, ";"
    Print #iFile, "; Query 3 - Identification"
    Print #iFile, "QUERY_TITLE3=MOP 3 - ID Quick-Look"
    Print #iFile, "QUERY_FIELDS3=Frequency, Emitter, Count(*) AS Count, Min(ReportTime) AS First_Time, Max(ReportTime) as Last_Time"
    Print #iFile, "QUERY3=(Msg_Type = 'MTHBACTREP' OR Msg_Type = 'MTANARSLT') AND Rpt_Type = 'SIG' GROUP BY Frequency, Emitter"
    Print #iFile, ";"
    Print #iFile, "; Query 4 - ATN"
    Print #iFile, "QUERY_TITLE4=MOP 4 - ATN Quick-Look"
    Print #iFile, "QUERY_FIELDS4=Frequency, Signal_ID, Flag, Count(*) AS Count, Min(ReportTime) AS First_Time"
    Print #iFile, "QUERY4=Msg_Type = 'MTANARSLT' AND Rpt_Type = 'SIG' GROUP BY Frequency, Signal_ID, Flag"
    Print #iFile, ";"
    Print #iFile, "; Query 5 - LOBs"
    Print #iFile, "QUERY_TITLE5=MOP 5 - Low/Mid Band LOB Quick-Look"
    Print #iFile, "QUERY_FIELDS5=Frequency, Signal_ID, Count(*) AS Count, Min(Bearing) as Min_Brng, Max(Bearing) AS Max_Brng, Avg(Bearing) AS Avg_Brng, Min(ReportTime) AS First_Time, Max(ReportTime) AS Last_Time"
    Print #iFile, "QUERY5=(Msg_Type = 'MTANARSLT' OR Msg_Type = 'MTULDDATA' OR Msg_Type = 'MTLOBSETRSLT') AND Rpt_Type = 'VEC' GROUP BY Frequency, Signal_ID"
    Print #iFile, ";"
    Print #iFile, "; Query 6 - High Band LOBs"
    Print #iFile, "QUERY_TITLE6=MOP 5 - High Band LOB Quick-Look"
    Print #iFile, "QUERY_FIELDS6=Frequency, Signal_ID, Count(*) AS Count, Min(Bearing) as Min_Brng, Max(Bearing) AS Max_Brng, Avg(Bearing) as Avg_Brng, Min(ReportTime) AS First_Time, Max(ReportTime) AS Last_Time"
    Print #iFile, "QUERY6=Msg_Type = 'MTHBLOBUPD' AND Rpt_Type = 'VEC' GROUP BY Frequency, Signal_ID"
    Print #iFile, ";"
    Print #iFile, "; Query 7 - Low/Mid Band Geolocation"
    Print #iFile, "QUERY_TITLE7=MOP 6 - Low/Mid Band Geolocation Quick-Look"
    Print #iFile, "QUERY_FIELDS7=Frequency, Signal_ID, Count(*) AS Count, Avg(Latitude) as Avg_Latitude, Avg(Longitude) as Avg_Longitude, Min(ReportTime) AS First_Time, Max(ReportTime) AS Last_Time"
    Print #iFile, "QUERY7=Msg_Type = 'MTFIXRSLT' AND Rpt_Type = 'GEO' GROUP BY Frequency, Signal_ID"
    Print #iFile, ";"
    Print #iFile, "; Query 8 - High Band Geolocation"
    Print #iFile, "QUERY_TITLE8=MOP 6 - High Band Geolocation Quick-Look"
    Print #iFile, "QUERY_FIELDS8=Frequency, Signal_ID, Count(*) AS Count, Avg(Latitude) as Avg_Latitude, Avg(Longitude) as Avg_Longitude, Min(ReportTime) AS First_Time, Max(ReportTime) AS Last_Time"
    Print #iFile, "QUERY8=(Msg_Type = 'MTHBSIGUPD' OR Msg_Type = 'MTHBGSREP' OR Msg_Type = 'MTHBSEMISTAT' OR Msg_Type = 'MTHBSELASK') AND Rpt_Type = 'GEO' GROUP BY Frequency, Signal_ID"
    Print #iFile, ";"
    Print #iFile, "; Query 9 - Effectiveness"
    Print #iFile, "QUERY_TITLE9=MOP 7 - Jam Activity Quick-Look"
    Print #iFile, "QUERY_FIELDS9=ReportTime, Msg_Type, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, Status, Tag, Flag, Other_Data"
    Print #iFile, "QUERY9=Msg_Type = 'MTJAMSTAT' OR Msg_Type = 'MTRUNMODE' OR Msg_Type = 'MTHBDYNRSP' OR Msg_Type = 'MTHBSELJAM' OR Msg_Type = 'MTHBXMTRSTAT'"
    Print #iFile, "QUERY_SORT9=ReportTime"
    Print #iFile, ";"
    Print #iFile, "; Query 10 - DPS-1 Ech"
    Print #iFile, "QUERY_TITLE10=DPS - 1 - Ech"
    Print #iFile, "QUERY_FIELDS10=ReportTime, Msg_Type, Origin AS Ech_Type, Origin_ID AS Ech_Type_ID, Latitude, Longitude, Allegiance, IFF, Flag AS Priority"
    Print #iFile, "QUERY10=Msg_Type = 'MTLHECHMOD'"
    Print #iFile, "QUERY_SORT10=ReportTime"
    Print #iFile, ";"
    Print #iFile, "; Query 11 - DPS-2 Trk"
    Print #iFile, "QUERY_TITLE11=DPS - 2 - Trk"
    Print #iFile, "QUERY_FIELDS11=*"
    Print #iFile, "QUERY11=(Msg_Type = 'MTLHTRKUPD' OR Msg_Type = 'MTHBLOBUPD') AND Rpt_Type = 'TRK'"
    Print #iFile, "QUERY_SORT11=Target_ID, ReportTime"
    Print #iFile, ";"
    Print #iFile, "; Query 12 - DPS-3 S/A"
    Print #iFile, "QUERY_TITLE12=DPS - 3 - S / A"
    Print #iFile, "QUERY_FIELDS12=ReportTime, Msg_Type, Origin as Subject, Origin_ID as Subject_ID, Parent as Object, Parent_ID as Object_ID, Status as Type, Other_Data"
    Print #iFile, "QUERY12=Msg_Type = 'MTLHCORRELATE'"
    Print #iFile, "QUERY_SORT12=ReportTime"
    Print #iFile, ";"
    Print #iFile, "; Query 13 - DPS-4 Pri"
    Print #iFile, "QUERY_TITLE13=DPS-4 - Pri."
    Print #iFile, "QUERY_FIELDS13=ReportTime, Msg_Type, Emitter, Emitter_ID, Signal, Signal_ID, Frequency, Status as Active, Tag as Variant, Flag as Operator"
    Print #iFile, "QUERY13=Msg_Type = 'MTSIGUPD'"
    Print #iFile, "QUERY_SORT13=Frequency, ReportTime"
    Print #iFile, ";"
    Print #iFile, "; Query 14 - DPS-5 TIBS"
    Print #iFile, ";QUERY_TITLE14=DPS-5 - TIBS"
    Print #iFile, ";QUERY_FIELDS14=*"
    Print #iFile, ";QUERY14=Msg_Type = 'MTLHTRKUPD' AND Origin = 'TIBS'"
    Print #iFile, ";QUERY_SORT14=Target_ID, ReportTime"
    Print #iFile, ";"
    Print #iFile, "; Query 15 - High band activity histogram"
    Print #iFile, "QUERY_TITLE15=High-band Activity Histogram"
    Print #iFile, "QUERY_FIELDS15=Frequency, Count(Frequency) AS Count, Min(ReportTime) AS First_Time, Max(ReportTime) AS Last_Time"
    Print #iFile, "QUERY15=Msg_Type = 'MTHBACTREP' GROUP BY Frequency"
    Print #iFile, ";"
    Print #iFile, "; Query 16 - Low/Mid Band Activity Histogram"
    Print #iFile, "QUERY_TITLE16=Low/Mid Band Activity Histogram"
    Print #iFile, "QUERY_FIELDS16=Frequency, Count(Frequency) AS Count, Min(ReportTime) AS First_Time, Max(ReportTime) AS Last_Time"
    Print #iFile, "QUERY16=(Msg_Type = 'MTSIGALARM' OR Msg_Type = 'MTANAALARM' OR Msg_Type = 'MTSDALARM' OR Msg_Type = 'MTDFALARM' OR Msg_Type = 'MTDFSDALARM') AND Rpt_Type = 'SIG' AND Frequency > 0 GROUP BY Frequency"
    Print #iFile, ";"
    Print #iFile, "[RCV List]"
    Print #iFile, ";RCV<Radio><Class>=<Sigtype>=<Modtype>"
    Print #iFile, "R0=UNKNOWN"
    Print #iFile, "C0=UNKNOWN"
    Print #iFile, "R1=AM"
    Print #iFile, "R2=FM"
    Print #iFile, "R3=LSB"
    Print #iFile, "R4=USB"
    Print #iFile, "R5=AR"
    Print #iFile, "R6=RF"
    Print #iFile, "R7=ISB"
    Print #iFile, "R8=M1"
    Print #iFile, "R9=M5"
    Print #iFile, "R10=M3"
    Print #iFile, "R11=M6"
    Print #iFile, "R13=M2"
    Print #iFile, "C1=VOICE"
    Print #iFile, "C2=COMMERCIALTV"
    Print #iFile, "C3=COMMERCIALAUDIO"
    Print #iFile, "C4=TS"
    Print #iFile, "C5=58"
    Print #iFile, "C6=FJ"
    Print #iFile, "C7=HB"
    Print #iFile, "C8=SW"
    Print #iFile, "C9=NL"
    Print #iFile, "C10=MK"
    Print #iFile, "C11=BC"
    Print #iFile, "C12=KT"
    Print #iFile, "C13=MT"
    Print #iFile, "C14=BK"
    Print #iFile, "C15=BN"
    Print #iFile, "C16=BS"
    Print #iFile, "C17=LC"
    Print #iFile, "C18=PB"
    Print #iFile, "C19=TD"
    Print #iFile, "C20=TN"
    Print #iFile, "C21=AN"
    Print #iFile, "C22=NR"
    Print #iFile, "C23=CM"
    Print #iFile, "C24=WA"
    Print #iFile, "C25=WF"
    Print #iFile, "C26=WS"
    Print #iFile, "C27=WQ"
    Print #iFile, "C28=WT"
    Print #iFile, "C29=WK"
    Print #iFile, "C30=WP"
    Print #iFile, "C31=WB"
    Print #iFile, "C34=PS"
    Print #iFile, "C35=ET"
    Print #iFile, ";"
    Print #iFile, ";RUNMODE<mode>=<description>"
    Print #iFile, "[Runmode]"
    Print #iFile, "RUNMODE1=PROM"
    Print #iFile, "RUNMODE2=IDLE"
    Print #iFile, "RUNMODE3=STANDBY"
    Print #iFile, "RUNMODE4=SEARCH"
    Print #iFile, "RUNMODE5=JAM"
    Print #iFile, "RUNMODE0=UNKNOWN"
    Print #iFile, ";"
    Print #iFile, ";JAMSTAT<status>=<description>"
    Print #iFile, "[JAMSTAT]"
    Print #iFile, "JAMSTAT4=NOJAM[INACTIVE]"
    Print #iFile, "JAMSTAT8=NOJAM[RESOURCES](NS)"
    Print #iFile, "JAMSTAT10=NOJAM[RESOURCES](S)"
    Print #iFile, "JAMSTAT13=JAM[QUESTIONABLE](NS)"
    Print #iFile, "JAMSTAT15=JAM[QUESTIONABLE](S)"
    Print #iFile, "JAMSTAT29=JAM[EFFECTIVE](NS)"
    Print #iFile, "JAMSTAT31=JAM[EFFECTIVE](S)"
    Print #iFile, ";"
    Print #iFile, ";CALL_IFF<IFF>=<Text>"
    Print #iFile, ";DAS_<CALL IFF TEXT>=<DAS ID #>"
    Print #iFile, ";IFF<DAS IFF ID>=<TEXT>"
    Print #iFile, "[IFF]"
    Print #iFile, "CALL_IFF0=UNDEF"
    Print #iFile, "CALL_IFF1=UNKNOWN"
    Print #iFile, "CALL_IFF2=HOSTILE"
    Print #iFile, "CALL_IFF3=FRIENDLY"
    Print #iFile, "CALL_IFF4=USE_SIGNAL"
    Print #iFile, "CALL_IFF5=NOT_SIGNAL"
    Print #iFile, "CALL_IFF6=USE_ID"
    Print #iFile, "DAS_UNDEF=0"
    Print #iFile, "DAS_UNKNOWN=0"
    Print #iFile, "DAS_HOSTILE=2"
    Print #iFile, "DAS_FRIENDLY=1"
    Print #iFile, "DAS_USE_SIGNAL=4"
    Print #iFile, "DAS_NOT_SIGNAL=5"
    Print #iFile, "DAS_USE_ID=6"
    Print #iFile, "IFF0=UNKNOWN"
    Print #iFile, "IFF1=FRIENDLY"
    Print #iFile, "IFF2=HOSTILE"
    Print #iFile, "IFF3=NEUTRAL"
    Print #iFile, "IFF4=USE_SIGNAL"
    Print #iFile, "IFF5=NOT_SIGNAL"
    Print #iFile, "IFF6=USE_ID"
    Print #iFile, ";"
    Print #iFile, "[HBSig]"
    Print #iFile, "HBSIG1=R"
    Print #iFile, "HBSIG2=T"
    Print #iFile, "HBSIG3=A3"
    Print #iFile, "HBSIG4=BT"
    Print #iFile, "HBSIG5=B5"
    Print #iFile, "HBSIG6=32"
    Print #iFile, "HBSIG7=33"
    Print #iFile, "HBSIG8=A4"
    Print #iFile, "HBSIG9=A37"
    Print #iFile, "HBSIG10=MI"
    Print #iFile, "HBSIG11=MA"
    Print #iFile, "HBSIG12=UNKNOWN"
    Print #iFile, ";"
    Print #iFile, "[MAP]"
    Print #iFile, ";MAP<index>=<AID Start>,<AID End>,<SigID Start>,<SigID End>"
    Print #iFile, "MAP1=0,95,0,7"
    Print #iFile, "MAP2=96,347,8,63"
    Print #iFile, "MAP3=348,369,192,199"
    Print #iFile, "MAP4=370,373,200,207"
    Print #iFile, "MAP5=374,377,208,299"
    Print #iFile, "MAP6=378,380,64,71"
    Print #iFile, "MAP7=381,383,72,79"
    Print #iFile, "MAP8=384,384,80,87"
    Print #iFile, "MAP9=385,386,128,135"
    Print #iFile, "MAP10=387,388,88,127"
    Print #iFile, "MAP11=389,390,136,191"
    Print #iFile, ";"
    Print #iFile, ";"
    Print #iFile, "[XMTRSTAT]"
    Print #iFile, "XMTRSTAT0=NA"
    Print #iFile, "XMTRSTAT1=ON"
    Print #iFile, "XMTRSTAT2=OFF"
    Print #iFile, "XMTRSTAT4=LOW"
    Print #iFile, ";"
    Print #iFile, "[HBFUNC]"
    Print #iFile, "HBFUNC0=DETECT"
    Print #iFile, "HBFUNC1=SEARCH"
    Print #iFile, "HBFUNC2=MAP"
    Print #iFile, "HBFUNC3=ASK"
    Print #iFile, "HBFUNC4=JAM"
    Print #iFile, "HBFUNC5=SCAN"
    Print #iFile, "HBFUNC6=SEARCH_ONCE"
    Print #iFile, ";"
    Print #iFile, "[HBOPT]"
    Print #iFile, "HBOPT1=NORM"
    Print #iFile, "HBOPT2=COV"
    Print #iFile, "HBOPT3=DF_BEARING"
    Print #iFile, "HBOPT4=HYP"
    Print #iFile, "HBOPT5=ELLIP"
    Print #iFile, "HBOPT6=RP"
    Print #iFile, "HBOPT7=AUTO"
    Print #iFile, "HBOPT8=AUTO_WO_PL"
    Print #iFile, "HBOPT9=MANUAL"
    Print #iFile, "HBOPT10=OI"
    Print #iFile, "HBOPT11=OR"
    Print #iFile, "HBOPT12=ADN"
    Print #iFile, "HBOPT13=ADC"
    Print #iFile, "HBOPT14=MPK"
    Print #iFile, "HBOPT15=AOI"
    Print #iFile, "HBOPT16=APK"
    Print #iFile, "HBOPT17=OD"
    Print #iFile, "HBOPT18=OS"
    Print #iFile, "HBOPT19=DME"
    Print #iFile, "HBOPT20=MG1"
    Print #iFile, "HBOPT21=AG1"
    Print #iFile, "HBOPT22=SCAN"
    Print #iFile, "HBOPT23=DWELL"
    Print #iFile, "HBOPT24=MAP_DWELL"
    Print #iFile, ";"
    Print #iFile, "[LMBSig]"
    Print #iFile, ";LMBSIG1=TEST"
    Print #iFile, ";"
    Print #iFile, "[Signal]"
    Print #iFile, "SIG-1=NEW_SIGNAL"
    Print #iFile, ";"
    Print #iFile, "[Emitters]"
    Print #iFile, "Emitter0=UNK_EMITTER"
    Print #iFile, ";"
    Print #iFile, "[Ground Sites]"
    Print #iFile, "SITE0=UNKNOWN_SITE"
    Print #iFile, ";"
    Print #iFile, "[Miscellaneous operations]"
    Print #iFile, "UPDATE_INTERVAL=1000"
    '
    '+v1.5TE
    Print #iFile, "ADVSQL=0"
    '-v1.5
    '
    '+v1.6TE
    Print #iFile, "WIZARDINTRO=0"
    Print #iFile, "UseGPS=0"
    '-v1.6
    Print #iFile, ";"
    Print #iFile, "[OPERATOR]"
    Print #iFile, "OPERATOR -1 = Automatic"
    Print #iFile, "OPERATOR1=Amn Gump"
    '
    '+v1.5TE
    Print #iFile, ";"
    Print #iFile, ";+v1.5"
    Print #iFile, "[Export]"
    Print #iFile, "; File export formats"
    Print #iFile, "; Text_Delimiter = character to put on either side of a text value (blank for none)"
    Print #iFile, "; Text_Space = character(s) to replace spaces within a text value (blank for space)"
    Print #iFile, "; Text_Blank = string to use if a text field is blank"
    Print #iFile, "; Long_Format = the format string for long integer values"
    Print #iFile, "; Dbl_Format = the format string for floating point values"
    Print #iFile, "; TSecs = 1 to export TSecs, 0 to export Date/Time"
    Print #iFile, "; Time_Format = the format string for date/time values"
    Print #iFile, "Text_Delimiter="
    Print #iFile, "Text_Space=_"
    Print #iFile, "Text_Blank=Unknown"
    Print #iFile, "Long_Format=0"
    Print #iFile, "Dbl_Format=0.00000"
    Print #iFile, "TSecs=1"
    Print #iFile, "Time_Format=mm/dd/yyyy hh:nn:ss.000"
    Print #iFile, ";-v1.5"
    '-v1.5
    '
    '+v1.6TE
    Print #iFile, ";"
    Print #iFile, ";+v1.6"
    Print #iFile, "[Versions]"
    Print #iFile, "CCOS1 = 2.0"
    Print #iFile, "CCOS2 = 2.2"
    Print #iFile, "CCOS3 = 2.3"
    Print #iFile, "CCOS4 = 2.31"
    Print #iFile, ";-v1.6"
    '-v1.6
    '
    '+v1.6TE
    Print #iFile, ";"
    Print #iFile, ";+v1.6"
    Print #iFile, "[Origin]"
    Print #iFile, "; Origin<ID>=<Subsystem>"
    Print #iFile, "ORIGIN512 = DLOAD_DWS"
    Print #iFile, "ORIGIN768 = DU0"
    Print #iFile, "ORIGIN1024 = DU1"
    Print #iFile, "ORIGIN1280 = DU2"
    Print #iFile, "ORIGIN1536 = DU3"
    Print #iFile, "ORIGIN1792 = DU4"
    Print #iFile, "ORIGIN2048 = DU5"
    Print #iFile, "ORIGIN2304 = DU6"
    Print #iFile, "ORIGIN2560 = DU7"
    Print #iFile, "ORIGIN2816 = DU8"
    Print #iFile, "ORIGIN3072 = DU9"
    Print #iFile, "ORIGIN3328 = DU10"
    Print #iFile, "ORIGIN3584 = DU11"
    Print #iFile, "ORIGIN3840 = DU12"
    Print #iFile, "ORIGIN4096 = DU13"
    Print #iFile, "ORIGIN4352 = DWS1"
    Print #iFile, "ORIGIN4608 = DWS2"
    Print #iFile, "ORIGIN4864 = DWS3"
    Print #iFile, "ORIGIN5120 = DWS4"
    Print #iFile, "ORIGIN5376 = DWS5"
    Print #iFile, "ORIGIN5632 = DWS6"
    Print #iFile, "ORIGIN5888 = DWS7"
    Print #iFile, "ORIGIN6144 = CCUA"
    Print #iFile, "ORIGIN6400 = CCUB"
    Print #iFile, "ORIGIN6656 = CCU"
    Print #iFile, "ORIGIN6912 = SPARE0"
    Print #iFile, "ORIGIN7168 = DF"
    Print #iFile, "ORIGIN7424 = EXCITER"
    Print #iFile, "ORIGIN7680 = ANALYSIS"
    Print #iFile, "ORIGIN7936 = ACQUISITION1"
    Print #iFile, "ORIGIN8192 = ACQUISITION2"
    Print #iFile, "ORIGIN8448 = BOTH_ACQ"
    Print #iFile, "ORIGIN8704 = HIGHBAND"
    Print #iFile, "ORIGIN8960 = DPS"
    Print #iFile, "ORIGIN9216 = TNA"
    Print #iFile, "ORIGIN9472 = TRC1"
    Print #iFile, "ORIGIN9728 = TRC2"
    Print #iFile, "ORIGIN9984 = TRC3"
    Print #iFile, "ORIGIN10240 = TRC4"
    Print #iFile, "ORIGIN10496 = TRC5"
    Print #iFile, "ORIGIN10752 = SUBA"
    Print #iFile, "ORIGIN11008 = SUBB"
    Print #iFile, "ORIGIN11264 = SUBC"
    Print #iFile, "ORIGIN11520 = SUBD"
    Print #iFile, "ORIGIN11776 = ACQ_DF_ANA_EXC"
    Print #iFile, "ORIGIN12032 = ACQ_DF_ANA"
    Print #iFile, "ORIGIN12288 = ACQ_DF"
    Print #iFile, "ORIGIN12544 = ACQ_DF_EXC"
    Print #iFile, "ORIGIN12800 = ACQ_ANA_EXC"
    Print #iFile, "ORIGIN13056 = ACQ_ANA"
    Print #iFile, "ORIGIN13312 = ACQ_EXC"
    Print #iFile, "ORIGIN13568 = DF_ANA_EXC"
    Print #iFile, "ORIGIN13824 = DF_ANA"
    Print #iFile, "ORIGIN14080 = DF_EXC"
    Print #iFile, "ORIGIN14336 = ANA_EXC"
    Print #iFile, "ORIGIN14592 = HBS_DPS"
    Print #iFile, "ORIGIN14848 = CCU_DWS_ACQ"
    Print #iFile, "ORIGIN15104 = CCU_DPS"
    Print #iFile, "ORIGIN15360 = CCU_DWS_DPS"
    Print #iFile, "ORIGIN15616 = CCU_DWS_HBS_DPS"
    Print #iFile, "ORIGIN15872 = CCU_DWS_DF_DPS"
    Print #iFile, "ORIGIN16128 = ACQ_DPS"
    Print #iFile, "ORIGIN16384 = CCU_DWS"
    Print #iFile, "ORIGIN16640 = ACQ_ANA_DF_EXC_DWS_DPS"
    Print #iFile, "ORIGIN16896 = HBS_DWS_DPS"
    Print #iFile, "ORIGIN17152 = All"
    Print #iFile, "ORIGIN17408 = DWS_DPS"
    Print #iFile, "ORIGIN17664 = ACQ_TEST"
    Print #iFile, "ORIGIN17920 = DWS_TEST"
    Print #iFile, "ORIGIN18176 = VME_TEST"
    Print #iFile, "ORIGIN18432 = UUT_TEST"
    Print #iFile, "ORIGIN39424 = HBE_ENODE"
    Print #iFile, "ORIGIN39680 = SEHDP_ENODE"
    Print #iFile, "ORIGIN39936 = SEHDF1_ENODE"
    Print #iFile, "ORIGIN40192 = SEHDF2_ENODE"
    Print #iFile, "ORIGIN40448 = SEHAQ1_ENODE"
    Print #iFile, "ORIGIN40704 = SEHAQ2_ENODE"
    Print #iFile, "ORIGIN40960 = SEHBE_ENODE"
    Print #iFile, ";-v1.6"
    '-v1.6
    '
    ' Close the file
    Close iFile
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Create_CCAT_Token_File (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Create_New_Session
' AUTHOR:   Tom Elkins
' PURPOSE:  Clears out the old tree view structures (if necessary) and adds a new
'           Session node as the root.
' INPUT:    None
' OUTPUT:   None
' NOTES:    This routine can be expanded in future versions for session management
'           functions.
Public Sub Create_New_Session()
    Dim nodSession As Node  ' The new session node
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Create_New_Session (Start)"
    '-v1.6.1
    '
    ' Log the event
    'basCCAT.WriteLogEntry "CCAT: Create_New_Session"
    '
    ' Check for existing nodes
    If frmMain.tvTreeView.Nodes.Count > 0 Then
        '
        ' Clear out the existing nodes
        frmMain.tvTreeView.Nodes.Clear
    End If
    '
    ' Create a new session node
    ' Since this is the root node, there is no relative or relationship values;
    ' therefore, the first two arguments to the Add method are blank.
    '   Relative:       None
    '   Relationship:   None
    '   Unique Key:     "SESSION"
    '   Text:           "Session " + current date and time
    '   Default Icon:   "Session" image
    '   Selected Icon:  None
    Set nodSession = frmMain.tvTreeView.Nodes.Add(, , gsSESSION, "Session " & Date, gsSESSION)
    nodSession.Tag = gsSESSION  ' Set the tag property for use in other routines
    '
    ' Since there is no expansion [+] indicator in front of the session node, we
    ' must explicitely expand the node; otherwise, when children nodes are added they
    ' will not appear in the tree view.
    nodSession.Expanded = True
    '
    ' Configure the interface for the session
    frmMain.ChangeMode gsSESSION
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Create_New_Session (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Remove_Database
' AUTHOR:   Tom Elkins
' PURPOSE:  Remove a database node from the tree view, and any reference to it in the
'           list view
' INPUT:    "sDB_Name" is the name of the database being removed
' OUTPUT:   None
' NOTES:
Public Sub Remove_Database(sDB_Name As String)
    Dim nodKenny As Node        ' Node to be removed
    Dim itmKenny As ListItem    ' List Item to be removed
    '
    ' Log the event
    '+v1.6.1TE
    'basCCAT.WriteLogEntry "CCAT: Remove_Database: " & sDB_Name
    basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Remove_Database (Start)"
    basCCAT.WriteLogEntry "ARGUMENTS: " & sDB_Name
    '-v1.6.1
    '
    ' Check for the existence of the node in the Tree View
    If frmMain.blnNodeExists(sDB_Name) Then
        '
        ' Find the specified node in the tree view
        Set nodKenny = frmMain.tvTreeView.Nodes(sDB_Name)
        '
        ' Search for the specified item in the list view
        Set itmKenny = frmMain.lvListView.FindItem(nodKenny.Text, lvwText)
        '
        ' See if anything was found
        If Not itmKenny Is Nothing Then
            '
            ' Make sure the item is a database item
            If itmKenny.Tag = gsDATABASE Then
                '
                ' Log the event
                basCCAT.WriteLogEntry "          Removing Item from List View"
                '
                ' Remove the item
                frmMain.lvListView.ListItems.Remove itmKenny.Index
            End If
        End If
        '
        ' Log the event
        basCCAT.WriteLogEntry "          Removing Node from Tree View"
        '
        ' Remove the database node from the tree view
        frmMain.tvTreeView.Nodes.Remove nodKenny.Index
    Else
        '
        ' Log the event
        basCCAT.WriteLogEntry "          Specified Database does not exist in the Session!"
        '
        ' Inform the user that something went wrong
        MsgBox "Specified Database does not exist in the Session!" & vbCr & sDB_Name, vbOKOnly, "Remove Failed"
    End If
    '
    '+v1.6.1TE
    basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Remove_Database (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Write_DAS_Header
' AUTHOR:   Tom Elkins
' PURPOSE:  Writes header and comment records to the specified DAS file
' INPUT:    "iFile" is the file number where the records will be written
' OUTPUT:   None
' NOTES:
Public Sub Write_DAS_Header(iFile As Integer)
    Dim rsTable As Recordset    ' Access tables within the database
    Dim iArchive As Integer     ' Store the archive identifier
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Write_DAS_Header (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & iFile
    End If
    '-v1.6.1
    '
    ' From the COMPASS CALL Data Analysis Plan, the DAS Header consists of
    ' a header identifier "#", file type (MTF, STF, SIG, or ERR), the name
    ' of the data source, and the module that created the file
    '
    ' Write header identifier
    Print #iFile, "#";
    '
    ' Write the file type.  The type of file selected by the user is equivalent to
    ' the index of the file type selected in the "Save As..." dialog.
    Select Case basCCAT.guExport.iFile_Type
        '
        ' Signal file
        Case 1:
            Print #iFile, "SIG,";
        '
        ' Moving target file
        Case 2, 3, 4:
            Print #iFile, "MTF,";
        '
        ' Stationary target file
        Case 5:
            Print #iFile, "STF,";
        '
        ' Event file
        Case 6:
            Print #iFile, "EVT,";
    End Select
    '
    ' Write the full data path, which is stored in the list view caption
    ' Write the application's name
    Print #iFile, frmMain.lblTitle(1).Caption; ","; App.Path & "\" & App.EXEName & ".exe"
    '
    ' Write DAS Comment records
    Print #iFile, "# Application " & basCCAT.sGet_Version
    Print #iFile, "# Output created on "; Now
    Print #iFile, "# CCAT Database: "; guCurrent.sName
    '
    ' Open the archives table and extract the record for the parent archive
    Set rsTable = guCurrent.DB.OpenRecordset("SELECT * FROM " & TBL_ARCHIVES & " WHERE ID = " & guCurrent.iArchive)
    '
    ' Write info about the archive to the DAS file
    Print #iFile, "# Archive identifier: "; rsTable!Name; " (ID = "; rsTable!ID; ")"
    Print #iFile, "# Original archive file: "; rsTable!Archive
    Print #iFile, "# Archive start time: "; rsTable!Start
    Print #iFile, "# Archive end time: "; rsTable!End
    Print #iFile, "# Filtered archive file: "; rsTable!Analysis_File
    Print #iFile, "# Archive processed on: "; rsTable!Processed
    Print #iFile, "# Archive Date: "; rsTable!Date
    '
    ' Set the time offset
    guCurrent.uArchive.dOffset_Time = (DatePart("y", rsTable!Date) - 1) * 86400#
    '
    ' Write the SQL query that extracted the data
    ' The string is split up so the lines are not too long.
    Print #iFile, "# SQL Statement: "
    Print #iFile, "#     SELECT "; guCurrent.uSQL.sFields
    Print #iFile, "#     FROM " & rsTable!Name & basDatabase.TBL_DATA
    If guCurrent.uSQL.sFilter <> "" Then Print #iFile, "#     WHERE " & guCurrent.uSQL.sFilter
    If guCurrent.uSQL.sOrder <> "" Then Print #iFile, "#     ORDER BY " & guCurrent.uSQL.sOrder
    '
    ' Close the archives table
    rsTable.Close
    '
    ' Write the number of matching records
    Print #iFile, "# Records meeting criteria: "; frmMain.Data1.Recordset.RecordCount
    '
    ' Print field header
    Print #iFile, "#" & guCurrent.uSQL.sFields
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Write_DAS_Header (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: uPopulate_DAS_Structure
' AUTHOR:   Tom Elkins
' PURPOSE:  Populates a DAS record structure with data contained in the specified record
' INPUT:    "rsData" is the record containing the data to be used
' OUTPUT:   a DAS Structure containing the data from the specified record
' NOTES:
Public Function uPopulate_DAS_Structure(rsData As Recordset) As DAS_MASTER_RECORD
    Dim sToken As String
    Dim lTokenLen As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basCCAT.uPopulate_DAS_Structure (Start)"
    '-v1.6.1
    '
    ' Use structure-level addressing
    With uPopulate_DAS_Structure
        '
        ' Copy the record data to the structure
        On Error Resume Next
        .dAltitude = rsData!Altitude
        .dBearing = rsData!Bearing
        .dElevation = rsData!Elevation
        .dFrequency = rsData!Frequency
        .dHeading = rsData!Heading
        .dLatitude = rsData!Latitude
        .dLongitude = rsData!Longitude
        .dPRI = rsData!PRI
        .dRange = rsData!Range
        .dSpeed = rsData!Speed
        .dReportTime = rsData!ReportTime
        .dXX = rsData!XX
        .dXY = rsData!XY
        .dYY = rsData!YY
        .lIFF = rsData!IFF
        .lCommon_ID = rsData!Common_ID
        .lEmitter_ID = rsData!Emitter_ID
        .lFlag = rsData!Flag
        .lOrigin_ID = rsData!Origin_ID
        .lParent_ID = rsData!Parent_ID
        .lSignal_ID = rsData!Signal_ID
        .lStatus = rsData!Status
        .lTag = rsData!Tag
        .lTarget_ID = rsData!Target_ID
        .sAllegiance = rsData!Allegiance
        If IsNumeric(.sAllegiance) Or (InStr(1, .sAllegiance, "UNKNOWN") > 0) Or .sAllegiance = "" Then
            .sAllegiance = basCCAT.GetAlias("IFF", "IFF" & .lIFF, "UNKNOWN")
        End If
        .sEmitter = rsData!Emitter
        If IsNumeric(.sEmitter) Or (InStr(1, .sEmitter, "UNKNOWN") > 0) Or .sEmitter = "" Then
            .sEmitter = basCCAT.GetAlias("Emitters", "Emitter" & .lEmitter_ID, "UNKNOWN")
        End If
        .sSupplemental = rsData!Other_Data
        .sMsg_Type = rsData!Msg_Type
        .sOrigin = rsData!Origin
        If IsNumeric(.sOrigin) Or (InStr(1, .sOrigin, "UNKNOWN") > 0) Or .sOrigin = "" Then
            .sOrigin = basCCAT.GetAlias("Origin", "ORIGIN" & .lOrigin_ID, "UNKNOWN")
        End If
        .sParent = rsData!Parent
        If IsNumeric(.sParent) Or (InStr(1, .sParent, "UNKNOWN") > 0) Or .sParent = "" Then
            .sParent = basCCAT.GetAlias("Parent", "PARENT" & .lParent_ID, "UNKNOWN")
        End If
        .sReport_Type = rsData!Report_Type
        .sSignal = rsData!Signal
        If IsNumeric(.sSignal) Or (InStr(1, .sSignal, "UNKNOWN") > 0) Or .sSignal = "" Then
            .sSignal = basCCAT.GetAlias("Signal", "SIG" & .lSignal_ID, "UNKNOWN")
        End If
        On Error GoTo 0
    End With
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basCCAT.uPopulate_DAS_Structure (End)"
    '-v1.6.1
    '
End Function
'
' ROUTINE:  Write_DAS_Record
' AUTHOR:   Tom Elkins
' PURPOSE:  Writes a data record to the specified DAS file
' INPUT:    "iFile" is the file number where the records will be written
'           "uRecord" is a DAS record structure containing the data to be output
'           "iType" is the type of DAS record to be output (SIG, MTF, STF, or EVT)
' OUTPUT:   None
' NOTES:
Public Sub Write_DAS_Record(iFile As Integer, uRecord As DAS_MASTER_RECORD)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Write_DAS_Record (Start)"
    '-v1.6.1
    '
    ' Write the elements of the record to the file in the proper format
    Print #iFile, Format(uRecord.dReportTime, "0.000");
    Print #iFile, ","; uRecord.sMsg_Type;
    If Len(uRecord.sReport_Type) = 0 Then uRecord.sReport_Type = gaDAS_Rec_Type(basCCAT.guExport.iRec_Type)
    Print #iFile, ","; uRecord.sReport_Type;
    Print #iFile, ","; uRecord.sOrigin;
    Print #iFile, ","; Format(uRecord.lOrigin_ID, "0");
    '
    '+v1.6TE
    Print #iFile, ","; Format(uRecord.lTarget_ID, "0");
    '-v1.6
    '
    ' Write position data for MTF/STF only
    If basCCAT.guExport.iFile_Type = giDAS_MTF Or basCCAT.guExport.iFile_Type = giDAS_STF Then
        '
        ' Target ID and position
        '+v1.6TE
        'Print #iFile, ","; Format(uRecord.lTarget_ID, "0");
        '-v1.6
        Print #iFile, ","; Format(uRecord.dLatitude, "0.00000");
        Print #iFile, ","; Format(uRecord.dLongitude, "0.00000");
        Print #iFile, ","; Format(uRecord.dAltitude, "0.00000");
        '
        ' Moving target file only
        If basCCAT.guExport.iFile_Type = giDAS_MTF Then
            '
            ' Target heading/speed
            Print #iFile, ","; Format(uRecord.dHeading, "0.000");
            Print #iFile, ","; Format(uRecord.dSpeed, "0.000");
        Else
            '
            ' Parent info for stationary target
            Print #iFile, ","; uRecord.sParent;
            Print #iFile, ","; Format(uRecord.lParent_ID, "0");
        End If
    End If
    '
    ' Write the next set of fields for all files except Event files
    If basCCAT.guExport.iFile_Type <> giDAS_EVT Then
        '
        ' Write signal/identification/status fields
        Print #iFile, ","; uRecord.sAllegiance;
        Print #iFile, ","; Format(uRecord.lIFF, "0");
        Print #iFile, ","; uRecord.sEmitter;
        Print #iFile, ","; Format(uRecord.lEmitter_ID, "0");
        Print #iFile, ","; uRecord.sSignal;
        Print #iFile, ","; Format(uRecord.lSignal_ID, "0");
        Print #iFile, ","; Format(uRecord.dFrequency, "0.00000");
        Print #iFile, ","; Format(uRecord.dPRI, "0.000");
        Print #iFile, ","; Format(uRecord.lStatus, "0");
        Print #iFile, ","; Format(uRecord.lTag, "0");
        Print #iFile, ","; Format(uRecord.lFlag, "0");
        '
        '+v1.6TE
        'Print #iFile, ","; Format(uRecord.lCommon_ID, "0");
        '-v1.6
    End If
    '
    '+v1.6TE
    ' Write the common identifier field
    Print #iFile, ","; Format(uRecord.lCommon_ID, "0");
    '-v1.6
    '
    '+v1.6TE
    ' Write the range for event files
    If basCCAT.guExport.iFile_Type = giDAS_EVT Then
        Print #iFile, ","; Format(uRecord.dRange, "0.00000");
    End If
    '-v1.6
    '
    ' Add the LOB-specific fields
    If basCCAT.guExport.iRec_Type = giREC_VEC Then
        '
        ' Range/bearing/elevation
        Print #iFile, ","; Format(uRecord.dRange, "0.00000");
        Print #iFile, ","; Format(uRecord.dBearing, "0.000");
        Print #iFile, ","; Format(uRecord.dElevation, "0.000");
    End If
    '
    ' Add the geolocation-specific fields
    If basCCAT.guExport.iRec_Type = giREC_GEO Then
        '
        ' Covariance matrix
        Print #iFile, ","; Format(uRecord.dXX, "0.00000");
        Print #iFile, ","; Format(uRecord.dXY, "0.00000");
        Print #iFile, ","; Format(uRecord.dYY, "0.00000");
    End If
    '
    ' See if there is anything in the supplemental field
    If Len(uRecord.sSupplemental) > 0 Then
        '
        ' Write it
        Print #iFile, ","; uRecord.sSupplemental
    Else
        Print #iFile,
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Write_DAS_Record (End)"
    '-v1.6.1
    '
End Sub
'
' FUNCTION: iExtract_ArchiveID
' AUTHOR:   Tom Elkins
' PURPOSE:  Extracts the archive ID reference from a Tree/List View key
' INPUT:    "sKey" is the Tree/List View key value
' OUTPUT:   The archive ID
' NOTES:
Public Function iExtract_ArchiveID(sKey As String) As Integer
    Dim sTmp As String          ' Copy the string so we can tear it apart
    Dim iDelimiter As Integer   ' Location of various delimiters
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : basCCAT.iExtract_ArchiveID (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sKey
    End If
    '-v1.6.1
    '
    ' Copy the string
    sTmp = sKey
    '
    ' Look for a message delimiter
    iDelimiter = InStr(1, sTmp, SEP_MESSAGE)
    '+v1.7BB
    If iDelimiter = 0 Then iDelimiter = InStr(1, sTmp, SEP_TOC_MSG)
    '-v1.7BB
    '
    ' Move the delimiter location to the end if it does not exist
    If iDelimiter = 0 Then iDelimiter = Len(sTmp) + 1
    '
    ' Strip off the message ID
    sTmp = Left(sTmp, iDelimiter - 1)
    '
    ' Look for an archive delimiter
    iDelimiter = InStr(1, sTmp, SEP_ARCHIVE)
    '
    ' Move the delimiter location to the end if it does not exist
    If iDelimiter = 0 Then iDelimiter = Len(sTmp)
    '
    ' Strip off the database name
    sTmp = Mid(sTmp, iDelimiter + 1)
    '
    ' Convert to integer and return
    iExtract_ArchiveID = CInt(Val(sTmp))
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basCCAT.iExtract_ArchiveID (End)"
    '-v1.6.1
    '
End Function
'
' FUNCTION: iExtract_MessageID
' AUTHOR:   Tom Elkins
' PURPOSE:  Extracts the message ID reference from a Tree/List View key
' INPUT:    "sKey" is the Tree/List View key value
' OUTPUT:   The message ID
' NOTES:
Public Function iExtract_MessageID(sKey As String) As Integer
    Dim sTmp As String          ' Copy the string so we can tear it apart
    Dim iDelimiter As Integer   ' Location of various delimiters
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : basCCAT.iExtract_MessageID (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sKey
    End If
    '-v1.6.1
    '
    ' Copy the string
    sTmp = sKey
    '
    ' Look for a message delimiter
    iDelimiter = InStr(1, sTmp, SEP_MESSAGE)
    '+v1.7BB
    If iDelimiter = 0 Then iDelimiter = InStr(1, sTmp, SEP_TOC_MSG)
    '-v1.7BB
    '
    ' Move the delimiter location to the end if it does not exist
    If iDelimiter = 0 Then iDelimiter = Len(sTmp)
    '
    ' Strip off the message ID
    sTmp = Mid(sTmp, iDelimiter + 1)
    '
    ' Convert to integer and return
    iExtract_MessageID = CInt(Val(sTmp))
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basCCAT.iExtract_MessageID (End)"
    '-v1.6.1
    '
End Function
'
' FUNCTION: sTSecs_To_Human_Time
' AUTHOR:   Tom Elkins
' PURPOSE:  Converts total seconds to a human-readible format
' INPUT:    "dTSecs" is the floating point value for the total number of seconds
' OUTPUT:   String value giving the time in a DDD:HH:MM:SS.SSS format
' NOTES:
Public Function sTSecs_To_Human_Time(dTsecs As Double) As String
    Dim dRemainder As Double
    Dim iVal As Integer
    Dim sTmp As String
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : basCCAT.sTSecs_To_Human_Time (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS:" & dTsecs
    End If
    '-v1.6.1
    '
    ' Copy the number so it can be decomposed
    dRemainder = dTsecs
    '
    ' Compute the JDay
    iVal = dRemainder \ 86400#
    '
    ' Store in the string
    sTmp = Format(iVal + 1, "000") & ":"
    '
    ' Remove the day
    dRemainder = dRemainder - (iVal * 86400#)
    '
    ' Extract the hours
    iVal = dRemainder \ 3600#
    '
    ' Store in the string
    sTmp = sTmp & Format(iVal, "00") & ":"
    '
    ' Remove the hours
    dRemainder = dRemainder - (iVal * 3600#)
    '
    ' Extract the minutes
    iVal = dRemainder \ 60#
    '
    ' Store in the string
    sTmp = sTmp & Format(iVal, "00") & ":"
    '
    ' Remove the minutes
    dRemainder = dRemainder - (iVal * 60#)
    '
    ' Store in the string
    sTmp = sTmp & Format(dRemainder, "00.000")
    '
    ' Return the string
    sTSecs_To_Human_Time = sTmp
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basCCAT.sTSecs_To_Human_Time (End)"
    '-v1.6.1
    '
End Function
'
' ROUTINE:  Read_Session_File
' AUTHOR:   Tom Elkins
' PURPOSE:  Reads a list of database names that the user had open in the last session,
'           and adds them to the current session
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Sub Read_Session_File()
    Dim iFile As Integer
    Dim sDB As String
    Dim bContinue As Boolean
    '
    ' Log the event
    basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Read_Session_File (Start)"
    '
    ' Check for the existence of a session file
    If Dir(App.Path & "\" & App.EXEName & ".ses") <> "" Then
        '
        ' Find a free file ID
        iFile = FreeFile
        '
        ' Open the session file
        Open App.Path & "\" & App.EXEName & ".ses" For Input As iFile
        '
        ' Set the processing flag
        bContinue = True
        '
        ' Loop through all of the records
        While bContinue And Not EOF(iFile)
            '
            ' Read a line from the file
            Line Input #iFile, sDB
            '
            ' Check for the existence of the specified file
            On Error Resume Next
            If Dir(sDB) <> "" Then
                '
                ' Check for errors
                If Err.Number = 0 Then
                    '
                    ' Open the specified database
                    basDatabase.Open_Existing_Database sDB
                Else
                    '
                    ' Log the error
                    basCCAT.WriteLogEntry "          Invalid Session Database entry: " & sDB
                End If
            Else
                basCCAT.WriteLogEntry "          Session Database not found: " & sDB
                '
                ' Inform the user
                bContinue = (MsgBox("Cannot find database file " & sDB & vbCr & "Do you wish to continue processing the session file?", vbYesNo Or vbQuestion, "Error Processing Session File") = vbYes)
            End If
            On Error GoTo 0
        Wend
        '
        ' Close the session file
        Close iFile
    End If
    '
    '+v1.6.1TE
    basCCAT.WriteLogEntry "ROUTINE  : basCCAT.Read_Session_File (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.6TE
''+v1.5
'' FUNCTION: lGetHelpID
'' AUTHOR:   Tom Elkins
'' PURPOSE:  Extracts the help file context ID from the INI map
'' INPUT:    sName = A string appended to the prefix "IDH_" that matches an entry in the INI file
'' OUTPUT:   A long integer set to the context ID in the help file that corresponds to the Name
'' NOTES:
'Public Function lGetHelpID(sName As String) As Long
'    lGetHelpID = lGetININum("Help Map", "IDH_" & sName, 0, gsCCAT_INI_Path)
'End Function
''-v1.5
'-v1.6
'
'+v1.5
' ROUTINE:  WriteLogEntry
' AUTHOR:   Tom Elkins
' PURPOSE:  Writes the specified text to a log file
' INPUT:    "sLogText" is the text to be written
' OUTPUT:   None
' NOTES:    The entry is time-stamped
Public Sub WriteLogEntry(sLogText As String)
    '
    '+v1.6.1TE
    giLog_File = FreeFile
    Open App.Path & "\" & App.EXEName & ".log" For Append As giLog_File
    '-v1.6.1
    '
    ' Time-stamp and write the text to the log file
    '+v1.6TE
    'Print #basCCAT.giLog_File, Time; sLogText
    Print #basCCAT.giLog_File, Now; sLogText
    '-v1.6
    '
    '+v1.6.1TE
    Close giLog_File
End Sub
'-v1.5
'
'+v1.5
' FUNCTION: dHumanTimeToTSecs
' AUTHOR:   Tom Elkins
' PURPOSE:  Parses a time string in the form of "DDD:HH:MM:SS" to its components, then computes
'           TSecs
' INPUT:    "sDDDHHMMSS" is a string containing the human-readable time
' OUTPUT:   A double representing the equivalent TSecs for the specified time string
' NOTES:
Public Function dHumanTimeToTSecs(sDDDHHMMSS As String) As Double
    Dim sTimeParts() As String  ' Array to hold the components of the string
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : basCCAT.dHumanTimeToTSecs (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & sDDDHHMMSS
    End If
    '-v1.6.1
    '
    '
    sTimeParts = Split(sDDDHHMMSS, ":")
    dHumanTimeToTSecs = (CDbl(Val(sTimeParts(0))) - 1) * 86400# + (CDbl(Val(sTimeParts(1))) * 3600#) + (CDbl(Val(sTimeParts(2))) * 60#) + CDbl(Val(sTimeParts(3)))
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basCCAT.dHumanTimeToTSecs (End)"
    '-v1.6.1
    '
End Function
'-v1.5
'
'+v1.6TE
' ROUTINE:  WriteToken
' AUTHOR:   Tom Elkins
' PURPOSE:  Writes an entry to the CCAT.INI file
' INPUT:    "strSection" is a string containing the section of the INI file to write to
'           "strKey" is the key code
'           "strValue" is the value to write to the INI file
' OUTPUT:   None
' NOTES:    If the write failed, the user is informed
Public Sub WriteToken(strSection As String, strKey As String, strValue As String)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.WriteToken (Start)"
    '-v1.6.1
    '
    ' Write the string to the INI file.  If there is an error, the function returns 0 (False)
    If lPutINIString(strSection, strKey, strValue, gsCCAT_INI_Path) = 0 Then
        '
        ' Write an entry to the log file and inform the user
        basCCAT.WriteLogEntry "Error writing to token file " & gsCCAT_INI_Path & " - [" & strSection & "] " & strKey & " = " & strValue
        If MsgBox("Could not write the entry to the token file", vbRetryCancel Or vbExclamation, "Error - WriteToken") = vbRetry Then WriteToken strSection, strKey, strValue
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.WriteToken (End)"
    '-v1.6.1
    '
End Sub
'-v1.6
'
'+v1.6TE
' ROUTINE:  PopulateCCOSVersions
' AUTHOR:   Tom Elkins
' PURPOSE:  Populates a combo box with the current known list of CCOS versions
' INPUT:    "ctlList" is a ComboBox from a form that will be populated
' OUTPUT:   None - the specified combo box is altered directly
' NOTES:    This routine will be moved to the external Archive Processor, since it will
'           know what versions it can process.
Public Sub PopulateCCOSVersions(ctlList As ComboBox)
    Dim pintVersion As Integer
    Dim psngVersion As Single
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basCCAT.PopulateCCOSVersions (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & ctlList.Name
    End If
    '-v1.6.1
    '
    ' Add a blank default entry and initialize the variables
    ctlList.AddItem ""
    psngVersion = 1#
    pintVersion = 1
    '
    ' Loop while there are entries in the INI file
    While psngVersion <> 0
        '
        ' Read the version entry from the INI file
        psngVersion = CSng(Val(basCCAT.GetAlias("Versions", "CCOS" & pintVersion, 0)))
        '
        ' Add the entry to the list
        If psngVersion > 0 Then ctlList.AddItem "CCOS v" & psngVersion
        '
        ' Update the count
        pintVersion = pintVersion + 1
    Wend
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.PopulateCCOSVersions (End)"
    '-v1.6.1
    '
End Sub
'-v1.6
'
'+v1.6TE
' ROUTINE:  PopulateMessageList
' AUTHOR:   Tom Elkins
' PURPOSE:  Populates a listview control with the current known list of CCOS messages
' INPUT:    "lvMsg" is a Listview control from a form that will be populated
' OUTPUT:   None - the specified listview is altered directly
' NOTES:    This routine will be moved to the external Archive Processor, since it will
'           know what messages it can process.
Public Sub PopulateMessageList(lvMsg As ListView)
    Dim pintMsg As Integer      ' The current message loop index
    Dim pstrMsg As String       ' The message name
    Dim pstrMsgDesc As String   ' The message description
    Dim plngMsg As Long         ' The message ID
    Dim pitmMsg As ListItem     ' The pointer to the item in the list view
    Const pintSUB_ID = 1
    Const pintSUB_DESC = 2
    Const pintDEFAULT_NUM = 0
    Const pintDEFAULT_ID = -9999
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.PopulateMessageList (Start)"
    '-v1.6.1
    '
    ' Loop through all of the messages in the INI file
    For pintMsg = 1 To basCCAT.GetNumber("Message List", "CC_MESSAGES", pintDEFAULT_NUM)
        '
        ' Get the information about the message
        pstrMsg = basCCAT.GetAlias("Message List", "CC_MSG" & pintMsg, "")
        plngMsg = basCCAT.GetNumber("Message ID", pstrMsg & "ID", pintDEFAULT_ID)
        pstrMsgDesc = basCCAT.GetAlias("Message Descriptions", "CC_MSG_DESC" & plngMsg, "UNKNOWN MESSAGE DESCRIPTION")
        '
        ' See if we have a valid message
        If plngMsg <> pintDEFAULT_ID And pstrMsg <> "" Then
            '
            ' Add the message to the list view
            Set pitmMsg = lvMsg.ListItems.Add(, "MSG" & plngMsg, pstrMsg)
            pitmMsg.Checked = True
            pitmMsg.SubItems(pintSUB_ID) = plngMsg
            pitmMsg.SubItems(pintSUB_DESC) = pstrMsgDesc
        End If
    Next pintMsg
    Set pitmMsg = Nothing
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.PopulateMessageList (End)"
    '-v1.6.1
    '
End Sub
'-v1.6
'
' FUNCTION: GetNumber
' AUTHOR:   Tom Elkins
' PURPOSE:  Replaces the token file method to use INI files
' INPUT:    "Section" is the [<name>] portion of the INI file
'           "Token" is the string to be replaced
'           "DefaultValue" is the number to be returned if the token does not exist
' OUTPUT:   "GetNumber" is the number found in the INI file for the specified token
' NOTES:
Public Function GetNumber(Section As String, Token As String, DefaultValue As Long) As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "FUNCTION : basCCAT.GetNumber(" & Section & ", " & Token & ", " & DefaultValue & ")"
    '-v1.6.1
    '
    '
    ' Search the INI file for the specified token
    GetNumber = lGetININum(Section, Token, DefaultValue, gsCCAT_INI_Path)
End Function
'
' FUNCTION: GetAlias
' AUTHOR:   Tom Elkins
' PURPOSE:  Replaces the token file method to use INI files
' INPUT:    "Section" is the [<name>] portion of the INI file
'           "Token" is the string to be replaced
'           "DefaultValue" is the string to be returned if the token does not exist
' OUTPUT:   "GetAlias" is the string found in the INI file for the specified token
' NOTES:
Public Function GetAlias(Section As String, Token As String, DefaultValue As String) As String
    Dim sToken As String        ' Returned value buffer
    Dim lTokenLen As Long       ' Length of returned value
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "FUNCTION : basCCAT.GetAlias(" & Section & ", " & Token & ", " & DefaultValue & ")"
    End If
    '-v1.6.1
    '
    '
    ' Set the buffer to all spaces
    sToken = String(255, " ")
    '
    ' Get the token from the INI file
    lTokenLen = lGetINIString(Section, Token, DefaultValue, sToken, 255, gsCCAT_INI_Path)
    '
    ' Return the specified number of characters from the buffer
    GetAlias = Left(sToken, lTokenLen)
End Function
'
'+v1.6TE
' ROUTINE:  ShowHelp
' AUTHOR:   Tom Elkins
' PURPOSE:  Displays the help file at a specified topic
' INPUT:    "frmActive" is the current active form
'           "lngTopic_ID" is the ID of the help topic to be displayed
' OUTPUT:   None
' NOTES:
Public Sub ShowHelp(frmActive As Form, Optional lngTopic_ID As Long = IDH_GUI_MAIN)
    Dim lngRet As Long
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : basCCAT.ShowHelp (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & frmActive.Name & ", " & lngTopic_ID
    End If
    '-v1.6.1
    '
    '
    lngRet = HtmlHelp(frmActive.hwnd, App.HelpFile, HH_HELP_CONTEXT, lngTopic_ID)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : basCCAT.ShowHelp (End)"
    '-v1.6.1
    '
End Sub
'-v1.6
'
'+v1.6.1TE
' PROPERTY: Verbose
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets/returns the state of the verbose reporting flag
' VALUE:    Boolean value - TRUE if verbose reporting is on, FALSE if verbose reporting is off
' NOTES:    Used to determine the level of output in the log file
Public Property Get Verbose() As Boolean
    Verbose = mblnVerbose
End Property
'
Public Property Let Verbose(blnState As Boolean)
    mblnVerbose = blnState
End Property
'-v1.6.1
