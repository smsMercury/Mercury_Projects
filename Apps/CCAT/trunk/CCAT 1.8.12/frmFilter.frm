VERSION 5.00
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Filter Options"
   ClientHeight    =   4245
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   9780
   HelpContextID   =   400
   Icon            =   "frmFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSave 
      Cancel          =   -1  'True
      Caption         =   "Save"
      Height          =   375
      Left            =   1080
      TabIndex        =   98
      ToolTipText     =   "Saves the query as a Stored Query"
      Top             =   3825
      Width           =   900
   End
   Begin VB.CommandButton btnHide 
      Caption         =   "<== Hide"
      Height          =   375
      Left            =   4290
      TabIndex        =   96
      ToolTipText     =   "Hide this extension panel"
      Top             =   3825
      Width           =   1095
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "Accept"
      Height          =   375
      Left            =   6750
      TabIndex        =   93
      ToolTipText     =   "Accept the changes made"
      Top             =   3825
      Width           =   1095
   End
   Begin VB.CheckBox chkManual 
      Caption         =   "Manual SQL Entry"
      Height          =   195
      Left            =   150
      TabIndex        =   90
      ToolTipText     =   "Edit the SQL query manually"
      Top             =   2475
      Width           =   1605
   End
   Begin VB.CheckBox chkAssistant 
      Caption         =   "Use Query Assistant"
      Height          =   195
      Left            =   150
      TabIndex        =   89
      ToolTipText     =   "Get help building a query"
      Top             =   90
      Width           =   1725
   End
   Begin VB.Frame fraBuilder 
      Height          =   2310
      Left            =   45
      TabIndex        =   79
      Top             =   105
      Width           =   3225
      Begin VB.CommandButton btnEditSort 
         Caption         =   "Add"
         Height          =   300
         Left            =   2565
         TabIndex        =   88
         ToolTipText     =   "Show sort field selector"
         Top             =   1635
         Width           =   555
      End
      Begin VB.CommandButton btnEditFilter 
         Caption         =   "Add"
         Height          =   300
         Left            =   2565
         TabIndex        =   87
         ToolTipText     =   "Show filter builder"
         Top             =   960
         Width           =   555
      End
      Begin VB.CommandButton btnEditFields 
         Caption         =   "Edit"
         Height          =   300
         Left            =   2565
         TabIndex        =   86
         ToolTipText     =   "Show field selection panel"
         Top             =   285
         Width           =   555
      End
      Begin VB.TextBox txtSort 
         Height          =   570
         Left            =   540
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   85
         Text            =   "frmFilter.frx":000C
         ToolTipText     =   "Fields to sort data by"
         Top             =   1635
         Width           =   2025
      End
      Begin VB.TextBox txtFilter 
         Height          =   570
         Left            =   540
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   84
         Text            =   "frmFilter.frx":0012
         ToolTipText     =   "Conditions in which to display"
         Top             =   960
         Width           =   2025
      End
      Begin VB.TextBox txtFields 
         Height          =   570
         Left            =   540
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   83
         Text            =   "frmFilter.frx":0018
         ToolTipText     =   "List of fields to display"
         Top             =   285
         Width           =   2025
      End
      Begin VB.Label lblBuilder 
         AutoSize        =   -1  'True
         Caption         =   "Sort"
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   82
         Top             =   1665
         Width           =   285
      End
      Begin VB.Label lblBuilder 
         AutoSize        =   -1  'True
         Caption         =   "Filter"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   81
         Top             =   1005
         Width           =   330
      End
      Begin VB.Label lblBuilder 
         AutoSize        =   -1  'True
         Caption         =   "Fields"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   80
         Top             =   330
         Width           =   405
      End
   End
   Begin VB.Frame fraQuery 
      Height          =   1245
      Left            =   45
      TabIndex        =   77
      Top             =   2520
      Width           =   3225
      Begin VB.TextBox txtSQL 
         Height          =   825
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   78
         Text            =   "frmFilter.frx":0024
         ToolTipText     =   "Enter a SQL statement"
         Top             =   300
         Width           =   3075
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   3180
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   60
      Width           =   5685
      Begin VB.CheckBox chkCustom 
         Caption         =   "User-Specified Field Selection"
         Height          =   330
         Left            =   180
         TabIndex        =   92
         ToolTipText     =   "Choose your own fields from the list below"
         Top             =   1065
         Width           =   2415
      End
      Begin VB.CheckBox chkPredefined 
         Caption         =   "Use Predefined Field Selections"
         Height          =   300
         Left            =   180
         TabIndex        =   91
         ToolTipText     =   "Choose one of the predefined field lists"
         Top             =   0
         Width           =   2550
      End
      Begin VB.Frame fraFields1 
         Height          =   570
         Left            =   60
         TabIndex        =   34
         Top             =   45
         Width           =   5580
         Begin VB.OptionButton optField 
            Caption         =   "Signal"
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   38
            ToolTipText     =   "Select fields for a DAS signal file"
            Top             =   255
            Width           =   885
         End
         Begin VB.OptionButton optField 
            Caption         =   "Event"
            Height          =   225
            Index           =   1
            Left            =   1095
            TabIndex        =   37
            ToolTipText     =   "Choose fields for a DAS Event file"
            Top             =   255
            Width           =   990
         End
         Begin VB.OptionButton optField 
            Caption         =   "Moving Target"
            Height          =   225
            Index           =   2
            Left            =   2115
            TabIndex        =   36
            ToolTipText     =   "Choose fields for a DAS MTF"
            Top             =   255
            Width           =   1365
         End
         Begin VB.OptionButton optField 
            Caption         =   "Stationary Target"
            Height          =   225
            Index           =   3
            Left            =   3645
            TabIndex        =   35
            ToolTipText     =   "Choose fields for a DAS STF"
            Top             =   255
            Width           =   1785
         End
      End
      Begin VB.Frame fraColumns 
         Caption         =   "Col"
         Height          =   2565
         Left            =   60
         TabIndex        =   44
         Top             =   1140
         Width           =   5580
         Begin VB.CheckBox chkField 
            Caption         =   "Report Time"
            Height          =   285
            Index           =   0
            Left            =   225
            TabIndex        =   76
            Top             =   285
            Width           =   1170
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Message Type"
            Height          =   285
            Index           =   1
            Left            =   225
            TabIndex        =   75
            Top             =   570
            Width           =   1365
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Report Type"
            Height          =   285
            Index           =   2
            Left            =   225
            TabIndex        =   74
            Top             =   855
            Width           =   1185
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Origin"
            Height          =   285
            Index           =   3
            Left            =   225
            TabIndex        =   73
            Top             =   1125
            Width           =   1230
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Origin ID"
            Height          =   285
            Index           =   4
            Left            =   225
            TabIndex        =   72
            Top             =   1395
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Target ID"
            Height          =   285
            Index           =   5
            Left            =   225
            TabIndex        =   71
            Top             =   1680
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Latitude"
            Height          =   285
            Index           =   6
            Left            =   225
            TabIndex        =   70
            Top             =   1965
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Longitude"
            Height          =   285
            Index           =   7
            Left            =   225
            TabIndex        =   69
            Top             =   2250
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Altitude"
            Height          =   285
            Index           =   8
            Left            =   1695
            TabIndex        =   68
            Top             =   285
            Width           =   1170
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Heading"
            Height          =   285
            Index           =   9
            Left            =   1695
            TabIndex        =   67
            Top             =   570
            Width           =   1170
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Speed"
            Height          =   285
            Index           =   10
            Left            =   1695
            TabIndex        =   66
            Top             =   855
            Width           =   1170
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Parent"
            Height          =   285
            Index           =   11
            Left            =   1695
            TabIndex        =   65
            Top             =   1125
            Width           =   1170
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Parent ID"
            Height          =   285
            Index           =   12
            Left            =   1695
            TabIndex        =   64
            Top             =   1395
            Width           =   1170
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Allegiance"
            Height          =   285
            Index           =   13
            Left            =   1695
            TabIndex        =   63
            Top             =   1680
            Width           =   1170
         End
         Begin VB.CheckBox chkField 
            Caption         =   "IFF Code"
            Height          =   285
            Index           =   14
            Left            =   1695
            TabIndex        =   62
            Top             =   1965
            Width           =   1170
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Emitter"
            Height          =   285
            Index           =   15
            Left            =   1695
            TabIndex        =   61
            Top             =   2250
            Width           =   1170
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Emitter ID"
            Height          =   285
            Index           =   16
            Left            =   2865
            TabIndex        =   60
            Top             =   285
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Signal"
            Height          =   285
            Index           =   17
            Left            =   2865
            TabIndex        =   59
            Top             =   570
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Signal ID"
            Height          =   285
            Index           =   18
            Left            =   2865
            TabIndex        =   58
            Top             =   855
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Frequency"
            Height          =   285
            Index           =   19
            Left            =   2865
            TabIndex        =   57
            Top             =   1125
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "PRI"
            Height          =   285
            Index           =   20
            Left            =   2865
            TabIndex        =   56
            Top             =   1395
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Status"
            Height          =   285
            Index           =   21
            Left            =   2865
            TabIndex        =   55
            Top             =   1680
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Tag"
            Height          =   285
            Index           =   22
            Left            =   2865
            TabIndex        =   54
            Top             =   1965
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Flag"
            Height          =   285
            Index           =   23
            Left            =   2865
            TabIndex        =   53
            Top             =   2250
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Common ID"
            Height          =   285
            Index           =   24
            Left            =   4110
            TabIndex        =   52
            Top             =   285
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Range"
            Height          =   285
            Index           =   25
            Left            =   4110
            TabIndex        =   51
            Top             =   570
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Bearing"
            Height          =   285
            Index           =   26
            Left            =   4110
            TabIndex        =   50
            Top             =   855
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Elevation"
            Height          =   285
            Index           =   27
            Left            =   4110
            TabIndex        =   49
            Top             =   1125
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "XX"
            Height          =   285
            Index           =   28
            Left            =   4110
            TabIndex        =   48
            Top             =   1395
            Width           =   1215
         End
         Begin VB.CheckBox chkField 
            Caption         =   "XY"
            Height          =   285
            Index           =   29
            Left            =   4110
            TabIndex        =   47
            Top             =   1680
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "YY"
            Height          =   285
            Index           =   30
            Left            =   4110
            TabIndex        =   46
            Top             =   1965
            Width           =   1245
         End
         Begin VB.CheckBox chkField 
            Caption         =   "Supplemental"
            Height          =   285
            Index           =   31
            Left            =   4110
            TabIndex        =   45
            Top             =   2250
            Width           =   1260
         End
      End
      Begin VB.Frame Frame1 
         Height          =   510
         Left            =   60
         TabIndex        =   39
         Top             =   495
         Width           =   5580
         Begin VB.OptionButton optRpt 
            Caption         =   "Geolocations/Fixes"
            Height          =   225
            Index           =   3
            Left            =   3645
            TabIndex        =   43
            ToolTipText     =   "Choose fields for a GEO report"
            Top             =   180
            Width           =   1725
         End
         Begin VB.OptionButton optRpt 
            Caption         =   "Lines of Bearing"
            Height          =   225
            Index           =   2
            Left            =   2115
            TabIndex        =   42
            ToolTipText     =   "Choose fields for a VEC report"
            Top             =   180
            Width           =   1440
         End
         Begin VB.OptionButton optRpt 
            Caption         =   "Tracks"
            Height          =   225
            Index           =   1
            Left            =   1095
            TabIndex        =   41
            ToolTipText     =   "Choose fields for a TRK report"
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton optRpt 
            Caption         =   "Dots"
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   40
            ToolTipText     =   "Choose fields for a DOT report"
            Top             =   180
            Width           =   885
         End
      End
   End
   Begin VB.CommandButton btnExecute 
      Caption         =   "Execute"
      Height          =   375
      Left            =   2115
      TabIndex        =   4
      ToolTipText     =   "Executes the query shown above"
      Top             =   3825
      Width           =   900
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   150
      Width           =   5685
      Begin VB.ComboBox cmbSortField 
         Height          =   315
         Left            =   2265
         TabIndex        =   32
         Text            =   "Combo1"
         ToolTipText     =   "Select a field to sort by"
         Top             =   1875
         Width           =   1545
      End
      Begin VB.CheckBox chkSort 
         Caption         =   "Then sort by this field"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   1215
         Width           =   1980
      End
      Begin VB.CheckBox chkSort 
         Caption         =   "Then sort by this field"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   705
         Width           =   1980
      End
      Begin VB.CheckBox chkSort 
         Caption         =   "Sort records by this field"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   225
         Width           =   1980
      End
      Begin VB.Label lblSortField 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   315
         Index           =   2
         Left            =   2295
         TabIndex        =   16
         Top             =   1185
         Width           =   1545
      End
      Begin VB.Label lblSortField 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   315
         Index           =   1
         Left            =   2295
         TabIndex        =   14
         Top             =   675
         Width           =   1545
      End
      Begin VB.Label lblSortField 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         Height          =   315
         Index           =   0
         Left            =   2295
         TabIndex        =   12
         Top             =   195
         Width           =   1545
      End
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Erases the query"
      Top             =   3825
      Width           =   900
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   5685
      Begin VB.Frame fraWhere 
         Height          =   3660
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   75
         Width           =   5520
         Begin VB.CheckBox chkWhere 
            Caption         =   "Select only records that meet the following criteria"
            Height          =   225
            Left            =   165
            TabIndex        =   97
            Top             =   225
            Width           =   3795
         End
         Begin VB.CheckBox chkOR1 
            Caption         =   "OR"
            Height          =   285
            Left            =   825
            TabIndex        =   95
            Top             =   180
            Width           =   540
         End
         Begin VB.CheckBox chkAND1 
            Caption         =   "AND"
            Height          =   285
            Left            =   165
            TabIndex        =   94
            Top             =   195
            Width           =   645
         End
         Begin VB.CheckBox chkAnd4 
            Caption         =   "AND"
            Height          =   285
            Left            =   165
            TabIndex        =   28
            Top             =   2880
            Width           =   645
         End
         Begin VB.CheckBox chkOR4 
            Caption         =   "OR"
            Height          =   285
            Left            =   825
            TabIndex        =   27
            Top             =   2880
            Width           =   540
         End
         Begin VB.CheckBox chkAnd3 
            Caption         =   "AND"
            Height          =   285
            Left            =   165
            TabIndex        =   23
            Top             =   1980
            Width           =   645
         End
         Begin VB.CheckBox chkOR3 
            Caption         =   "OR"
            Height          =   285
            Left            =   825
            TabIndex        =   22
            Top             =   1980
            Width           =   540
         End
         Begin VB.CheckBox chkAnd2 
            Caption         =   "AND"
            Height          =   285
            Left            =   165
            TabIndex        =   18
            Top             =   1080
            Width           =   645
         End
         Begin VB.CheckBox chkOR2 
            Caption         =   "OR"
            Height          =   285
            Left            =   825
            TabIndex        =   17
            Top             =   1080
            Width           =   540
         End
         Begin VB.TextBox txtValue 
            Height          =   315
            Left            =   2970
            TabIndex        =   7
            Text            =   "txtValue"
            ToolTipText     =   "Enter the value for the field"
            Top             =   540
            Width           =   2445
         End
         Begin VB.ComboBox cmbOperator 
            Height          =   315
            Left            =   1785
            TabIndex        =   6
            Text            =   "cmbOperator"
            ToolTipText     =   "Choose an operator"
            Top             =   540
            Width           =   1110
         End
         Begin VB.ComboBox cmbFields 
            Height          =   315
            Left            =   165
            TabIndex        =   5
            Text            =   "cmbFields"
            ToolTipText     =   "Select a field"
            Top             =   540
            Width           =   1545
         End
         Begin VB.Label lblWhereField4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblWhereField4"
            Height          =   315
            Left            =   165
            TabIndex        =   31
            Top             =   3240
            Width           =   1545
         End
         Begin VB.Label lblCond4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCond4"
            Height          =   315
            Left            =   1785
            TabIndex        =   30
            Top             =   3240
            Width           =   1110
         End
         Begin VB.Label lblVal4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVal4"
            Height          =   315
            Left            =   2970
            TabIndex        =   29
            Top             =   3240
            Width           =   2445
         End
         Begin VB.Label lblVal3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVal3"
            Height          =   315
            Left            =   2970
            TabIndex        =   26
            Top             =   2340
            Width           =   2445
         End
         Begin VB.Label lblCond3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCond3"
            Height          =   315
            Left            =   1785
            TabIndex        =   25
            Top             =   2340
            Width           =   1110
         End
         Begin VB.Label lblWhereField3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblWhereField3"
            Height          =   315
            Left            =   165
            TabIndex        =   24
            Top             =   2340
            Width           =   1545
         End
         Begin VB.Label lblWhereField2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblWhereField2"
            Height          =   315
            Left            =   165
            TabIndex        =   21
            Top             =   1440
            Width           =   1545
         End
         Begin VB.Label lblCond2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCond2"
            Height          =   315
            Left            =   1785
            TabIndex        =   20
            Top             =   1440
            Width           =   1110
         End
         Begin VB.Label lblVal2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVal2"
            Height          =   315
            Left            =   2970
            TabIndex        =   19
            Top             =   1440
            Width           =   2445
         End
         Begin VB.Label lblVal1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblVal1"
            Height          =   315
            Left            =   2970
            TabIndex        =   10
            Top             =   540
            Width           =   2445
         End
         Begin VB.Label lblCond1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblCond1"
            Height          =   315
            Left            =   1785
            TabIndex        =   9
            Top             =   540
            Width           =   1110
         End
         Begin VB.Label lblWhereField1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblWhereField2"
            Height          =   315
            Left            =   165
            TabIndex        =   8
            Top             =   540
            Width           =   1545
         End
      End
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' COPYRIGHT (C) 1999-2001 Mercury Solutions, Inc.
'
' FORM:     frmFilter
' AUTHOR:   Tom Elkins
' PURPOSE:  Provide an interface for the user to filter the data
' REVISIONS:
'   v1.3.1  TAE Added automatic placement of quotes around text values
'   v1.3.2  TAE Removed automatic placement of quotes around text values
'   v1.4.0  TAE Added button to save queries; however, code does not work
'           TAE Tied some message boxes to help file and added help button
'   v1.5.0  TAE Added context-sensitive help to the form.  Pressing F1 will bring up the help
'               file to the page for the current form, including the assistants.
'           TAE Repaired the code to save queries.
'           TAE Moved the clear query routine to a public sub so it can be called externally
'   v1.6.0  TAE Updated variables names to comply with programming standard
'           TAE Added constants for DAS fields and file types
'           TAE Updated Event format field selection
'   v1.6.1  TAE Added verbose logging calls
Option Explicit
'
' Form Constants
Const mintFILTER_MIN_WIDTH = 3400   ' Minimum width for the form
Const mintFILTER_MAX_WIDTH = 9000   ' Maximum width for the form
Const mintFILTER_POS_HIDE = -20000  ' "Hidden" position for form extensions
Const mintFILTER_POS_SHOW = 3180    ' "Visible" position for form extensions
Const mintPIC_FIELDS = 1            ' Index for the field selection extension
Const mintPIC_FILTER = 2            ' Index for the filter creation extension
Const mintPIC_SORT = 3              ' Index for the sort creation extension
Const mintNONE = 0
'
' Form-level Variables
Dim mblnModify_SQL As Boolean       ' Flag to indicate that the SQL query should be modified
Dim mblnModify_Fields As Boolean    ' Flag to indicate that the field list should be modified
Dim mblnModify_Filter As Boolean    ' Flag to indicate that the filter should be modified
Dim mblnModify_Sort As Boolean      ' Flag to indicate that the sort list should be modified
Dim mintCurrent_Pic As Integer      ' Currently selected extension
'
' EVENT:    btnAccept_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Process the changes made by the user
' TRIGGER:  User clicked on the "Accept" button
' INPUT:    None
' OUTPUT:   None
' NOTES:    Actions are based on the currently selected extension
Private Sub btnAccept_Click()
    Dim pintField As Integer            ' Current field index
    Dim pstrField_List As String        ' Selected field list
    Dim pintNum_Not_Selected As Integer ' Number of items not selected
    Dim pintNum_Selected As Integer     ' Number of items selected
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnAccept Click (Start)"
    '-v1.6.1
    '
    ' Action is determined by the currently selected extension
    Select Case mintCurrent_Pic
        '
        ' Field list selection
        Case mintPIC_FIELDS:
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmFilter.btnAccept Click (Field Selection)"
            '-v1.6.1
            '
            '
            ' Initialize the variables
            pintNum_Not_Selected = mintNONE
            pintNum_Selected = mintNONE
            pstrField_List = ""
            '
            ' Loop through the field check box array
            For pintField = frmFilter.chkField.LBound To frmFilter.chkField.UBound
                '
                ' See if the user selected the field
                If frmFilter.chkField(pintField).Value = vbChecked Then
                    '
                    ' Add the field to the list and increment the counter
                    pstrField_List = pstrField_List & frmFilter.chkField(pintField).Tag & ", "
                    pintNum_Selected = pintNum_Selected + 1
                Else
                    '
                    ' Keep track of the number of fields not selected
                    pintNum_Not_Selected = pintNum_Not_Selected + 1
                End If
            Next pintField
            '
            ' Make sure some fields were checked
            If pintNum_Selected > mintNONE Then
                '
                ' If any fields were selected, there will be an extraneous ", " at the
                ' beginning of the field list.  Remove the extra characters.
                pstrField_List = Left(pstrField_List, Len(pstrField_List) - 2)
                '
                ' If all fields were selected, use the SQL shortcut
                If pintNum_Not_Selected = mintNONE Then pstrField_List = "*"
                '
                ' Indicate that the field list should be modified
                mblnModify_Fields = True
                '
                ' Replace the existing field list with the new one
                frmFilter.txtFields.Text = pstrField_List
            End If
            '
            ' Reset the field modification flag
            mblnModify_Fields = False
        '
        ' Filter modification extension
        Case mintPIC_FILTER:
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmFilter.btnAccept Click (Filter Selection)"
            '-v1.6.1
            '
            ' Set the filter modification flag
            mblnModify_Filter = True
            '
            ' See if this first clause is used
            If frmFilter.chkWhere.Value = vbChecked And frmFilter.chkWhere.Visible Then
                '
                ' Make sure there is a clause
                If Len(frmFilter.lblWhereField1.Caption) > 0 And Len(frmFilter.lblCond1.Caption) > 0 And Len(frmFilter.lblVal1.Caption) > 0 Then
                    '
                    ' Set the filter to the clause
                    frmFilter.txtFilter.Text = frmFilter.lblWhereField1.Caption & " " & frmFilter.lblCond1.Caption & " " & frmFilter.lblVal1.Caption
                End If
            End If
            '
            ' See if this is an addition to the filter string
            If frmFilter.chkAND1.Value = vbChecked And frmFilter.chkAND1.Visible Then
                '
                ' Make sure there is a clause
                If Len(frmFilter.lblWhereField1.Caption) > 0 And Len(frmFilter.lblCond1.Caption) > 0 And Len(frmFilter.lblVal1.Caption) > 0 Then
                    '
                    ' Set the filter to the clause
                    frmFilter.txtFilter.Text = frmFilter.txtFilter.Text & " AND " & frmFilter.lblWhereField1.Caption & " " & frmFilter.lblCond1.Caption & " " & frmFilter.lblVal1.Caption
                End If
            End If
            '
            If frmFilter.chkOR1.Value = vbChecked And frmFilter.chkOR1.Visible Then
                '
                ' Make sure there is a clause
                If Len(frmFilter.lblWhereField1.Caption) > 0 And Len(frmFilter.lblCond1.Caption) > 0 And Len(frmFilter.lblVal1.Caption) > 0 Then
                    '
                    ' Set the filter to the clause
                    frmFilter.txtFilter.Text = frmFilter.txtFilter.Text & " OR " & frmFilter.lblWhereField1.Caption & " " & frmFilter.lblCond1.Caption & " " & frmFilter.lblVal1.Caption
                End If
            End If
            '
            ' See if the second clause is used
            If frmFilter.chkAnd2.Value = vbChecked And frmFilter.chkAnd2.Enabled Then
                '
                ' Make sure there is a clause
                If Len(frmFilter.lblWhereField2.Caption) > 0 And Len(frmFilter.lblCond2.Caption) > 0 And Len(frmFilter.lblVal2.Caption) > 0 Then
                    '
                    ' Set the filter to the clause
                    frmFilter.txtFilter.Text = frmFilter.txtFilter.Text & " AND " & frmFilter.lblWhereField2.Caption & " " & frmFilter.lblCond2.Caption & " " & frmFilter.lblVal2.Caption
                End If
            End If
            '
            If frmFilter.chkOR2.Value = vbChecked And frmFilter.chkOR2.Enabled Then
                '
                ' Make sure there is a clause
                If Len(frmFilter.lblWhereField2.Caption) > 0 And Len(frmFilter.lblCond2.Caption) > 0 And Len(frmFilter.lblVal2.Caption) > 0 Then
                    '
                    ' Set the filter to the clause
                    frmFilter.txtFilter.Text = frmFilter.txtFilter.Text & " OR " & frmFilter.lblWhereField2.Caption & " " & frmFilter.lblCond2.Caption & " " & frmFilter.lblVal2.Caption
                End If
            End If
            '
            ' See if the third clause is used
            If frmFilter.chkAnd3.Value = vbChecked And frmFilter.chkAnd3.Enabled Then
                '
                ' Make sure there is a clause
                If Len(frmFilter.lblWhereField3.Caption) > 0 And Len(frmFilter.lblCond3.Caption) > 0 And Len(frmFilter.lblVal3.Caption) > 0 Then
                    '
                    ' Set the filter to the clause
                    frmFilter.txtFilter.Text = frmFilter.txtFilter.Text & " AND " & frmFilter.lblWhereField3.Caption & " " & frmFilter.lblCond3.Caption & " " & frmFilter.lblVal3.Caption
                End If
            End If
            '
            If frmFilter.chkOR3.Value = vbChecked And frmFilter.chkOR3.Enabled Then
                '
                ' Make sure there is a clause
                If Len(frmFilter.lblWhereField3.Caption) > 0 And Len(frmFilter.lblCond3.Caption) > 0 And Len(frmFilter.lblVal3.Caption) > 0 Then
                    '
                    ' Set the filter to the clause
                    frmFilter.txtFilter.Text = frmFilter.txtFilter.Text & " AND " & frmFilter.lblWhereField3.Caption & " " & frmFilter.lblCond3.Caption & " " & frmFilter.lblVal3.Caption
                End If
            End If
            '
            ' See if the fourth clause is used
            If frmFilter.chkAnd4.Value = vbChecked And frmFilter.chkAnd4.Enabled Then
                '
                ' Make sure there is a clause
                If Len(frmFilter.lblWhereField4.Caption) > 0 And Len(frmFilter.lblCond4.Caption) > 0 And Len(frmFilter.lblVal4.Caption) > 0 Then
                    '
                    ' Set the filter to the clause
                    frmFilter.txtFilter.Text = frmFilter.txtFilter.Text & " AND " & frmFilter.lblWhereField4.Caption & " " & frmFilter.lblCond4.Caption & " " & frmFilter.lblVal4.Caption
                End If
            End If
            '
            If frmFilter.chkOR4.Value = vbChecked And frmFilter.chkOR4.Enabled Then
                '
                ' Make sure there is a clause
                If Len(frmFilter.lblWhereField4.Caption) > 0 And Len(frmFilter.lblCond4.Caption) > 0 And Len(frmFilter.lblVal4.Caption) > 0 Then
                    '
                    ' Set the filter to the clause
                    frmFilter.txtFilter.Text = frmFilter.txtFilter.Text & " AND " & frmFilter.lblWhereField4.Caption & " " & frmFilter.lblCond4.Caption & " " & frmFilter.lblVal4.Caption
                End If
            End If
            '
            ' Reset the filter modification flag
            mblnModify_Filter = False
        '
        ' Sort list creation extension
        Case mintPIC_SORT:
            '
            '+v1.6.1TE
            If basCCAT.Verbose Then basCCAT.WriteLogEntry "INFO     : frmFilter.btnAccept Click (Sort Selection)"
            '-v1.6.1
            '
            ' Set the sort string modification flag
            mblnModify_Sort = True
            '
            ' Loop through the sort check box array
            For pintField = frmFilter.chkSort.LBound To frmFilter.chkSort.UBound
                '
                ' See if the box was checked and enabled and there is a field
                If frmFilter.chkSort(pintField).Value = vbChecked And frmFilter.chkSort(pintField).Enabled And Len(frmFilter.lblSortField(pintField).Caption) > 0 Then
                    '
                    ' Add the field to the list
                    frmFilter.txtSort.Text = frmFilter.txtSort.Text & ", " & frmFilter.lblSortField(pintField).Caption
                End If
            Next pintField
            '
            ' See if there is extraneous characters at the beginning of the list
            If Mid(frmFilter.txtSort.Text, 1, 1) = "," Then
                '
                ' Remove the extra characters
                frmFilter.txtSort.Text = Mid(frmFilter.txtSort.Text, 3)
            End If
            '
            ' Reset the sort modification flag
            mblnModify_Sort = False
    End Select
    '
    '+v1.5
    ' Change the help context to point to the basic form
    Me.HelpContextID = basCCAT.IDH_GUI_FILTER
    '-v1.5
    '
    ' Hide the extension and return the form to normal
    frmFilter.Shrink_Window
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT     : frmFilter.btnAccept Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnClear_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Deletes the existing SQL query
' TRIGGER:  User clicked on the "Clear" button
' INPUT:    None
' OUTPUT:   None
' NOTES:    This wipes out the SQL query that was used to generate the
'           current dataset.  The user will have to rerun the query by
'           clicking on the node again.
Private Sub btnClear_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnClear Click (Start)"
    '-v1.6.1
    '
    '+v1.5
    frmFilter.ClearQuery
    '    '
    '    ' Blank out all strings
    '    frmFilter.txtFields.Text = ""
    '    guCurrent.uSQL.sFields = ""
    '    frmFilter.txtFilter.Text = ""
    '    guCurrent.uSQL.sFilter = ""
    '    frmFilter.txtSort.Text = ""
    '    guCurrent.uSQL.sOrder = ""
    '    frmFilter.txtSQL.Text = ""
    '    guCurrent.uSQL.sQuery = ""
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnClear Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnEditFields_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Show the field selection extension
' TRIGGER:  User clicked on the "Edit" button next to the Field list
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnEditFields_Click()
    Dim pintField As Integer   ' Current field index
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnEditFields Click (Start)"
    '-v1.6.1
    '
    '
    ' Reset the field extension
    frmFilter.chkCustom.Value = vbUnchecked
    frmFilter.chkPredefined.Value = vbUnchecked
    '
    ' Loop through the field check boxes
    For pintField = frmFilter.chkField.LBound To frmFilter.chkField.UBound
        '
        ' Look for the field name in the current field list
        If InStr(1, UCase(guCurrent.uSQL.sFields), UCase(frmFilter.chkField(pintField).Tag)) > 0 Or guCurrent.uSQL.sFields = "*" Then
            '
            ' Mark the field as selected
            frmFilter.chkField(pintField).Value = vbChecked
        Else
            '
            ' Mark the field as unselected
            frmFilter.chkField(pintField).Value = vbUnchecked
        End If
    Next pintField
    '
    ' Select the field selection extension
    mintCurrent_Pic = mintPIC_FIELDS
    '
    '+v1.5
    ' Change help context to point to the field assistant page
    Me.HelpContextID = basCCAT.IDH_GUI_FILTER_FIELDS
    '-v1.5
    '
    ' Modify the form to show the extension
    frmFilter.Expand_Window
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnEditFields Click (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Expand_Window
' AUTHOR:   Tom Elkins
' PURPOSE:  Alter the form to include an extension and disable the controls
'           on the main part of the form
' TRIGGER:  The user clicked on the "Edit" button next to the field list
'           The user clicked on the "Add" button next to the filter list
'           The user clicked on the "Add" button next to the sort list
' INPUT:    None
' OUTPUT:   None
' NOTES:
Friend Sub Expand_Window()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmFilter.Expand_Window (Start)"
    '-v1.6.1
    '
    ' Disable the controls on the main part of the form
    frmFilter.chkAssistant.Enabled = False
    frmFilter.lblBuilder(0).Enabled = False
    frmFilter.lblBuilder(1).Enabled = False
    frmFilter.lblBuilder(2).Enabled = False
    frmFilter.txtFields.Enabled = False
    frmFilter.txtFilter.Enabled = False
    frmFilter.txtSort.Enabled = False
    frmFilter.btnEditFields.Enabled = False
    frmFilter.btnEditFilter.Enabled = False
    frmFilter.btnEditSort.Enabled = False
    frmFilter.chkManual.Enabled = False
    frmFilter.btnClear.Enabled = False
    frmFilter.btnExecute.Enabled = False
    frmFilter.btnAccept.Enabled = True
    '
    ' Move the requested extension into place
    frmFilter.picOptions(mintCurrent_Pic).Left = mintFILTER_POS_SHOW
    '
    ' Resize the form to show the extension
    frmFilter.Width = mintFILTER_MAX_WIDTH
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmFilter.Expand_Window (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Shrink_Window
' AUTHOR:   Tom Elkins
' PURPOSE:  Return the form to its normal size and re-enable the controls
'           on the main part of the form
' TRIGGER:  The user clicked on the "Accept" button
'           The user clicked on the "Hide" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Friend Sub Shrink_Window()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmFilter.Shrink_Window (Start)"
    '-v1.6.1
    '
    ' Re-enable the controls
    frmFilter.chkAssistant.Enabled = True
    frmFilter.lblBuilder(0).Enabled = True
    frmFilter.lblBuilder(1).Enabled = True
    frmFilter.lblBuilder(2).Enabled = True
    frmFilter.txtFields.Enabled = True
    frmFilter.txtFilter.Enabled = True
    frmFilter.txtSort.Enabled = True
    frmFilter.btnEditFields.Enabled = True
    frmFilter.btnEditFilter.Enabled = True
    frmFilter.btnEditSort.Enabled = True
    frmFilter.chkManual.Enabled = True
    frmFilter.btnClear.Enabled = True
    frmFilter.btnExecute.Enabled = True
    frmFilter.btnAccept.Enabled = False
    '
    ' Hide the extension
    frmFilter.picOptions(mintCurrent_Pic).Left = mintFILTER_POS_HIDE
    '
    ' Resize the window back to its original size
    frmFilter.Width = mintFILTER_MIN_WIDTH
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.Shrink_Window (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnEditFilter_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Show the filter modification extension
' TRIGGER:  The user clicked on the "Add" button next to the filter list
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnEditFilter_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnEditFilter Click (Start)"
    '-v1.6.1
    '
    ' Determine if there is an existing filter clause
    If Len(frmFilter.txtFilter) > 0 Then
        '
        ' Make the AND/OR check boxes visible and hide the WHERE box
        frmFilter.chkWhere.Visible = False
        frmFilter.chkAND1.Visible = True
        frmFilter.chkOR1.Visible = True
    Else
        '
        ' Make the WHERE box visible and hide the first set of AND/OR check boxes
        frmFilter.chkWhere.Visible = True
        frmFilter.chkAND1.Visible = False
        frmFilter.chkOR1.Visible = False
    End If
    '
    ' Select the filter modification extension
    mintCurrent_Pic = mintPIC_FILTER
    '
    '+v1.5
    ' Change help context to point to the filter assistant page
    Me.HelpContextID = basCCAT.IDH_GUI_FILTER_FILTER
    '-v1.5
    '
    ' Display the extension
    frmFilter.Expand_Window
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnEditFilter Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnEditSort_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Show the sort list modification extension
' TRIGGER:  The user clicked on the "Add" button next to the sort list
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnEditSort_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnEditSort Click (Start)"
    '-v1.6.1
    '
    ' See if there is an existing sort list
    If Len(frmFilter.txtSort.Text) > 0 Then
        '
        ' Change the caption of the first check box
        frmFilter.chkSort(0).Caption = "Then sort by this field"
    Else
        '
        ' Change the caption of the first check box
        frmFilter.chkSort(0).Caption = "Sort records by this field"
    End If
    '
    ' Select the sort modification extension
    mintCurrent_Pic = mintPIC_SORT
    '
    '+v1.5
    ' Change help context to point to the sort assistant page
    Me.HelpContextID = basCCAT.IDH_GUI_FILTER_SORT
    '-v1.5
    '
    ' Display the extension
    frmFilter.Expand_Window
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnEditSort Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnExecute_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Execute the current SQL query and update the DBGrid
' TRIGGER:  The user clicked on the "Execute" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnExecute_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnExecute Click (Start)"
    '-v1.6.1
    '
    ' Trap any error
    On Error GoTo Bad_SQL
    '
    ' Change the mouse to the "busy" look
    Screen.MousePointer = vbHourglass
    '
    ' Execute the query
    frmMain.Data1.RecordSource = guCurrent.uSQL.sQuery
    '
    ' Update the display
    frmMain.Data1.Refresh
    '
    ' See if there were any records returned
    If frmMain.Data1.Recordset.RecordCount > 0 Then
        '
        ' Move to the end of the record set to get an accurate record count
        frmMain.Data1.Recordset.MoveLast
        frmMain.Data1.Recordset.MoveFirst
    End If
    '
    ' Return the mouse pointer to normal
    Screen.MousePointer = vbDefault
    '
    ' Update the status bar
    '+v1.5
    'frmMain.sbStatusBar.Panels(1).Text = frmMain.Data1.Recordset.RecordCount & " records retrieved from database"
    frmMain.UpdateStatusText frmMain.Data1.Recordset.RecordCount & " records retrieved from database"
    '-v1.5
    '
    ' Hide the filter form
    frmFilter.Hide
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnExecute Click (End)"
    '-v1.6.1
    '
    ' Leave the subroutine
    Exit Sub
'
' Error handler
Bad_SQL:
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmFilter.btnEditSort Click"
    '-v1.6.1
    '
    ' Restore error reporting
    On Error GoTo 0
    '
    ' Restore the mouse
    Screen.MousePointer = vbDefault
    '
    ' Inform the user there was an error
    '+v1.5
    'MsgBox "Syntax error in the SQL query.  Please correct and retry.", vbOKOnly Or vbMsgBoxHelpButton, "SQL Error", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, IDH_FilterData
    MsgBox "Syntax error in the SQL query.  Please correct and retry.", vbOKOnly Or vbMsgBoxHelpButton, "SQL Error", App.Path & DAS_HELP_PATH & CCAT_HELP_FILE, basCCAT.IDH_DB_FILTERING
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : frmFilter.btnEditSort Click (" & Me.txtSQL.Text & ")"
    '-v1.6.1
    '
    '
    ' Return the user to the form
    If frmFilter.chkAssistant.Value = vbChecked Then
        '
        ' Move the cursor to the field selection box
        frmFilter.txtFields.SetFocus
    Else
        '
        ' Highlight the SQL query text and allow the user to edit it
        frmFilter.txtSQL.SelStart = 0
        frmFilter.txtSQL.SetFocus
    End If
End Sub
'
' EVENT:    btnHide_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Hide the form extension and return the form to its original shape
' TRIGGER:  The user clicked on the "Hide" button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnHide_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnHide Click (Start)"
    '-v1.6.1
    '
    '+v1.5
    ' Change help context to point to the basic form
    Me.HelpContextID = basCCAT.IDH_GUI_FILTER
    '-v1.5
    '
    ' Return the window to its original shape and re-enable the controls
    frmFilter.Shrink_Window
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnHide Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    btnSave_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Saves a query in the token file for later re-use
' TRIGGER:  The user clicked on the Save button
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub btnSave_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnSave Click (Start)"
    '-v1.6.1
    '
    If MsgBox("Are you sure you want to save this query?", vbYesNo, "Save Query") = vbYes Then
        frmFilter.SaveCustomFilter
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.btnSave Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkAnd1_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the filter extension for the changes
' TRIGGER:  The user clicked on the first AND check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkAnd1_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAnd1 Click (Start)"
    '-v1.6.1
    '
    ' See if the item was checked
    If frmFilter.chkAND1.Value = vbChecked Then
        '
        ' Uncheck the OR box
        frmFilter.chkOR1.Value = vbUnchecked
        '
        ' Enable the AND/OR boxes on the next line
        frmFilter.chkAnd2.Enabled = True
        frmFilter.chkOR2.Enabled = True
        '
        ' Enable the labels for the first clause
        frmFilter.lblWhereField1.Enabled = True
        frmFilter.lblCond1.Enabled = True
        frmFilter.lblVal1.Enabled = True
        '
        ' Move the field selection combo box over the field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField1.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField1.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField1.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator selection combo box over the operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond1.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond1.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond1.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the value label
        frmFilter.txtValue.Left = frmFilter.lblVal1.Left
        frmFilter.txtValue.Top = frmFilter.lblVal1.Top
        frmFilter.txtValue.Text = frmFilter.lblVal1.Caption
        frmFilter.txtValue.Visible = True
    Else
        '
        ' Uncheck the AND and OR boxes, and reconfigure the form
        frmFilter.chkAnd2.Value = vbUnchecked
        chkAnd2_Click
        frmFilter.chkOR2.Value = vbUnchecked
        chkOR2_Click
        '
        ' Disable the second AND/OR boxes
        frmFilter.chkAnd2.Enabled = False
        frmFilter.chkOR2.Enabled = False
        '
        ' Disable the labels for the first clause
        frmFilter.lblWhereField1.Enabled = False
        frmFilter.lblCond1.Enabled = False
        frmFilter.lblVal1.Enabled = False
        '
        ' Hide the combo boxes
        frmFilter.cmbFields.Visible = False
        frmFilter.cmbOperator.Visible = False
        frmFilter.txtValue.Visible = False
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAnd1 Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkAnd2_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the filter extension for the changes
' TRIGGER:  The user clicked on the second AND check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkAnd2_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAnd2 Click (Start)"
    '-v1.6.1
    '
    ' See if the item was checked
    If frmFilter.chkAnd2.Value = vbChecked Then
        '
        ' Uncheck the OR box
        frmFilter.chkOR2.Value = vbUnchecked
        '
        ' Enable the AND/OR boxes on the next line
        frmFilter.chkAnd3.Enabled = True
        frmFilter.chkOR3.Enabled = True
        '
        ' Enable the labels for the second clause
        frmFilter.lblWhereField2.Enabled = True
        frmFilter.lblCond2.Enabled = True
        frmFilter.lblVal2.Enabled = True
        '
        ' Move the field selection combo box over the field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField2.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField2.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField2.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator selection combo box over the operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond2.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond2.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond2.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the value label
        frmFilter.txtValue.Left = frmFilter.lblVal2.Left
        frmFilter.txtValue.Top = frmFilter.lblVal2.Top
        frmFilter.txtValue.Text = frmFilter.lblVal2.Caption
        frmFilter.txtValue.Visible = True
    Else
        '
        ' Uncheck the AND and OR boxes, and reconfigure the form
        frmFilter.chkAnd3.Value = vbUnchecked
        chkAnd3_Click
        frmFilter.chkOR3.Value = vbUnchecked
        chkOR3_Click
        '
        ' Disable the third AND/OR boxes
        frmFilter.chkAnd3.Enabled = False
        frmFilter.chkOR3.Enabled = False
        '
        ' Disable the labels for the second clause
        frmFilter.lblWhereField2.Enabled = False
        frmFilter.lblCond2.Enabled = False
        frmFilter.lblVal2.Enabled = False
        '
        ' Move the field combo box to the first clause
        frmFilter.cmbFields.Left = frmFilter.lblWhereField1.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField1.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField1.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box to the first clause
        frmFilter.cmbOperator.Left = frmFilter.lblCond1.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond1.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond1.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box to the first clause
        frmFilter.txtValue.Left = frmFilter.lblVal1.Left
        frmFilter.txtValue.Top = frmFilter.lblVal1.Top
        frmFilter.txtValue.Text = frmFilter.lblVal1.Caption
        frmFilter.txtValue.Visible = True
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAnd2 Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkAnd3_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the filter extension for the changes
' TRIGGER:  The user clicked on the third AND check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkAnd3_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAnd3 Click (Start)"
    '-v1.6.1
    '
    ' See if the item was checked
    If frmFilter.chkAnd3.Value = vbChecked Then
        '
        ' Uncheck the OR box
        frmFilter.chkOR3.Value = vbUnchecked
        '
        ' Enable the AND/OR boxes on the next line
        frmFilter.chkAnd4.Enabled = True
        frmFilter.chkOR4.Enabled = True
        '
        ' Enable the labels for the third clause
        frmFilter.lblWhereField3.Enabled = True
        frmFilter.lblCond3.Enabled = True
        frmFilter.lblVal3.Enabled = True
        '
        ' Move the field selection combo box over the field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField3.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField3.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField3.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator selection combo box over the operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond3.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond3.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond3.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the value label
        frmFilter.txtValue.Left = frmFilter.lblVal3.Left
        frmFilter.txtValue.Top = frmFilter.lblVal3.Top
        frmFilter.txtValue.Text = frmFilter.lblVal3.Caption
        frmFilter.txtValue.Visible = True
    Else
        '
        ' Uncheck the AND and OR boxes, and reconfigure the form
        frmFilter.chkAnd4.Value = vbUnchecked
        chkAnd4_Click
        frmFilter.chkOR4.Value = vbUnchecked
        chkOR4_Click
        '
        ' Disable the fourth AND/OR boxes
        frmFilter.chkAnd4.Enabled = False
        frmFilter.chkOR4.Enabled = False
        '
        ' Disable the labels for the third clause
        frmFilter.lblWhereField3.Enabled = False
        frmFilter.lblCond3.Enabled = False
        frmFilter.lblVal3.Enabled = False
        '
        ' Move the field combo box to the second clause
        frmFilter.cmbFields.Left = frmFilter.lblWhereField2.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField2.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField2.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box to the second clause
        frmFilter.cmbOperator.Left = frmFilter.lblCond2.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond2.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond2.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box to the second clause
        frmFilter.txtValue.Left = frmFilter.lblVal2.Left
        frmFilter.txtValue.Top = frmFilter.lblVal2.Top
        frmFilter.txtValue.Text = frmFilter.lblVal2.Caption
        frmFilter.txtValue.Visible = True
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAnd3 Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkAnd4_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the filter extension for the changes
' TRIGGER:  The user clicked on the fourth AND check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkAnd4_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAnd4 Click (Start)"
    '-v1.6.1
    '
    ' See if the item was checked
    If frmFilter.chkAnd4.Value = vbChecked Then
        '
        ' Uncheck the OR box
        frmFilter.chkOR4.Value = vbUnchecked
        '
        ' Enable the labels for the fourth clause
        frmFilter.lblWhereField4.Enabled = True
        frmFilter.lblCond4.Enabled = True
        frmFilter.lblVal4.Enabled = True
        '
        ' Move the field selection combo box over the field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField4.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField4.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField4.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator selection combo box over the operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond4.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond4.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond4.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the value label
        frmFilter.txtValue.Left = frmFilter.lblVal4.Left
        frmFilter.txtValue.Top = frmFilter.lblVal4.Top
        frmFilter.txtValue.Text = frmFilter.lblVal4.Caption
        frmFilter.txtValue.Visible = True
    Else
        '
        ' Disable the labels for the fourth clause
        frmFilter.lblWhereField4.Enabled = False
        frmFilter.lblCond4.Enabled = False
        frmFilter.lblVal4.Enabled = False
        '
        ' Move the field combo box to the third clause
        frmFilter.cmbFields.Left = frmFilter.lblWhereField3.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField3.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField3.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box to the third clause
        frmFilter.cmbOperator.Left = frmFilter.lblCond3.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond3.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond3.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box to the third clause
        frmFilter.txtValue.Left = frmFilter.lblVal3.Left
        frmFilter.txtValue.Top = frmFilter.lblVal3.Top
        frmFilter.txtValue.Text = frmFilter.lblVal3.Caption
        frmFilter.txtValue.Visible = True
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAnd4 Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkAssistant_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Enable the controls specific to the SQL assistant
' TRIGGER:  The user clicked on the "Use Query Assistant" check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkAssistant_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAssistant Click (Start)"
    '-v1.6.1
    '
    ' See if the box was checked
    If frmFilter.chkAssistant.Value = vbChecked Then
        '
        ' Uncheck the "Manual" box
        frmFilter.chkManual.Value = vbUnchecked
    End If
    '
    ' Enable/Disable the controls depending on whether the box was checked
    frmFilter.lblBuilder(0).Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    frmFilter.lblBuilder(1).Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    frmFilter.lblBuilder(2).Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    frmFilter.txtFields.Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    frmFilter.txtFilter.Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    frmFilter.txtSort.Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    frmFilter.btnEditFields.Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    frmFilter.btnEditFilter.Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    frmFilter.btnEditSort.Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    frmFilter.btnExecute.Enabled = (frmFilter.chkAssistant.Value = vbChecked)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkAssistant Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkCustom_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the field extension to allow the user to select fields
' TRIGGER:  The user clicked on the "User-Specified..." check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkCustom_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkCustom Click (Start)"
    '-v1.6.1
    '
    ' Set the default export file type to "User-Defined"
    guExport.iFile_Type = giUSR_TXT
    '
    ' Enable/Disable the predefined check box
    frmFilter.chkPredefined.Enabled = Not (frmFilter.chkCustom.Value = vbChecked)
    '
    ' Enable/Disable the field check boxes
    frmFilter.chkField(0).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(1).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(2).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(3).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(4).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(5).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(6).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(7).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(8).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(9).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(10).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(11).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(12).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(13).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(14).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(15).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(16).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(17).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(18).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(19).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(20).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(21).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(22).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(23).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(24).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(25).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(26).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(27).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(28).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(29).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(30).Enabled = frmFilter.chkCustom.Value
    frmFilter.chkField(31).Enabled = frmFilter.chkCustom.Value
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkCustom Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkManual_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Allow the user to manually enter/modify the SQL query
' TRIGGER:  The user clicked on the "Manual" check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkManual_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkManual Click (Start)"
    '-v1.6.1
    '
    ' See if the user checked the box
    If frmFilter.chkManual.Value = vbChecked Then
        '
        ' Uncheck the "Assistant" check box
        frmFilter.chkAssistant.Value = vbUnchecked
    End If
    '
    ' Enable/Disable the SQL text box and Execute button
    frmFilter.txtSQL.Enabled = (frmFilter.chkManual.Value = vbChecked)
    frmFilter.btnExecute.Enabled = (frmFilter.chkManual.Value = vbChecked)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkManual Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkOR1_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the filter extension for the changes
' TRIGGER:  The user clicked on the first OR check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkOR1_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkOr1 Click (Start)"
    '-v1.6.1
    '
    ' See if the user checked the box
    If frmFilter.chkOR1.Value = vbChecked Then
        '
        ' Uncheck the AND box
        frmFilter.chkAND1.Value = vbUnchecked
        '
        ' Enable the AND/OR boxes on the next line
        frmFilter.chkAnd2.Enabled = True
        frmFilter.chkOR2.Enabled = True
        '
        ' Enable the labels for the first clause
        frmFilter.lblWhereField1.Enabled = True
        frmFilter.lblCond1.Enabled = True
        frmFilter.lblVal1.Enabled = True
        '
        ' Move the field combo box over the field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField1.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField1.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField1.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box over the operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond1.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond1.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond1.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the value label
        frmFilter.txtValue.Left = frmFilter.lblVal1.Left
        frmFilter.txtValue.Top = frmFilter.lblVal1.Top
        frmFilter.txtValue.Text = frmFilter.lblVal1.Caption
        frmFilter.txtValue.Visible = True
    Else
        '
        ' Uncheck both AND/OR boxes and reconfigure
        frmFilter.chkAnd2.Value = vbUnchecked
        chkAnd2_Click
        frmFilter.chkOR2.Value = vbUnchecked
        chkOR2_Click
        '
        ' Disable the AND/OR boxes on the next line
        frmFilter.chkAnd2.Enabled = False
        frmFilter.chkOR2.Enabled = False
        '
        ' Disable the labels for the first clause
        frmFilter.lblWhereField1.Enabled = False
        frmFilter.lblCond1.Enabled = False
        frmFilter.lblVal1.Enabled = False
        '
        ' Hide the combo boxes
        frmFilter.cmbFields.Visible = False
        frmFilter.cmbOperator.Visible = False
        frmFilter.txtValue.Visible = False
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkOr1 Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkOR2_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the filter extension for the changes
' TRIGGER:  The user clicked on the second OR check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkOR2_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkOr2 Click (Start)"
    '-v1.6.1
    '
    ' See if the user checked the box
    If frmFilter.chkOR2.Value = vbChecked Then
        '
        ' Uncheck the AND box
        frmFilter.chkAnd2.Value = vbUnchecked
        '
        ' Enable the AND/OR boxes on the next line
        frmFilter.chkAnd3.Enabled = True
        frmFilter.chkOR3.Enabled = True
        '
        ' Enable the labels for the second clause
        frmFilter.lblWhereField2.Enabled = True
        frmFilter.lblCond2.Enabled = True
        frmFilter.lblVal2.Enabled = True
        '
        ' Move the field combo box over the field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField2.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField2.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField2.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box over the operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond2.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond2.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond2.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the value label
        frmFilter.txtValue.Left = frmFilter.lblVal2.Left
        frmFilter.txtValue.Top = frmFilter.lblVal2.Top
        frmFilter.txtValue.Text = frmFilter.lblVal2.Caption
        frmFilter.txtValue.Visible = True
    Else
        '
        ' Uncheck both AND/OR boxes and reconfigure
        frmFilter.chkAnd3.Value = vbUnchecked
        chkAnd3_Click
        frmFilter.chkOR3.Value = vbUnchecked
        chkOR3_Click
        '
        ' Disable the AND/OR boxes on the next line
        frmFilter.chkAnd3.Enabled = False
        frmFilter.chkOR3.Enabled = False
        '
        ' Disable the labels for the second clause
        frmFilter.lblWhereField2.Enabled = False
        frmFilter.lblCond2.Enabled = False
        frmFilter.lblVal2.Enabled = False
        '
        ' Move the field combo box over the previous field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField1.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField1.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField1.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box over the previous operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond1.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond1.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond1.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the previous value label
        frmFilter.txtValue.Left = frmFilter.lblVal1.Left
        frmFilter.txtValue.Top = frmFilter.lblVal1.Top
        frmFilter.txtValue.Text = frmFilter.lblVal1.Caption
        frmFilter.txtValue.Visible = True
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkOr2 Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkOR3_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the filter extension for the changes
' TRIGGER:  The user clicked on the third OR check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkOR3_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkOr3 Click (Start)"
    '-v1.6.1
    '
    ' See if the user checked the box
    If frmFilter.chkOR3.Value = vbChecked Then
        '
        ' Uncheck the AND box
        frmFilter.chkAnd3.Value = vbUnchecked
        '
        ' Enable the AND/OR boxes on the next line
        frmFilter.chkAnd4.Enabled = True
        frmFilter.chkOR4.Enabled = True
        '
        ' Enable the labels for the third clause
        frmFilter.lblWhereField3.Enabled = True
        frmFilter.lblCond3.Enabled = True
        frmFilter.lblVal3.Enabled = True
        '
        ' Move the field combo box over the field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField3.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField3.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField3.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box over the operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond3.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond3.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond3.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the value label
        frmFilter.txtValue.Left = frmFilter.lblVal3.Left
        frmFilter.txtValue.Top = frmFilter.lblVal3.Top
        frmFilter.txtValue.Text = frmFilter.lblVal3.Caption
        frmFilter.txtValue.Visible = True
    Else
        '
        ' Uncheck both AND/OR boxes and reconfigure
        frmFilter.chkAnd4.Value = vbUnchecked
        chkAnd4_Click
        frmFilter.chkOR4.Value = vbUnchecked
        chkOR4_Click
        '
        ' Disable the AND/OR boxes on the next line
        frmFilter.chkAnd4.Enabled = False
        frmFilter.chkOR4.Enabled = False
        '
        ' Disable the labels for the third clause
        frmFilter.lblWhereField3.Enabled = False
        frmFilter.lblCond3.Enabled = False
        frmFilter.lblVal3.Enabled = False
        '
        ' Move the field combo box over the previous field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField2.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField2.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField2.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box over the previous operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond2.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond2.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond2.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the previous value label
        frmFilter.txtValue.Left = frmFilter.lblVal2.Left
        frmFilter.txtValue.Top = frmFilter.lblVal2.Top
        frmFilter.txtValue.Text = frmFilter.lblVal2.Caption
        frmFilter.txtValue.Visible = True
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkOr3 Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkOR4_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configure the filter extension for the changes
' TRIGGER:  The user clicked on the fourth OR check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkOR4_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkOr4 Click (Start)"
    '-v1.6.1
    '
    ' See if the user checked the box
    If frmFilter.chkOR4.Value = vbChecked Then
        '
        ' Uncheck the AND box
        frmFilter.chkAnd4.Value = vbUnchecked
        '
        ' Enable the labels for the fourth clause
        frmFilter.lblWhereField4.Enabled = True
        frmFilter.lblCond4.Enabled = True
        frmFilter.lblVal4.Enabled = True
        '
        ' Move the field combo box over the field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField4.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField4.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField4.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box over the operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond4.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond4.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond4.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the value label
        frmFilter.txtValue.Left = frmFilter.lblVal4.Left
        frmFilter.txtValue.Top = frmFilter.lblVal4.Top
        frmFilter.txtValue.Text = frmFilter.lblVal4.Caption
        frmFilter.txtValue.Visible = True
    Else
        '
        ' Disable the labels for the fourth clause
        frmFilter.lblWhereField4.Enabled = False
        frmFilter.lblCond4.Enabled = False
        frmFilter.lblVal4.Enabled = False
        '
        ' Move the field combo box over the previous field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField3.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField3.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField3.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box over the previous operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond3.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond3.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond3.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the previous value label
        frmFilter.txtValue.Left = frmFilter.lblVal3.Left
        frmFilter.txtValue.Top = frmFilter.lblVal3.Top
        frmFilter.txtValue.Text = frmFilter.lblVal3.Caption
        frmFilter.txtValue.Visible = True
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkOr4 Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkPredefined_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Allow the user to select a predefined set of fields
' TRIGGER:  User clicked on the "Predefined" check box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkPredefined_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkPredefined Click (Start)"
    '-v1.6.1
    '
    ' Enable/Disable the Custom check box
    frmFilter.chkCustom.Enabled = Not (frmFilter.chkPredefined.Value = vbChecked)
    '
    ' Enable/Disable the predefined option buttons
    frmFilter.optField(0).Enabled = frmFilter.chkPredefined.Value
    frmFilter.optField(1).Enabled = frmFilter.chkPredefined.Value
    frmFilter.optField(2).Enabled = frmFilter.chkPredefined.Value
    frmFilter.optField(3).Enabled = frmFilter.chkPredefined.Value
    frmFilter.optRpt(0).Enabled = frmFilter.chkPredefined.Value
    frmFilter.optRpt(1).Enabled = frmFilter.chkPredefined.Value
    frmFilter.optRpt(2).Enabled = frmFilter.chkPredefined.Value
    frmFilter.optRpt(3).Enabled = frmFilter.chkPredefined.Value
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkPredefined Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkSort_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Allow the user to select a field for sorting
' TRIGGER:  User clicked on a "Sort" check box
' INPUT:    "intField" indicates which check box was selected
' OUTPUT:   None
' NOTES:
Private Sub chkSort_Click(intField As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmFilter.chkSort Click (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intField
    End If
    '-v1.6.1
    '
    ' See if the box was unchecked
    If frmFilter.chkSort(intField).Value = vbUnchecked Then
        '
        ' Remove the field name
        frmFilter.lblSortField(intField).Caption = ""
        '
        ' See if the box is not the last one
        If intField < frmFilter.chkSort.UBound Then
            '
            ' Uncheck the box below the current one
            frmFilter.chkSort(intField + 1).Value = vbUnchecked
            '
            ' Trigger this event for the next box
            chkSort_Click (intField + 1)
        End If
    End If
    '
    ' Determine if the sort field combo box should be visible
    frmFilter.cmbSortField.Visible = (frmFilter.chkSort(intField).Value = vbChecked)
    '
    ' Enable/Disable the next check box
    If intField < frmFilter.chkSort.UBound Then frmFilter.chkSort(intField + 1).Enabled = (frmFilter.chkSort(intField).Value = vbChecked)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkSort Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkSort_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Moves the field combo box over the appropriate field entry
' TRIGGER:  The user clicked on a "Sort" box
' INPUT:    "intField" indicates which check box was selected
' OUTPUT:   None
' NOTES:
Private Sub chkSort_GotFocus(intField As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmFilter.chkSort GotFocus (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intField
    End If
    '-v1.6.1
    '
    ' Move the combo box to the position of the appropriate field
    frmFilter.cmbSortField.Left = frmFilter.lblSortField(intField).Left
    frmFilter.cmbSortField.Top = frmFilter.lblSortField(intField).Top
    frmFilter.cmbSortField.Text = frmFilter.lblSortField(intField).Caption
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkSort GotFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    chkWhere_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Enables the user to create filter clauses
' TRIGGER:  User clicked on the "Select only records..." box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub chkWhere_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkWhere Click (Start)"
    '-v1.6.1
    '
    ' See if the user checked the box
    If frmFilter.chkWhere.Value = vbChecked Then
        '
        ' Enable the AND/OR boxes on the next line
        frmFilter.chkAnd2.Enabled = True
        frmFilter.chkOR2.Enabled = True
        '
        ' Enable the labels for the first clause
        frmFilter.lblWhereField1.Enabled = True
        frmFilter.lblCond1.Enabled = True
        frmFilter.lblVal1.Enabled = True
        '
        ' Move the field combo box over the field label
        frmFilter.cmbFields.Left = frmFilter.lblWhereField1.Left
        frmFilter.cmbFields.Top = frmFilter.lblWhereField1.Top
        frmFilter.cmbFields.Text = frmFilter.lblWhereField1.Caption
        frmFilter.cmbFields.Visible = True
        '
        ' Move the operator combo box over the operator label
        frmFilter.cmbOperator.Left = frmFilter.lblCond1.Left
        frmFilter.cmbOperator.Top = frmFilter.lblCond1.Top
        frmFilter.cmbOperator.Text = frmFilter.lblCond1.Caption
        frmFilter.cmbOperator.Visible = True
        '
        ' Move the value entry box over the value label
        frmFilter.txtValue.Left = frmFilter.lblVal1.Left
        frmFilter.txtValue.Top = frmFilter.lblVal1.Top
        frmFilter.txtValue.Text = frmFilter.lblVal1.Caption
        frmFilter.txtValue.Visible = True
    Else
        '
        ' Uncheck the second AND/OR boxes and reconfigure the form
        frmFilter.chkAnd2.Value = vbUnchecked
        chkAnd2_Click
        frmFilter.chkOR2.Value = vbUnchecked
        chkOR2_Click
        '
        ' Disable the AND/OR boxes for the second clause
        frmFilter.chkAnd2.Enabled = False
        frmFilter.chkOR2.Enabled = False
        '
        ' Disable the labels for the first clause
        frmFilter.lblWhereField1.Enabled = False
        frmFilter.lblCond1.Enabled = False
        frmFilter.lblVal1.Enabled = False
        '
        ' Hide the combo boxes
        frmFilter.cmbFields.Visible = False
        frmFilter.cmbOperator.Visible = False
        frmFilter.txtValue.Visible = False
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.chkWhere Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbFields_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the selected field for a clause
' TRIGGER:  User changed the selected field
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbFields_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbFields Change (Start)"
    '-v1.6.1
    '
    ' Determine which clause based on the location of the combo box
    Select Case frmFilter.cmbFields.Top
        '
        ' First row
        Case frmFilter.lblWhereField1.Top:
            frmFilter.lblWhereField1.Caption = frmFilter.cmbFields.Text
        '
        ' Second row
        Case frmFilter.lblWhereField2.Top:
            frmFilter.lblWhereField2.Caption = frmFilter.cmbFields.Text
        '
        ' Third row
        Case frmFilter.lblWhereField3.Top:
            frmFilter.lblWhereField3.Caption = frmFilter.cmbFields.Text
        '
        ' Fourth row
        Case frmFilter.lblWhereField4.Top:
            frmFilter.lblWhereField4.Caption = frmFilter.cmbFields.Text
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbFields Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbFields_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the selected field for a clause
' TRIGGER:  User clicked on a field name
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbFields_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbFields Click (Start)"
    '-v1.6.1
    '
    ' Trigger the field change event
    cmbFields_Change
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbFields Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbOperator_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the selected operator for a clause
' TRIGGER:  User changed the selected operator
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbOperator_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbOperator Change (Start)"
    '-v1.6.1
    '
    ' Determine which clause based on the location of the combo box
    Select Case frmFilter.cmbOperator.Top
        '
        ' First row
        Case frmFilter.lblCond1.Top:
            frmFilter.lblCond1.Caption = frmFilter.cmbOperator.Text
        '
        ' Second row
        Case frmFilter.lblCond2.Top:
            frmFilter.lblCond2.Caption = frmFilter.cmbOperator.Text
        '
        ' Third row
        Case frmFilter.lblCond3.Top:
            frmFilter.lblCond3.Caption = frmFilter.cmbOperator.Text
        '
        ' Fourth row
        Case frmFilter.lblCond4.Top:
            frmFilter.lblCond4.Caption = frmFilter.cmbOperator.Text
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbOperator Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbOperator_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the selected operator for a clause
' TRIGGER:  User clicked on an operator
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbOperator_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbOperator Click (Start)"
    '-v1.6.1
    '
    ' Trigger the operator change event
    cmbOperator_Change
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbOperator Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbSortField_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the selected field for a sort item
' TRIGGER:  User changed the selected field
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbSortField_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbSortField Change (Start)"
    '-v1.6.1
    '
    ' Determine which field based on the location of the combo box
    Select Case frmFilter.cmbSortField.Top
        '
        ' First row
        Case frmFilter.lblSortField(0).Top:
            frmFilter.lblSortField(0).Caption = frmFilter.cmbSortField.Text
        '
        ' Second row
        Case frmFilter.lblSortField(1).Top:
            frmFilter.lblSortField(1).Caption = frmFilter.cmbSortField.Text
        '
        ' Third row
        Case frmFilter.lblSortField(2).Top:
            frmFilter.lblSortField(2).Caption = frmFilter.cmbSortField.Text
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbSortField Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    cmbSortField_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the selected field for a sort item
' TRIGGER:  User clicked on a field
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub cmbSortField_Click()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbSortField Click (Start)"
    '-v1.6.1
    '
    ' Trigger the field change event
    cmbSortField_Change
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.cmbSortField Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Activate
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the controls for the current SQL query
' TRIGGER:  User requested the filter form
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Activate()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter Activate (Start)"
    '-v1.6.1
    '
    ' Set to default size
    frmFilter.Width = mintFILTER_MIN_WIDTH
    '
    ' Set the controls to the current SQL query
    frmFilter.txtFields.Text = guCurrent.uSQL.sFields
    frmFilter.txtFilter.Text = guCurrent.uSQL.sFilter
    frmFilter.txtSort.Text = guCurrent.uSQL.sOrder
    frmFilter.txtSQL.Text = guCurrent.uSQL.sQuery
    '
    ' Uncheck both check boxes and reconfigure the form
    frmFilter.chkAssistant.Value = vbUnchecked
    chkAssistant_Click
    frmFilter.chkManual.Value = vbUnchecked
    chkManual_Click
    '
    ' Enable the buttons
    frmFilter.btnClear.Enabled = True
    frmFilter.btnAccept.Enabled = False
    frmFilter.btnHide.Enabled = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter Activate (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Load
' AUTHOR:   Tom Elkins
' PURPOSE:  Builds and configures the form for use
' TRIGGER:  The first time the form is called
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Load()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter Load (Start)"
    '-v1.6.1
    '
    ' Hide all extensions
    frmFilter.picOptions(mintPIC_FIELDS).Left = mintFILTER_POS_HIDE
    frmFilter.picOptions(mintPIC_FILTER).Left = mintFILTER_POS_HIDE
    frmFilter.picOptions(mintPIC_SORT).Left = mintFILTER_POS_HIDE
    '
    ' Uncheck and reconfigure the field selection boxes
    frmFilter.chkPredefined.Value = vbUnchecked
    chkPredefined_Click
    '
    ' Set the default values for the field check boxes
    ' The actual DAS field name is stored in the Tag property
    frmFilter.chkField(0).Value = vbChecked
    frmFilter.chkField(0).Tag = "ReportTime"
    frmFilter.chkField(1).Value = vbChecked
    frmFilter.chkField(1).Tag = "Msg_Type"
    frmFilter.chkField(2).Value = vbChecked
    frmFilter.chkField(2).Tag = "Rpt_Type"
    frmFilter.chkField(3).Value = vbChecked
    frmFilter.chkField(3).Tag = "Origin"
    frmFilter.chkField(4).Value = vbChecked
    frmFilter.chkField(4).Tag = "Origin_ID"
    frmFilter.chkField(5).Value = vbChecked
    frmFilter.chkField(5).Tag = "Target_ID"
    frmFilter.chkField(6).Value = vbChecked
    frmFilter.chkField(6).Tag = "Latitude"
    frmFilter.chkField(7).Value = vbChecked
    frmFilter.chkField(7).Tag = "Longitude"
    frmFilter.chkField(8).Value = vbChecked
    frmFilter.chkField(8).Tag = "Altitude"
    frmFilter.chkField(9).Value = vbChecked
    frmFilter.chkField(9).Tag = "Heading"
    frmFilter.chkField(10).Value = vbChecked
    frmFilter.chkField(10).Tag = "Speed"
    frmFilter.chkField(11).Value = vbChecked
    frmFilter.chkField(11).Tag = "Parent"
    frmFilter.chkField(12).Value = vbChecked
    frmFilter.chkField(12).Tag = "Parent_ID"
    frmFilter.chkField(13).Value = vbChecked
    frmFilter.chkField(13).Tag = "Allegiance"
    frmFilter.chkField(14).Value = vbChecked
    frmFilter.chkField(14).Tag = "IFF"
    frmFilter.chkField(15).Value = vbChecked
    frmFilter.chkField(15).Tag = "Emitter"
    frmFilter.chkField(16).Value = vbChecked
    frmFilter.chkField(16).Tag = "Emitter_ID"
    frmFilter.chkField(17).Value = vbChecked
    frmFilter.chkField(17).Tag = "Signal"
    frmFilter.chkField(18).Value = vbChecked
    frmFilter.chkField(18).Tag = "Signal_ID"
    frmFilter.chkField(19).Value = vbChecked
    frmFilter.chkField(19).Tag = "Frequency"
    frmFilter.chkField(20).Value = vbChecked
    frmFilter.chkField(20).Tag = "PRI"
    frmFilter.chkField(21).Value = vbChecked
    frmFilter.chkField(21).Tag = "Status"
    frmFilter.chkField(22).Value = vbChecked
    frmFilter.chkField(22).Tag = "Tag"
    frmFilter.chkField(23).Value = vbChecked
    frmFilter.chkField(23).Tag = "Flag"
    frmFilter.chkField(24).Value = vbChecked
    frmFilter.chkField(24).Tag = "Common_ID"
    frmFilter.chkField(25).Value = vbChecked
    frmFilter.chkField(25).Tag = "Range"
    frmFilter.chkField(26).Value = vbChecked
    frmFilter.chkField(26).Tag = "Bearing"
    frmFilter.chkField(27).Value = vbChecked
    frmFilter.chkField(27).Tag = "Elevation"
    frmFilter.chkField(28).Value = vbChecked
    frmFilter.chkField(28).Tag = "XX"
    frmFilter.chkField(29).Value = vbChecked
    frmFilter.chkField(29).Tag = "XY"
    frmFilter.chkField(30).Value = vbChecked
    frmFilter.chkField(30).Tag = "YY"
    frmFilter.chkField(31).Value = vbChecked
    frmFilter.chkField(31).Tag = "Other_Data"
    '
    ' Uncheck the custom box and reconfigure the extension
    frmFilter.chkCustom.Value = vbUnchecked
    chkCustom_Click
    '
    ' Load the field names in the combo box
    frmFilter.cmbFields.AddItem "ReportTime"
    frmFilter.cmbFields.AddItem "Msg_Type"
    frmFilter.cmbFields.AddItem "Rpt_Type"
    frmFilter.cmbFields.AddItem "Origin"
    frmFilter.cmbFields.AddItem "Origin_ID"
    frmFilter.cmbFields.AddItem "Target_ID"
    frmFilter.cmbFields.AddItem "Latitude"
    frmFilter.cmbFields.AddItem "Longitude"
    frmFilter.cmbFields.AddItem "Altitude"
    frmFilter.cmbFields.AddItem "Heading"
    frmFilter.cmbFields.AddItem "Speed"
    frmFilter.cmbFields.AddItem "Parent"
    frmFilter.cmbFields.AddItem "Parent_ID"
    frmFilter.cmbFields.AddItem "Allegiance"
    frmFilter.cmbFields.AddItem "IFF"
    frmFilter.cmbFields.AddItem "Emitter"
    frmFilter.cmbFields.AddItem "Emitter_ID"
    frmFilter.cmbFields.AddItem "Signal"
    frmFilter.cmbFields.AddItem "Signal_ID"
    frmFilter.cmbFields.AddItem "Frequency"
    frmFilter.cmbFields.AddItem "PRI"
    frmFilter.cmbFields.AddItem "Status"
    frmFilter.cmbFields.AddItem "Tag"
    frmFilter.cmbFields.AddItem "Flag"
    frmFilter.cmbFields.AddItem "Common_ID"
    frmFilter.cmbFields.AddItem "Range"
    frmFilter.cmbFields.AddItem "Bearing"
    frmFilter.cmbFields.AddItem "Elevation"
    frmFilter.cmbFields.AddItem "XX"
    frmFilter.cmbFields.AddItem "XY"
    frmFilter.cmbFields.AddItem "YY"
    frmFilter.cmbFields.AddItem "Other_Data"
    '
    ' Load the operator choices in the combo box
    frmFilter.cmbOperator.AddItem "="
    frmFilter.cmbOperator.AddItem "<"
    frmFilter.cmbOperator.AddItem "<="
    frmFilter.cmbOperator.AddItem ">"
    frmFilter.cmbOperator.AddItem ">="
    frmFilter.cmbOperator.AddItem "between"
    '
    ' Hide the combo boxes and text entry boxes
    frmFilter.cmbFields.Visible = False
    frmFilter.cmbOperator.Visible = False
    frmFilter.txtValue.Text = ""
    frmFilter.txtValue.Visible = False
    '
    ' Enable/Disable the default filter check boxes
    frmFilter.chkAND1.Enabled = True
    frmFilter.chkAnd2.Enabled = False
    frmFilter.chkAnd3.Enabled = False
    frmFilter.chkAnd4.Enabled = False
    frmFilter.chkOR1.Enabled = True
    frmFilter.chkOR2.Enabled = False
    frmFilter.chkOR3.Enabled = False
    frmFilter.chkOR4.Enabled = False
    frmFilter.chkWhere.Enabled = True
    '
    ' Disable the clause labels
    frmFilter.lblWhereField1.Enabled = False
    frmFilter.lblWhereField2.Enabled = False
    frmFilter.lblWhereField3.Enabled = False
    frmFilter.lblWhereField4.Enabled = False
    frmFilter.lblCond1.Enabled = False
    frmFilter.lblCond2.Enabled = False
    frmFilter.lblCond3.Enabled = False
    frmFilter.lblCond4.Enabled = False
    frmFilter.lblVal1.Enabled = False
    frmFilter.lblVal2.Enabled = False
    frmFilter.lblVal3.Enabled = False
    frmFilter.lblVal4.Enabled = False
    '
    ' Set the default values for the clause labels
    frmFilter.lblWhereField1.Caption = ""
    frmFilter.lblWhereField2.Caption = ""
    frmFilter.lblWhereField3.Caption = ""
    frmFilter.lblWhereField4.Caption = ""
    frmFilter.lblCond1.Caption = ""
    frmFilter.lblCond2.Caption = ""
    frmFilter.lblCond3.Caption = ""
    frmFilter.lblCond4.Caption = ""
    frmFilter.lblVal1.Caption = ""
    frmFilter.lblVal2.Caption = ""
    frmFilter.lblVal3.Caption = ""
    frmFilter.lblVal4.Caption = ""
    '
    ' Uncheck all of the sort boxes
    frmFilter.chkSort(0).Value = vbUnchecked
    frmFilter.chkSort(1).Value = vbUnchecked
    frmFilter.chkSort(2).Value = vbUnchecked
    '
    ' Disable the second and third sort options
    frmFilter.chkSort(1).Enabled = False
    frmFilter.chkSort(2).Enabled = False
    '
    ' Set the default values for the sort fields
    frmFilter.lblSortField(0).Caption = ""
    frmFilter.lblSortField(1).Caption = ""
    frmFilter.lblSortField(2).Caption = ""
    '
    ' Enter the sort field options in the combo box
    frmFilter.cmbSortField.AddItem "ReportTime"
    frmFilter.cmbSortField.AddItem "Msg_Type"
    frmFilter.cmbSortField.AddItem "Rpt_Type"
    frmFilter.cmbSortField.AddItem "Origin"
    frmFilter.cmbSortField.AddItem "Origin_ID"
    frmFilter.cmbSortField.AddItem "Target_ID"
    frmFilter.cmbSortField.AddItem "Latitude"
    frmFilter.cmbSortField.AddItem "Longitude"
    frmFilter.cmbSortField.AddItem "Altitude"
    frmFilter.cmbSortField.AddItem "Heading"
    frmFilter.cmbSortField.AddItem "Speed"
    frmFilter.cmbSortField.AddItem "Parent"
    frmFilter.cmbSortField.AddItem "Parent_ID"
    frmFilter.cmbSortField.AddItem "Allegiance"
    frmFilter.cmbSortField.AddItem "IFF"
    frmFilter.cmbSortField.AddItem "Emitter"
    frmFilter.cmbSortField.AddItem "Emitter_ID"
    frmFilter.cmbSortField.AddItem "Signal"
    frmFilter.cmbSortField.AddItem "Signal_ID"
    frmFilter.cmbSortField.AddItem "Frequency"
    frmFilter.cmbSortField.AddItem "PRI"
    frmFilter.cmbSortField.AddItem "Status"
    frmFilter.cmbSortField.AddItem "Tag"
    frmFilter.cmbSortField.AddItem "Flag"
    frmFilter.cmbSortField.AddItem "Common_ID"
    frmFilter.cmbSortField.AddItem "Range"
    frmFilter.cmbSortField.AddItem "Bearing"
    frmFilter.cmbSortField.AddItem "Elevation"
    frmFilter.cmbSortField.AddItem "XX"
    frmFilter.cmbSortField.AddItem "XY"
    frmFilter.cmbSortField.AddItem "YY"
    frmFilter.cmbSortField.AddItem "Other_Data"
    frmFilter.cmbSortField.Visible = False
    '
    '+v1.5
    ' Set help context to the basic form
    Me.HelpContextID = basCCAT.IDH_GUI_FILTER
    '-v1.5
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter Load (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    Form_Resize
' AUTHOR:   Tom Elkins
' PURPOSE:  Moves the form to the center of the screen
' TRIGGER:  The form changed size
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub Form_Resize()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter Resize (Start)"
    '-v1.6.1
    '
    ' Center the form
    frmFilter.Move (Screen.Width - frmFilter.Width) / 2, (Screen.Height - frmFilter.Height) / 2
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter Resize (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    optField_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configures the selected fields based on the predefined option type
' TRIGGER:  User clicked on a predefined field selection
' INPUT:    "intFile_Type" indicates which option was selected
' OUTPUT:   None
' NOTES:
Private Sub optField_Click(intFile_Type As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmFilter.optField Click (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intFile_Type
    End If
    '-v1.6.1
    '
    '+v1.6TE
    Const FLD_TIM = 0
    Const FLD_MSG = 1
    Const FLD_RPT = 2
    Const FLD_ORG = 3
    Const FLD_OID = 4
    Const FLD_TGT = 5
    Const FLD_LAT = 6
    Const FLD_LON = 7
    Const FLD_ALT = 8
    Const FLD_HDG = 9
    Const FLD_SPD = 10
    Const FLD_PAR = 11
    Const FLD_PID = 12
    Const FLD_ALG = 13
    Const FLD_IFF = 14
    Const FLD_EMT = 15
    Const FLD_EID = 16
    Const FLD_SIG = 17
    Const FLD_SID = 18
    Const FLD_FRQ = 19
    Const FLD_PRI = 20
    Const FLD_STA = 21
    Const FLD_TAG = 22
    Const FLD_FLG = 23
    Const FLD_COM = 24
    Const FLD_RNG = 25
    Const FLD_AZM = 26
    Const FLD_ELV = 27
    Const FLD_GXX = 28
    Const FLD_GXY = 29
    Const FLD_GYY = 30
    Const FLD_SUP = 31
    Const RPT_DOT = 0
    Const RPT_TRK = 1
    Const RPT_VEC = 2
    Const RPT_GEO = 3
    Const DAS_SIG = 0
    Const DAS_EVT = 1
    Const DAS_MTF = 2
    Const DAS_STF = 3
    '-v1.6
    '
    ' Action is based on which option was selected
    Select Case intFile_Type
        '
        ' Signal
        Case DAS_SIG:
            '
            ' Set the default export file type
            guExport.iFile_Type = giDAS_SIG
            '
            ' Disable the report type options
            frmFilter.optRpt(RPT_DOT).Enabled = False
            frmFilter.optRpt(RPT_DOT).Value = False
            frmFilter.optRpt(RPT_TRK).Enabled = False
            frmFilter.optRpt(RPT_TRK).Value = False
            frmFilter.optRpt(RPT_VEC).Enabled = False
            frmFilter.optRpt(RPT_VEC).Value = False
            frmFilter.optRpt(RPT_GEO).Enabled = False
            frmFilter.optRpt(RPT_GEO).Value = False
            '
            ' Select/Unselect the appropriate fields for this type
            frmFilter.chkField(FLD_TIM).Value = vbChecked
            frmFilter.chkField(FLD_MSG).Value = vbChecked
            frmFilter.chkField(FLD_RPT).Value = vbChecked
            frmFilter.chkField(FLD_ORG).Value = vbChecked
            frmFilter.chkField(FLD_OID).Value = vbChecked
            frmFilter.chkField(FLD_TGT).Value = vbUnchecked
            frmFilter.chkField(FLD_LAT).Value = vbUnchecked
            frmFilter.chkField(FLD_LON).Value = vbUnchecked
            frmFilter.chkField(FLD_ALT).Value = vbUnchecked
            frmFilter.chkField(FLD_HDG).Value = vbUnchecked
            frmFilter.chkField(FLD_SPD).Value = vbUnchecked
            frmFilter.chkField(FLD_PAR).Value = vbUnchecked
            frmFilter.chkField(FLD_PID).Value = vbUnchecked
            frmFilter.chkField(FLD_ALG).Value = vbChecked
            frmFilter.chkField(FLD_IFF).Value = vbChecked
            frmFilter.chkField(FLD_EMT).Value = vbChecked
            frmFilter.chkField(FLD_EID).Value = vbChecked
            frmFilter.chkField(FLD_SIG).Value = vbChecked
            frmFilter.chkField(FLD_SID).Value = vbChecked
            frmFilter.chkField(FLD_FRQ).Value = vbChecked
            frmFilter.chkField(FLD_PRI).Value = vbChecked
            frmFilter.chkField(FLD_STA).Value = vbChecked
            frmFilter.chkField(FLD_TAG).Value = vbChecked
            frmFilter.chkField(FLD_FLG).Value = vbChecked
            frmFilter.chkField(FLD_COM).Value = vbChecked
            frmFilter.chkField(FLD_RNG).Value = vbUnchecked
            frmFilter.chkField(FLD_AZM).Value = vbUnchecked
            frmFilter.chkField(FLD_ELV).Value = vbUnchecked
            frmFilter.chkField(FLD_GXX).Value = vbUnchecked
            frmFilter.chkField(FLD_GXY).Value = vbUnchecked
            frmFilter.chkField(FLD_GYY).Value = vbUnchecked
            frmFilter.chkField(FLD_SUP).Value = vbUnchecked
        '
        ' Event
        Case DAS_EVT:
            '
            ' Set the default export file type
            guExport.iFile_Type = giDAS_EVT
            '
            ' Disable the report type options
            frmFilter.optRpt(RPT_DOT).Enabled = False
            frmFilter.optRpt(RPT_DOT).Value = False
            frmFilter.optRpt(RPT_TRK).Enabled = False
            frmFilter.optRpt(RPT_TRK).Value = False
            frmFilter.optRpt(RPT_VEC).Enabled = False
            frmFilter.optRpt(RPT_VEC).Value = False
            frmFilter.optRpt(RPT_GEO).Enabled = False
            frmFilter.optRpt(RPT_GEO).Value = False
            '
            ' Select the appropriate fields for this type
            frmFilter.chkField(FLD_TIM).Value = vbChecked
            frmFilter.chkField(FLD_MSG).Value = vbChecked
            frmFilter.chkField(FLD_RPT).Value = vbChecked
            frmFilter.chkField(FLD_ORG).Value = vbChecked
            frmFilter.chkField(FLD_OID).Value = vbChecked
            '
            '+v1.6TE
            'frmFilter.chkField(FLD_TGT).Value = vbUnchecked
            frmFilter.chkField(FLD_TGT).Value = vbChecked
            '-v1.6
            frmFilter.chkField(FLD_LAT).Value = vbUnchecked
            frmFilter.chkField(FLD_LON).Value = vbUnchecked
            frmFilter.chkField(FLD_ALT).Value = vbUnchecked
            frmFilter.chkField(FLD_HDG).Value = vbUnchecked
            frmFilter.chkField(FLD_SPD).Value = vbUnchecked
            frmFilter.chkField(FLD_PAR).Value = vbUnchecked
            frmFilter.chkField(FLD_PID).Value = vbUnchecked
            frmFilter.chkField(FLD_ALG).Value = vbUnchecked
            frmFilter.chkField(FLD_IFF).Value = vbUnchecked
            frmFilter.chkField(FLD_EMT).Value = vbUnchecked
            frmFilter.chkField(FLD_EID).Value = vbUnchecked
            frmFilter.chkField(FLD_SIG).Value = vbUnchecked
            frmFilter.chkField(FLD_SID).Value = vbUnchecked
            frmFilter.chkField(FLD_FRQ).Value = vbUnchecked
            frmFilter.chkField(FLD_PRI).Value = vbUnchecked
            frmFilter.chkField(FLD_STA).Value = vbUnchecked
            frmFilter.chkField(FLD_TAG).Value = vbUnchecked
            frmFilter.chkField(FLD_FLG).Value = vbUnchecked
            '
            '+v1.6TE
            'frmFilter.chkField(FLD_COM).Value = vbUnchecked
            'frmFilter.chkField(FLD_RNG).Value = vbUnchecked
            frmFilter.chkField(FLD_COM).Value = vbChecked
            frmFilter.chkField(FLD_RNG).Value = vbChecked
            '-v1.6
            frmFilter.chkField(FLD_AZM).Value = vbUnchecked
            frmFilter.chkField(FLD_ELV).Value = vbUnchecked
            frmFilter.chkField(FLD_GXX).Value = vbUnchecked
            frmFilter.chkField(FLD_GXY).Value = vbUnchecked
            frmFilter.chkField(FLD_GYY).Value = vbUnchecked
            frmFilter.chkField(FLD_SUP).Value = vbChecked
        '
        ' Moving Target
        Case DAS_MTF:
            '
            ' Set the default export file type
            guExport.iFile_Type = giDAS_MTF
            '
            ' Enable/Disable the appropriate report type options
            frmFilter.optRpt(RPT_DOT).Enabled = True
            frmFilter.optRpt(RPT_DOT).Value = True
            frmFilter.optRpt(RPT_TRK).Enabled = True
            frmFilter.optRpt(RPT_TRK).Value = False
            frmFilter.optRpt(RPT_VEC).Enabled = True
            frmFilter.optRpt(RPT_VEC).Value = False
            frmFilter.optRpt(RPT_GEO).Enabled = True
            frmFilter.optRpt(RPT_GEO).Value = False
            '
            ' Check the appropriate fields for this type
            frmFilter.chkField(FLD_TIM).Value = vbChecked
            frmFilter.chkField(FLD_MSG).Value = vbChecked
            frmFilter.chkField(FLD_RPT).Value = vbChecked
            frmFilter.chkField(FLD_ORG).Value = vbChecked
            frmFilter.chkField(FLD_OID).Value = vbChecked
            frmFilter.chkField(FLD_TGT).Value = vbChecked
            frmFilter.chkField(FLD_LAT).Value = vbChecked
            frmFilter.chkField(FLD_LON).Value = vbChecked
            frmFilter.chkField(FLD_ALT).Value = vbChecked
            frmFilter.chkField(FLD_HDG).Value = vbChecked
            frmFilter.chkField(FLD_SPD).Value = vbChecked
            frmFilter.chkField(FLD_PAR).Value = vbUnchecked
            frmFilter.chkField(FLD_PID).Value = vbUnchecked
            frmFilter.chkField(FLD_ALG).Value = vbChecked
            frmFilter.chkField(FLD_IFF).Value = vbChecked
            frmFilter.chkField(FLD_EMT).Value = vbChecked
            frmFilter.chkField(FLD_EID).Value = vbChecked
            frmFilter.chkField(FLD_SIG).Value = vbChecked
            frmFilter.chkField(FLD_SID).Value = vbChecked
            frmFilter.chkField(FLD_FRQ).Value = vbChecked
            frmFilter.chkField(FLD_PRI).Value = vbChecked
            frmFilter.chkField(FLD_STA).Value = vbChecked
            frmFilter.chkField(FLD_TAG).Value = vbChecked
            frmFilter.chkField(FLD_FLG).Value = vbChecked
            frmFilter.chkField(FLD_COM).Value = vbChecked
            frmFilter.chkField(FLD_RNG).Value = vbUnchecked
            frmFilter.chkField(FLD_AZM).Value = vbUnchecked
            frmFilter.chkField(FLD_ELV).Value = vbUnchecked
            frmFilter.chkField(FLD_GXX).Value = vbUnchecked
            frmFilter.chkField(FLD_GXY).Value = vbUnchecked
            frmFilter.chkField(FLD_GYY).Value = vbUnchecked
            frmFilter.chkField(FLD_SUP).Value = vbUnchecked
        '
        ' Stationary Target
        Case DAS_STF:
            '
            ' Set the default export file type
            guExport.iFile_Type = giDAS_STF
            '
            ' Enable/Disable the appropriate report type options
            frmFilter.optRpt(RPT_DOT).Enabled = True
            frmFilter.optRpt(RPT_DOT).Value = True
            frmFilter.optRpt(RPT_TRK).Enabled = True
            frmFilter.optRpt(RPT_TRK).Value = False
            frmFilter.optRpt(RPT_VEC).Enabled = True
            frmFilter.optRpt(RPT_VEC).Value = False
            frmFilter.optRpt(RPT_GEO).Enabled = True
            frmFilter.optRpt(RPT_GEO).Value = False
            '
            ' Check the appropriate fields for this type
            frmFilter.chkField(FLD_TIM).Value = vbChecked
            frmFilter.chkField(FLD_MSG).Value = vbChecked
            frmFilter.chkField(FLD_RPT).Value = vbChecked
            frmFilter.chkField(FLD_ORG).Value = vbChecked
            frmFilter.chkField(FLD_OID).Value = vbChecked
            frmFilter.chkField(FLD_TGT).Value = vbChecked
            frmFilter.chkField(FLD_LAT).Value = vbChecked
            frmFilter.chkField(FLD_LON).Value = vbChecked
            frmFilter.chkField(FLD_ALT).Value = vbChecked
            frmFilter.chkField(FLD_HDG).Value = vbUnchecked
            frmFilter.chkField(FLD_SPD).Value = vbUnchecked
            frmFilter.chkField(FLD_PAR).Value = vbChecked
            frmFilter.chkField(FLD_PID).Value = vbChecked
            frmFilter.chkField(FLD_ALG).Value = vbChecked
            frmFilter.chkField(FLD_IFF).Value = vbChecked
            frmFilter.chkField(FLD_EMT).Value = vbChecked
            frmFilter.chkField(FLD_EID).Value = vbChecked
            frmFilter.chkField(FLD_SIG).Value = vbChecked
            frmFilter.chkField(FLD_SID).Value = vbChecked
            frmFilter.chkField(FLD_FRQ).Value = vbChecked
            frmFilter.chkField(FLD_PRI).Value = vbChecked
            frmFilter.chkField(FLD_STA).Value = vbChecked
            frmFilter.chkField(FLD_TAG).Value = vbChecked
            frmFilter.chkField(FLD_FLG).Value = vbChecked
            frmFilter.chkField(FLD_COM).Value = vbChecked
            frmFilter.chkField(FLD_RNG).Value = vbUnchecked
            frmFilter.chkField(FLD_AZM).Value = vbUnchecked
            frmFilter.chkField(FLD_ELV).Value = vbUnchecked
            frmFilter.chkField(FLD_GXX).Value = vbUnchecked
            frmFilter.chkField(FLD_GXY).Value = vbUnchecked
            frmFilter.chkField(FLD_GYY).Value = vbUnchecked
            frmFilter.chkField(FLD_SUP).Value = vbUnchecked
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.optField Click (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    optRpt_Click
' AUTHOR:   Tom Elkins
' PURPOSE:  Configures the selected fields based on the selected report type
' TRIGGER:  User clicked on a report type
' INPUT:    "intRpt_Type" indicates which option was selected
' OUTPUT:   None
' NOTES:
Private Sub optRpt_Click(intRpt_Type As Integer)
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "EVENT    : frmFilter.optRpt Click (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & intRpt_Type
    End If
    '-v1.6.1
    '
    ' Save the output report type
    guExport.iRec_Type = intRpt_Type
    '
    ' Action is based on which option was selected
    Select Case intRpt_Type
        '
        ' DOT and TRK
        Case 0, 1:
            '
            ' Check the appropriate fields for this type
            frmFilter.chkField(25).Value = vbUnchecked
            frmFilter.chkField(26).Value = vbUnchecked
            frmFilter.chkField(27).Value = vbUnchecked
            frmFilter.chkField(28).Value = vbUnchecked
            frmFilter.chkField(29).Value = vbUnchecked
            frmFilter.chkField(30).Value = vbUnchecked
            frmFilter.chkField(31).Value = vbUnchecked
        '
        ' VEC - Lines of bearing
        Case 2:
            '
            ' Check the appropriate fields for this type
            frmFilter.chkField(25).Value = vbChecked
            frmFilter.chkField(26).Value = vbChecked
            frmFilter.chkField(27).Value = vbChecked
            frmFilter.chkField(28).Value = vbUnchecked
            frmFilter.chkField(29).Value = vbUnchecked
            frmFilter.chkField(30).Value = vbUnchecked
            frmFilter.chkField(31).Value = vbUnchecked
        '
        ' GEO - Geolocation
        Case 3:
            '
            ' Check the appropriate fields for this type
            frmFilter.chkField(25).Value = vbUnchecked
            frmFilter.chkField(26).Value = vbUnchecked
            frmFilter.chkField(27).Value = vbUnchecked
            frmFilter.chkField(28).Value = vbChecked
            frmFilter.chkField(29).Value = vbChecked
            frmFilter.chkField(30).Value = vbChecked
            frmFilter.chkField(31).Value = vbUnchecked
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.optRpt Click (End)"
    '-v1.6.1
    '
End Sub
'
' ROUTINE:  Create_SQL
' AUTHOR:   Tom Elkins
' PURPOSE:  Generates an SQL query based on a field list, filter clause, and sort list
' INPUT:    None
' OUTPUT:   None
' NOTES:    Uses the text stored in the guCurrent structure.  As elements of the
'           query are modified, the structure is updated.
Friend Sub Create_SQL()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmFilter.Create_SQL (Start)"
    '-v1.6.1
    '
    ' Create the field list and table portion
    frmFilter.txtSQL.Text = "SELECT " & guCurrent.uSQL.sFields & " FROM " & guCurrent.uSQL.sTable
    '
    ' If there is a filter, add the WHERE clause
    If Len(guCurrent.uSQL.sFilter) > 0 Then frmFilter.txtSQL.Text = frmFilter.txtSQL.Text & " WHERE " & guCurrent.uSQL.sFilter
    '
    ' If there is a sort list, add the ORDER BY clause
    If Len(guCurrent.uSQL.sOrder) > 0 Then frmFilter.txtSQL.Text = frmFilter.txtSQL.Text & " ORDER BY " & guCurrent.uSQL.sOrder
    '
    ' Update the query on the form
    guCurrent.uSQL.sQuery = frmFilter.txtSQL.Text
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmFilter.Create_SQL (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtFields_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Changes the field list for the SQL query
' TRIGGER:  The field list text box changed
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub txtFields_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFields Change (Start)"
    '-v1.6.1
    '
    ' See if the modification flag was set
    If mblnModify_Fields Then
        '
        ' Update the field list
        guCurrent.uSQL.sFields = frmFilter.txtFields.Text
        '
        ' Create the new query
        frmFilter.Create_SQL
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFields Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtFields_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets the field list modification flag
' TRIGGER:  The user entered the field list text box
' INPUT:    None
' OUTPUT:   None
' NOTES:    This ensures that the text will be updated only if the text box has the
'           focus, and not necessarily when the text box is updated externally.
Private Sub txtFields_GotFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFields GotFocus (Start)"
    '-v1.6.1
    '
    ' Set the modification flag
    mblnModify_Fields = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFields GotFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtFields_LostFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Resets the field list modification flag
' TRIGGER:  The user left the field list text box
' INPUT:    None
' OUTPUT:   None
' NOTES:    This ensures that the text will be updated only if the text box has the
'           focus, and not necessarily when the text box is updated externally.
Private Sub txtFields_LostFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFields LostFocus (Start)"
    '-v1.6.1
    '
    ' Reset the modification flag
    mblnModify_Fields = False
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFields LostFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtFilter_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the SQL filter clause
' TRIGGER:  The user modified the filter
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub txtFilter_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFilter Change (Start)"
    '-v1.6.1
    '
    ' See if the modification flag is set
    If mblnModify_Filter Then
        '
        ' Update the filter
        guCurrent.uSQL.sFilter = frmFilter.txtFilter.Text
        '
        ' Create the SQL query
        frmFilter.Create_SQL
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFilter Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtFilter_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets the filter modification flag
' TRIGGER:  The user entered the filter text box
' INPUT:    None
' OUTPUT:   None
' NOTES:    This ensures that the text will be updated only if the text box has the
'           focus, and not necessarily when the text box is updated externally.
Private Sub txtFilter_GotFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFilter GotFocus (Start)"
    '-v1.6.1
    '
    ' Set the modification flag
    mblnModify_Filter = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFilter GotFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtFilter_LostFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Resets the filter modification flag
' TRIGGER:  The user left the filter text box
' INPUT:    None
' OUTPUT:   None
' NOTES:    This ensures that the text will be updated only if the text box has the
'           focus, and not necessarily when the text box is updated externally.
Private Sub txtFilter_LostFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFilter LostFocus (Start)"
    '-v1.6.1
    '
    ' Reset the modification flag
    mblnModify_Filter = False
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtFilter LostFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtSort_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the SQL sort clause
' TRIGGER:  The user modified the sort
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub txtSort_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSort Change (Start)"
    '-v1.6.1
    '
    ' See if the modification flag is set
    If mblnModify_Sort Then
        '
        ' Update the sort list
        guCurrent.uSQL.sOrder = frmFilter.txtSort.Text
        '
        ' Create a new SQL query
        frmFilter.Create_SQL
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSort Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtSort_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets the sort modification flag
' TRIGGER:  The user entered the sort text box
' INPUT:    None
' OUTPUT:   None
' NOTES:    This ensures that the text will be updated only if the text box has the
'           focus, and not necessarily when the text box is updated externally.
Private Sub txtSort_GotFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSort GotFocus (Start)"
    '-v1.6.1
    '
    ' Set the modification flag
    mblnModify_Sort = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSort GotFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtSort_LostFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Resets the sort modification flag
' TRIGGER:  The user left the sort text box
' INPUT:    None
' OUTPUT:   None
' NOTES:    This ensures that the text will be updated only if the text box has the
'           focus, and not necessarily when the text box is updated externally.
Private Sub txtSort_LostFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSort LostFocus (Start)"
    '-v1.6.1
    '
    ' Rest the modification flag
    mblnModify_Sort = False
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSort LostFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtSQL_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Parses the SQL query and updates the elements
' TRIGGER:  The user modified the query
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub txtSQL_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSQL Change (Start)"
    '-v1.6.1
    '
    ' See if the modification flag is set
    If mblnModify_SQL Then
        '
        ' Parse the string
        basDatabase.Parse_SQL (frmFilter.txtSQL.Text)
        '
        ' Update the controls with the new elements
        frmFilter.txtFields.Text = guCurrent.uSQL.sFields
        frmFilter.txtFilter.Text = guCurrent.uSQL.sFilter
        frmFilter.txtSort.Text = guCurrent.uSQL.sOrder
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSQL Change (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtSQL_GotFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Sets the SQL modification flag
' TRIGGER:  The user entered the SQL text box
' INPUT:    None
' OUTPUT:   None
' NOTES:    This ensures that the text will be updated only if the text box has the
'           focus, and not necessarily when the text box is updated externally.
Private Sub txtSQL_GotFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSQL GotFocus (Start)"
    '-v1.6.1
    '
    ' Set the modification flag
    mblnModify_SQL = True
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSQL GotFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtSQL_LostFocus
' AUTHOR:   Tom Elkins
' PURPOSE:  Resets the SQL modification flag
' TRIGGER:  The user left the SQL text box
' INPUT:    None
' OUTPUT:   None
' NOTES:    This ensures that the text will be updated only if the text box has the
'           focus, and not necessarily when the text box is updated externally.
Private Sub txtSQL_LostFocus()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSQL LostFocus (Start)"
    '-v1.6.1
    '
    ' Reset the modification flag
    mblnModify_SQL = False
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtSQL LostFocus (End)"
    '-v1.6.1
    '
End Sub
'
' EVENT:    txtValue_Change
' AUTHOR:   Tom Elkins
' PURPOSE:  Updates the value of the filter clause
' TRIGGER:  The user entered a value in the value entry box
' INPUT:    None
' OUTPUT:   None
' NOTES:
Private Sub txtValue_Change()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtValue Change (Start)"
    '-v1.6.1
    '
    ' Determine which clause by the location of the entry box
    Select Case frmFilter.txtValue.Top
        '
        ' First clause
        Case frmFilter.lblVal1.Top:
            frmFilter.lblVal1.Caption = frmFilter.txtValue.Text
        '
        ' Second clause
        Case frmFilter.lblVal2.Top:
            frmFilter.lblVal2.Caption = frmFilter.txtValue.Text
        '
        ' Third clause
        Case frmFilter.lblVal3.Top:
            frmFilter.lblVal3.Caption = frmFilter.txtValue.Text
        '
        ' Fourth clause
        Case frmFilter.lblVal4.Top:
            frmFilter.lblVal4.Caption = frmFilter.txtValue.Text
    End Select
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmFilter.txtValue Change (End)"
    '-v1.6.1
    '
End Sub
'
'+v1.5
' ROUTINE:  SaveCustomFilter
' AUTHOR:   Tom Elkins
' PURPOSE:  Saves the current query to the INI file
' INPUT:    Optional "strName" - the name for the saved query
' OUTPUT:   None
' NOTES:    If "strName" is provided, there is no user interaction
Public Sub SaveCustomFilter(Optional strName As String)
    Dim plngMaxQueries As Long  ' The maximum # of queries in the INI file
    Dim plngLength As Long      ' The write status for the INI operation
    Dim pstrTemplate As String  ' The set of query tokens in the INI file
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then
        basCCAT.WriteLogEntry "ROUTINE  : frmFilter.SaveCustomFilter (Start)"
        basCCAT.WriteLogEntry "ARGUMENTS: " & strName
    End If
    '-v1.6.1
    '
    ' Get the current maximum query count
    plngMaxQueries = basCCAT.GetNumber("Queries", "Max_Queries", 0)
    '
    ' Add 1
    plngMaxQueries = plngMaxQueries + 1
    '
    ' Save the new maximum count
    plngLength = lPutINIString("Queries", "Max_Queries", CStr(plngMaxQueries), basCCAT.gsCCAT_INI_Path)
    '
    ' Get the name of the new query
    If strName = "" Then strName = InputBox("Enter the name for the new query", "New Query Name", "Query " & plngMaxQueries)
    '
    ' Write a spacer, header comment, and templates for the proper order of keys
    pstrTemplate = ";" & vbCrLf & "; Query" & plngMaxQueries & " - " & strName & vbCrLf & "; CCAT Save on " & Now & vbCrLf & "QUERY_TITLE" & plngMaxQueries & "=" & vbCrLf & "QUERY_FIELDS" & plngMaxQueries & "=*" & vbCrLf & "QUERY" & plngMaxQueries & "=" & vbCrLf & "QUERY_SORT" & plngMaxQueries
    plngLength = lPutINIString("Queries", pstrTemplate, "", gsCCAT_INI_Path)
    '
    ' Save the name to the ini file
    plngLength = lPutINIString("Queries", "QUERY_TITLE" & plngMaxQueries, strName, gsCCAT_INI_Path)
    '
    ' Save the field list
    If Len(guCurrent.uSQL.sFields) > 0 Then
        plngLength = lPutINIString("Queries", "QUERY_FIELDS" & plngMaxQueries, guCurrent.uSQL.sFields, gsCCAT_INI_Path)
    End If
    '
    ' Save the filter
    plngLength = lPutINIString("Queries", "QUERY" & plngMaxQueries, guCurrent.uSQL.sFilter, gsCCAT_INI_Path)
    '
    ' Save the sort list
    If Len(guCurrent.uSQL.sOrder) > 0 Then
        plngLength = lPutINIString("Queries", "QUERY_SORT" & plngMaxQueries, guCurrent.uSQL.sOrder, gsCCAT_INI_Path)
    End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmFilter.SaveCustomFilter (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
'
'+v1.5
' ROUTINE:  ClearQuery
' AUTHOR:   Tom Elkins
' PURPOSE:  Clears the query from the structure and the form
' INPUT:    None
' OUTPUT:   None
' NOTES:
Public Sub ClearQuery()
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmFilter.ClearQuery (Start)"
    '-v1.6.1
    '
    '
    ' Blank out all strings
    frmFilter.txtFields.Text = ""
    guCurrent.uSQL.sFields = ""
    frmFilter.txtFilter.Text = ""
    guCurrent.uSQL.sFilter = ""
    frmFilter.txtSort.Text = ""
    guCurrent.uSQL.sOrder = ""
    frmFilter.txtSQL.Text = ""
    guCurrent.uSQL.sQuery = ""
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "ROUTINE  : frmFilter.ClearQuery (End)"
    '-v1.6.1
    '
End Sub
'-v1.5
