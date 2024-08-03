VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCanRpt 
   Caption         =   "Canned Reports"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Report Scratchpad"
      Height          =   7815
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8655
      Begin VB.ComboBox ReportCombo 
         Height          =   315
         ItemData        =   "frmCanRpt.frx":0000
         Left            =   240
         List            =   "frmCanRpt.frx":0002
         TabIndex        =   5
         Text            =   "Select a Report Type"
         Top             =   7080
         Width           =   3015
      End
      Begin VB.CommandButton ClearCommand 
         Caption         =   "Clear Scratchpad"
         Height          =   615
         Left            =   5160
         TabIndex        =   4
         Top             =   6840
         Width           =   1575
      End
      Begin VB.CommandButton SaveCommand 
         Caption         =   "Save Scratchpad"
         Height          =   615
         Left            =   3480
         TabIndex        =   3
         Top             =   6840
         Width           =   1575
      End
      Begin VB.CommandButton CancelCommand 
         Caption         =   "Cancel/Exit"
         Height          =   615
         Left            =   6840
         TabIndex        =   2
         Top             =   6840
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   6135
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   10821
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmCanRpt.frx":0004
      End
      Begin VB.Label Label1 
         Caption         =   "Canned Reports"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   6720
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCanRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelCommand_Click()
   Unload Me
End Sub

Private Sub ClearCommand_Click()
    RichTextBox1.Text = ""
End Sub

Private Sub Form_Activate()

     
         With frmCanRpt.ReportCombo
            .Clear
            .AddItem "Run Mode Changes"
            .AddItem "Jam Periods"
            .AddItem "Target List"
            .AddItem "Tgt List During Jam"
            .AddItem "Hardware Status"
        End With
       
            frmCanRpt.Show , frmMain
           
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    'If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    :  (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
   
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Me.Hide

End Sub


Private Sub RptRunmode()
    Dim rsMessage As Recordset  ' Pointer to records in the Message table
    Dim sLastTime As String
    Dim sReport As String
    Dim sLastRunmode As String
    Dim sTime() As String
    Dim sMode() As String
    
        
    Set rsMessage = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_Data", dbOpenDynaset)
    
    rsMessage.FindFirst "Msg_Type = 'MTRUNMODE'"
    '
    ' Look for a match
    If (rsMessage.NoMatch = False) Then
        sTime = Split(rsMessage!ReportTime, " ", 2)
        sReport = "Runmode Changes for " & sTime(0) & vbCrLf & vbCrLf

        While (rsMessage.NoMatch = False)
            If (sLastTime <> rsMessage!ReportTime) Then
                sMode = Split(rsMessage!Other_Data, "SYS@")
                sMode = Split(sMode(1), ",")
                If (sMode(0) <> sLastRunmode) Then
                    sReport = sReport & rsMessage!ReportTime & " , " & sMode(0) & vbCrLf
                    sLastRunmode = sMode(0)
                End If
                sLastTime = rsMessage!ReportTime
            End If
            rsMessage.FindNext "Msg_Type = 'MTRUNMODE'"
        Wend
        RichTextBox1.Text = sReport
    End If
    rsMessage.Close
    
'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    'If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basTOC.Get_Message_Name (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
   ' Err.Raise Err.Number, "CCAT:Get_Message_Name", Err.Description
End Sub


Private Sub RptJamPeriod()
    Dim rsMessage As Recordset  ' Pointer to records in the Message table
    Dim sLastTime As String
    Dim sReport As String
    Dim sLastRunmode As String
    Dim blnWet As Boolean
    Dim sMode() As String
    Dim sTime() As String
    
    Set rsMessage = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_Data", dbOpenDynaset)
    
    rsMessage.FindFirst "Msg_Type = 'MTRUNMODE'"
    
    If (rsMessage.NoMatch = False) Then
        sTime = Split(rsMessage!ReportTime, " ", 2)
        sReport = "Jam Periods for " & sTime(0) & vbCrLf & vbCrLf
        blnWet = False
    
    '
    ' Look for a match

        While (rsMessage.NoMatch = False)
            If (sLastTime <> rsMessage!ReportTime) Then
                sMode = Split(rsMessage!Other_Data, "SYS@")
                sMode = Split(sMode(1), ",")
                sTime = Split(rsMessage!ReportTime, " ", 2)
                If (sMode(0) <> sLastRunmode) Then
                    If (sMode(0) = "JAM") Then
                        blnWet = True
                        sReport = sReport & sTime(1) & ", - ,"
                    ElseIf (blnWet = True) Then
                        blnWet = False
                        sReport = sReport & sTime(1) & vbCrLf
                    End If
                    sLastRunmode = sMode(0)
                End If
                sLastTime = rsMessage!ReportTime
            End If
            rsMessage.FindNext "Msg_Type = 'MTRUNMODE'"
        Wend
        RichTextBox1.Text = sReport
    End If
    rsMessage.Close

'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    'If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basTOC.Get_Message_Name (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
   ' Err.Raise Err.Number, "CCAT:Get_Message_Name", Err.Description
End Sub


Private Sub RptTgtList4Jam()
    Dim rsMessage As Recordset  ' Pointer to records in the Message table
    Dim rsMessage2 As Recordset  ' Pointer to records in the Message table
    Dim sLastTime As String
    Dim sLastTime2 As String
    Dim sReport As String
    Dim sLastRunmode As String
    Dim blnWet As Boolean
    Dim sMode() As String
    Dim sTime() As String
    Dim sJamStart As String
    Dim sJamStop As String
    Dim blnJamLoop As Boolean
    Dim blnNewList As Boolean
    Dim i As Integer
    Dim iTgt As Integer
    Dim lTgtTotal As Long
    Dim sTgtList(1 To 7) As String
    Dim iTgtList(1 To 7) As Integer
    Dim lTgtListTotal(1 To 7) As Long
    Dim lNumRpts As Long
    Dim lTgtJam As Long
    Dim lTgtListJam(1 To 7) As Long
    Dim lNumRptsJam As Long
    Dim sJamSummary As String
    Dim iJamPeriod As Integer
   
    
    
    sTgtList(1) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 4, "UNKNOWN")
    sTgtList(2) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 8, "UNKNOWN")
    sTgtList(3) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 10, "UNKNOWN")
    sTgtList(4) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 13, "UNKNOWN")
    sTgtList(5) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 15, "UNKNOWN")
    sTgtList(6) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 29, "UNKNOWN")
    sTgtList(7) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 31, "UNKNOWN")
    
    
    
    Set rsMessage = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_Data", dbOpenDynaset)
    Set rsMessage2 = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_Data", dbOpenDynaset)
    
    rsMessage.FindFirst "Msg_Type = 'MTRUNMODE'"
    
    If (rsMessage.NoMatch = False) Then
        rsMessage2.FindFirst "Msg_Type = 'MTJAMSTAT'"
        sTime = Split(rsMessage!ReportTime, " ", 2)
        iJamPeriod = 0
        sJamSummary = sJamSummary & vbCrLf & "Jam Period Averages" & vbCrLf
        sJamSummary = sJamSummary & "Period"
        For i = UBound(iTgtList) To LBound(iTgtList) Step -1
            sJamSummary = sJamSummary & "," & sTgtList(i)
        Next i
        sJamSummary = sJamSummary & "," & "Total Tgts" & vbCrLf

        sReport = "Target List During Jam Periods for " & sTime(0) & vbCrLf & vbCrLf
        sReport = sReport & "Time"
        For i = UBound(iTgtList) To LBound(iTgtList) Step -1
            sReport = sReport & "," & sTgtList(i)
        Next i
        sReport = sReport & "," & "Total Tgts" & vbCrLf
        blnWet = False
    
    '
    ' Look for a match

        While (rsMessage.NoMatch = False)
            If (sLastTime <> rsMessage!ReportTime) Then
                sMode = Split(rsMessage!Other_Data, "SYS@")
                sMode = Split(sMode(1), ",")
                sTime = Split(rsMessage!ReportTime, " ", 2)
                If (sMode(0) <> sLastRunmode) Then
                    If (sMode(0) = "JAM") Then
                        blnWet = True
                        sJamStart = rsMessage!ReportTime
                        blnJamLoop = True
                        'sReport = sReport & vbCrLf & "Jam Period" & vbCrLf
                        'sReport = sReport & sTime(1) & ", - ,"
                    ElseIf (blnWet = True) Then
                        blnWet = False
                        sJamStop = rsMessage!ReportTime
                        lTgtJam = 0
                        lNumRptsJam = 0
                        For i = LBound(iTgtList) To UBound(iTgtList)
                            lTgtListJam(i) = 0
                        Next i
                        
                        'sReport = sReport & sTime(1) & vbCrLf
                        While ((rsMessage2.NoMatch = False) And (blnJamLoop = True))
                            If (rsMessage2!ReportTime > sJamStart) Then
                                If (rsMessage2!ReportTime < sJamStop) Then
                                    If (sLastTime2 <> rsMessage2!ReportTime) Then
                                        If (sLastTime2 <> "") Then
                                            sTime = Split(rsMessage2!ReportTime, " ", 2)
                                            sReport = sReport & sTime(1)
                                            'sReport = sReport & vbCrLf & sTime(1) & vbCrLf
                                            'sReport = sReport & "Total Signals : " & iTgt & vbCrLf
                                            lNumRpts = lNumRpts + 1
                                            lNumRptsJam = lNumRptsJam + 1
                                            For i = UBound(iTgtList) To LBound(iTgtList) Step -1
                                                sReport = sReport & ","
                                                If (iTgtList(i) > 0) Then
                                                    sReport = sReport & iTgtList(i)
                                                    lTgtListTotal(i) = lTgtListTotal(i) + iTgtList(i)
                                                    lTgtListJam(i) = lTgtListJam(i) + iTgtList(i)
                                                    iTgtList(i) = 0
                                                End If
                                            Next i
                                            sReport = sReport & "," & iTgt & vbCrLf
                                            lTgtTotal = lTgtTotal + iTgt
                                            lTgtJam = lTgtJam + iTgt
                                            iTgt = 0
                                        End If
                                    End If
                                    
                                                                         
                                     Select Case rsMessage2!Status
                                     
                                        Case 4:
                                                iTgt = iTgt + 1
                                                iTgtList(1) = iTgtList(1) + 1
                                        Case 8:
                                                iTgt = iTgt + 1
                                                iTgtList(2) = iTgtList(2) + 1
                                        
                                        Case 10:
                                                iTgt = iTgt + 1
                                                iTgtList(3) = iTgtList(3) + 1
                                        
                                        Case 13:
                                                iTgt = iTgt + 1
                                                iTgtList(4) = iTgtList(4) + 1
                                        
                                        Case 15:
                                                iTgt = iTgt + 1
                                                iTgtList(5) = iTgtList(5) + 1
                                        
                                        Case 29:
                                                iTgt = iTgt + 1
                                                iTgtList(6) = iTgtList(6) + 1
                                        
                                        Case 31:
                                                iTgt = iTgt + 1
                                                iTgtList(7) = iTgtList(7) + 1
                                        
                                        
                                    End Select
                                Else
                                    blnJamLoop = False
                                End If
                            End If
                            sLastTime2 = rsMessage2!ReportTime
                            rsMessage2.FindNext "Msg_Type = 'MTJAMSTAT'"
                        Wend
                        iJamPeriod = iJamPeriod + 1
                        sJamSummary = sJamSummary & iJamPeriod
                        For i = UBound(lTgtListJam) To LBound(lTgtListJam) Step -1
                            sJamSummary = sJamSummary & ","
                            If (lTgtListJam(i) > 0) Then
                                sJamSummary = sJamSummary & Format(lTgtListJam(i) / lNumRptsJam, "0.00")
                            End If
                        Next i
                        sJamSummary = sJamSummary & "," & lTgtJam & vbCrLf
                    End If
                    sLastRunmode = sMode(0)
                End If
                sLastTime = rsMessage!ReportTime
            End If
            rsMessage.FindNext "Msg_Type = 'MTRUNMODE'"
        Wend
        sReport = sReport & vbCrLf & "Totals" & vbCrLf
        sReport = sReport & sJamSummary & vbCrLf
        sReport = sReport & "Total Signals , " & lTgtTotal & vbCrLf
        If (lNumRpts > 0) Then
            sReport = sReport & "Average Targets Per Report , " & Format(lTgtTotal / lNumRpts, "0.00") & vbCrLf & vbCrLf
        End If
        For i = LBound(iTgtList) To UBound(iTgtList)
            lTgtListTotal(i) = lTgtListTotal(i)
            If (lTgtListTotal(i) > 0) Then
                sReport = sReport & "Total " & sTgtList(i) & " , " & lTgtListTotal(i) & vbCrLf
            End If
        Next i

        RichTextBox1.Text = sReport
    End If
    rsMessage.Close

'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    'If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basTOC.Get_Message_Name (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
   ' Err.Raise Err.Number, "CCAT:Get_Message_Name", Err.Description
End Sub

Private Sub RptTgtList()
    Dim rsMessage As Recordset  ' Pointer to records in the Message table
    Dim sLastTime As String
    Dim sReport As String
    Dim sLastRunmode As String
    Dim blnNewList As Boolean
    Dim i As Integer
    Dim iTgt As Integer
    Dim lTgtTotal As Long
    Dim sMode() As String
    Dim sTime() As String
    Dim sTgtList(1 To 7) As String
    Dim iTgtList(1 To 7) As Integer
    Dim lTgtListTotal(1 To 7) As Long
    Dim lNumRpts As Long
    
    
    sTgtList(1) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 4, "UNKNOWN")
    sTgtList(2) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 8, "UNKNOWN")
    sTgtList(3) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 10, "UNKNOWN")
    sTgtList(4) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 13, "UNKNOWN")
    sTgtList(5) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 15, "UNKNOWN")
    sTgtList(6) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 29, "UNKNOWN")
    sTgtList(7) = basCCAT.GetAlias("JAMSTAT", "JAMSTAT" & 31, "UNKNOWN")
    
    Set rsMessage = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_Data", dbOpenDynaset)
    
    rsMessage.FindFirst "Msg_Type = 'MTJAMSTAT'"
    
   If (rsMessage.NoMatch = False) Then
        sTime = Split(rsMessage!ReportTime, " ", 2)
        sReport = "LMB Tgt List Summary for " & sTime(0) & vbCrLf & vbCrLf
    
    '
    ' Look for a match

        While (rsMessage.NoMatch = False)
            If (sLastTime <> rsMessage!ReportTime) Then
                If (sLastTime <> "") Then
                    sTime = Split(rsMessage!ReportTime, " ", 2)
                    If (UBound(sTime) > 0) Then
                        sReport = sReport & vbCrLf & sTime(1) & vbCrLf
                    End If
                    sReport = sReport & "Total Signals : " & iTgt & vbCrLf
                    lNumRpts = lNumRpts + 1
                    For i = LBound(iTgtList) To UBound(iTgtList)
                        If (iTgtList(i) > 0) Then
                            sReport = sReport & sTgtList(i) & " : " & iTgtList(i) & vbCrLf
                            lTgtListTotal(i) = lTgtListTotal(i) + iTgtList(i)
                            iTgtList(i) = 0
                        End If
                    Next i
                    lTgtTotal = lTgtTotal + iTgt
                    iTgt = 0
                 End If
             End If
             
             Select Case rsMessage!Status
             
                Case 4:
                        iTgt = iTgt + 1
                        iTgtList(1) = iTgtList(1) + 1
                Case 8:
                        iTgt = iTgt + 1
                        iTgtList(2) = iTgtList(2) + 1
                
                Case 10:
                        iTgt = iTgt + 1
                        iTgtList(3) = iTgtList(3) + 1
                
                Case 13:
                        iTgt = iTgt + 1
                        iTgtList(4) = iTgtList(4) + 1
                
                Case 15:
                        iTgt = iTgt + 1
                        iTgtList(5) = iTgtList(5) + 1
                
                Case 29:
                        iTgt = iTgt + 1
                        iTgtList(6) = iTgtList(6) + 1
                
                Case 31:
                        iTgt = iTgt + 1
                        iTgtList(7) = iTgtList(7) + 1
                
                
            End Select
            
            sLastTime = rsMessage!ReportTime
            rsMessage.FindNext "Msg_Type = 'MTJAMSTAT'"
        Wend
        sReport = sReport & vbCrLf & "Totals" & vbCrLf
        sReport = sReport & "Total Signals , " & lTgtTotal & vbCrLf
        sReport = sReport & "Average Targets Per Report , " & lTgtTotal / lNumRpts & vbCrLf & vbCrLf
        For i = LBound(iTgtList) To UBound(iTgtList)
            lTgtListTotal(i) = lTgtListTotal(i) + iTgtList(i)
            If (lTgtListTotal(i) > 0) Then
                sReport = sReport & "Total " & sTgtList(i) & " , " & lTgtListTotal(i) & vbCrLf
            End If
        Next i
        
        RichTextBox1.Text = sReport
    End If
    rsMessage.Close

'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    'If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basTOC.Get_Message_Name (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
   ' Err.Raise Err.Number, "CCAT:Get_Message_Name", Err.Description
End Sub

Private Sub RptHWStatus()
    Dim rsMessage As Recordset  ' Pointer to records in the Message table
    Dim sLastTime As String
    Dim sReport As String
    Dim blnLogStat As Boolean
    Dim sStatus() As String
    Dim sTime() As String
    
    Set rsMessage = guCurrent.DB.OpenRecordset(guCurrent.sArchive & "_Data", dbOpenDynaset)
    
    rsMessage.FindFirst "Msg_Type = 'MTSYSCONFIG'"
    
    If (rsMessage.NoMatch = False) Then
        sTime = Split(rsMessage!ReportTime, " ", 2)
        sReport = "Hardware Status for " & sTime(0) & vbCrLf & vbCrLf
        blnLogStat = True
    
    '
    ' Look for a match

        While (rsMessage.NoMatch = False)
            If (sLastTime <> rsMessage!ReportTime) Then
                blnLogStat = True
                sTime = Split(rsMessage!ReportTime, " ", 2)
                sReport = sReport & vbCrLf
                
            End If
            If (blnLogStat = True) Then
                sStatus = Split(rsMessage!Other_Data, "@")

                If (sStatus(0) = "TECH") Then
                    blnLogStat = False
                End If
                sReport = sReport & sTime(1) & " , " & sStatus(0) & " , " & sStatus(1) & vbCrLf
                sLastTime = rsMessage!ReportTime
            End If
            rsMessage.FindNext "Msg_Type = 'MTSYSCONFIG'"
        Wend
    Else
        sReport = "Hardware Status : No Reports " & vbCrLf & vbCrLf
    End If
    RichTextBox1.Text = sReport
    rsMessage.Close

'
'
ERR_HANDLER:
    '
    '+v1.6.1TE
    'If basCCAT.Verbose Then basCCAT.WriteLogEntry "ERROR    : basTOC.Get_Message_Name (Error #" & Err.Number & " - " & Err.Description & ")"
    '-v1.6.1
    '
   ' Err.Raise Err.Number, "CCAT:Get_Message_Name", Err.Description
End Sub
Private Sub ReportCombo_Click()

    MousePointer = vbHourglass
    Select Case ReportCombo.Text

        Case "Run Mode Changes":
            RptRunmode
        
        Case "Jam Periods":
            RptJamPeriod
        
        Case "Target List":
            RptTgtList
            
        Case "Tgt List During Jam":
            RptTgtList4Jam
        
        Case "Hardware Status":
            RptHWStatus

    End Select
    MousePointer = vbDefault

    
End Sub

Private Sub SaveCommand_Click()
'
' EVENT:    SaveCommand_Click
' AUTHOR:   Brad Brown
' PURPOSE:  Saves data to an external file
' INPUT:    None
' OUTPUT:   None
' NOTES:


   Dim strNewFile As String

    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmMain.SaveCommand_Click Click (Start)"
    '-v1.6.1
    '
    ' Use control-level addressing
    With frmMain.dlgCommonDialog
        '
        ' Change the title on the dialog
        .DialogTitle = "Save data as..."
        '
        ' Blank out the file name
        .FileName = ""
        '
        ' Set the filters to the various export types
        .Filter = "Comma-delimited Text (*.csv)|*.csv|Text (*.txt)|*.txt"
        '
        ' Set the flags
        .Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
        '
        ' Set the default file type
        '.FilterIndex = guExport.iFile_Type
        '
        ' Give it to the user
        .ShowSave
        '
        ' Save the filename
        strNewFile = .FileName
    End With
    '
    ' Check for a filename
    If Len(strNewFile) > 0 Then
        ' Log the event
        basCCAT.WriteLogEntry "frmCanRpt: SaveCommand_Click: Output File = " & strNewFile
        '
        RichTextBox1.SaveFile strNewFile, rtfText
        
     End If
    '
    '+v1.6.1TE
    If basCCAT.Verbose Then basCCAT.WriteLogEntry "EVENT    : frmCanRpt: SaveCommand_Click (End)"
    '-v1.6.1
    '
End Sub








