Attribute VB_Name = "GenModule"
Option Explicit

Private ErrMsg As String
Private Const SND_ASYNC = &H1     ' Play asynchronously
Private Const SND_NODEFAULT = &H2 ' Don't use default sound
Private Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file
Private Const SND_PURGE = &H40
Private SoundBuffer() As Byte

'
'**************************************************************************************************
' This program processes the annual workitems. It gets the next available workitem in the Annuals Robot
' workset and checks whether the workitem belongs to one of the classes mentioned in the settings.
' If the workitem doesn't belong to any listed class, the workitem will be routed to Annuals Workstep.
' If the workitem belongs to any one of classes, then the workitem is processed.
' Each workitem will be opened in Read/Write mode and call an Orcale stored procedure which will
' take the file number, workitem-ID, annual Type (clean or Manaual), Receive Date, Year of Filing and
' user-id (in this case the user-id which is being used by this program), the return values from the stored
' procedure are CompanyName, Return Code and Return Message.
' If the error code is "00" from the stored procedure, then the company name is stored in the workitem and
' then defaulted to filing queue, if the workitem is scanned in ANNUALS-C batch. If the workitem is scanned in
' ANNUALS-M batch, then workitem is routed to Annuals workstep. -Ravi Gavarasana
'
'*****************************************************************************************************************
'
'*****************************************************************************************************************
' Change Description                                        | Date of change        | Changed by
'*****************************************************************************************************************
' Rewrote the program for BRIMS integration                 | 11/12/2001            | Ravi (Covansys)
'-----------------------------------------------------------|-----------------------|----------------------------
' Mark the process status as Completed instead of Approved  | 12/05/2002            | Ravi (Covansys)
' If the processing of the clean annual is successful       |                       |
'-----------------------------------------------------------|-----------------------|----------------------------
' Changed program to include folder procesing also          | 03/20/2002            | Ravi (Covansys)
'-----------------------------------------------------------|-----------------------|----------------------------
' Changed program to include resub processing. Annual resubs| 03/25/2002            | Ravi
' will be routed to annuals robot queue and this program    |                       |
' will notify BRIMS that a new resub annual has been        |                       |
' received.                                                 |                       |
'-----------------------------------------------------------|-----------------------|----------------------------
'                                                           |                       |
'*****************************************************************************************************************
'
Public Sub Main()
'
' Read the registry and get the settings, if settings are missing then show the options form.
' If options form is loaded then, save settings and get the settings.
' If settings are ok, then logon to workflow and process documents.
'
    Dim objfrmOptions As frmOptions
    Dim objAboutForm As frmAbout
    Dim intLoopCnt As Integer
    Dim strWeekDay As String
    Dim strMsg As String
    Dim dtEndTime As Date
    Dim blnDoNotProceed As Boolean
        
    On Error GoTo Main_Error
    

    If App.PrevInstance Then
        ErrMsg = "A previous instance of this application is running." & vbCrLf
        ErrMsg = ErrMsg & "Cannot run two instances. Terminating application."
        MsgBox ErrMsg, vbApplicationModal + vbCritical, App.EXEName
        Exit Sub
    End If
    
    Set objAboutForm = New frmAbout
    Load objAboutForm
    objAboutForm.cmdOK.Visible = False
    objAboutForm.lblMsg.Caption = "Connecting. Please wait...."
    DoEvents
    objAboutForm.Show
'Create LogFile
    CreateLogFile

' Check whether all required values are there or not.
    If Not CheckSettingValues Then
        If Not objAboutForm Is Nothing Then
            Unload objAboutForm
            Set objAboutForm = Nothing
        End If
        DoEvents
        Set objMainForm = New frmMain
        DoEvents
        Load objMainForm
        With objMainForm
            .AddItemToList "Check Settings failed. Please check Properties."
            DoEvents
            .StatusBar1.Panels(1).Text = "0 out of " & Trim$(Str$(mintSleepTime)) & " minutes passed."
            .StatusBar1.Panels(2).Text = vbNullString
            .mnuStart.Enabled = False
            .tmrSleep.Enabled = True
            .mnuStop.Enabled = True
            .mnuStartNow.Enabled = True
             DoEvents
            .mnuProperties.Enabled = False
            .mnuEnd.Enabled = False
            .Caption = "AnnRobot - Sleeping"
        End With
        DoEvents
        objMainForm.Show
        DoEvents
        Exit Sub
    End If
    
'Check whether the robot is specified to run today or not.
    strWeekDay = Trim$(Str$(Weekday(Date)))
    dtEndTime = CDate(gstrEndtime)
    If InStr(1, gstrRundays, strWeekDay, vbTextCompare) = 0 Then
       'Stop the program if it not scheduled to run today.
        blnDoNotProceed = True
        Select Case strWeekDay
            Case Is = 1
                strMsg = "Sunday"
            Case Is = 2
                strMsg = "Monday"
            Case Is = 3
                strMsg = "Tuesday"
            Case Is = 4
                strMsg = "Wednesday"
            Case Is = 5
                strMsg = "Thursday"
            Case Is = 6
                strMsg = "Friday"
            Case Is = 7
                strMsg = "Saturday"
            Case Else
                strMsg = "Unknown"
        End Select
        DoEvents
        strMsg = "Robot was specified NOT to run on " & strMsg & "." '  & vbCrLf & "Shutting down the robot."""
    ElseIf DateDiff("n", dtEndTime, Time) > 0 Then
        blnDoNotProceed = True      'Stop the program once it reaches the end time
        strMsg = "Robot was specified Not to run after " & gstrEndtime & "."
    End If
        
        
    If blnDoNotProceed Then
        WriteToLogFile strMsg
        If Not objAboutForm Is Nothing Then
            objAboutForm.lblMsg.Caption = strMsg
            DoEvents
            objAboutForm.Show
            DoEvents
            Sleep 3500
            strMsg = "Shutting down the robot...."
            objAboutForm.lblMsg.Caption = strMsg
            WriteToLogFile strMsg
            DoEvents
            Sleep 1000
            DoEvents
            Unload objAboutForm
            Set objAboutForm = Nothing
            DestroyGlobalObjects
        End If
        Exit Sub
    End If
    
' Log onto the domain.
   
    If Not LogonToWorkflow Then
        If Not objAboutForm Is Nothing Then
            Unload objAboutForm
            Set objAboutForm = Nothing
        End If
        DoEvents
        Set objMainForm = New frmMain
        DoEvents
        Load objMainForm
        With objMainForm
            DoEvents
            .mnuStart.Enabled = False
            .tmrSleep.Enabled = True
             DoEvents
            .mnuStartNow.Enabled = True
            .mnuStop.Enabled = True
            .mnuProperties.Enabled = False
            .mnuEnd.Enabled = False
            .Caption = "AnnRobot - Sleeping"
             DoEvents
            .StatusBar1.Panels(1).Picture = LoadResPicture(2, vbResIcon)
             DoEvents
             DoEvents
            .AddItemToList "Login Failed. Check the log file for the reason"
             DoEvents
            .StatusBar1.Panels(1).Text = "Login failed. Check log."
             DoEvents
            .StatusBar1.Panels(2).Text = vbNullString
             DoEvents
        End With
        DoEvents
        objMainForm.Show
        DoEvents
        DoEvents
        Exit Sub
    End If
        
    objAboutForm.lblMsg.Caption = "Validating worksets and classes"
    DoEvents
    If Not CheckWorksetsNClasses Then
        If Not objAboutForm Is Nothing Then
            Unload objAboutForm
            Set objAboutForm = Nothing
        End If
        DoEvents
        Set objMainForm = New frmMain
        DoEvents
        Load objMainForm
        With objMainForm
            .AddItemToList FormatDateTime$(Now, vbGeneralDate) & " - " & " Worksets and Classes check failed, See log file."
             DoEvents
             DoEvents
            .StatusBar1.Panels(1).Picture = LoadResPicture(2, vbResIcon)
             DoEvents
            .StatusBar1.Panels(1).Text = "0 out of " & Trim$(Str$(mintSleepTime)) & " minutes passed."
            .StatusBar1.Panels(2).Text = vbNullString
            .mnuStart.Enabled = False
            .tmrSleep.Enabled = True
            .mnuStartNow.Enabled = True
             DoEvents
            .mnuStop.Enabled = True
            .mnuProperties.Enabled = False
            .mnuEnd.Enabled = False
            .Caption = "AnnRobot - Sleeping"
        End With
        DoEvents
        objMainForm.Show
        DoEvents
        Exit Sub
    End If
    
    Set objMainForm = New frmMain
    DoEvents
    DoEvents
    DoEvents
    Load objMainForm
    DoEvents
    DoEvents
    objMainForm.StatusBar1.Panels(1).Text = "Robot started"
    objMainForm.StatusBar1.Panels(2).Text = ""
    objMainForm.StatusBar1.Panels(3).Text = "Clean:" & Trim$(Str$(gCleanCnt))
    objMainForm.StatusBar1.Panels(4).Text = "Manual:" & Trim$(Str$(gManualCnt))
    DoEvents
    If Not objAboutForm Is Nothing Then
        Unload objAboutForm
        Set objAboutForm = Nothing
    End If
    DoEvents
    objMainForm.Show
    DoEvents
    gJustStarted = True
    DoEvents
    objMainForm.StartApp
    DoEvents
    Exit Sub

Main_Error:
    Err.Source = "Main Procedure"
    ErrMsg = Err.Description & vbCrLf
    ErrMsg = ErrMsg & "Error No: " & Err.Number & " Error Src: " & Err.Source
    MsgBox ErrMsg, vbApplicationModal + vbCritical, App.EXEName
    End
End Sub
Public Function GetRegistrySettings() As Boolean
    
'
' Get required values from registry.
'
    Dim strAppName As String
    Dim strSectionName As String
    Dim strKeyName As String
    Dim intTmpCnt As Integer
    Dim tmpstrWorksetName As String
    Dim strClass As String
    Dim intCounter As Integer
    Dim vClassCombo As Variant
    Dim strManualCnt As String
    Dim strCleanCnt As String

    
    On Error GoTo GetSettings_Error
    
    GetRegistrySettings = False
    
' Get userid
    strAppName = App.EXEName
    
    strSectionName = "Logon"
    
    strKeyName = "UserID"
    mstrUserID = GetSetting(strAppName, strSectionName, strKeyName, vbNullString)
    
' Get Password

    strKeyName = "Password"
    mstrPassword = GetSetting(strAppName, strSectionName, strKeyName, vbNullString)
    EncryptDecrypt mstrPassword
    
' Get Domain.

    strKeyName = "Domain"
    mstrDomain = GetSetting(strAppName, strSectionName, strKeyName, vbNullString)
    
' Get Database SID

    strKeyName = "Database"
    mstrDatabase = GetSetting(strAppName, strSectionName, strKeyName, vbNullString)
    
    'strKeyName = "ODBC"
    'mstrODBC = GetSetting(strAppName, strSectionName, strKeyName, vbNullString)

' Get Sleep Time

    strKeyName = "SleepTime"
    mintSleepTime = CInt(GetSetting(strAppName, strSectionName, strKeyName, "90"))

    strSectionName = "Worksets"
    
 ' Get the count of worksets.
   
    strKeyName = "ListCount"
    mintWorksetCnt = CInt(Trim$(GetSetting(strAppName, strSectionName, strKeyName, 0)))
    
' Get all the worksets.
    
    intTmpCnt = -1
    For intCounter = 0 To mintWorksetCnt - 1
        strKeyName = "Workset" & Trim$(Str$(intCounter + 1))
        tmpstrWorksetName = Trim$(GetSetting(strAppName, strSectionName, strKeyName, vbNullString))
        If tmpstrWorksetName > " " Then
            intTmpCnt = intTmpCnt + 1
            ReDim Preserve mstrWorksets(intTmpCnt) As String
            mstrWorksets(intTmpCnt) = tmpstrWorksetName
        End If
    Next
    
    mintWorksetCnt = intTmpCnt + 1  'inttmpcnt was initialized with -1
    
    strSectionName = "Classes"
    
' Get the count of document classes

    strKeyName = "DocClassCount"
    mintDocCLassCount = CInt(Trim$(GetSetting(strAppName, strSectionName, strKeyName, 0)))
    
' Get all the document classes

    intTmpCnt = -1
    For intCounter = 0 To mintDocCLassCount - 1
        strKeyName = "DocClass" & Trim$(Str$(intCounter + 1))
        strClass = Trim$(GetSetting(strAppName, strSectionName, strKeyName, vbNullString))
        If strClass > " " Then
            intTmpCnt = intTmpCnt + 1
            ReDim Preserve mstrDocClass(intTmpCnt) As String
            mstrDocClass(intTmpCnt) = strClass
            DoEvents
        End If
    Next
    
    mintDocCLassCount = intTmpCnt + 1  'intTmpCnt was initialized with -1
    
' Get the count of document classes

    strKeyName = "FoldClassCount"
    mintFoldCLassCount = CInt(Trim$(GetSetting(strAppName, strSectionName, strKeyName, 0)))
    
' Get all the document classes

    intTmpCnt = -1
    For intCounter = 0 To mintFoldCLassCount - 1
        strKeyName = "FoldClass" & Trim$(Str$(intCounter + 1))
        strClass = Trim$(GetSetting(strAppName, strSectionName, strKeyName, vbNullString))
        If strClass > " " Then
            intTmpCnt = intTmpCnt + 1
            ReDim Preserve mstrFoldClass(intTmpCnt) As String
            mstrFoldClass(intTmpCnt) = strClass
            DoEvents
        End If
    Next
    
    mintFoldCLassCount = intTmpCnt + 1  'intTmpCnt was initialized with -1
   
    
'Section Name
    strSectionName = "Preferences"

'Key Name EndTime
    strKeyName = "EndTime"
    gstrEndtime = GetSetting(strAppName, strSectionName, strKeyName, "06:00:00 PM")
    DoEvents

'Key Name RunDays
    strKeyName = "RunDays"
    gstrRundays = GetSetting(strAppName, strSectionName, strKeyName, "234567")      ' Default to run on mon,tue,wed,
    DoEvents                                                                        ' thu,fri and sat


'Key Name CleanCount
    strKeyName = "CleanCount"
    strCleanCnt = GetSetting(strAppName, strSectionName, strKeyName, "0")
    gCleanCnt = CLng(strCleanCnt)
    
'Key Name ManualCount
    strKeyName = "ManualCount"
    strManualCnt = GetSetting(strAppName, strSectionName, strKeyName, "0")
    gManualCnt = CLng(strManualCnt)
    
    GetRegistrySettings = True
    Exit Function

GetSettings_Error:

    DoEvents
    Err.Source = "GetSettings"
    ErrMsg = Err.Description & vbCrLf
    ErrMsg = ErrMsg & "Error No: " & Err.Number & " Error Src: " & Err.Source
    MsgBox ErrMsg, vbApplicationModal + vbCritical, App.EXEName
    DoEvents

End Function

Public Sub EncryptDecrypt(strWord As String)
'
' Encrypt or Decrypt a string by using Xor with 256 multipiled by 0.8. I have used 0.8, can use any value < 1 - Ravi
'

    Dim CharNum As Integer
    Dim ICnt As Integer
    Dim strENword As String
    
    If strWord = vbNullString Then Exit Sub
    
    For ICnt = 1 To Len(strWord)
        CharNum = Asc(Mid$(strWord, ICnt, 1))
        strENword = strENword & Chr$(CharNum Xor Int(256 * 0.8))
    Next
    strWord = strENword     ' since default is by reference, the original value changes.
    
End Sub


Public Function CheckSettingValues() As Boolean
'
' Check whether all settings are there or not. If one is missing then show options form
' Return true is everything is ok or else false
'
    Dim objfrmOptions As frmOptions
    
    On Error GoTo CheckSetting_Error
        
    
    GetRegistrySettings
    CheckSettingValues = True
    
    If Not mstrUserID > " " Then CheckSettingValues = False
    
    If Not mstrPassword > " " Then CheckSettingValues = False
    
    If Not mstrDomain > " " Then CheckSettingValues = False
    
    If Not mstrDatabase > " " Then CheckSettingValues = False
    
    If mintWorksetCnt = 0 Then CheckSettingValues = False
    
    If mintDocCLassCount = 0 Then CheckSettingValues = False
    
    If Not CheckSettingValues Then
        Set objfrmOptions = New frmOptions
        Load objfrmOptions
        objfrmOptions.Show vbModal
        Set objfrmOptions = Nothing
        GetRegistrySettings
        CheckSettingValues = True           'Assuming that they entered all values
    End If
    
    DoEvents
    Exit Function
    
CheckSetting_Error:

    Err.Source = "Check Setting Values"
    ErrMsg = Err.Description & vbCrLf
    ErrMsg = ErrMsg & "Error No: " & Err.Number & " Error Src: " & Err.Source
    WriteToLogFile ErrMsg
    CheckSettingValues = False
    If Not objfrmOptions Is Nothing Then
        Unload objfrmOptions
        Set objfrmOptions = Nothing
    End If

End Function

Public Sub CreateLogFile()
'
' Create log file depending on the week of the day. If the file create date is same as today then append
' records else create a new one.
'
    Dim strFileName As String
    Dim objFile As File
    
    On Error GoTo CreateLog_Error
    
    If objFSO Is Nothing Then Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    strFileName = "AR" & Format$(Date$, "MMDDYY")
    
    strFileName = App.Path & "\Logs\" & strFileName & ".Log"
    DoEvents
    DoEvents
    If Not objFSO.FolderExists(App.Path & "\Logs") Then objFSO.CreateFolder App.Path & "\Logs"
    DoEvents
    DoEvents
    DoEvents
    If objFSO.FileExists(strFileName) Then
        Set objFile = objFSO.GetFile(strFileName)
        DoEvents
        DoEvents
        Sleep 1000
        DoEvents
        Set objCurFile = objFile.OpenAsTextStream(ForAppending, TristateFalse)
        DoEvents
        DoEvents
        Sleep 1000
        DoEvents
    Else
        Set objCurFile = objFSO.CreateTextFile(strFileName, True, False)
        DoEvents
        DoEvents
        Set objFile = objFSO.GetFile(strFileName)
        DoEvents
        Sleep 1000
    End If
    DoEvents
    
    If Not objFile Is Nothing Then gLogFileName = objFile.ShortName
    Set objFile = Nothing
    Exit Sub

CreateLog_Error:

    DoEvents
    MsgBox Err.Description
    Err.Clear
    Set objCurFile = Nothing
    Set objFile = Nothing
    DoEvents
    DoEvents
    DoEvents
    
End Sub


Public Sub InitializeVars()
'
' Initialize all global variables
'
    On Error Resume Next
    
    gcurWorkset = mstrWorksets(0)        ' Set the current workset as the first workset
    gblnRunning = False                  ' Indicate that the merge process hasnt started.
    gblnStopRequested = False            ' Boolean which indicates whether the stop button was pressed.
    
End Sub


Public Sub WriteToLogFile(ByVal StrMessage As String)
'
' Write all events to the current log file.
'
    On Error GoTo Write_Error
    
    DoEvents
    If objCurFile Is Nothing Then CreateLogFile
    DoEvents
    StrMessage = FormatDateTime$(Now, vbGeneralDate) & " - " & StrMessage
    DoEvents
    If Not objCurFile Is Nothing Then objCurFile.WriteLine StrMessage
    DoEvents
    Sleep 1000
    DoEvents
    Exit Sub
    
Write_Error:
    DoEvents
    Err.Clear
    Exit Sub
    
    
End Sub

Public Function ProcessError() As String
'
' Trap the error info and pass it back as  string
'
    Dim strTemp As String
    
    ProcessError = vbNullString
    strTemp = "Error No: " & Err.Number & "Error Src: " & Err.Source
    If Err.Number = 0 Then
        Exit Function
    ElseIf objCALMaster.LastError.Code > 0 Then
        ProcessError = objCALMaster.LastError.Message & vbCrLf
    Else
        ProcessError = Err.Description & vbCrLf
    End If
    ProcessError = ProcessError & strTemp
    Err.Clear
    DoEvents
    
End Function



Public Sub ShowFormDialog(strMsg As String)
'
'  Show the dialog form, to inform user.
'
    If Not objDialogForm Is Nothing Then
        Unload objDialogForm
        DoEvents
        Set objDialogForm = Nothing
        DoEvents
    End If
    Set objDialogForm = New frmDialog
    DoEvents
    Load objDialogForm
    DoEvents
    objDialogForm.lblMsg = strMsg
    DoEvents
    objDialogForm.Show
    DoEvents
    
End Sub

Public Function ConnectToDataBase() As Boolean
'
' Input: Global userid, password and database sid.
' Output: True or False depending on whether the connection is successful or not.
' Connect to the database using the ODBC DSN that was given in the settings.
'

    On Error GoTo Connect_Error
        
    Dim strConnectString As String
    Dim objADOError As ADODB.Error
    Dim strMsg As String
    
   ' strConnectString = "DSN=" & Trim$(mstrODBC)
    strConnectString = "Provider=OraOLEDB.Oracle.1;Data Source=" & Trim$(mstrDatabase) & ";DistribTx=0"
    Set objADOConnection = New ADODB.Connection
    DoEvents
    DoEvents
    objADOConnection.ConnectionString = strConnectString
    DoEvents
    objADOConnection.CursorLocation = adUseServer
    objADOConnection.ConnectionTimeout = 60   '60 seconds
    objADOConnection.Open strConnectString, mstrUserID, mstrPassword, adConnectUnspecified
    DoEvents
    DoEvents
    Sleep 1000
    ConnectToDataBase = True
    Exit Function

Connect_Error:
    
    
    For Each objADOError In objADOConnection.Errors
        strMsg = objADOError.Description & " Native Error: " & objADOError.NativeError
        WriteToLogFile strMsg
        DoEvents
        If Not objMainForm Is Nothing Then objMainForm.AddItemToList strMsg
    Next
    ConnectToDataBase = False

End Function


Public Sub LogOffFromDatabase()
'
' Disconnect from Database
'
    On Error Resume Next
    
    If Not objADOConnection Is Nothing Then
        objADOConnection.Close
        DoEvents
        Set objADOConnection = Nothing
        DoEvents
    End If
    
End Sub

Public Function GetProcessingInfoFromDB(strFileNumber As String, intFilingYear As Integer, dtReceiveDate As Date, _
                                        strAnnualType As String, strWorkitemID As String, strCompanyName As String, strErrCode As String, strErrMsg As String) As String

' Input:  Filenumber, Filing Year, Workitem ID, Receive Date
' Output: Company Name, Error Code and Error Message.
' Purpose: Call the Oracle Stored procedure by passing the Filenumber, Year of filing, Receive Date, Type of Annual
' ie Whether Clean or Manual and userid. The stored procuder will return the company name, error code and error message
' If its a clean annual, the return code will indicate, that the oracle database has been updated.
' If its a manaul annual, then just take the company name and return.

    On Error GoTo DB_Error
    
    Dim objADOCommand As ADODB.Command
    Dim objADOParameter As ADODB.Parameter
    Dim objADOError As ADODB.Error
    Dim strLastChar As String * 1
    
    RemoveZeroes strFileNumber
    gblnGPLLP = False
    Set objADOCommand = New ADODB.Command
    DoEvents
    Sleep 1000
    'strCompanyName = Space$(500)
    'strErrMsg = Space$(200)
    DoEvents
    With objADOCommand
        .ActiveConnection = objADOConnection
        .CommandText = "ESWM.BRIMS_INTEGRATION.BRIMS_ANNUAL_ROBOT"
        .CommandType = adCmdStoredProc
         DoEvents
        
         Set objADOParameter = .CreateParameter("p_file_number", adVarChar, adParamInput, 15, strFileNumber)
        .Parameters.Append objADOParameter
         DoEvents
         
         Set objADOParameter = .CreateParameter("p_year_of_filing", adInteger, adParamInput, , intFilingYear)
        .Parameters.Append objADOParameter
         DoEvents
   
         Set objADOParameter = .CreateParameter("p_annual_type", adVarChar, adParamInput, 1, strAnnualType)
        .Parameters.Append objADOParameter
         DoEvents
          
         Set objADOParameter = .CreateParameter("p_workitem_id", adVarChar, adParamInput, 20, strWorkitemID)
        .Parameters.Append objADOParameter
         DoEvents
          
         Set objADOParameter = .CreateParameter("p_received_dt", adDate, adParamInput, , dtReceiveDate)
        .Parameters.Append objADOParameter
         DoEvents
          
         Set objADOParameter = .CreateParameter("p_company_name", adVarChar, adParamOutput, 500, strCompanyName)
        .Parameters.Append objADOParameter
         DoEvents
          
         Set objADOParameter = .CreateParameter("p_err_code", adVarChar, adParamOutput, 2, strErrCode)
        .Parameters.Append objADOParameter
         DoEvents
          
         Set objADOParameter = .CreateParameter("p_err_msg", adVarChar, adParamOutput, 200, strErrMsg)
        .Parameters.Append objADOParameter
         DoEvents
          
        .Execute , , adCmdStoredProc
         DoEvents
         Sleep 1000
         DoEvents
         DoEvents

         strErrCode = vbNullString
         strErrMsg = vbNullString
         strCompanyName = vbNullString
         
         If .Parameters.Item("p_err_code").Value > " " Then strErrCode = Trim$(.Parameters.Item("p_err_code").Value)
         If .Parameters.Item("p_err_msg").Value > " " Then strErrMsg = Trim$(.Parameters.Item("p_err_msg").Value)
         If .Parameters.Item("p_company_name").Value > " " Then
            strCompanyName = Trim$(.Parameters.Item("p_company_name").Value)        'Remove any < and > chars
            strCompanyName = Replace$(strCompanyName, "<", vbNullString, , , vbTextCompare)
            strCompanyName = Replace$(strCompanyName, ">", vbNullString, , , vbTextCompare)
            strLastChar = Right$(strCompanyName, 1)                                 ' Remove "*" at the end of the company name
            Do Until strLastChar <> "*"
                strCompanyName = Left$(strCompanyName, Len(strCompanyName) - 1)
                If Len(strCompanyName) < 2 Then Exit Do
                strLastChar = Right$(strCompanyName, 1)
            Loop
         End If
         DoEvents
         If strErrCode > "19" Then  'GPLLP Send to manual annuals
            gblnGPLLP = True
            gblnFolderGPLLP = True
            If strErrCode = "20" Then
                strErrCode = "00"
            ElseIf strErrCode = "21" Then
                strErrCode = "01"
            ElseIf strErrCode = "22" Then
                strErrCode = "02"
            ElseIf strErrCode = "31" Then
                strErrCode = "11"
            End If
         End If
         If (Not strErrCode > " ") And strCompanyName > " " Then strErrCode = "01"      ' Update document with company name but send it to manual review
         If ((strErrCode = "00") Or (strErrCode = "01")) And (Not strCompanyName > " ") Then strErrCode = "11"   'If the company name is blank, then send it manual processing
         
         DoEvents
         
       '  MsgBox "strErrCode: " & strErrCode & vbCrLf & "strCompanyName: " & strCompanyName & vbCrLf & "strErrMsg: " & strErrMsg
       '  If Trim$(strErrCode) = "00" Then
       '      For Each objADOParameter In .Parameters
       '          MsgBox "Name: " & objADOParameter.Name & "   Value: " & objADOParameter.Value & "   Type: " & objADOParameter.Type
       '      Next
       '  Else
       '      MsgBox strErrMsg & " " & strErrCode
       '  End If
       
    End With
    DoEvents
    Exit Function

DB_Error:

    For Each objADOError In objADOConnection.Errors
        ErrMsg = objADOError.Description
        MsgBox ErrMsg
        WriteToLogFile ErrMsg & "-" & objADOError.NativeError
    Next
    strErrCode = "10"
    GetProcessingInfoFromDB = "10"
    Err.Clear
   
End Function

Public Sub DestroyGlobalObjects()
'
' Set all global objects to nothing
'
    On Error Resume Next
    
    DoEvents
    If Not objCurFile Is Nothing Then objCurFile.Close
    DoEvents
    Set objCurFile = Nothing
    DoEvents
    Set objCALMaster = Nothing
    DoEvents
    Set objCALClient = Nothing
    DoEvents
    Set objCALClientList = Nothing
    DoEvents
    Set objADOConnection = Nothing
    DoEvents
    Err.Clear
    
End Sub

Public Sub BeginPlaySound(ByVal ResourceId As Integer)
    On Error Resume Next
    SoundBuffer = LoadResData(ResourceId, "SOUND")
    sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub

Public Sub EndPlaySound()
    On Error Resume Next
    sndPlaySound ByVal vbNullString, 0&
End Sub


Private Sub RemoveZeroes(strT8 As String)
'
'Remove leading zeroes in the string.
'
    On Error GoTo Remove_Error
    
    Dim strTemp As String
    Dim intA As Integer
    Dim ICnt As Integer
    Dim strFileSuffix As String
    Dim strFileNo As String
    
    strTemp = strT8
    For ICnt = 1 To Len(strT8)
        If IsNumeric(Left$(strTemp, 1)) Then
            strTemp = Right$(strTemp, Len(strTemp) - 1)
        Else
            intA = Len(strT8) - ICnt + 1
            Exit For
        End If
    Next ICnt
    
    DoEvents
    strFileSuffix = Right$(strT8, intA)
    strFileNo = Format$(Left$(strT8, Len(strT8) - intA), "#")
    strT8 = strFileNo & strFileSuffix
    strT8 = UCase$(strT8)
    DoEvents
    Exit Sub

Remove_Error:
    DoEvents
    Err.Clear
    
End Sub

