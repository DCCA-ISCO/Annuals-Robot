Attribute VB_Name = "CALXRoutines"
Option Explicit
Private ErrMsg As String

Public Function LogonToWorkflow() As Boolean
'
' Logon to workflow with userid, password and domain info.
' Clear any new workitems that are on the client list.
' Remove any workitems that are WIP as "No".
' Create Log file and log the message as "Logged onto Workflow"

    On Error GoTo Logon_Error
    
    Dim strMsg As String
    Dim sngTime As Single
    
    InitializeVars                             ' Initialize all global variables
    Set objCALClient = objCALMaster.CreateClient(mstrUserID, mstrPassword, mstrDomain)  'Create CAL Client
    sngTime = Timer
    Do Until Timer > sngTime + 1
        DoEvents
    Loop
    Set objCALClientList = objCALClient.ClientList
    objCALClientList.Clear calClearAbortNew
    DoEvents
    strMsg = "Logged onto domain " & mstrDomain & "."
    DoEvents
    DoEvents
    WriteToLogFile strMsg
    DoEvents
    If Not objMainForm Is Nothing Then objMainForm.AddItemToList strMsg
    DoEvents
    DoEvents
    If ConnectToDataBase Then
        strMsg = "Connected to database " & mstrDatabase & "."
        LogonToWorkflow = True
    Else
        strMsg = "Failed to connect to database " & mstrDatabase & "."
        LogonToWorkflow = False
    End If
    WriteToLogFile strMsg
    DoEvents
    If Not objMainForm Is Nothing Then
        objMainForm.AddItemToList strMsg
        objMainForm.Refresh
        DoEvents
        Sleep 1000
    End If
    DoEvents
    Exit Function
    
Logon_Error:

    Err.Source = "Logon"
    If Err.Number = 1000 Then
        If objCALMaster.LastError.Code > 0 Then
            ErrMsg = Err.Source & ": " & objCALMaster.LastError.Message
        Else
            ErrMsg = Err.Source & ": " & Err.Description
        End If
    Else
        ErrMsg = Err.Source & ": " & Err.Description
    End If
    WriteToLogFile ErrMsg
    If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
    LogonToWorkflow = False
    
End Function

Public Function CheckWorksetsNClasses() As Boolean
'
' Check whether the classes that are entered on the options form are existing in the workflow or not.
' Check whethe the Worksets are existing or not
'
    Dim objCALWorkitemClasses As CALWorkitemClasses
    Dim objCALWorkitemClass As CALWorkitemClass
    Dim objCALQueues As CALQueues
    Dim objCALQueue As CALQueue
    
    Dim intCounter As Integer
    Dim blnWorksetExists As Boolean
    
    On Error GoTo Check_Error
    
    CheckWorksetsNClasses = True
'
' Check the classes.
'
    
    Set objCALWorkitemClasses = objCALClient.WorkitemClasses
    For intCounter = 0 To mintDocCLassCount - 1
    
        Set objCALWorkitemClass = objCALWorkitemClasses.Find(mstrDocClass(intCounter)) ' , intCounter))
        If objCALWorkitemClass Is Nothing Then
            ErrMsg = mstrDocClass(intCounter) & " document class is not found in the domain " & mstrDomain & vbCrLf
            WriteToLogFile ErrMsg
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
            DoEvents
            CheckWorksetsNClasses = False
            Exit Function
        End If
        
    Next
    DoEvents
    
    For intCounter = 0 To mintFoldCLassCount - 1
    
        Set objCALWorkitemClass = objCALWorkitemClasses.Find(mstrFoldClass(intCounter)) ' , intCounter))
        If objCALWorkitemClass Is Nothing Then
            ErrMsg = mstrFoldClass(intCounter) & " folder class is not found in the domain " & mstrDomain & vbCrLf
            WriteToLogFile ErrMsg
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
            DoEvents
            CheckWorksetsNClasses = False
            Exit Function
        End If
        
    Next

    Set objCALWorkitemClasses = Nothing
    Set objCALWorkitemClass = Nothing
    
' Check the worksets next.

    Set objCALQueues = objCALClient.Queues(calQueueListWorkset + calQueueListUserOnly)
    DoEvents
    For intCounter = 0 To mintWorksetCnt - 1
        blnWorksetExists = False
        For Each objCALQueue In objCALQueues
            If objCALQueue.Name = mstrWorksets(intCounter) Then
                blnWorksetExists = True
                Exit For
            End If
        Next
        If Not blnWorksetExists Then
            ErrMsg = "Workset " & mstrWorksets(intCounter) & " does not exist in domain-" & mstrDomain & vbCrLf
            ErrMsg = "Please key in the correct workset in the options form."
            WriteToLogFile ErrMsg
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
            DoEvents
            CheckWorksetsNClasses = False
            Exit Function
        End If
    Next
    
   Exit Function

Check_Error:

    Err.Source = "Check Worksets N Classes"
    ErrMsg = Err.Description & vbCrLf
    ErrMsg = ErrMsg & "Error No: " & Err.Number & " Error Src: " & Err.Source
    WriteToLogFile ErrMsg
    CheckWorksetsNClasses = False
    
End Function


Public Sub ProcessAnnuals()
'
' Process Annuals from the current workset. Get the next workitem that is available in the workset.
' Check whether that workitem is OK to process, and if it ok, then process that workitem. If not OK, then
' send it to Annuals workstep for manual processing.
'
    On Error GoTo ProcessResub_Error
    
    Dim blnProcessDone As Boolean
    Dim intNextQueue As Integer
    Dim eErrorStatus As CALProcessStatus
    Dim dtEndTime As Date
    Dim strMsg As String
    Dim strWeekDay As String
    Dim blnStop As Boolean
    
    Dim objCALQueue As CALQueue
    Dim objCALClientListitem As CALClientListItem
    
    If IsDate(gstrEndtime) Then
        dtEndTime = CDate(gstrEndtime)
    Else
        dtEndTime = #6:00:00 PM#
    End If
    DoEvents
    
    Set objCALClientListitem = Nothing
    Set objCALQueue = Nothing
    DoEvents
    
    Set objCALQueue = New CALQueue
    DoEvents
    With objCALQueue
        .Client = objCALClient
        .Type = calQueueTypeWorkset
        .Name = gcurWorkset
    End With
    
    DoEvents
    gJustStarted = False    'This will be true, when the app starts the first time
                            'and executes the Main procedure
    Do Until blnProcessDone
    
        gblnRunning = True
        DoEvents
        strWeekDay = Trim$(Str$(Weekday(Date)))
        If DateDiff("n", dtEndTime, Time) > 0 Then
            blnStop = True      'Stop the program once it reaches the end time
            strMsg = "Robot was specified not to run after " & gstrEndtime & ". Shutting down the Robot."
        ElseIf InStr(1, gstrRundays, strWeekDay, vbTextCompare) = 0 Then
            blnStop = True      'Stop the program if it not scheduled to run today.
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
            strMsg = "Robot was specified not to run on " & strMsg & ". Shutting down the Robot."
        End If
        
        If blnStop Then
            gblnStopRequested = True
            WriteToLogFile strMsg
            If Not objMainForm Is Nothing Then
                objMainForm.AddItemToList strMsg
                objMainForm.UnloadRequested = True
                objMainForm.mnuStop_Click
            End If
            LogoffWorkflow      'LogOff From workflow
            gblnRunning = False
            DoEvents
            Set objCALQueue = Nothing
            DoEvents
            Exit Sub
        End If
        DoEvents

        If gblnStopRequested Then
            LogoffWorkflow              'Log off from workflow
            gblnRunning = False
            DoEvents
            Set objCALQueue = Nothing
            DoEvents
            Exit Sub
        End If
        DoEvents
        
'Get the next workitem available and store the pointer in objCALClientListItem

        DoEvents
        eErrorStatus = GetNextWorkitem(objCALClientListitem, objCALQueue)
        
        If eErrorStatus = icSuccess Then
            DoEvents
            If objCALClientListitem.Info.Type = calObjTypeDocument Then
                eErrorStatus = ProcessWorkitem(objCALClientListitem)
            Else
                eErrorStatus = ProcessFolderWorkitem(objCALClientListitem)
            End If
            If eErrorStatus = icCriticalError Then
                blnProcessDone = True
                LogoffWorkflow
            End If
        ElseIf eErrorStatus = icQueueEmpty Then     'Go for the next queue.
            intNextQueue = intNextQueue + 1
            If intNextQueue < mintWorksetCnt Then   'If more queues exist
                gcurWorkset = mstrWorksets(intNextQueue)
                Set objCALQueue = Nothing
                Set objCALClientListitem = Nothing
                DoEvents
                Set objCALQueue = New CALQueue
                DoEvents
                 With objCALQueue
                    .Client = objCALClient
                    .Type = calQueueTypeWorkset
                    .Name = gcurWorkset
                End With
                DoEvents
            Else
                blnProcessDone = True
                LogoffWorkflow
            End If
        Else
            blnProcessDone = True
            LogoffWorkflow
        End If
    Loop
    
    DoEvents
    gblnRunning = False
    Set objCALClientListitem = Nothing
    Set objCALQueue = Nothing
    Exit Sub
    
ProcessResub_Error:

    DoEvents
    Err.Source = "ProcessAnnuals"
    WriteToLogFile (ProcessError)           ' ProcessError is a function which returns a string
    blnProcessDone = False                    ' containing the error info.
    Set objCALClientListitem = Nothing
    Set objCALQueue = Nothing
    LogoffWorkflow
    
End Sub

Private Function GetNextWorkitem(objCALClientListitem As CALClientListItem, objCALQueue As CALQueue) As CALProcessStatus
'
' Retrieve the next workitem in the queue. Check whether the retrieved workitem is a document or not.
' If it is a folder put it in error and get the next workitem.


    Dim blnGotWorkitem As Boolean
    Dim blnWorkitemOK As Integer
    Dim eErrorStatus As CALProcessStatus
    Dim strWorkitemID As String
    Dim objCALSendQ As CALQueue
  
    On Error Resume Next
    
    Err.Source = "RetrieveNextItem"
    blnGotWorkitem = False
    GetNextWorkitem = icSuccess
    DoEvents
    Sleep 2000
    DoEvents
    If objCALClient Is Nothing Then
        If Not LogonToWorkflow Then
            GetNextWorkitem = icCriticalError
            Exit Function
        ElseIf objCALClient Is Nothing Then
            GetNextWorkitem = icCriticalError
            Exit Function
        End If
    End If
    
    ' Loop till you get a valid workitem from the queue.
    
    Do Until blnGotWorkitem

        DoEvents
        Set objCALClientListitem = Nothing
        Set objCALClientListitem = objCALQueue.GetNext(calGetNextCheckAccess)
        DoEvents
        eErrorStatus = CheckCALError
        
        Select Case eErrorStatus
        
            Case Is = icQueueEmpty
                GetNextWorkitem = icQueueEmpty
                Set objCALClientListitem = Nothing
                ErrMsg = gcurWorkset & " queue is empty."
                WriteToLogFile ErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                Exit Function
                
            Case Is = icWorkitemInList
                
                If Not objCALClientListitem Is Nothing Then
                    Set objCALSendQ = New CALQueue
                    With objCALSendQ
                        .Client = objCALClient
                        .Name = objCALClientListitem.Info.SourceWorkstep        'Send it back to the workstep from where it was retrieved
                        .Type = calQueueTypeWorkstep
                    End With
                    DoEvents
                    
                    objCALClientListitem.Send objCALQueue, calSendDiscard
                    If CheckCALError <> icSuccess Then
                        Set objCALSendQ = Nothing
                        Set objCALQueue = Nothing
                        Set objCALClientListitem = Nothing
                        GetNextWorkitem = icCriticalError
                        ErrMsg = "Critical Error occured. Processing Suspended (GetNextWorkitem, icWorkitemInList)"
                        WriteToLogFile ErrMsg
                        If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                        Exit Function
                    End If
                End If
                
                Set objCALSendQ = Nothing
                Set objCALClientListitem = Nothing
                DoEvents
                
            Case Is = icSuccess
                DoEvents
                
            Case Else
                Set objCALSendQ = Nothing
                Set objCALQueue = Nothing
                Set objCALClientListitem = Nothing
                GetNextWorkitem = icCriticalError
                ErrMsg = "Critical Error occured. Processing Suspended (GetNextWorkitem, Else)"
                WriteToLogFile ErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                Exit Function
                
        End Select

        If Not objCALClientListitem Is Nothing Then
            strWorkitemID = objCALClientListitem.Info.Name
            eErrorStatus = VerifyThisWorkitem(objCALClientListitem)
            If eErrorStatus = icVerifyOK Then
                blnGotWorkitem = True
            ElseIf eErrorStatus = icVerifyNotOK Then
                ErrMsg = strWorkitemID & " verify not OK. (GetNextWorkitem)."
                WriteToLogFile ErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                blnGotWorkitem = False
            ElseIf eErrorStatus = icCriticalError Then
                Set objCALSendQ = Nothing
                Set objCALQueue = Nothing
                Set objCALClientListitem = Nothing
                GetNextWorkitem = icCriticalError
                ErrMsg = "Critical Error occured. Processing Suspended (GetNextWorkitem, AfterVerification)"
                WriteToLogFile ErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                Exit Function
            End If
        End If
        
    Loop
    
    Err.Clear
    DoEvents
    
End Function

Private Function VerifyThisWorkitem(objCALClientListitem As CALClientListItem) As CALProcessStatus

' Input: CALClientListItem
' Output: Enumrated value of calProcessStatus
' Verify whether the workitem in hand is a folder or not. If it is a folder, then verify whether the folder
' contains child items with the class combination defined in the options form. In some cases there wont be a resub,
' but still it is a good workitem. If the class combination does not match with the one that was specified, then put
' this workitem in error.
'
    On Error Resume Next
    
    Dim strComment As String
    Dim blnWorkitemFound As Boolean
    Dim eErrorStatus As CALProcessStatus
    Dim intCntr As Integer
    Dim strWorkitemID As String
    Dim objCALDocument As CALDocument
    Dim objCALFolder As CALFolder

    
    strWorkitemID = objCALClientListitem.Info.Name
    
    If objCALClientListitem.Info.Type = calObjTypeDocument Then 'If the workitem is a document
    
' Open the Document in ReadOnly mode.

        Set objCALDocument = objCALClientListitem.Open(calOpenReadOnly)
        eErrorStatus = CheckCALError
        
        Select Case eErrorStatus
        
            Case Is = icSuccess
                DoEvents
                
            Case Is = icWorkitemNotInList
                ErrMsg = strWorkitemID & " open failed. (VerifyThisWorkItem - icWorkitemNotInList)."
                WriteToLogFile ErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                Set objCALDocument = Nothing
                Set objCALClientListitem = Nothing
                VerifyThisWorkitem = icVerifyNotOK
                
            Case Is = icWorkitemOpen        'Already opened
                Set objCALDocument = objCALClientListitem.OpenedItem
                DoEvents
            
            Case Else
                ErrMsg = "Critical Error occured. Processing Suspended (VerifyThisWorkitem, AfterVerification)"
                WriteToLogFile ErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                VerifyThisWorkitem = icCriticalError
                Set objCALDocument = Nothing
                Set objCALClientListitem = Nothing
                Exit Function
                
        End Select
        
        If objCALDocument Is Nothing Then
            Set objCALClientListitem = Nothing
            VerifyThisWorkitem = icVerifyNotOK
            Exit Function
        End If
'
' Check whether the document belongs to the one of the classes listed. If it is then exit out, otherwise put
' the workitem in error.

        For intCntr = 0 To mintDocCLassCount - 1
            If objCALDocument.Class = mstrDocClass(intCntr) Then
                blnWorkitemFound = True
                objCALDocument.Close calCloseAbortChanges + calCloseRetainLock
                DoEvents
                Sleep 1000
                VerifyThisWorkitem = icVerifyOK
                DoEvents
                Exit For
            End If
        Next
        
        If Not blnWorkitemFound Then
            objCALDocument.Close calCloseAbortChanges + calCloseRetainLock
            DoEvents
            Sleep 1000
            DoEvents
            strComment = "Workitem does not belong to any of the specified classes"
            If PlaceWorkitemInError(objCALClientListitem, strComment) <> icCriticalError Then
                VerifyThisWorkitem = icVerifyNotOK
            Else
                VerifyThisWorkitem = icCriticalError
            End If
        End If
    ElseIf objCALClientListitem.Info.Type = calObjTypeFolder Then   ' If the workitem is a folder

' Open the Folder in ReadOnly mode.

        Set objCALFolder = objCALClientListitem.Open(calOpenReadOnly)
        eErrorStatus = CheckCALError
        
        Select Case eErrorStatus
        
            Case Is = icSuccess
                DoEvents
                
            Case Is = icWorkitemNotInList   'Shouldn't happen but may be
                ErrMsg = strWorkitemID & " open failed. (VerifyThisWorkItem - icWorkitemNotInList)."
                WriteToLogFile ErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                Set objCALFolder = Nothing
                Set objCALClientListitem = Nothing
                VerifyThisWorkitem = icVerifyNotOK
                
            Case Is = icWorkitemOpen        'Already opened
                Set objCALFolder = objCALClientListitem.OpenedItem
                DoEvents
            
            Case Else                       'Something else
                ErrMsg = "Critical Error occured. Processing Suspended (VerifyThisWorkitem, AfterVerification)"
                WriteToLogFile ErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                VerifyThisWorkitem = icCriticalError
                Set objCALFolder = Nothing
                Set objCALClientListitem = Nothing
                Exit Function
                
        End Select
        
        If objCALFolder Is Nothing Then
            Set objCALClientListitem = Nothing
            VerifyThisWorkitem = icVerifyNotOK
            Exit Function
        End If
'
' Check whether the folder belongs to the one of the classes listed. If it exists then exit out, otherwise put
' the workitem in error.

        For intCntr = 0 To mintFoldCLassCount - 1
            If objCALFolder.Class = mstrFoldClass(intCntr) Then
                blnWorkitemFound = True
                objCALFolder.Close calCloseAbortChanges + calCloseRetainLock
                DoEvents
                Sleep 1000
                VerifyThisWorkitem = icVerifyOK
                DoEvents
                Exit For
            End If
        Next
        
        If Not blnWorkitemFound Then
            objCALFolder.Close calCloseAbortChanges + calCloseRetainLock
            DoEvents
            Sleep 1000
            DoEvents
            strComment = "Workitem does not belong to any of the specified classes"
            If PlaceWorkitemInError(objCALClientListitem, strComment) <> icCriticalError Then
                VerifyThisWorkitem = icVerifyNotOK
            Else
                VerifyThisWorkitem = icCriticalError
            End If
        End If
        
    Else
        strComment = "Cannot process. Workitem is not a document or a folder"
        If PlaceWorkitemInError(objCALClientListitem, strComment) <> icCriticalError Then
            VerifyThisWorkitem = icVerifyNotOK
        Else
            VerifyThisWorkitem = icCriticalError
        End If
    End If
    
    Set objCALDocument = Nothing
    Set objCALFolder = Nothing
    Err.Clear
    DoEvents
    
End Function

Private Function ProcessWorkitem(objCALClientListitem As CALClientListItem, Optional blnChildWorkitem As Boolean = False) As CALProcessStatus
'
' Input: CALClientListItem
' Output: Process status of the function
' This function will take the calclientlistitem and then calls the oracle stored procedure.
' Oracle stored procedure in turn will return Company Name, Transaction Code, Transaction Description, Return Code
' & Return Message.
' The company name, transaction code and transaction description are transfered to the objcalclientlistitem.
' If the return code is zero then the annual is routed to filing workstep, else it is routed to annuals workstep.
'
    
    On Error Resume Next
    
    Dim objCALDocument As CALDocument
    Dim objCALFields As CALFormFields
    Dim objCALField As CALFormField
    Dim objCALSendQueue As CALQueue
    
    Dim eErrorStatus As CALProcessStatus
    Dim strFileNumber As String
    Dim strReceiveDT As String
    Dim strFilingYear As String
    Dim strBRAMessage As String
    Dim strCompanyName As String
    Dim strWorkitemID As String
    Dim strAnnualType As String * 1
    Dim strBatchType As String
    Dim strNoChange As String
    Dim strSignFound As String
    Dim strErrCode As String
    Dim strErrMsg As String
    Dim strTemp As String
    
    strWorkitemID = objCALClientListitem.Info.Name
    If Not objMainForm Is Nothing Then
        With objMainForm
            .StatusBar1.Panels(1).Text = "Processing Workitem"
            .StatusBar1.Panels(2).Text = strWorkitemID
            .StatusBar1.Panels(3).Text = "Clean:" & Trim$(Str$(gCleanCnt))
            .StatusBar1.Panels(4).Text = "Manual:" & Trim$(Str$(gManualCnt))
            .Refresh
             DoEvents
        End With
    End If
    
    If (objCALClientListitem.Info.Status And calObjStatusReserved) = calObjStatusReserved Then
        Set objCALDocument = objCALClientListitem.OpenedItem
    ElseIf (objCALClientListitem.Info.Status And calObjStatusOpen) = calObjStatusOpen Then
        Set objCALDocument = objCALClientListitem.OpenedItem
        DoEvents
        objCALDocument.Close calCloseAbortChanges + calCloseRetainLock
        DoEvents
        Sleep 1000
        DoEvents
        Set objCALDocument = objCALClientListitem.Open(calOpenReadWrite)
    Else
        Set objCALDocument = objCALClientListitem.Open(calOpenReadWrite)
    End If
    
    eErrorStatus = CheckCALError
    If eErrorStatus = icSuccess Then
        DoEvents
    Else
        Set objCALDocument = Nothing
        Set objCALClientListitem = Nothing
        DoEvents
        ProcessWorkitem = icSuccess
        Exit Function
    End If

    If objCALDocument Is Nothing Then
        Set objCALClientListitem = Nothing
        ProcessWorkitem = icSuccess
        Exit Function
    End If
    
    
    DoEvents
    If (objCALDocument.Info.Status And calObjStatusReserved) = calObjStatusReserved Then
        DoEvents
    Else
        objCALDocument.Close calCloseAbortChanges
        Set objCALDocument = Nothing
        Set objCALClientListitem = Nothing
        DoEvents
        Sleep 1000
        ProcessWorkitem = icSuccess
        DoEvents
        Exit Function
    End If
    
       
    ' Get the required fields from workitem.
       
    strBRAMessage = vbNullString
    
    Set objCALFields = objCALDocument.FormFields(calFieldsNoViews)
    DoEvents
    If Not objCALFields Is Nothing Then
        
        Set objCALField = Nothing                       'Get the captured barcode
        Set objCALField = objCALFields.Find("WFCAPTURED_BARCODE")
        If Not objCALField Is Nothing Then
            If objCALField.Value = "REA" Then
                Set objCALField = Nothing
                Set objCALFields = Nothing
                'Set objCALDocument = Nothing
                ProcessWorkitem = ProcessResubWorkitem(objCALClientListitem, objCALDocument)
                Exit Function
            End If
        End If
        DoEvents
        
        Set objCALField = Nothing                       'Get file number
        Set objCALField = objCALFields.Find("WFT8")
        If Not objCALField Is Nothing Then strFileNumber = objCALField.Value
        DoEvents
        
        Set objCALField = Nothing                       ' Get Receive Date
        Set objCALField = objCALFields.Find("WFRECEIVE_DATE")
        If Not objCALField Is Nothing Then strReceiveDT = objCALField.Value
        DoEvents
        If Not strReceiveDT > " " Then
            Set objCALField = Nothing                   ' if not then get Scan Date
            Set objCALField = objCALFields.Find("WFSCAN_DATE")
            If Not objCALField Is Nothing Then strReceiveDT = objCALField.Value
            DoEvents
        End If
                
        Set objCALField = Nothing
        Set objCALField = objCALFields.Find("WFYEAR")   ' Get Year of filing
        If Not objCALField Is Nothing Then strFilingYear = objCALField.Value
        DoEvents
        
        Set objCALField = Nothing
        Set objCALField = objCALFields.Find("WFBATCH_TYPE")   ' Get Year of filing
        If Not objCALField Is Nothing Then strBatchType = objCALField.Value
        DoEvents

        Set objCALField = Nothing
        Set objCALField = objCALFields.Find("WFNO_CHANGE")   ' Get Year of filing
        If Not objCALField Is Nothing Then strNoChange = objCALField.Value
        DoEvents

        Set objCALField = Nothing
        Set objCALField = objCALFields.Find("WFSIGNATURE_FOUND")   ' Get Year of filing
        If Not objCALField Is Nothing Then strSignFound = objCALField.Value
        DoEvents

        strBRAMessage = vbNullString
        If Not strFileNumber > " " Then strBRAMessage = "File number is blank. Routing to Annuals queue."
        
        If Not strReceiveDT > " " Then
            strBRAMessage = "Receive Date is blank. Routing to Annuals queue."
        ElseIf Not IsDate(strReceiveDT) Then
            strBRAMessage = "Receive Date is not a valid date. Routing to Annuals queue."
        ElseIf CDate(strReceiveDT) > Date Then
            strBRAMessage = "Receive Date greater than current date. Routing to Annuals queue."
        End If
        
        If Not strFilingYear > " " Then
            strBRAMessage = "Filing Year is blank. Routing to Annuals queue."
        ElseIf CInt(strFilingYear) > Year(Date) Then
            strBRAMessage = "Filing Year is greater than current year. Routing to Annuals queue."
        End If
               
        If strBRAMessage > " " Then
        
            WriteToLogFile strBRAMessage
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList strBRAMessage
            DoEvents
            strBRAMessage = "*Note by AnnRobot on " & FormatDateTime$(Now, vbGeneralDate) & vbCrLf & strBRAMessage
            Set objCALField = Nothing
            Set objCALField = objCALFields.Find("WFBRA_NOTES")   ' Get BRA Notes
            If Not objCALField Is Nothing Then
                If objCALField.Value > " " Then
                    objCALField.Value = strBRAMessage & vbCrLf & vbCrLf & objCALField.Value
                Else
                    objCALField.Value = strBRAMessage
                End If
            End If
            DoEvents
            If Not blnChildWorkitem Then
                Set objCALField = Nothing
                Set objCALField = objCALFields.Find("WFT8_1")   ' Get the unused T8 field and insert "ANN", to indicate for the post processing rules, to route the workitem to Annualsworkstep
                If Not objCALField Is Nothing Then objCALField.Value = "ANN"
                DoEvents
            End If
            Set objCALField = Nothing                           ' Enter user-id
            Set objCALField = objCALFields.Find("WFDOCI_USER")
            If Not objCALField Is Nothing Then objCALField.Value = UCase$(objCALClient.UserName)
            DoEvents
            DoEvents
            Set objCALField = Nothing                           ' Enter Current Date
            Set objCALField = objCALFields.Find("WFDOCI_DATE")
            If Not objCALField Is Nothing Then objCALField.Value = Format$(Date$, "YYYY-MM-DD")
            DoEvents
            strTemp = FormatDateTime$(Now, vbLongTime)
            strTemp = Replace(strTemp, " ", "", , , vbTextCompare)
            Set objCALField = Nothing                           ' Enter Current Time
            Set objCALField = objCALFields.Find("WFDOCI_TIME")
            If Not objCALField Is Nothing Then objCALField.Value = strTemp
            DoEvents
            objCALDocument.Save
            DoEvents
            objCALDocument.Close calCloseSaveChanges
            DoEvents
            Sleep 1000
            DoEvents
            Set objCALDocument = Nothing
            DoEvents
            If Not blnChildWorkitem Then
                objCALClientListitem.SendToDefault calSendDiscard
                Set objCALClientListitem = Nothing
            End If
            DoEvents
            eErrorStatus = CheckCALError
            If eErrorStatus <> icCriticalError Then
                ProcessWorkitem = icSuccess
            Else
                ProcessWorkitem = icCriticalError
            End If
            Exit Function
        End If
    Else            'If objCALFormFields is nothing, then do this
        If (objCALDocument.Info.Status And calObjStatusOpen) = calObjStatusOpen Then
            objCALDocument.Close calCloseAbortChanges
            DoEvents
            Sleep 1000
        End If
        Set objCALDocument = Nothing
        DoEvents
        Set objCALSendQueue = New CALQueue
        DoEvents
        With objCALSendQueue
            .Client = objCALClient
            .Name = objCALClientListitem.Info.SourceWorkstep
            .Type = calQueueTypeWorkstep
        End With
        DoEvents
        objCALClientListitem.Send objCALSendQueue, calSendDiscard
        DoEvents
        eErrorStatus = CheckCALError
        Set objCALSendQueue = Nothing
        Set objCALClientListitem = Nothing
        If eErrorStatus <> icCriticalError Then
            ProcessWorkitem = icSuccess
        Else
            ProcessWorkitem = icCriticalError
        End If
        Exit Function
    End If
    
    Set objCALField = Nothing
    Set objCALFields = Nothing
    
    DoEvents
    ' A workitem is considered as a clean annual if its batch type is "Annuals-c", No changes are made to the annual
    ' and Signature found
    If Trim$(strBatchType) = "ANNUALS-C" And Trim$(strNoChange) = "1" And Trim$(strSignFound) = "1" Then
        strAnnualType = "C"
    Else
        strAnnualType = "M"
    End If
    
    'Remove any trailing zeroes in the file number.
    
    
    ' Call Oracle Stored Procedure.
    
    GetProcessingInfoFromDB strFileNumber, CInt(strFilingYear), CDate(strReceiveDT), strAnnualType, _
                            strWorkitemID, strCompanyName, strErrCode, strErrMsg
    
    Select Case strErrCode
        
        Case Is = "00", "01", "11"            ' Brims return code is either "00" - Success or "01" - Business Rules Failed, No update done on brims side
        
            eErrorStatus = UpdateCALFields(objCALDocument, strAnnualType, strCompanyName, strReceiveDT, strFileNumber, blnChildWorkitem, strErrCode, strErrMsg)
            If eErrorStatus = icSuccess Then            ' Successfully updated the document, send it to default
                objCALDocument.Save
                DoEvents
                objCALDocument.Close calCloseSaveChanges
                DoEvents
                Sleep 1000
                DoEvents
                Set objCALDocument = Nothing
                DoEvents
                DoEvents
                If Not blnChildWorkitem Then
                    objCALClientListitem.SendToDefault calSendDiscard
                    Set objCALClientListitem = Nothing
                End If
                DoEvents
                DoEvents
                eErrorStatus = CheckCALError
                If eErrorStatus <> icCriticalError Then
                    If strErrCode = "00" And strAnnualType = "C" Then
                        gCleanCnt = gCleanCnt + 1
                    Else
                        gManualCnt = gManualCnt + 1
                    End If
                    If Not objMainForm Is Nothing Then
                        With objMainForm
                            .StatusBar1.Panels(1).Text = "Current Counts"
                            .StatusBar1.Panels(2).Text = vbNullString
                            .StatusBar1.Panels(3).Text = "Clean:" & Trim$(Str$(gCleanCnt))
                            .StatusBar1.Panels(4).Text = "Manual:" & Trim$(Str$(gManualCnt))
                            .Refresh
                             DoEvents
                        End With
                    End If
                    ProcessWorkitem = icSuccess
                Else
                    strErrMsg = "Critical error occurred. Shutting down."
                    ProcessWorkitem = icCriticalError
                    
                End If
                WriteToLogFile strErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
            Else
                objCALDocument.Close calCloseAbortChanges           ' Not Successful, send it back to the queue
                DoEvents
                Sleep 1000
                Set objCALDocument = Nothing
                DoEvents
                If Not blnChildWorkitem Then
                    Set objCALSendQueue = New CALQueue
                    DoEvents
                    With objCALSendQueue
                        .Client = objCALClient
                        .Name = objCALClientListitem.Info.SourceWorkstep
                        .Type = calQueueTypeWorkstep
                    End With
                    DoEvents
                    objCALClientListitem.Send objCALSendQueue, calSendDiscard
                    Set objCALSendQueue = Nothing
                    Set objCALClientListitem = Nothing
                End If
                DoEvents
                eErrorStatus = CheckCALError
                If eErrorStatus <> icCriticalError Then
                    ProcessWorkitem = icSuccess
                    strErrMsg = strWorkitemID & "- was not succesful. Sending back to the queue."
                    WriteToLogFile strErrMsg
                    If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
                Else
                    strErrMsg = "BRIMS returned error code 01. Terminating the process due to 01 code."
                    WriteToLogFile strErrMsg
                    If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
                    ProcessWorkitem = icCriticalError
                End If
            End If
            
        Case Is = "02"                      'Return the workitem to the queue and try later
            
            If blnChildWorkitem Then
                objCALDocument.Close calCloseAbortChanges
                eErrorStatus = CheckCALError
                Set objCALDocument = Nothing
                If eErrorStatus <> icCriticalError Then
                    strErrMsg = strWorkitemID & "- was not succesful. Sending the folder back to queue."
                    ProcessWorkitem = icTryAgain
                Else
                    strErrMsg = "Critial error occurred. Shutting down."
                    ProcessWorkitem = icCriticalError
                End If
                WriteToLogFile strErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
                strErrMsg = vbNullString
                Exit Function
            End If
            
            strErrCode = vbNullString
            Set objCALFields = objCALDocument.FormFields(calFieldsNoViews)
            DoEvents
            If Not objCALFields Is Nothing Then
                Set objCALField = Nothing                       'Get the value in WFT8_1 and if the value is blank
                Set objCALField = objCALFields.Find("WFT8_1")   ' then put ERR and send it back to the queue. If it is ERR then set it ANN and
                If Not objCALField Is Nothing Then             ' route it to Annuals queue. Do not resend the document more than once.
                    strErrCode = objCALField.Value
                    strErrCode = IIf((strErrCode = "ERR"), "ANN", "ERR")
                    DoEvents
                    objCALField.Value = strErrCode
                    If strErrCode = "ANN" Then
                        DoEvents
                        Set objCALField = Nothing
                        Set objCALField = objCALFields.Find("WFBRA_NOTES")
                        strErrMsg = "*Note by AnnRobot on " & FormatDateTime$(Now, vbGeneralDate) & vbCrLf
                        strErrMsg = strErrMsg & "Failed Updating BRIMS after 2 attempts. Routing to Annuals" & vbCrLf & vbCrLf
                        If Not objCALField Is Nothing Then objCALField.Value = strErrMsg & objCALField.Value
                        DoEvents
                    End If
                    objCALDocument.Save
                    DoEvents
                    Sleep 1000
                    DoEvents
                End If
            End If
                
            Set objCALField = Nothing
            Set objCALFields = Nothing
            DoEvents
            objCALDocument.Close calCloseAbortChanges
            DoEvents
            Sleep 1000
            DoEvents
            Set objCALDocument = Nothing
            DoEvents
            DoEvents
            objCALClientListitem.SendToDefault calSendDiscard
            DoEvents
            eErrorStatus = CheckCALError
            If eErrorStatus <> icCriticalError Then
                ProcessWorkitem = icSuccess
            Else
                ProcessWorkitem = icCriticalError
            End If
            Set objCALClientListitem = Nothing
            If strErrCode = "ERR" Then
                strErrMsg = strWorkitemID & "- was not succesful. Sending back to the queue."
            Else
                strErrMsg = strWorkitemID & "-Failed after 2 attempts. Routing to Annuals."
            End If
            WriteToLogFile strErrMsg
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
            Sleep 1000
            DoEvents
            strErrCode = vbNullString
            DoEvents

        Case Else
        
            If blnChildWorkitem Then
                objCALDocument.Close calCloseAbortChanges
                eErrorStatus = CheckCALError
                Set objCALDocument = Nothing
                If eErrorStatus <> icCriticalError Then
                    strErrMsg = "Unknown error occured when trying to contact BRIMS. Shutting down."
                    ProcessWorkitem = icTryAgain
                Else
                    strErrMsg = "Critial error occurred. Shutting down."
                    ProcessWorkitem = icCriticalError
                End If
                WriteToLogFile strErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
                strErrMsg = vbNullString
                Exit Function
            End If

            If (objCALDocument.Info.Status And calObjStatusOpen) = calObjStatusOpen Then    'Critical error on BRIMS side, so logoff
                objCALDocument.Close calCloseAbortChanges
                DoEvents
                Sleep 1000
            End If
            Set objCALDocument = Nothing
            DoEvents
            With objCALSendQueue
                .Client = objCALClient
                .Name = objCALClientListitem.Info.SourceWorkstep
                .Type = calQueueTypeWorkstep
            End With
            DoEvents
            objCALClientListitem.Send objCALSendQueue, calSendDiscard
            DoEvents
            Set objCALClientListitem = Nothing
            Set objCALSendQueue = Nothing
            DoEvents
            ProcessWorkitem = icCriticalError
            strErrMsg = "Unknown error occured when trying to contact BRIMS. Shutting down."
            WriteToLogFile strErrMsg
            DoEvents
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg

    End Select
    
    DoEvents
    Sleep 1000
    DoEvents
            
End Function

Private Function UpdateCALFields(objCALDocument As CALDocument, strAnnualType As String, strCompanyName As String, _
                                strReceiveDT As String, strFileNumber As String, blnChildItem As Boolean, strErrCode As String, strMsg As String) As CALProcessStatus
'
' Input: CALDocument, Annual type, Company Name, Receive Date, Error Code returned from stored procedure.
' Output: CALProcessStatus
' Update the CAL form fields depending on the Error code and the type of annual (clean or manual)
'
    On Error GoTo Update_Error
    
    Dim objCALField As CALFormField
    Dim objCALFields As CALFormFields
    Dim strBRAMsg As String
    Dim strDocName As String
    Dim strTmpMsg As String
    Dim strTempCompanyName As String
    Dim strTempCompanyName1 As String
    
    
    strDocName = Trim$(objCALDocument.Info.Name)
    
    Set objCALFields = objCALDocument.FormFields(calFieldsNoViews)
    DoEvents
    
    If objCALFields Is Nothing Then
        UpdateCALFields = icVerifyNotOK
        Exit Function
    End If
        
    Set objCALField = Nothing
    Set objCALField = objCALFields.Find("WFCOMPANY_NAME")
    If Not objCALField Is Nothing Then strTempCompanyName = objCALField.Value
    
    Set objCALField = Nothing
    Set objCALField = objCALFields.Find("WFCOMPANY_NAME1")
    If Not objCALField Is Nothing Then strTempCompanyName1 = objCALField.Value
    
    strTempCompanyName = strTempCompanyName & strTempCompanyName1
    strTempCompanyName = UCase$(strTempCompanyName)
    If strTempCompanyName > " " Then strCompanyName = strTempCompanyName
    
    If strCompanyName > " " Then
        strTmpMsg = strCompanyName          'Remove the end astericks
        Do While Right$(strTmpMsg, 1) = "*"
            strTmpMsg = Left$(strTmpMsg, Len(strTmpMsg) - 1)
        Loop
        strTmpMsg = Replace$(strTmpMsg, "<", vbNullString, , , vbTextCompare)   'Remove the < and > symbols
        strTmpMsg = Replace$(strTmpMsg, ">", vbNullString, , , vbTextCompare)
        strCompanyName = UCase$(strTmpMsg)
        strTmpMsg = vbNullString
        If Len(strCompanyName) > 250 Then
            Set objCALField = Nothing
            Set objCALField = objCALFields.Find("WFCOMPANY_NAME")
            If Not objCALField Is Nothing Then objCALField.Value = Left$(strCompanyName, 250)
            DoEvents
            Set objCALField = Nothing
            Set objCALField = objCALFields.Find("WFCOMPANY_NAME1")
            If Not objCALField Is Nothing Then objCALField.Value = Mid$(strCompanyName, 251)
        Else
            Set objCALField = Nothing
            Set objCALField = objCALFields.Find("WFCOMPANY_NAME")
            If Not objCALField Is Nothing Then objCALField.Value = strCompanyName
        End If
    End If
    
    DoEvents
    Set objCALField = Nothing   'If receive date is blank, then enter receive date or else format it
    Set objCALField = objCALFields.Find("WFRECEIVE_DATE")
    strReceiveDT = Format$(strReceiveDT, "YYYY-MM-DD")
    If Not objCALField Is Nothing Then
        If Not objCALField.Value > " " Then
            objCALField.Value = strReceiveDT
        Else
            strTmpMsg = objCALField.Value
            If IsDate(strTmpMsg) Then
                strTmpMsg = Format$(strTmpMsg, "YYYY-MM-DD")
                objCALField.Value = strTmpMsg
                strTmpMsg = vbNullString
            End If
        End If
    End If

    DoEvents
    Set objCALField = Nothing
    Set objCALField = objCALFields.Find("WFT8")
    If Not objCALField Is Nothing Then
        If strFileNumber > " " Then objCALField.Value = strFileNumber
    End If

    Set objCALField = Nothing                           ' Suffix for GPLLP
    Set objCALField = objCALFields.Find("WFT8_2")
    If gblnGPLLP Then
        If Not objCALField Is Nothing Then objCALField.Value = "K5"  'GP w/ LLP Route To GPLLP Queue.
    Else
        If Not objCALField Is Nothing Then objCALField.Value = Right(strFileNumber, 2)
    End If
    DoEvents

    DoEvents
    Set objCALField = Nothing                           ' Enter user-id
    Set objCALField = objCALFields.Find("WFDOCI_USER")
    If Not objCALField Is Nothing Then objCALField.Value = UCase$(objCALClient.UserName)
    DoEvents
    
    DoEvents
    Set objCALField = Nothing       ' Enter Current Date
    Set objCALField = objCALFields.Find("WFDOCI_DATE")
    If Not objCALField Is Nothing Then objCALField.Value = Format$(Date$, "YYYY-MM-DD")
    DoEvents
    
    DoEvents
    strBRAMsg = FormatDateTime$(Now, vbLongTime)    'Did not feel like using another variable
    strBRAMsg = Replace(strBRAMsg, " ", "", , , vbTextCompare)
    Set objCALField = Nothing                           ' Enter Current Time
    Set objCALField = objCALFields.Find("WFDOCI_TIME")
    If Not objCALField Is Nothing Then objCALField.Value = strBRAMsg
    DoEvents


    If strErrCode = "00" Then
    
        If strAnnualType = "C" Then
            strBRAMsg = "Annual successfully processed. Marking Process status as Completed."
            strTmpMsg = vbCrLf & "BRIMS Msg: " & strMsg
            DoEvents
            Set objCALField = Nothing
            Set objCALField = objCALFields.Find("WFPROCESS_STATUS")
            If Not objCALField Is Nothing Then objCALField.Value = "Completed"
            DoEvents
            
            Set objCALField = Nothing                           ' Indicate that the workitem is being routed to Filing
            Set objCALField = objCALFields.Find("WFT8_1")
            If Not objCALField Is Nothing Then objCALField.Value = "FIL"
            DoEvents
        Else
            strBRAMsg = "Company Name Retrived. Routing to Annuals queue."
            strTmpMsg = vbCrLf & "BRIMS Msg: " & strMsg
            DoEvents
            Set objCALField = Nothing                           ' Indicate that the workitem is being routed to Annuals Q
            Set objCALField = objCALFields.Find("WFT8_1")
            If Not objCALField Is Nothing Then objCALField.Value = "ANN"
            DoEvents
        End If
        
    ElseIf strErrCode = "01" Then   'Business rules failed on BRIMS
        
        If strAnnualType = "C" Then
            strBRAMsg = "Failed updating BRIMS. Routing to Annuals queue."
            strTmpMsg = vbCrLf & "BRIMS Msg: " & strMsg
        Else
            strBRAMsg = "Company Name Retrived. Routing to Annuals queue."
            strTmpMsg = vbCrLf & "BRIMS Msg: " & strMsg
        End If
    
        DoEvents
        Set objCALField = Nothing                           ' Indicate that the workitem is being routed to Annuals Q
        Set objCALField = objCALFields.Find("WFT8_1")
        If Not objCALField Is Nothing Then objCALField.Value = "ANN"
        DoEvents
    Else                            'Something happened on BRIMS side, no company name, send it for Manual processing
        strBRAMsg = "Failed retrieving the company name. Routing to Annuals queue."
        strTmpMsg = vbCrLf & "BRIMS Msg: " & strMsg
        DoEvents
        Set objCALField = Nothing                           ' Indicate that the workitem is being routed to Annuals Q
        Set objCALField = objCALFields.Find("WFT8_1")
        If Not objCALField Is Nothing Then objCALField.Value = "ANN"
        DoEvents

    End If
    
    strMsg = strDocName & "-" & strBRAMsg
    strBRAMsg = "*Note by AnnRobot on " & FormatDateTime$(Now, vbGeneralDate) & vbCrLf & strBRAMsg & strTmpMsg
    DoEvents
    Set objCALField = Nothing                           ' Write the Message in the BRA notes
    Set objCALField = objCALFields.Find("WFBRA_NOTES")
    If Not objCALField Is Nothing Then
        If objCALField.Value > " " Then
            objCALField.Value = strBRAMsg & vbCrLf & vbCrLf & objCALField.Value
        Else
            objCALField.Value = strBRAMsg
        End If
    End If
    If blnChildItem Then
        Set objCALField = Nothing   'If this is a document in a folder, then undo the wft8_1 value
        Set objCALField = objCALFields.Find("WFT8_1")
        If Not objCALField Is Nothing Then objCALField.Value = vbNullString
        DoEvents

    End If
    
    DoEvents
    UpdateCALFields = icSuccess
    Set objCALField = Nothing
    Set objCALFields = Nothing
    Exit Function
    
Update_Error:
    
    Err.Clear
    Set objCALField = Nothing
    Set objCALFields = Nothing
    strMsg = vbNullString
    UpdateCALFields = icVerifyNotOK
                    
End Function

Private Function ProcessFolderWorkitem(objCALClientListitem As CALClientListItem) As CALProcessStatus
'
'Input: CALClientListItem
'OutPut: Process Status
'Open the folder and for each child item found, check whether the child item belongs to one of the document
'classes. If the captured barcode is "ANN", process it, other wise do not process it.

    Dim strWorkitemID As String
    Dim strChildWorkitemID As String
    Dim blnWorkitemNotOK As Boolean
    Dim blnTryAgain As Boolean
    Dim intCounter As Integer
    
    Dim objCALFolder As CALFolder
    Dim objCALFolderChildren As CALFolderChildren
    Dim objCALFolderChildrenItems As CALFolderChildrenItems
    Dim objCALFolderChild As CALFolderChild
    Dim objCALID As CALObjID
    Dim objCALClientListItemChild As CALClientListItem
    Dim objCALDocument As CALDocument
    Dim objCALFormFields As CALFormFields
    Dim objCALFormField As CALFormField
    Dim eErrorStatus As CALProcessStatus
    
    On Error Resume Next
    
    gblnFolderGPLLP = False   'GP/LLP
    strWorkitemID = objCALClientListitem.Info.Name
    If Not objMainForm Is Nothing Then
        With objMainForm
            .StatusBar1.Panels(1).Text = "Processing folder"
            .StatusBar1.Panels(2).Text = strWorkitemID
            .StatusBar1.Panels(3).Text = "Clean:" & Trim$(Str$(gCleanCnt))
            .StatusBar1.Panels(4).Text = "Manual:" & Trim$(Str$(gManualCnt))
            .Refresh
             DoEvents
        End With
    End If
    
    If (objCALClientListitem.Info.Status And calObjStatusReserved) = calObjStatusReserved Then
        Set objCALFolder = objCALClientListitem.OpenedItem
    ElseIf (objCALClientListitem.Info.Status And calObjStatusOpen) = calObjStatusOpen Then
        Set objCALFolder = objCALClientListitem.OpenedItem
        DoEvents
        objCALFolder.Close calCloseAbortChanges + calCloseRetainLock
        DoEvents
        DoEvents
        Set objCALFolder = objCALClientListitem.Open(calOpenReadWrite)
    Else
        Set objCALFolder = objCALClientListitem.Open(calOpenReadWrite)
    End If
    
    eErrorStatus = CheckCALError
    If eErrorStatus = icSuccess Then
        DoEvents
    Else
        Set objCALFolder = Nothing
        Set objCALClientListitem = Nothing
        DoEvents
        ProcessFolderWorkitem = icSuccess
        Exit Function
    End If

    If objCALFolder Is Nothing Then
        Set objCALClientListitem = Nothing
        ProcessFolderWorkitem = icSuccess
        Exit Function
    End If
    
    DoEvents
    If (objCALFolder.Info.Status And calObjStatusReserved) = calObjStatusReserved Then
        DoEvents
    Else
        objCALFolder.Close calCloseAbortChanges
        Set objCALFolder = Nothing
        ErrMsg = "Unable to reserve the folder."
        If PlaceWorkitemInError(objCALClientListitem, ErrMsg) <> icCriticalError Then
            ProcessFolderWorkitem = icSuccess
        Else
            ProcessFolderWorkitem = icCriticalError
        End If
        Set objCALClientListitem = Nothing
        DoEvents
        Sleep 1000
        DoEvents
        DoEvents
        Exit Function
    End If
    
    blnWorkitemNotOK = False
    Set objCALFolderChildren = objCALFolder.Children
    
    If objCALFolderChildren Is Nothing Then
        ErrMsg = "Folder " & strWorkitemID & " not OK. Sending to next workstep."
        blnWorkitemNotOK = True
    Else
        Set objCALFolderChildrenItems = objCALFolderChildren.DocumentItems
        If objCALFolderChildrenItems Is Nothing Then
            ErrMsg = "Folder " & strWorkitemID & " has no child items. Sending to next workstep."
            blnWorkitemNotOK = True
        ElseIf objCALFolderChildrenItems.Count = 0 Then
            blnWorkitemNotOK = True
            ErrMsg = "Folder " & strWorkitemID & " has no child items. Sending to next workstep."
        End If
    End If
    
    If blnWorkitemNotOK Then
        Set objCALFolderChildren = Nothing
        Set objCALFolderChildrenItems = Nothing
        objCALFolder.Close calCloseAbortChanges
        Set objCALFolder = Nothing
        objCALClientListitem.SendToDefault calSendDiscard
        If ErrMsg > " " Then
            WriteToLogFile ErrMsg
            DoEvents
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
        End If
        DoEvents
        eErrorStatus = CheckCALError
        If eErrorStatus <> icCriticalError Then
            ProcessFolderWorkitem = icSuccess
        Else
            ProcessFolderWorkitem = icCriticalError
        End If
        Set objCALClientListitem = Nothing
        Exit Function
    End If
    
    For Each objCALFolderChild In objCALFolderChildrenItems
        blnTryAgain = False
        If objCALFolderChild.Info.Type = calObjTypeDocument Then
            
            For intCounter = 0 To mintDocCLassCount - 1
            
                If objCALFolderChild.Class = mstrDocClass(intCounter) Then
                
                    Set objCALID = objCALFolderChild.Info.ObjID
                    Set objCALClientListItemChild = objCALClientList.Add(objCALID, 0)
                    
                    eErrorStatus = CheckCALError
                    Select Case eErrorStatus
                        Case Is = icSuccess
                            DoEvents
                        Case Is = icWorkitemInList
                            Set objCALClientListItemChild = objCALClientList.Find(objCALID)
                            DoEvents
                        Case Else
                            ErrMsg = "Unable to retrieve child item " & objCALFolderChild.Info.Name
                            
                            Set objCALClientListItemChild = Nothing
                            Set objCALFolderChild = Nothing
                            Set objCALFolderChildrenItems = Nothing
                            objCALFolder.Close calCloseAbortChanges
                            Set objCALFolder = Nothing
                            If PlaceWorkitemInError(objCALClientListitem, ErrMsg) <> icCriticalError Then
                                ProcessFolderWorkitem = icSuccess
                            Else
                                ProcessFolderWorkitem = icCriticalError
                            End If
                            Set objCALClientListitem = Nothing
                            Exit Function
                    End Select
                    
                    If objCALClientListItemChild Is Nothing Then
                        ErrMsg = "Unable to retrieve child item "
                        If Not objCALFolderChild Is Nothing Then ErrMsg = ErrMsg & objCALFolderChild.Info.Name
                        Set objCALClientListItemChild = Nothing
                        Set objCALFolderChild = Nothing
                        Set objCALFolderChildrenItems = Nothing
                        objCALFolder.Close calCloseAbortChanges
                        Set objCALFolder = Nothing
                        If PlaceWorkitemInError(objCALClientListitem, ErrMsg) <> icCriticalError Then
                            ProcessFolderWorkitem = icSuccess
                        Else
                            ProcessFolderWorkitem = icCriticalError
                        End If
                        Set objCALClientListitem = Nothing
                        Exit Function
                    End If
                    
                    If (objCALClientListItemChild.Info.Status And calObjStatusReserved) = calObjStatusReserved Then
                        Set objCALDocument = objCALClientListItemChild.OpenedItem
                    ElseIf (objCALClientListItemChild.Info.Status And calObjStatusOpen) = calObjStatusOpen Then
                        Set objCALDocument = objCALClientListItemChild.OpenedItem
                        objCALDocument.Close calCloseAbortChanges
                        Set objCALDocument = Nothing
                        Set objCALDocument = objCALClientListItemChild.Open(calOpenReadWrite)
                    Else
                        Set objCALDocument = objCALClientListItemChild.Open(calOpenReadWrite)
                    End If
                    
                    If objCALDocument Is Nothing Then
                        ErrMsg = "Unable to open document "
                        If Not objCALFolderChild Is Nothing Then ErrMsg = ErrMsg & objCALFolderChild.Info.Name
                        Set objCALClientListItemChild = Nothing
                        Set objCALFolderChild = Nothing
                        Set objCALFolderChildrenItems = Nothing
                        objCALFolder.Close calCloseAbortChanges
                        Set objCALFolder = Nothing
                        If PlaceWorkitemInError(objCALClientListitem, ErrMsg) <> icCriticalError Then
                            ProcessFolderWorkitem = icSuccess
                        Else
                            ProcessFolderWorkitem = icCriticalError
                        End If
                        Set objCALClientListitem = Nothing
                        Exit Function
                    End If
                    
                    strChildWorkitemID = objCALDocument.Info.Name
                    
                    Set objCALFormFields = objCALDocument.FormFields(calFieldsNoViews)
                    DoEvents
                    Set objCALFormField = objCALFormFields.Find("WFCAPTURED_BARCODE")
                    If objCALFormField.Value <> "ANN" Then
                        objCALDocument.Close calCloseAbortChanges
                        Set objCALDocument = Nothing
                        objCALClientList.Remove objCALClientListItemChild
                        DoEvents
                        Set objCALClientListItemChild = Nothing
                        DoEvents
                    Else
                      ' Commented out for cutover.
                        'Set objCALFormField = objCALFormFields.Find("WFDOCI_USER")  ' If processed by AnnRobot, then do not process again
                        'If objCALFormField.Value = UCase$(objCALClient.UserName) Then
                        '    ErrMsg = strChildWorkitemID & " already processed by robot. Skipping to next workitem."
                        '    WriteToLogFile ErrMsg
                        '    DoEvents
                        '    If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
                        '    Set objCALFormField = Nothing
                        '    objCALDocument.Close calCloseAbortChanges
                        '    Set objCALDocument = Nothing
                        '    objCALClientList.Remove objCALClientListItemChild
                        '    DoEvents
                        '    Set objCALClientListItemChild = Nothing
                        '    DoEvents
                        '    Set objCALFormField = Nothing
                        'Else
                            Set objCALDocument = Nothing
                            eErrorStatus = ProcessWorkitem(objCALClientListItemChild, True)
                            If eErrorStatus = icCriticalError Then
                                objCALClientList.Remove objCALClientListItemChild
                                Set objCALFolderChild = Nothing
                                Set objCALFolderChildrenItems = Nothing
                                Set objCALClientListItemChild = Nothing
                                Set objCALFolder = Nothing
                                Set objCALClientListitem = Nothing
                                ProcessFolderWorkitem = icCriticalError
                                Exit Function
                            Else
                                objCALClientList.Remove objCALClientListItemChild
                                If eErrorStatus = icTryAgain Then
                                    Set objCALFormField = Nothing
                                    Set objCALFormFields = Nothing
                                    blnTryAgain = True
                                    Exit For
                                End If

                            End If
                            DoEvents
                        'End If
                    End If
                    Set objCALFormField = Nothing
                    Set objCALFormFields = Nothing
                    Set objCALClientListItemChild = Nothing
                End If
                Err.Clear
            Next
        End If
    Next
    
    Set objCALFolderChild = Nothing
    Set objCALClientListItemChild = Nothing
    Set objCALFolderChildrenItems = Nothing
    Set objCALFormFields = objCALFolder.FormFields(calFieldsNoViews)
    Set objCALFormField = objCALFormFields.Find("WFSUSPEND")
    

'Send it back to the queue by setting the Suspend Field to "R"
' Route rules will verify this field and if "R" will route it back to Annuals robot

    If blnTryAgain Then
        If Not objCALFormField Is Nothing Then objCALFormField.Value = "R"
        ErrMsg = "Folder " & strWorkitemID & " is routed back to queue."
        WriteToLogFile ErrMsg
        DoEvents
        If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
 'If GP/LLP Route to Annuals GPLLP Workstep.  -- Ryee
    ElseIf gblnFolderGPLLP Then
        If Not objCALFormField Is Nothing Then objCALFormField.Value = "G"
        ErrMsg = "Folder witn GP/LLP Item" & strWorkitemID & " is routed Annuals GPLLP queue."
        WriteToLogFile ErrMsg
        DoEvents
        If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
    Else
        If Not objCALFormField Is Nothing Then objCALFormField.Value = vbNullString
        ErrMsg = "Folder " & strWorkitemID & " is sent to Annuals queue."
        WriteToLogFile ErrMsg
        DoEvents
        If Not objMainForm Is Nothing Then objMainForm.AddItemToList ErrMsg
        
    End If
        
    objCALFolder.Close calCloseSaveChanges
    Set objCALFormFields = Nothing
    Set objCALFormField = Nothing
    Set objCALFolder = Nothing
    eErrorStatus = CheckCALError
    If eErrorStatus = icCriticalError Then
        ProcessFolderWorkitem = icCriticalError
        Set objCALClientListitem = Nothing
        Exit Function
    End If
    
    DoEvents
    objCALClientListitem.SendToDefault calSendDiscard
    Sleep 1000
    eErrorStatus = CheckCALError
    If eErrorStatus = icCriticalError Then
        ProcessFolderWorkitem = icCriticalError
    Else
        ProcessFolderWorkitem = icSuccess
    End If
    Set objCALClientListitem = Nothing
    DoEvents
    Err.Clear
    
End Function

Public Sub LogoffWorkflow()
'
' Logoff from workflow. If workflow doesn't respond then set objCALClientList, objCALClient and obJCALMaster to nothing.

    On Error Resume Next
    
    Dim strTemp As String
    Dim strSectionName As String
    Dim strKeyName As String
    Dim lngTotalCnt As Long
    
    LogOffFromDatabase          'Logoff from database
    DoEvents
    If Not objMainForm Is Nothing Then
        With objMainForm
             DoEvents
            .StatusBar1.Panels(1).Text = "Logging of from domain " & mstrDomain & "."
            .StatusBar1.Panels(2).Text = vbNullString
            .Refresh
             DoEvents
        End With
    End If
    DoEvents
    Set objCALClientList = Nothing
    DoEvents
    If Not objCALClient Is Nothing Then objCALClient.Destroy calDestroyAbortChanges
    DoEvents
    
'Save the processing counts and preferences. These preferences change during the course of the program.
'Section Name
    strSectionName = "Preferences"
    
'Key Name CleanCount
    strKeyName = "CleanCount"
    strTemp = Trim$(Str$(gCleanCnt))
    SaveSetting App.EXEName, strSectionName, strKeyName, strTemp
    DoEvents
    
'Key Name ManualCount
    strKeyName = "ManualCount"
    strTemp = Trim$(Str$(gManualCnt))
    SaveSetting App.EXEName, strSectionName, strKeyName, strTemp
    DoEvents
    
'Write the totals to the file and the form.
    lngTotalCnt = gCleanCnt + gManualCnt
    DoEvents
    strTemp = "Counts-Clean:" & Trim$(Str$(gCleanCnt)) & " Manual:" & strTemp & "  Total:" & Trim$(Str$(lngTotalCnt))
    WriteToLogFile strTemp
    If Not objMainForm Is Nothing Then objMainForm.AddItemToList strTemp
    DoEvents
    strTemp = "Logged off from domain " & mstrDomain & "."
    WriteToLogFile strTemp
    If Not objMainForm Is Nothing Then objMainForm.AddItemToList strTemp
    DoEvents
    
'Close the file
    If Not objCurFile Is Nothing Then objCurFile.Close
    DoEvents
    Sleep 1000
    DoEvents
    DoEvents
    DestroyGlobalObjects
    gblnRunning = False
    DoEvents
    Err.Clear
    
End Sub

Private Function PlaceWorkitemInError(objCALClientListitem As CALClientListItem, strComment As String) As CALProcessStatus
'
' This Procedure places the ClientListItem in Error in that workstep
'
        Dim eErrorStatus As CALProcessStatus
        Dim objSendCALQ As CALQueue
        Dim strWorkitemID As String
        Dim strMsg As String
        
        On Error Resume Next
        
        If objCALClientListitem Is Nothing Then
            PlaceWorkitemInError = icSuccess
            Exit Function
        End If
        
        strWorkitemID = objCALClientListitem.Info.Name
        
        objCALClientListitem.PlaceInError calPlaceInErrorDiscard, strComment
        eErrorStatus = CheckCALError
        Select Case eErrorStatus
            Case Is = icSuccess
                strComment = strWorkitemID & " - " & strComment
                WriteToLogFile strComment
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList strComment
                Set objCALClientListitem = Nothing
                PlaceWorkitemInError = eErrorStatus
            Case Is = icCriticalError
                strMsg = "Critical Error occured. Processing Suspended (PlaceWorkitemInError)."
                WriteToLogFile strMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList strMsg
                Set objCALClientListitem = Nothing
                PlaceWorkitemInError = icCriticalError
            Case Is = icPlaceErrorFailed
                If Not objCALClientListitem Is Nothing Then
                    Set objSendCALQ = New CALQueue
                    With objSendCALQ
                        .Client = objCALClient
                        .Name = objCALClientListitem.Info.SourceWorkstep
                        .Type = calQueueTypeWorkstep
                    End With
                    objCALClientListitem.Send objSendCALQ, calSendDiscard
                    If CheckCALError <> icCriticalError Then
                        strMsg = strWorkitemID & " sent back to " & gcurWorkset & " (PlaceWorkitemInError)."
                        WriteToLogFile strMsg
                        If Not objMainForm Is Nothing Then objMainForm.AddItemToList strMsg
                        Set objCALClientListitem = Nothing
                        PlaceWorkitemInError = icSuccess
                    Else
                        strMsg = "Critical Error occured. Processing Suspended (PlaceWorkitemInError- " & strWorkitemID & " )."
                        WriteToLogFile strMsg
                        If Not objMainForm Is Nothing Then objMainForm.AddItemToList strMsg
                        Set objCALClientListitem = Nothing
                        PlaceWorkitemInError = icCriticalError
                    End If
                Else
                    strMsg = strWorkitemID & " could not place in error.(PlaceWorkitemInError)."
                    WriteToLogFile strMsg
                    If Not objMainForm Is Nothing Then objMainForm.AddItemToList strMsg
                    Set objCALClientListitem = Nothing
                    PlaceWorkitemInError = icSuccess
                End If
            Case Else
                strMsg = strWorkitemID & " could not place in error.(PlaceWorkitemInError)."
                WriteToLogFile strMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList strMsg
                Set objCALClientListitem = Nothing
                PlaceWorkitemInError = icSuccess
        End Select
        strComment = vbNullString
        strMsg = vbNullString

End Function


Private Function CheckCALError() As CALProcessStatus
'
' Check the error code and if it is a CAL API error, then pass the appropriate error value (enum of  CALProcessStatus)
' if error is by VB or CAL, then pass critical error.

    If Err.Number <> 0 Then
        If Err.Number = 1000 Then
            If objCALMaster.LastError.Code <> 0 Then
                Select Case objCALMaster.LastError.Code
                    Case Is = 10
                        CheckCALError = icInvalidQueue
                    Case Is = 18
                        CheckCALError = icQueueEmpty
                    Case Is = 26
                        CheckCALError = icWorkitemInList
                    Case Is = 30
                        CheckCALError = icWorkitemNotInList
                    Case Is = 96
                        CheckCALError = icWorkitemisWIP
                    Case Is = 115
                        CheckCALError = icWorkitemOpen
                    Case Is = 116
                        CheckCALError = icWorkitemNew
                    Case Is = 35
                        CheckCALError = icSaveFailed
                    Case Is = 149
                        CheckCALError = icPlaceErrorFailed
                    Case Is = 88, 119, 121, 179, 200, 202
                        CheckCALError = icSetPageError
                    Case Is = 112, 181
                        CheckCALError = icCloseFailed
                    Case Is = 123, 138                       ' 123- workitem was placed in error by route engine
                        CheckCALError = icSuccess            ' 138 - No markups found
                    Case Else
                        CheckCALError = icCriticalError
                End Select
            Else
                CheckCALError = icCriticalError
            End If
        ElseIf Err.Number > 1000 And Err.Number <= 1100 Then
            Select Case Err.Number
                Case Is = 1013
                    CheckCALError = icSaveFailed
                Case Else
                    CheckCALError = icCriticalError
            End Select
        Else
            Select Case Err.Number
                Case Is = 1104, 1172
                    CheckCALError = icSetPageError
                Case Else
                    CheckCALError = icCriticalError
                    'MsgBox Err.Number
            End Select
        End If
    Else
        CheckCALError = icSuccess
    End If
    
    Err.Clear
            
    
End Function


Private Function ProcessResubWorkitem(objCALClientListitem As CALClientListItem, objCALDocument As CALDocument) As CALProcessStatus

'
' Input: CALClientListItem
' Output: Process status of the function
' This function will take the calclientlistitem and then calls the oracle stored procedure.
' Oracle stored procedure in turn will return Company Name, Transaction Code, Transaction Description, Return Code
' & Return Message.
' The company name is transfered to the objcalclientlistitem and a new note is keyed into the workitem.
' If the return code is "00", "01" or "11", then the resub is routed to
' suspend queue, else put into error.
'
    
    On Error Resume Next
    
    'Dim objCALDocument As CALDocument
    Dim objCALFields As CALFormFields
    Dim objCALField As CALFormField
    Dim objCALSendQueue As CALQueue
    Dim objCALNotes As CALNotes
    
    Dim eErrorStatus As CALProcessStatus
    Dim strFileNumber As String
    Dim strReceiveDT As String
    Dim strFilingYear As String
    Dim strCompanyName As String
    Dim strWorkitemID As String
    Dim strAnnualType As String * 1
    Dim strErrCode As String
    Dim strErrMsg As String
    Dim strTemp As String
    

    strWorkitemID = objCALClientListitem.Info.Name
    
    If objCALDocument Is Nothing Then
    
        If (objCALClientListitem.Info.Status And calObjStatusReserved) = calObjStatusReserved Then
             Set objCALDocument = objCALClientListitem.OpenedItem
        ElseIf (objCALClientListitem.Info.Status And calObjStatusOpen) = calObjStatusOpen Then
            Set objCALDocument = objCALClientListitem.OpenedItem
            DoEvents
            objCALDocument.Close calCloseAbortChanges + calCloseRetainLock
            DoEvents
            Sleep 1000
            DoEvents
            Set objCALDocument = objCALClientListitem.Open(calOpenReadWrite)
        Else
            Set objCALDocument = objCALClientListitem.Open(calOpenReadWrite)
        End If
    
        eErrorStatus = CheckCALError
        If eErrorStatus = icSuccess Then
            DoEvents
        Else
            Set objCALDocument = Nothing
            Set objCALClientListitem = Nothing
            DoEvents
            ProcessResubWorkitem = icSuccess
            Exit Function
        End If
    End If
    
    
    DoEvents
    If (objCALDocument.Info.Status And calObjStatusReserved) = calObjStatusReserved Then
        DoEvents
    Else
        objCALDocument.Close calCloseAbortChanges
        Set objCALDocument = Nothing
        Set objCALClientListitem = Nothing
        DoEvents
        Sleep 1000
        ProcessResubWorkitem = icSuccess
        DoEvents
        Exit Function
    End If
    
    ' Get the required fields from workitem.
       
    strErrMsg = vbNullString
    
    Set objCALFields = objCALDocument.FormFields(calFieldsNoViews)
    DoEvents
    If Not objCALFields Is Nothing Then
        
        Set objCALField = Nothing                       'Get file number
        Set objCALField = objCALFields.Find("WFT8")
        If Not objCALField Is Nothing Then strFileNumber = objCALField.Value
        DoEvents
        
        Set objCALField = Nothing                       ' Get Receive Date
        Set objCALField = objCALFields.Find("WFRECEIVE_DATE")
        If Not objCALField Is Nothing Then strReceiveDT = objCALField.Value
        DoEvents
        
        If Not strReceiveDT > " " Then strReceiveDT = Format$(Now, "YYYY-MM-DD")
                
        Set objCALField = Nothing
        Set objCALField = objCALFields.Find("WFYEAR")   ' Get Year of filing
        If Not objCALField Is Nothing Then strFilingYear = objCALField.Value
        DoEvents
        
        Set objCALField = Nothing
        Set objCALField = objCALFields.Find("WFMATCH_ID")   ' Get Year of filing
        If Not objCALField Is Nothing Then strWorkitemID = objCALField.Value
        DoEvents
       
        If Not strWorkitemID > " " Then strErrMsg = "Match ID is blank. Please key in match-id"
        
        strErrMsg = vbNullString
        If Not strFileNumber > " " Then strErrMsg = "File number is blank. Routing to Annuals queue."
        
        If Not strReceiveDT > " " Then
            strErrMsg = "Receive Date is blank."
        ElseIf Not IsDate(strReceiveDT) Then
            strErrMsg = "Receive Date is not a valid date."
        ElseIf CDate(strReceiveDT) > Date Then
            strErrMsg = "Receive Date greater than current date."
        End If
        
        If Not strFilingYear > " " Then
            strErrMsg = "Filing Year is blank."
        ElseIf CInt(strFilingYear) > Year(Date) Then
            strErrMsg = "Filing Year is greater than current year."
        End If
               
        If strErrMsg > " " Then
        
            Set objCALField = Nothing
            Set objCALFields = Nothing
            If Not objCALDocument Is Nothing Then
                objCALDocument.Close calCloseAbortChanges
                DoEvents
                Set objCALDocument = Nothing
            End If
            WriteToLogFile strErrMsg
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
            DoEvents
            If PlaceWorkitemInError(objCALClientListitem, strErrMsg) <> icCriticalError Then
                ProcessResubWorkitem = icSuccess
            Else
                ProcessResubWorkitem = icCriticalError
            End If
            Exit Function
        End If
            
    Else            'If objCALFormFields is nothing, then do this
        If (objCALDocument.Info.Status And calObjStatusOpen) = calObjStatusOpen Then
            objCALDocument.Close calCloseAbortChanges
            DoEvents
            Sleep 1000
        End If
        Set objCALDocument = Nothing
        DoEvents
        Set objCALSendQueue = New CALQueue
        DoEvents
        With objCALSendQueue
            .Client = objCALClient
            .Name = objCALClientListitem.Info.SourceWorkstep
            .Type = calQueueTypeWorkstep
        End With
        DoEvents
        objCALClientListitem.Send objCALSendQueue, calSendDiscard
        DoEvents
        eErrorStatus = CheckCALError
        Set objCALSendQueue = Nothing
        Set objCALClientListitem = Nothing
        If eErrorStatus <> icCriticalError Then
            ProcessResubWorkitem = icSuccess
            Exit Function
        Else
            ProcessResubWorkitem = icCriticalError
            Exit Function
        End If
    End If
    
    
    ' A workitem is considered as a clean annual if its batch type is "Annuals-c", No changes are made to the annual
    ' and Signature found
    
    
    ' Call Oracle Stored Procedure.
    strAnnualType = "M"
    strErrMsg = vbNullString
    strErrCode = vbNullString
    strCompanyName = vbNullString
    
    GetProcessingInfoFromDB strFileNumber, CInt(strFilingYear), CDate(strReceiveDT), strAnnualType, _
                            strWorkitemID, strCompanyName, strErrCode, strErrMsg
    
        
    Select Case strErrCode
        
        Case Is = "00", "01", "11"            ' Brims return code is either "00" - Success or "01" - Business Rules Failed, No update done on brims side
        
        
            If Len(strCompanyName) > 250 Then
                Set objCALField = Nothing
                Set objCALField = objCALFields.Find("WFCOMPANY_NAME")
                If Not objCALField Is Nothing Then objCALField.Value = Left$(strCompanyName, 250)
                DoEvents
                Set objCALField = Nothing
                Set objCALField = objCALFields.Find("WFCOMPANY_NAME1")
                If Not objCALField Is Nothing Then objCALField.Value = Mid$(strCompanyName, 251)
            Else
                Set objCALField = Nothing
                Set objCALField = objCALFields.Find("WFCOMPANY_NAME")
                If Not objCALField Is Nothing Then objCALField.Value = strCompanyName
            End If
            strTemp = "Note on " & Format$(Now, "MM/DD/YYYY")
            If strErrMsg = vbNullString Then strErrMsg = "Updated BRIMS system."
            Set objCALNotes = objCALDocument.Notes
            objCALNotes.Add calNoteFirst, strTemp, strErrMsg
            DoEvents
            objCALDocument.Save
            DoEvents
            objCALDocument.Close calCloseSaveChanges
            DoEvents
            
            eErrorStatus = CheckCALError
            If eErrorStatus = icSuccess Then
                DoEvents
            ElseIf eErrorStatus = icCriticalError Then
                strErrMsg = "Critical error occurred. Shutting down."
                ProcessResubWorkitem = icCriticalError
                WriteToLogFile strErrMsg
                If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
                Set objCALClientListitem = Nothing
                Err.Clear
                Exit Function
            Else
                strErrMsg = "Unable to send the workitem to default workstep."
                If PlaceWorkitemInError(objCALClientListitem, strErrMsg) <> icCriticalError Then
                    ProcessResubWorkitem = icSuccess
                Else
                    ProcessResubWorkitem = icCriticalError
                    strErrMsg = "Critical error occurred. Shutting down."
                    WriteToLogFile strErrMsg
                    If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
                    Set objCALClientListitem = Nothing
                End If
                Err.Clear
                Exit Function
            End If
            
            objCALClientListitem.SendToDefault calSendDiscard
            eErrorStatus = CheckCALError
            If eErrorStatus <> icCriticalError Then
                ProcessResubWorkitem = icSuccess
                strErrMsg = strWorkitemID & " (resub) send to suspend queue for further processing."
            Else
                ProcessResubWorkitem = icCriticalError
                strErrMsg = "Critical error occurred. Shutting down."
            End If
            WriteToLogFile strErrMsg
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
            Set objCALClientListitem = Nothing
            
        Case "02"
            
            strErrCode = vbNullString
            Set objCALFields = objCALDocument.FormFields(calFieldsNoViews)
            DoEvents
            If Not objCALFields Is Nothing Then
                Set objCALField = Nothing                                   'Get the value in WFT8_1 and if the value is blank
                Set objCALField = objCALFields.Find("WFBARCODE_FOR_WORD")   ' then put ERR and send it back to the queue. If it is ERR then set it ANN and
                If Not objCALField Is Nothing Then                          ' route it to Annuals queue. Do not resend the document more than once.
                    strErrCode = objCALField.Value
                    strErrCode = IIf((strErrCode = "ERR"), "ANN", "ERR")
                    DoEvents
                    objCALField.Value = strErrCode
                    objCALDocument.Save
                    DoEvents
                    Sleep 1000
                    DoEvents
                End If

            End If
            
            Set objCALField = Nothing
            Set objCALFields = Nothing
            DoEvents
            objCALDocument.Close calCloseAbortChanges
            DoEvents
            Sleep 1000
            DoEvents
            Set objCALDocument = Nothing
            DoEvents
            DoEvents
            objCALClientListitem.SendToDefault calSendDiscard
            DoEvents
            eErrorStatus = CheckCALError
            If eErrorStatus <> icCriticalError Then
                ProcessResubWorkitem = icSuccess
                If strErrCode = "ERR" Then
                    strErrMsg = strWorkitemID & "- was not succesful. Sending back to the queue."
                Else
                    strErrMsg = strWorkitemID & "-Failed after 2 attempts. Routing to suspend queue."
                End If
            Else
                ProcessResubWorkitem = icCriticalError
                strErrMsg = "Critical error occurred. Shutting down."
            End If
            Set objCALClientListitem = Nothing
            WriteToLogFile strErrMsg
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg
            DoEvents
            
        Case Else
        

            If (objCALDocument.Info.Status And calObjStatusOpen) = calObjStatusOpen Then    'Critical error on BRIMS side, so logoff
                objCALDocument.Close calCloseAbortChanges
                DoEvents
                Sleep 1000
            End If
            Set objCALDocument = Nothing
            DoEvents
            With objCALSendQueue
                .Client = objCALClient
                .Name = objCALClientListitem.Info.SourceWorkstep
                .Type = calQueueTypeWorkstep
            End With
            DoEvents
            objCALClientListitem.Send objCALSendQueue, calSendDiscard
            DoEvents
            Set objCALClientListitem = Nothing
            Set objCALSendQueue = Nothing
            DoEvents
            ProcessResubWorkitem = icCriticalError
            strErrMsg = "Unknown error occured when trying to contact BRIMS. Shutting down."
            WriteToLogFile strErrMsg
            DoEvents
            If Not objMainForm Is Nothing Then objMainForm.AddItemToList strErrMsg

    End Select
            
    DoEvents
    Sleep 1000
    DoEvents
    Err.Clear
    
End Function






