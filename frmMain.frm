VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Annuals Robot"
   ClientHeight    =   3825
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7905
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrStop 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3210
      Top             =   1665
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   3405
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5469
            MinWidth        =   5469
            Picture         =   "frmMain.frx":08CA
            Key             =   "Status"
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3916
            MinWidth        =   3916
            Key             =   "Workitem"
            Object.ToolTipText     =   "Current Workitem Number"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2240
            MinWidth        =   2240
            Key             =   "Clean"
            Object.ToolTipText     =   "Count of Clean Annuals Sent to Filing queue"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2240
            MinWidth        =   2240
            Key             =   "Manual"
            Object.ToolTipText     =   "Count of Manual and Clean sent to Annuals queue"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   7575
      Begin VB.ListBox lstMsg 
         Height          =   2595
         Left            =   135
         TabIndex        =   2
         Top             =   240
         Width           =   7320
      End
   End
   Begin VB.Timer tmrSleep 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6195
      Top             =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuStop 
         Caption         =   "S&top"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuStartNow 
         Caption         =   "Process &Now"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Properties"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuLogFile 
         Caption         =   "&View Current Log File"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "&End"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnUnloadRequested As Boolean
Private mblnStartApp As Boolean
Private mStrCaption As String
Private objDialogForm As frmDialog
Private intTimePassed As Integer         ' Variable which keeps track of minutes passed

Private Sub Form_Load()
'
' Load this form
'
    With Me
         mStrCaption = .Caption
         DoEvents
        .mnuStop.Enabled = False
        .Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
         DoEvents
        .StatusBar1.Panels(1).Picture = LoadResPicture(4, vbResIcon)
         DoEvents
         DoEvents
         DoEvents
        .StatusBar1.Panels(1).Text = "Processing. Please wait..."
         DoEvents
        .mnuEnd.Enabled = False
        .mnuStartNow.Enabled = False
         DoEvents
         DoEvents
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
' Check whether the application is processing workitems. If it is then, wait until the current
' workitem was done processing.
'
    If (UnloadMode = vbFormCode) Then
        DoEvents
    Else
        If blnUnloadRequested Then
            Cancel = 1
            Exit Sub
        End If
        If gblnRunning Then
            blnUnloadRequested = True
            Call mnuStop_Click
            Cancel = 1      ' Do not unload until the processing is done ie gblnrunning is false
        ElseIf gJustStarted Then
            Cancel = 1
        End If
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyGlobalObjects    'Destroy all global objects
    Set objFSO = Nothing
End Sub

Private Sub mnuAbout_Click()
'
' Show the About form
'
    Dim objAboutForm As frmAbout
    
    On Error Resume Next
    
    Set objAboutForm = New frmAbout
    Load objAboutForm
    objAboutForm.Show vbModal
    Set objAboutForm = Nothing
    DoEvents
    
End Sub

Private Sub mnuEnd_Click()
'
' Logoff and unload the form.
'
    LogoffWorkflow
    Unload Me
    
End Sub



Private Sub mnuLogFile_Click()
    
    Dim strLog As String
    
    On Error GoTo Log_Error
   
    If Not objFSO Is Nothing Then
        strLog = App.Path & "\Logs\" & gLogFileName
        If objFSO.FileExists(strLog) Then
            strLog = "Notepad.exe """ & strLog & """"
            Shell strLog, vbNormalNoFocus
        Else
            strLog = "Cannot find log file."
            MsgBox strLog, vbApplicationModal + vbCritical, App.EXEName
        End If
    End If
    Exit Sub

Log_Error:
    strLog = Err.Description & vbCrLf
    strLog = strLog & "Error No: " & Err.Number & " Error Src: View Log"
    MsgBox strLog, vbApplicationModal + vbInformation, App.EXEName
    Err.Clear
        
End Sub

Private Sub mnuProperties_Click()
'
' Show the options form.
'
    Dim objOptionsForm As frmOptions
    
    On Error Resume Next
    Me.Hide
    Set objOptionsForm = New frmOptions
    Load objOptionsForm
    objOptionsForm.Show vbModal
    DoEvents
    Me.Show
    
End Sub

Public Sub mnuStart_Click()
'
' Disable the start menu item and logon to workflow. If logon is successful, start processing workitems.
'
    On Error Resume Next
    
    With Me
        .tmrSleep.Enabled = False
        .mnuStartNow.Enabled = False
         DoEvents
        .mnuProperties.Enabled = False
        .mnuLogFile.Enabled = False
        .mnuStart.Enabled = False
        .mnuEnd.Enabled = False
         DoEvents
        .StatusBar1.Panels(1).Picture = LoadResPicture(4, vbResIcon)
         DoEvents
         DoEvents
        .StatusBar1.Panels(1).Text = "Processing. Please wait...."
         DoEvents
        .Caption = mStrCaption & " -Processing"
         DoEvents
         DoEvents
         .Refresh
         DoEvents
    End With
    
    DoEvents
    gcurWorkset = mstrWorksets(0)
    gblnStopRequested = False
    DoEvents
    
    If Not CheckSettingValues Then
        With Me
            .Caption = mStrCaption & " -Sleeping"
            .AddItemToList "Registry settings check failed. Please verify."
             DoEvents
             DoEvents
            .StatusBar1.Panels(1).Picture = LoadResPicture(2, vbResIcon)
             DoEvents
             DoEvents
            .StatusBar1.Panels(1).Text = "0 out of " & Trim$(Str$(mintSleepTime)) & " minutes passed."
            .StatusBar1.Panels(2).Text = vbNullString
            .mnuStart.Enabled = False
            .tmrSleep.Enabled = True
             DoEvents
            .mnuStartNow.Enabled = True
            .mnuLogFile.Enabled = True
            .mnuStop.Enabled = True
            .mnuProperties.Enabled = False
            .Refresh
            DoEvents
        End With
        Exit Sub
    End If

    If objCALClient Is Nothing Then
        If Not LogonToWorkflow Then
            With Me
                .Caption = mStrCaption & " -Sleeping"
                 DoEvents
                .StatusBar1.Panels(1).Picture = LoadResPicture(2, vbResIcon)
                 DoEvents
                 DoEvents
                .StatusBar1.Panels(1).Text = "0 out of " & Trim$(Str$(mintSleepTime)) & " minutes passed."
                 DoEvents
                .StatusBar1.Panels(2).Text = vbNullString
                .mnuStart.Enabled = False
                .tmrSleep.Enabled = True
                .mnuStartNow.Enabled = True
                 DoEvents
                .mnuStop.Enabled = True
                .mnuLogFile.Enabled = True
                .mnuProperties.Enabled = False
                .Refresh
                 DoEvents
            End With
            Exit Sub
        End If
        DoEvents
        DoEvents
    End If
        
    
    If Not CheckWorksetsNClasses Then
        With Me
            .Caption = mStrCaption & " -Sleeping"
            .AddItemToList "Worksets or Classes check failed. Please check Properties."
             DoEvents
             DoEvents
            .StatusBar1.Panels(1).Picture = LoadResPicture(2, vbResIcon)
             DoEvents
             DoEvents
            .StatusBar1.Panels(1).Text = "0 out of " & Trim$(Str$(mintSleepTime)) & " minutes passed."
             DoEvents
            .StatusBar1.Panels(2).Text = vbNullString
            .mnuStart.Enabled = False
            .tmrSleep.Enabled = True
            .mnuStop.Enabled = True
            .mnuLogFile.Enabled = True
            .mnuProperties.Enabled = False
            .mnuStartNow.Enabled = True
            .Refresh
             DoEvents
        End With
        Exit Sub
    End If
    
    Me.mnuStop.Enabled = True
    DoEvents
    ProcessAnnuals           ' Main routine which process the merging of images
    DoEvents
    
    If gblnStopRequested Then
        gblnStopRequested = False
        Me.Caption = mStrCaption & " -Stopped"
        DoEvents
    Else
        With Me
            .Caption = mStrCaption & " -Sleeping"
            .StatusBar1.Panels(1).Picture = LoadResPicture(2, vbResIcon)
             DoEvents
            .StatusBar1.Panels(1).Text = "0 out of " & Trim$(Str$(mintSleepTime)) & " minutes passed."
             DoEvents
            .StatusBar1.Panels(2).Text = vbNullString
            .mnuStart.Enabled = False
            .tmrSleep.Enabled = True
            .mnuStop.Enabled = True
            .mnuLogFile.Enabled = True
            .mnuProperties.Enabled = False
            .mnuStartNow.Enabled = True
             DoEvents
            .Refresh
             DoEvents
        End With
    End If
    
End Sub

Private Sub mnuStartNow_Click()
    
    DoEvents
    mnuStartNow.Enabled = False
    Call mnuStop_Click
    DoEvents
    Sleep 500
    DoEvents
    Sleep 500
    DoEvents
    Call mnuStart_Click
    
End Sub

Public Sub mnuStop_Click()
'
' Stop processing the resub folders. Disable stop menu and enable start menu.
'
    On Error Resume Next
    
    Dim strMsg As String
    
    With Me
        .MousePointer = vbCustom
         DoEvents
         intTimePassed = 0
         DoEvents
        .MouseIcon = LoadResPicture(101, vbResCursor)
         DoEvents
         DoEvents
         DoEvents
         DoEvents
        .mnuFile.Enabled = False
         DoEvents
        .mnuHelp.Enabled = False
         DoEvents
        .mnuStop.Enabled = False
         DoEvents
         gblnStopRequested = True
        .mnuLogFile.Enabled = True
         DoEvents
        .tmrSleep.Enabled = False
         DoEvents
        .mnuStartNow.Enabled = False
         DoEvents
        .StatusBar1.Panels(1).Picture = LoadResPicture(2, vbResIcon)
         DoEvents
        .StatusBar1.Panels(1).Text = "User Intervention - Stopping"
         DoEvents
        .Caption = mStrCaption & " -Stopping"
         DoEvents
         DoEvents
        .tmrStop.Enabled = True
         DoEvents
         If gblnRunning Then ShowFormDialog "Shutting down. Please wait...."
    End With
    DoEvents
    Me.StatusBar1.Panels(1).Picture = LoadResPicture(3, vbResIcon)
    Me.Refresh
    DoEvents
    
End Sub

Private Sub tmrSleep_Timer()
'
' If the time past is equal to the sleep time then call mnuStart procedure, else keep ticking
'
    On Error Resume Next
    
    

    If intTimePassed >= mintSleepTime - 1 Then
        intTimePassed = 0
        With Me
            .tmrSleep.Enabled = False
            .mnuStartNow.Enabled = False
             DoEvents
            .mnuStart_Click
        End With
    Else
        intTimePassed = intTimePassed + 1
        Me.StatusBar1.Panels(1).Text = Trim$(Str$(intTimePassed)) & " out of " & Trim$(Str$(mintSleepTime)) & " minutes passed."
        DoEvents
    End If
    
End Sub

Private Sub tmrStop_Timer()
'
' This timer is used to display the status of stopping of the Application. The timer interval is
' set at one second and is enabled immediately after the stop button is pressed, and every 1 second it will
' whether the processing of merging pages is done for the workitem that is in hand currently (gblnrunning will be set
' to false in the ProcessAnnuals procedure). If it was processed ie glbnrunning is set to false then procedure
' will close the dialog form and enable the file menu.


    On Error Resume Next
    
    Dim strMsg As String
    
    If mblnStartApp Then            'This is run only once.
        With Me
            mblnStartApp = False
            DoEvents
           .tmrStop.Enabled = False
            DoEvents
           .mnuStart_Click
        End With
        Exit Sub
    End If                          ' Till here.
    
    If Not gblnRunning Then
        With Me
            .mnuFile.Enabled = True
             DoEvents
            .mnuHelp.Enabled = True
             DoEvents
            .tmrStop.Enabled = False
             DoEvents
            .StatusBar1.Panels(1).Picture = LoadResPicture(3, vbResIcon)
             DoEvents
            .StatusBar1.Panels(1).Text = "User Intervention - Stopped"
             DoEvents
             strMsg = App.EXEName & " -User Intervention"
             DoEvents
            .AddItemToList strMsg
             DoEvents
            .StatusBar1.Panels(2).Text = vbNullString
            .Caption = mStrCaption & " -Stopped"
             DoEvents
            .mnuStart.Enabled = True
             DoEvents
            .mnuEnd.Enabled = True
             DoEvents
            .mnuProperties.Enabled = True
             DoEvents
            .MousePointer = vbDefault
             DoEvents
            .Refresh
        End With
        
        DoEvents
        If Not objDialogForm Is Nothing Then
            Unload objDialogForm
            DoEvents
            Set objDialogForm = Nothing
        End If
        
        DoEvents
        If blnUnloadRequested Then
            blnUnloadRequested = False
            DoEvents
            Unload Me
        End If
        
    End If
    DoEvents
    Sleep 1000
    
End Sub

Public Sub AddItemToList(ByVal strMsg As String)
'
' Display the respective message on the screen. If the list box contains more than 50 entries, then clear all
' of the messages except the last two.
'

    On Error Resume Next
    
    Dim strTemp1 As String
    Dim strTemp2 As String
    
    strMsg = FormatDateTime(Now, vbLongTime) & "-" & strMsg
    DoEvents
    If Me.lstMsg.ListCount = 50 Then
        strTemp1 = Me.lstMsg.List(Me.lstMsg.ListCount - 2)
        strTemp2 = Me.lstMsg.List(Me.lstMsg.ListCount - 1)
        Me.lstMsg.Clear
        DoEvents
        DoEvents
        DoEvents
        Me.lstMsg.AddItem strTemp1
        DoEvents
        DoEvents
        Me.lstMsg.AddItem strTemp2
        DoEvents
        DoEvents
    End If
    Me.lstMsg.AddItem strMsg
    DoEvents
    DoEvents
    DoEvents
    Me.lstMsg.ListIndex = Me.lstMsg.ListCount - 1
    DoEvents
    DoEvents
    
End Sub

Public Sub StartApp()
'
' This is used to start the application as soon as the program is launched. This sub is used only once as soon
' the app is launched.
'
    Dim strMsg As String
    
    On Error Resume Next
    
    mblnStartApp = True
    With Me
         DoEvents
        .StatusBar1.Panels(1).Picture = LoadResPicture(4, vbResIcon)
         DoEvents
         DoEvents
         strMsg = "Logged onto domain " & mstrDomain & "."
         DoEvents
        .AddItemToList strMsg
         DoEvents
         DoEvents
        .StatusBar1.Panels(1).Text = "Processing. Please wait..."
         DoEvents
        .tmrStop.Enabled = True           ' Using the same timer for starting purposes also.
        .Refresh
    End With
    
End Sub

Private Sub ShowFormDialog(strMsg As String)
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

Public Property Let UnloadRequested(blnValue As Boolean)
    blnUnloadRequested = blnValue
End Property

