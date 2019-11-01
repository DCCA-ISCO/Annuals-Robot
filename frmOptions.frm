VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOptions 
   Caption         =   "Annuals Robot Settings"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameOpt 
      Caption         =   "Preferences"
      Height          =   3120
      Index           =   3
      Left            =   375
      TabIndex        =   37
      Top             =   650
      Width           =   6000
      Begin VB.Frame Frame1 
         Caption         =   "Run Days"
         Height          =   1350
         Left            =   240
         TabIndex        =   44
         Top             =   1500
         Width           =   5565
         Begin VB.CheckBox chkRunDay 
            Caption         =   "Saturday"
            Height          =   300
            Index           =   7
            Left            =   4300
            TabIndex        =   51
            Top             =   300
            Value           =   1  'Checked
            Width           =   1000
         End
         Begin VB.CheckBox chkRunDay 
            Caption         =   "Friday"
            Height          =   300
            Index           =   6
            Left            =   2900
            TabIndex        =   50
            Top             =   800
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkRunDay 
            Caption         =   "Thursday"
            Height          =   300
            Index           =   5
            Left            =   2900
            TabIndex        =   49
            Top             =   300
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkRunDay 
            Caption         =   "Wednesday"
            Height          =   300
            Index           =   4
            Left            =   1500
            TabIndex        =   48
            Top             =   800
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkRunDay 
            Caption         =   "Tuesday"
            Height          =   300
            Index           =   3
            Left            =   1500
            TabIndex        =   47
            Top             =   300
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkRunDay 
            Caption         =   "Monday"
            Height          =   300
            Index           =   2
            Left            =   100
            TabIndex        =   46
            Top             =   800
            Value           =   1  'Checked
            Width           =   1200
         End
         Begin VB.CheckBox chkRunDay 
            Caption         =   "Sunday"
            Height          =   300
            Index           =   1
            Left            =   100
            TabIndex        =   45
            Top             =   300
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   225
         Picture         =   "frmOptions.frx":08CA
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   39
         Top             =   360
         Width           =   500
      End
      Begin MSComCtl2.DTPicker DTPEndTime 
         Height          =   315
         Left            =   980
         TabIndex        =   38
         Top             =   1020
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   22806530
         CurrentDate     =   37284.75
      End
      Begin VB.Label lblEndTime 
         Alignment       =   1  'Right Justify
         Caption         =   "End Time:"
         Height          =   270
         Left            =   180
         TabIndex        =   41
         Top             =   1050
         Width           =   760
      End
      Begin VB.Label Label13 
         Caption         =   $"frmOptions.frx":0D0C
         Height          =   600
         Left            =   870
         TabIndex        =   40
         Top             =   330
         Width           =   4650
      End
   End
   Begin VB.Frame frameOpt 
      Caption         =   "Classes"
      Height          =   3120
      Index           =   2
      Left            =   -10000
      TabIndex        =   31
      Top             =   650
      Visible         =   0   'False
      Width           =   6000
      Begin VB.ListBox lstFoldClasses 
         Height          =   1035
         Left            =   2550
         TabIndex        =   17
         Top             =   1980
         Width           =   2280
      End
      Begin VB.TextBox txtFoldClass 
         Height          =   325
         Left            =   2550
         TabIndex        =   13
         ToolTipText     =   "Enter Original workitem class"
         Top             =   1305
         Width           =   2280
      End
      Begin VB.CommandButton cmdClassesRemove 
         Caption         =   "Remove"
         Height          =   390
         Left            =   4960
         TabIndex        =   15
         Top             =   1800
         Width           =   850
      End
      Begin VB.CommandButton cmdClassesAdd 
         Caption         =   "Add"
         Height          =   390
         Left            =   4960
         TabIndex        =   14
         Top             =   1300
         Width           =   850
      End
      Begin VB.ListBox lstDocClasses 
         Height          =   1035
         Left            =   195
         TabIndex        =   16
         Top             =   1980
         Width           =   2280
      End
      Begin VB.PictureBox picIcon1 
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   125
         Picture         =   "frmOptions.frx":0DAD
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   32
         Top             =   250
         Width           =   500
      End
      Begin VB.TextBox txtDocClass 
         Height          =   325
         Left            =   195
         TabIndex        =   12
         ToolTipText     =   "Enter Original workitem class"
         Top             =   1305
         Width           =   2280
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Existing Folder classes"
         Height          =   285
         Left            =   2550
         TabIndex        =   43
         Top             =   1700
         Width           =   2280
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "New Folder Class"
         Height          =   270
         Left            =   2550
         TabIndex        =   42
         Top             =   1005
         Width           =   2280
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Existing Document Classes"
         Height          =   285
         Left            =   195
         TabIndex        =   35
         Top             =   1700
         Width           =   2280
      End
      Begin VB.Label lblDocClass 
         Alignment       =   1  'Right Justify
         Caption         =   "New Document Class"
         Height          =   270
         Left            =   360
         TabIndex        =   34
         Top             =   1005
         Width           =   1710
      End
      Begin VB.Label lblClassesMsg 
         Caption         =   $"frmOptions.frx":11EF
         Height          =   585
         Left            =   735
         TabIndex        =   33
         Top             =   240
         Width           =   5025
      End
   End
   Begin VB.Frame frameOpt 
      Caption         =   "Worksets"
      Height          =   3120
      Index           =   1
      Left            =   -10000
      TabIndex        =   26
      Top             =   650
      Width           =   6000
      Begin VB.TextBox txtNewWorkset 
         Height          =   325
         Left            =   1995
         TabIndex        =   8
         Top             =   1065
         Width           =   2460
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   225
         Picture         =   "frmOptions.frx":1286
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   27
         Top             =   450
         Width           =   480
      End
      Begin VB.ListBox lstWorksets 
         Height          =   1230
         Left            =   1995
         TabIndex        =   11
         Top             =   1500
         Width           =   2460
      End
      Begin VB.CommandButton cmdAddWorkset 
         Caption         =   "Add"
         Height          =   375
         Left            =   4875
         TabIndex        =   9
         Top             =   1050
         Width           =   870
      End
      Begin VB.CommandButton cmdRemoveWorkset 
         Caption         =   "Remove"
         Height          =   375
         Left            =   4875
         TabIndex        =   10
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label lblAddWorkset 
         Caption         =   $"frmOptions.frx":16C8
         Height          =   630
         Left            =   915
         TabIndex        =   30
         Top             =   315
         Width           =   4650
      End
      Begin VB.Label lblExistWorkset 
         Alignment       =   1  'Right Justify
         Caption         =   "Exisiting Worksets:"
         Height          =   270
         Left            =   270
         TabIndex        =   29
         Top             =   1515
         Width           =   1650
      End
      Begin VB.Label lblNewWorkset 
         Alignment       =   1  'Right Justify
         Caption         =   "New Workset:"
         Height          =   270
         Left            =   270
         TabIndex        =   28
         Top             =   1110
         Width           =   1650
      End
   End
   Begin VB.Frame frameOpt 
      Caption         =   "Logon"
      Height          =   3120
      Index           =   0
      Left            =   -10000
      TabIndex        =   19
      Top             =   650
      Width           =   6000
      Begin VB.TextBox txtODBC 
         Height          =   325
         Left            =   4260
         TabIndex        =   6
         Top             =   1755
         Width           =   1600
      End
      Begin VB.TextBox txtUserID 
         Height          =   325
         Left            =   1260
         TabIndex        =   3
         Top             =   1260
         Width           =   1600
      End
      Begin VB.TextBox txtPassword 
         Height          =   325
         IMEMode         =   3  'DISABLE
         Left            =   1260
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1755
         Width           =   1600
      End
      Begin VB.TextBox txtDomain 
         Height          =   325
         Left            =   4260
         TabIndex        =   5
         Top             =   1260
         Width           =   1600
      End
      Begin VB.PictureBox picIcon 
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   225
         Picture         =   "frmOptions.frx":175A
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   20
         Top             =   450
         Width           =   500
      End
      Begin VB.ComboBox cmbSleepTime 
         Height          =   315
         ItemData        =   "frmOptions.frx":1B9C
         Left            =   4260
         List            =   "frmOptions.frx":1BBE
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2430
         Width           =   1600
      End
      Begin VB.Label lblDataBase 
         Alignment       =   1  'Right Justify
         Caption         =   "Database SID:"
         Height          =   270
         Left            =   3120
         TabIndex        =   36
         Top             =   1785
         Width           =   1095
      End
      Begin VB.Label lblUserID 
         Alignment       =   1  'Right Justify
         Caption         =   "User ID:"
         Height          =   270
         Left            =   295
         TabIndex        =   25
         Top             =   1290
         Width           =   915
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   270
         Left            =   295
         TabIndex        =   24
         Top             =   1785
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Domain:"
         Height          =   270
         Left            =   3330
         TabIndex        =   23
         Top             =   1290
         Width           =   885
      End
      Begin VB.Label lblLogonMsg 
         Caption         =   $"frmOptions.frx":1BF1
         Height          =   645
         Left            =   1020
         TabIndex        =   22
         Top             =   375
         Width           =   4650
      End
      Begin VB.Label lblSleepTime 
         Alignment       =   1  'Right Justify
         Caption         =   "Sleep Time (Mins) :"
         Height          =   270
         Left            =   2610
         TabIndex        =   21
         Top             =   2505
         Width           =   1470
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   3795
      Left            =   90
      TabIndex        =   18
      Top             =   225
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   6694
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Logon"
            Key             =   "Logon"
            Object.ToolTipText     =   "Logon Properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Worksets"
            Key             =   "Worksets"
            Object.ToolTipText     =   "Worksets"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Classes"
            Key             =   "Classes"
            Object.ToolTipText     =   "Class Combination"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Preferences"
            Key             =   "Preferences"
            Object.ToolTipText     =   "Preferences"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2745
      TabIndex        =   0
      Top             =   4100
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   4155
      TabIndex        =   1
      Top             =   4100
      Width           =   1125
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   390
      Left            =   5490
      TabIndex        =   2
      Top             =   4100
      Width           =   1125
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ErrMsg As String

Private Sub cmdAddWorkset_Click()
'
' Check whether the one that is being added already exists in the list or not.
' If not, then add, else error message

    Dim intCnt As Integer
    
    On Error Resume Next
        
    For intCnt = 0 To lstWorksets.ListCount - 1
        If StrComp(txtNewWorkset.Text, lstWorksets.List(intCnt), vbBinaryCompare) = 0 Then
            ErrMsg = "Cannot add this workset. This workset is already included."
            MsgBox ErrMsg, vbApplicationModal + vbExclamation, App.EXEName
            cmdAddWorkset.Enabled = False
            Exit Sub
        End If
    Next
        
    ErrMsg = "Are you sure, you want to add this workset?"
    If MsgBox(ErrMsg, vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, App.EXEName) = vbYes Then
        lstWorksets.AddItem txtNewWorkset.Text
        DoEvents
        cmdAddWorkset.Enabled = False
        txtNewWorkset.Text = vbNullString
    End If
    
End Sub

Private Sub cmdApply_Click()
'
' Check the values in the options form and then save the settings if ok.
'
   If CheckSettings Then SaveRegistrySettings
   
End Sub

Private Sub cmdCancel_Click()
'
' Unload the form without saving any values
'
    Unload Me
End Sub

Private Sub cmdClassesAdd_Click()
'
' Confirm the addition and then add the combination to the list view control/
'
    On Error GoTo Add_Error
    
    Dim strText As String
    
    If txtDocClass.Text > " " Then
        ErrMsg = "Do you want to add this document class?"
        If MsgBox(ErrMsg, vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, App.EXEName) = vbYes Then
        lstDocClasses.AddItem UCase$(txtDocClass.Text)
            txtDocClass.Text = vbNullString
            Me.Refresh
        End If
    End If
    If txtFoldClass.Text > " " Then
        ErrMsg = "Do you want to add this folder class?"
        If MsgBox(ErrMsg, vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, App.EXEName) = vbYes Then
        lstFoldClasses.AddItem UCase$(txtFoldClass.Text)
            txtFoldClass.Text = vbNullString
            Me.Refresh
        End If
    End If
    DoEvents
    Exit Sub

Add_Error:

    Err.Source = "ListBox Add"
    ErrMsg = Err.Description & vbCrLf
    ErrMsg = ErrMsg & "Error No: " & Err.Number & " Error Src: " & Err.Source
    MsgBox ErrMsg, vbApplicationModal + vbExclamation, App.EXEName
    DoEvents
    
End Sub

Private Sub cmdClassesRemove_Click()
'
' Confirm with the user and remove the combination.
'
    Dim objListItem As MSComctlLib.ListItem
    Dim intListIndex As Integer
    
    On Error GoTo Remove_Error
    
    If lstDocClasses.ListIndex = -1 And lstFoldClasses.ListIndex = -1 Then
        MsgBox "Please select a document/fodler class and then click on remove", vbApplicationModal + vbInformation, App.EXEName
        Exit Sub
    End If
    
    If lstDocClasses.ListIndex <> -1 Then
        ErrMsg = "Are you sure, you want to remove this document class?."
        If MsgBox(ErrMsg, vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, App.EXEName) = vbYes Then
            lstDocClasses.RemoveItem lstDocClasses.ListIndex
        End If
        Me.Refresh
    End If
    
    If lstFoldClasses.ListIndex <> -1 Then
        ErrMsg = "Are you sure, you want to remove this folder class?."
        If MsgBox(ErrMsg, vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, App.EXEName) = vbYes Then
            lstFoldClasses.RemoveItem lstFoldClasses.ListIndex
        End If
        Me.Refresh
    End If
    
    Exit Sub
    
Remove_Error:
    DoEvents
    Err.Source = "ListClass Remove"
    ErrMsg = Err.Description & vbCrLf
    ErrMsg = ErrMsg & "Error No: " & Err.Number & " Error Src: " & Err.Source
    MsgBox ErrMsg, vbApplicationModal + vbExclamation, App.EXEName
    DoEvents
    
End Sub

Private Sub CmdOK_Click()
'
' Save settings and close
'
    If CheckSettings Then
        SaveRegistrySettings
        Unload Me
    End If
    
End Sub

Private Sub cmdRemoveWorkset_Click()
    On Error Resume Next
    
    If lstWorksets.ListIndex <> -1 Then
        ErrMsg = "Are you sure, you want to remove this workset?."
        If MsgBox(ErrMsg, vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton2, App.EXEName) = vbYes Then
            lstWorksets.RemoveItem lstWorksets.ListIndex
            DoEvents
            cmdRemoveWorkset.Enabled = False
        End If
    End If
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
' Navigate to each tab, if CTRL + Tab is pressed.

    Dim intTabIndex As Integer
    
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
    
        intTabIndex = tbsOptions.SelectedItem.Index
        If intTabIndex = tbsOptions.Tabs.Count Then
            ' Last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            ' Increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(intTabIndex + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
'
' Enable first frame and disable the others
'
    Dim intTabIndex As Integer
    
    EnableControls
    If Not LoadRegistrySettings Then
        ErrMsg = "Fatal: Could not load settings from registry." & vbCrLf
        ErrMsg = ErrMsg & "Cannot continue."
        MsgBox ErrMsg, vbApplicationModal + vbCritical, App.EXEName
        DoEvents
        End
    End If
       
    For intTabIndex = 0 To tbsOptions.Tabs.Count - 1
        If intTabIndex = 0 Then
            frameOpt(intTabIndex).Visible = True
            frameOpt(intTabIndex).Enabled = True
            frameOpt(intTabIndex).Left = 375
        Else
            frameOpt(intTabIndex).Left = 375
            frameOpt(intTabIndex).Visible = False
            frameOpt(intTabIndex).Enabled = False
        End If
    Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Beep
        Cancel = 1      'do not close
    End If
End Sub

Private Sub lstWorksets_Click()
    cmdRemoveWorkset.Enabled = True
End Sub


Private Sub tbsOptions_Click()
'
' Show and enable the corresponding frame, based on the tab selected. Hide other frames.
'
    Dim intTabIndex As Integer
    
    On Error GoTo OptionsClick_Error
    
    For intTabIndex = 0 To tbsOptions.Tabs.Count - 1
        If intTabIndex = tbsOptions.SelectedItem.Index - 1 Then
            frameOpt(intTabIndex).Visible = True
            frameOpt(intTabIndex).Enabled = True
        Else
            frameOpt(intTabIndex).Visible = False
            frameOpt(intTabIndex).Enabled = False
        End If
    Next
    Exit Sub
    
OptionsClick_Error:
    
    DoEvents
    Err.Source = "OptionsClick"
    ErrMsg = Err.Description & vbCrLf
    ErrMsg = ErrMsg & "Error No: " & Err.Number & " Error Src: " & Err.Source
    MsgBox ErrMsg, vbApplicationModal + vbInformation, App.EXEName
    DoEvents
    
End Sub

Private Function LoadRegistrySettings() As Boolean
    
'
' Get required values from registry.
'
    Dim intCounter As Integer
    Dim strRunday As String
    
    On Error GoTo LoadSettings_Error
    

    LoadRegistrySettings = False
    If Not GetRegistrySettings Then Exit Function
    
' Load Logon settings

    txtUserID.Text = mstrUserID
    txtPassword.Text = mstrPassword
    txtDomain.Text = mstrDomain
    txtODBC.Text = mstrDatabase
    cmbSleepTime.Text = mintSleepTime
    
'Load preferences

    DTPEndTime.Value = FormatDateTime$(gstrEndtime, vbLongTime)
    DoEvents

' Load Workset settings

    For intCounter = 0 To mintWorksetCnt - 1
        lstWorksets.AddItem mstrWorksets(intCounter)
    Next
    
' Load document classes into the list view

    For intCounter = 0 To mintDocCLassCount - 1
        lstDocClasses.AddItem mstrDocClass(intCounter)
    Next
    
' Load folder classes into the list view

    For intCounter = 0 To mintFoldCLassCount - 1
        lstFoldClasses.AddItem mstrFoldClass(intCounter)
    Next

'Load run days.

    For intCounter = 1 To 7
        strRunday = Trim$(Str$(intCounter))
        If InStr(1, gstrRundays, strRunday, vbTextCompare) > 0 Then
            chkRunDay(intCounter).Value = 1
        Else
            chkRunDay(intCounter).Value = 0
        End If
    Next

    Me.Refresh

    LoadRegistrySettings = True
    Exit Function

LoadSettings_Error:
    Err.Source = "LoadSettings"
    ErrMsg = Err.Description & vbCrLf
    ErrMsg = ErrMsg & "Error No: " & Err.Number & " Error Src: " & Err.Source
    MsgBox ErrMsg, vbApplicationModal + vbCritical, App.EXEName
    DoEvents
    
End Function


Private Sub SaveRegistrySettings()
    
'
' Save the options into the registry. Since we are using SaveSetting command, the values will be stored in
' My Computer\HKEY_CURRENT_USER\Software\VB and VBA Program Settings\AnnRobot. Can be saved in some other hive
' using API
'
    Dim strAppName As String
    Dim strSectionName As String
    Dim strKeyName As String
    Dim intCounter As Integer
    Dim strSettingValue As String
    Dim objListItem As MSComctlLib.ListItem
    
    
    On Error Resume Next
    
' Applicatio Name.

    strAppName = App.EXEName
    
' Save userid
    strSectionName = "Logon"
    
    strKeyName = "UserID"
    SaveSetting strAppName, strSectionName, strKeyName, mstrUserID
    
' Save Password

    strKeyName = "Password"
    EncryptDecrypt mstrPassword             'Encrypt the password.
    SaveSetting strAppName, strSectionName, strKeyName, mstrPassword
    
' Save Domain.

    strKeyName = "Domain"
    SaveSetting strAppName, strSectionName, strKeyName, mstrDomain
    DoEvents
    
' Save the Database SID

    strKeyName = "Database"
    SaveSetting strAppName, strSectionName, strKeyName, mstrDatabase
    DoEvents

' Save Sleep Time.

    strKeyName = "SleepTime"
    SaveSetting strAppName, strSectionName, strKeyName, mintSleepTime
    DoEvents
    
    strSectionName = "Worksets"
    
' Delete the old one, if it exists. Runtime error occurs if no key exists but resume next takes care of it.

    DeleteSetting strAppName, strSectionName
    
' Save the count of worksets.

    strKeyName = "ListCount"
    strSettingValue = Trim$(Str$(lstWorksets.ListCount))
    SaveSetting strAppName, strSectionName, strKeyName, strSettingValue
    
' Now save each workset name.

    For intCounter = 1 To lstWorksets.ListCount
        strKeyName = "Workset" & Trim$(Str$(intCounter))
        strSettingValue = Trim$(lstWorksets.List(intCounter - 1))
        SaveSetting strAppName, strSectionName, strKeyName, strSettingValue
        DoEvents
    Next

' Save each document class entered in the Class tab. first save the count and then the each class

    strSectionName = "Classes"
    
'Delete the old one, if it exists. Runtime error occurs if no key exists but resume next takes care of it.

    DeleteSetting strAppName, strSectionName
     
'Save the total document classes count first.
   
    strKeyName = "DocClassCount"
    strSettingValue = Trim$(Str$(lstDocClasses.ListCount))
    SaveSetting strAppName, strSectionName, strKeyName, strSettingValue
    
'Now save each document class as DocClass1, DocClass2 etc..

    For intCounter = 1 To lstDocClasses.ListCount
        strKeyName = "DocClass" & Trim$(Str$(intCounter))
        strSettingValue = Trim$(lstDocClasses.List(intCounter - 1))
        SaveSetting strAppName, strSectionName, strKeyName, strSettingValue
        DoEvents
    Next
    
'Save the total folder classes count.
   
    strKeyName = "FoldClassCount"
    strSettingValue = Trim$(Str$(lstFoldClasses.ListCount))
    SaveSetting strAppName, strSectionName, strKeyName, strSettingValue
    
'Now save each fodler class as FoldClass1, FoldClass2 etc..

    For intCounter = 1 To lstFoldClasses.ListCount
        strKeyName = "FoldClass" & Trim$(Str$(intCounter))
        strSettingValue = Trim$(lstFoldClasses.List(intCounter - 1))
        SaveSetting strAppName, strSectionName, strKeyName, strSettingValue
        DoEvents
    Next
    
'Save the preferences
    
    gstrEndtime = FormatDateTime$(DTPEndTime.Value, vbLongTime)
    
'Section Name
    strSectionName = "Preferences"
    
'Key Name EndTime
    strKeyName = "EndTime"
    SaveSetting App.EXEName, strSectionName, strKeyName, gstrEndtime
    DoEvents
    
    
'Save the run days.
    gstrRundays = vbNullString
    For intCounter = 1 To 7
        If chkRunDay(intCounter).Value = 1 Then     'Checked
            gstrRundays = gstrRundays & Trim$(Str$(intCounter))
        End If
    Next
    
    If gstrRundays = vbNullString Then gstrRundays = "234567"   ' Default to run on mon,tue,wed,thu,fri and sat
    
'Key Name EndTime
    strKeyName = "RunDays"
    SaveSetting App.EXEName, strSectionName, strKeyName, gstrRundays
    DoEvents
    
    
    Err.Clear
    DoEvents

    
End Sub


Private Function CheckSettings() As Boolean
'
' Check the values entered on the form.
'
    Dim strRundays As String
    Dim intCounter As Integer
    
    CheckSettings = False
    
    If txtUserID.Text > " " Then
        mstrUserID = txtUserID.Text
    Else
        ErrMsg = "Please enter user-id"
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        DoEvents
        txtUserID.SetFocus
        DoEvents
        MsgBox ErrMsg, vbApplicationModal + vbInformation, App.EXEName
        Exit Function
    End If
    
    If txtPassword.Text > " " Then
        mstrPassword = txtPassword.Text
    Else
        ErrMsg = "Please enter password"
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        DoEvents
        txtPassword.SetFocus
        DoEvents
        MsgBox ErrMsg, vbApplicationModal + vbInformation, App.EXEName
        Exit Function
    End If
    
    If txtDomain.Text > " " Then
        mstrDomain = UCase$(txtDomain.Text)
    Else
        ErrMsg = "Please enter Domain"
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        DoEvents
        txtDomain.SetFocus
        DoEvents
        MsgBox ErrMsg, vbApplicationModal + vbInformation, App.EXEName
        Exit Function
    End If

    If txtODBC.Text > " " Then
        mstrDatabase = UCase$(txtODBC.Text)
    Else
        ErrMsg = "Please enter the Database name (SID)"
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        DoEvents
        txtODBC.SetFocus
        DoEvents
        MsgBox ErrMsg, vbApplicationModal + vbInformation, App.EXEName
        Exit Function
    End If

    If CInt(cmbSleepTime.Text) = 0 Then
        mintSleepTime = 30          ' Default to 90 min.
    Else
        mintSleepTime = CInt(cmbSleepTime.Text)
    End If
    
    
    If lstWorksets.ListCount < 1 Then
        ErrMsg = "You need to specify, atleast one workset."
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(2)
        DoEvents
        txtNewWorkset.SetFocus
        DoEvents
        MsgBox ErrMsg, vbApplicationModal + vbInformation, App.EXEName
        Exit Function
    End If
    
    If lstDocClasses.ListCount < 1 Then 'Atleast one document class has to be specified.
        ErrMsg = "You need to specify, atleast one document class in the classes tab" '& vbCrLf
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(3)
        DoEvents
        txtDocClass.SetFocus
        DoEvents
        MsgBox ErrMsg, vbApplicationModal + vbInformation, App.EXEName
        DoEvents
        Exit Function
    End If
    
    'Check the run days.
    strRundays = vbNullString
    For intCounter = 1 To 7
        If chkRunDay(intCounter).Value = 1 Then     'Checked
            strRundays = strRundays & Trim$(Str$(intCounter))
        End If
    Next
    
    If strRundays = vbNullString Then
        ErrMsg = "No run days specified. Atleast one run day has to be specified."
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(4)
        DoEvents
        MsgBox ErrMsg, vbApplicationModal + vbInformation, App.EXEName
        DoEvents
        Exit Function
    End If
    
    CheckSettings = True
    
End Function

Private Sub txtDocClass_Change()
    If txtDocClass.Text > " " Then
        cmdClassesAdd.Enabled = True
    Else
        cmdClassesAdd.Enabled = False
    End If
End Sub

Private Sub txtFoldClass_Change()
    If txtFoldClass.Text > " " Then
        cmdClassesAdd.Enabled = True
    Else
        cmdClassesAdd.Enabled = False
    End If
End Sub

Private Sub txtNewWorkset_Change()

    If txtNewWorkset.Text > " " Then
        cmdAddWorkset.Enabled = True
    Else
        cmdAddWorkset.Enabled = False
    End If
    
End Sub

Private Sub EnableControls()

    On Error Resume Next
    
    cmdAddWorkset.Enabled = False
    cmdRemoveWorkset.Enabled = False
    cmdClassesAdd.Enabled = False
    cmbSleepTime.ListIndex = 1
    
End Sub


