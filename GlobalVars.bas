Attribute VB_Name = "GlobalVars"
Option Explicit
'
' Global Variables
'
' Intrinsic Variables.
Public mstrUserID As String
Public mstrPassword As String
Public mstrDomain As String
Public mstrDatabase As String
Public mintWorksetCnt As Integer
Public mstrWorksets() As String
Public mintDocCLassCount As Integer
Public mintFoldCLassCount As Integer
Public mstrDocClass() As String
Public mstrFoldClass() As String
Public gLogFileName As String
Public mintSleepTime As Integer
Public gcurWorkset As String
Public gblnRunning As Boolean
Public gJustStarted As Boolean
Public gblnStopRequested As Boolean
Public gstrEndtime As String
Public gstrRundays As String
Public gManualCnt As Long
Public gCleanCnt As Long
Public gblnGPLLP As Boolean
Public gblnFolderGPLLP As Boolean
Public eCALStatus As CALProcessStatus

' Object Variables
Public objMainForm As frmMain
Public objDialogForm As frmDialog
Public objFSO As FileSystemObject
Public objCurFile As TextStream
Public objADOConnection As ADODB.Connection

'
' Global Workflow variables
'
Public objCALMaster As New CALMaster
Public objCALClient As CALClient
Public objCALClientList As CALClientList

Public Enum CALProcessStatus
    icSuccess
    icVerifyOK
    icVerifyNotOK
    icInvalidDomain
    icInvalidWorkset
    icQueueEmpty
    icWorkitemInList
    icWorkitemNotInList
    icWorkitemOpen
    icWorkitemisWIP
    icPlaceErrorFailed
    icSaveFailed
    icSetPageError
    icCloseFailed
    icInvalidQueue
    icWorkitemNew
    icCriticalError
    icTryAgain
End Enum

'
' *******************************************************************************************
' All API functions are declared here.
' *******************************************************************************************
'
' Sleep API Function.
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
' High level sound support API
Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (lpszSoundName As Any, ByVal uFlags As Long) As Long

