VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3540
   ClientLeft      =   11370
   ClientTop       =   1050
   ClientWidth     =   7110
   ControlBox      =   0   'False
   Icon            =   "frmabout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   -15
      Picture         =   "frmabout.frx":0442
      ScaleHeight     =   3600
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   -30
      Width           =   7185
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6030
         MouseIcon       =   "frmabout.frx":9D57
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   2750
         Width           =   690
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   405
         Left            =   2670
         TabIndex        =   8
         Top             =   2300
         Width           =   4365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         Caption         =   "Copyright DCCA 2001"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5325
         TabIndex        =   6
         Top             =   3135
         Width           =   1740
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000012&
         Caption         =   "Business Registration "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   480
         Left            =   2505
         TabIndex        =   5
         Top             =   30
         Width           =   4680
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Height          =   900
         Left            =   5745
         TabIndex        =   4
         Top             =   2220
         Width           =   1410
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   390
         Left            =   4350
         TabIndex        =   3
         Top             =   1700
         Width           =   2715
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Annuals Robot"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   780
         Left            =   3450
         TabIndex        =   2
         Top             =   1035
         Width           =   3645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Height          =   405
         Left            =   4335
         TabIndex        =   1
         Top             =   615
         Width           =   2715
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdOK_Click()

    On Error Resume Next
    
    If Me.cmdOK.Visible Then EndPlaySound
    Unload Me
End Sub


Private Sub Form_Activate()
' Play sound only once.
    Static blnOnce As Boolean
    
    On Error Resume Next
    
    If Not blnOnce And Me.cmdOK.Visible Then
        BeginPlaySound 102
        blnOnce = True
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    On Error Resume Next
    
    If KeyAscii = 27 Then  'Escape key
        If Me.cmdOK.Visible Then EndPlaySound
        DoEvents
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    DoEvents
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & "      "
    DoEvents
    DoEvents
    
End Sub
