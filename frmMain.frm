VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "HardDisk Serial Number Example"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   2880
      TabIndex        =   16
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CheckBox chkSysFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "Entire Volume Compressed?"
      Height          =   255
      Index           =   5
      Left            =   1860
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CheckBox chkSysFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "Indiv File Compression?"
      Height          =   255
      Index           =   4
      Left            =   1860
      TabIndex        =   14
      Top             =   2220
      Width           =   2295
   End
   Begin VB.CheckBox chkSysFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "Access Control Support?"
      Height          =   255
      Index           =   3
      Left            =   1860
      TabIndex        =   13
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CheckBox chkSysFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "UniCode Filenames?"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   2520
      Width           =   1755
   End
   Begin VB.CheckBox chkSysFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "Case Sensitive?"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   2220
      Width           =   1755
   End
   Begin VB.CheckBox chkSysFlag 
      Alignment       =   1  'Right Justify
      Caption         =   "Case is preserved?"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   1755
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Thanks to William Bailey for sharing this code!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   17
      Top             =   2940
      Width           =   2535
   End
   Begin VB.Label lblStats 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   9
      Top             =   1500
      Width           =   2415
   End
   Begin VB.Label lblStats 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label lblStats 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   7
      Top             =   900
      Width           =   2415
   End
   Begin VB.Label lblStats 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "File System:"
      Height          =   195
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   1560
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Component Length:"
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   1260
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Volume Serial Number:"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Volume Name:"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   660
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Select &Drive:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents hdInfo As clsHDSerNo
Attribute hdInfo.VB_VarHelpID = -1
Private Sub Drive1_Change()
    ClearAll
    If hdInfo.GetRegistrationNumber(Drive1.Drive) <> 0 Then ShowInfo
End Sub
Private Sub Form_Load()
Static Inited As Boolean
    If Not Inited Then
        Set hdInfo = New clsHDSerNo
        Inited = True
        Drive1_Change   'Get info for the default/first drive
    End If
End Sub
Private Sub hdInfo_NoVolumeInformationAvailable()
    ClearAll
    lblStats(0).Caption = "No Volume Info Available"
    lblStats(0).ForeColor = vbRed
End Sub
Private Sub ClearAll()
Dim i As Long
    For i = 0 To 3
        lblStats(i).Caption = ""
        lblStats(i).ForeColor = vbBlack
    Next
    For i = 0 To 5
        chkSysFlag(i).Value = 0
    Next
End Sub
Private Sub ShowInfo()
    lblStats(0).Caption = hdInfo.VolumeName
    lblStats(1).Caption = hdInfo.VolumeSerialNumber
    lblStats(2).Caption = hdInfo.ComponentLength
    lblStats(3).Caption = hdInfo.FileSystem
    chkSysFlag(0).Value = IIf(hdInfo.CasePreserved, 1, 0)
    chkSysFlag(1).Value = IIf(hdInfo.CaseSensitive, 1, 0)
    chkSysFlag(2).Value = IIf(hdInfo.Unicode, 1, 0)
    chkSysFlag(3).Value = IIf(hdInfo.AccessControlSupported, 1, 0)
    chkSysFlag(4).Value = IIf(hdInfo.FileCompression, 1, 0)
    chkSysFlag(5).Value = IIf(hdInfo.VolumeCompression, 1, 0)
End Sub
Private Sub cmdExit_Click()
    Set hdInfo = Nothing
    Unload Me
End Sub
