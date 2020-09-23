VERSION 5.00
Begin VB.Form frmDownloadItOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DownloadIt! Options"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CheckBox chkDownloadDrop 
         Caption         =   "Auto Start Download On Drag && Drop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox txtRollback 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Text            =   "1024"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox chkRollback 
         Caption         =   "Rollback On Resume"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CheckBox chkClipboard 
         Caption         =   "Get URL From Clipboard On Load"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         Caption         =   "Tip: 1024 bytes = 1 KB"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Bytes Back"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   1680
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmDownloadItOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CRegister As CRegister

Private Sub Check1_Click()

End Sub

Private Sub Command1_Click()
SaveOptions
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
ScreenSettings
LoadOptions
End Sub

Private Sub TxtRollback_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = 8 Or KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub

Private Sub chkRollback_Click()
ScreenSettings
End Sub
Private Sub ScreenSettings()
txtRollback.Enabled = (chkRollback = vbChecked)
txtRollback.BackColor = IIf(chkRollback.Value, vbWhite, &H8000000F)
End Sub
Private Sub LoadOptions()
    Set CRegister = New CRegister
    
    txtRollback.Text = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Rollback Amount", 1024)
    chkRollback = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Rollback", 1)
    chkClipboard = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Clipboard", 1)
    chkDownloadDrop = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Download Drop", 1)
'    chkOntop = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Ontop", 0)
    
    Set CRegister = Nothing
End Sub
Private Sub SaveOptions()
    Set CRegister = New CRegister
    
    CRegister.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Rollback Amount", txtRollback
    CRegister.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Rollback", chkRollback.Value
    CRegister.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Clipboard", chkClipboard.Value
    CRegister.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Download Drop", chkDownloadDrop.Value
'    CRegister.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Ontop", chkOntop.Value

    Set CRegister = Nothing
End Sub
