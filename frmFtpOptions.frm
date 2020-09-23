VERSION 5.00
Begin VB.Form frmFtpOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FTP Options"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   2670
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
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
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtTimeout 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "30"
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkPasv 
         Alignment       =   1  'Right Justify
         Caption         =   "Use PASV Mode"
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
         TabIndex        =   2
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "sec."
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
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Timeout:"
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
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmFtpOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy Use", vbUnchecked
Dim cRegister As cRegister

Private Sub Command1_Click()
SaveOptions
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    LoadOptions
End Sub
Private Sub LoadOptions()
    Set cRegister = New cRegister
    txtTimeout.Text = cRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "FTP Timeout", "30")
    chkPasv = cRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "FTP PASV", 1)
    Set cRegister = Nothing
End Sub
Private Sub SaveOptions()
    Set cRegister = New cRegister
    cRegister.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "FTP Timeout", txtTimeout
    cRegister.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "FTP PASV", chkPasv.Value
    Set cRegister = Nothing
End Sub
Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = 8 Or KeyAscii = 13 Then
Else
KeyAscii = 0
End If
End Sub
