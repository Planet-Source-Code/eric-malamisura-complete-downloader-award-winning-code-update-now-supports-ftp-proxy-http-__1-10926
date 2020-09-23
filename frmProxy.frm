VERSION 5.00
Begin VB.Form frmProxy 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proxy Settings"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Enabled         =   0   'False
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
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proxy Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox chkFTPPRoxy 
         Alignment       =   1  'Right Justify
         Caption         =   "FTP Through HTTP Proxy"
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
         TabIndex        =   4
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox txtProxyPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtProxyIP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkProxy 
         Alignment       =   1  'Right Justify
         Caption         =   "Use HTTP Proxy"
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
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Port"
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
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Adress"
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
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkProxy_Click()
    cmdSave.Enabled = True
    ScreenSettings
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    SaveProxySettings
    cmdSave.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
    ScreenSettings
    LoadProxySettings
End Sub

Private Sub txtProxyIP_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtProxyPort_Change()
    cmdSave.Enabled = True
End Sub

Private Sub ScreenSettings()
    txtProxyIP.Enabled = (chkProxy = vbChecked)
    txtProxyIP.BackColor = IIf(txtProxyIP.Enabled, &H80000005, &H8000000F)
    txtProxyPort.Enabled = (chkProxy = vbChecked)
    txtProxyPort.BackColor = IIf(txtProxyPort.Enabled, &H80000005, &H8000000F)
    chkFTPPRoxy.Enabled = (chkProxy = vbChecked)
End Sub

Private Sub LoadProxySettings()
    Dim CRegister                       As CRegister
    Set CRegister = New CRegister
    txtProxyIP = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy IP", "")
    txtProxyPort = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy Port", "")
    chkProxy = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy Use", vbUnchecked)
    chkFTPPRoxy = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy FTP", vbUnchecked)
    Set CRegister = Nothing
End Sub

Private Sub SaveProxySettings()
    Dim CRegister                       As CRegister
    Set CRegister = New CRegister
    Call CRegister.REGSaveSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy IP", txtProxyIP)
    Call CRegister.REGSaveSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy Port", txtProxyPort)
    Call CRegister.REGSaveSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy Use", chkProxy)
    Call CRegister.REGSaveSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy FTP", chkFTPPRoxy)
    Set CRegister = Nothing
End Sub
