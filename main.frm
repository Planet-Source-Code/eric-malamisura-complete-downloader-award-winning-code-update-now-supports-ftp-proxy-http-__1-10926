VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "DownloadIt!  Beta 6 (Updated Aug 20)"
   ClientHeight    =   2475
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5820
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckDownload 
      Left            =   0
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer tmrUpdateProgress 
      Interval        =   1
      Left            =   0
      Top             =   1920
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Tag             =   "http://tucows.erols.com/files4/bzfinst.exe"
      Text            =   "ftp://ftp.microsoft.com/ls-lr.zip"
      Top             =   240
      Width           =   5775
   End
   Begin VB.Timer tmrTimeLeft 
      Interval        =   1000
      Left            =   0
      Top             =   1560
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "&Download"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame fraDownloadProgress 
      Caption         =   "&File Download Progress"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   5775
      Begin VB.PictureBox picDownloadProgress 
         FillColor       =   &H00C00000&
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   195
         ScaleWidth      =   5475
         TabIndex        =   2
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label lblSize 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblRecieve 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblSpeed 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblElapsed 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblRemaining 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Elapsed Time:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Time Remaining:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   8
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Recieved Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter the url in which the file is located:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2835
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewInstanceFile 
         Caption         =   "&New Instance"
      End
      Begin VB.Menu line2file 
         Caption         =   "-"
      End
      Begin VB.Menu showheader 
         Caption         =   "&Show Header"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSettingsProxy 
         Caption         =   "&Proxy Server"
      End
      Begin VB.Menu mnuftpoptions 
         Caption         =   "&Ftp Options"
      End
      Begin VB.Menu mnuline1settings 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDownloadOptions 
         Caption         =   "&DownloadIt! Options"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutDownloader 
         Caption         =   "&About Downloader"
      End
      Begin VB.Menu mnuElucidOnWeb 
         Caption         =   "&Elucid Software Webpage"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sDATA                         As String
Private Percent                         As Integer
Private BeginTransfer                   As Single

Private Header                          As Variant
Private Status                          As String
Private TransferRate                    As Single

Private bFTPThroughProxy                As Boolean
Private WithEvents CFTPConnection       As CFTPConnection
Attribute CFTPConnection.VB_VarHelpID = -1
Private bFTPDownload                    As Boolean
Private bDownloadPaused                 As Boolean
Private bDownloadComplete               As Boolean

Public Function ConvertTime(ByVal TheTime As Single) As String
    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function

Public Function StartUpdate(ByVal strURL As String)
    Dim Pos                             As Integer
    Dim LENGTH                          As Integer
    Dim NextPos                         As Integer
    Dim LENGTH2                         As Integer
    Dim POS2                            As Integer
    Dim POS3                            As Integer
    BytesAlreadySent = 1
    If strURL = "" Then
        Exit Function
    End If
    URL = strURL
    Pos = InStr(strURL, "://") 'Record position of ://
    LENGTH2 = Len("://") 'Record the length of it
    LENGTH = Len(strURL) 'Length of the entire url
    If InStr(strURL, "://") Then  ' check if they entered the http:// or ftp://
        strURL = Right(strURL, LENGTH - LENGTH2 - Pos + 1) ' remove http:// or ftp://
    End If
    If InStr(strURL, "/") Then 'looks for the first / mark going from left to right
        POS2 = InStr(strURL, "/") 'gets the position of the / mark
        '-----------------GET THE FILENAME-------------
        Dim strFile                     As String
        strFile = strURL 'load the variables into each other
        Do Until InStr(strFile, "/") = 0 'Do the loop until all is left is the filename
            LENGTH2 = Len(strFile) 'get the length of the filename every time its passed over by the loop
            POS3 = InStr(strFile, "/") 'find the / mark
            strFile = Right(strURL, LENGTH2 - POS3) 'slash it down removing everything before the / mark including the / mark...
        Loop
        
            If InStr(strFile, ":") Then
                Filename = Left(strFile, InStr(strFile, ":") - 1)
            Else
                Filename = strFile
            End If
            
        '----------------END GET FILE NAME--------------
        If Not bProxy Then
            strSvrURL = Left(strURL, POS2 - 1) 'removes everything after the / mark leaving just the server name as the end result
        End If
    End If
    '-----------END TRIM THE URL FOR THE SERVER NAME-----------
End Function

Public Sub Reset()
    CloseSocket
    m_sDATA = ""
    Percent = 0
    BeginTransfer = 0
    BytesAlreadySent = 0
    BytesRemaining = 0
    Status = ""
    Header = ""
    RESUMEFILE = False
    UpdateProgress picDownloadProgress, 0
    cmdDownload.Enabled = True
    cmdPause.Enabled = False
    cmdStop.Enabled = False
End Sub

Public Sub CloseSocket()
    Do Until sckDownload.State = 0
        sckDownload.Close
        sckDownload.LocalPort = 0
        Close #1
    Loop
End Sub

Private Sub CFtpConnection_DownloadProgress(lBytes As Long)

BytesAlreadySent = lBytes

 If RESUMEFILE = False Then
        'This is pretty straightforward if you ever taken math before you can tell what im doing!
        TransferRate = Format(Int(BytesAlreadySent / (Timer - BeginTransfer)) / 1000, "####.00")
    Else
        'If you dont subtract the difference you will get a really large and odd download speed hehe.
        TransferRate = Format(Int((BytesAlreadySent - FileLength) / (Timer - BeginTransfer)) / 1000, "####.00")
    End If

End Sub

Private Sub CFtpConnection_ReplyMessage(sMessage As String)
    frmHeader.txtHeader.SelText = sMessage
End Sub

Private Sub CFtpConnection_StateChanged(State As FTP_CONNECTION_STATES)
    Select Case State
        Case FTP_CONNECTION_RESOLVING_HOST
            frmHeader.txtHeader.SelText = "FTP_CONNECTION_RESOLVING_HOST" & vbNewLine
        Case FTP_CONNECTION_HOST_RESOLVED
            frmHeader.txtHeader.SelText = "FTP_CONNECTION_HOST_RESOLVED" & vbNewLine
        Case FTP_CONNECTION_CONNECTED
            frmHeader.txtHeader.SelText = "FTP_CONNECTION_CONNECTED" & vbNewLine
        Case FTP_CONNECTION_AUTHENTICATION
            frmHeader.txtHeader.SelText = "FTP_CONNECTION_AUTHENTICATION" & vbNewLine
        Case FTP_USER_LOGGED
            frmHeader.txtHeader.SelText = "FTP_USER_LOGGED" & vbNewLine
        Case FTP_ESTABLISHING_DATA_CONNECTION
            frmHeader.txtHeader.SelText = "FTP_ESTABLISHING_DATA_CONNECTION" & vbNewLine
        Case FTP_DATA_CONNECTION_ESTABLISHED
            frmHeader.txtHeader.SelText = "FTP_DATA_CONNECTION_ESTABLISHED" & vbNewLine
        Case FTP_RETRIEVING_DIRECTORY_INFO
            frmHeader.txtHeader.SelText = "FTP_RETRIEVING_DIRECTORY_INFO" & vbNewLine
        Case FTP_DIRECTORY_INFO_COMPLETED
            frmHeader.txtHeader.SelText = "FTP_DIRECTORY_INFO_COMPLETED" & vbNewLine
        Case FTP_TRANSFER_STARTING
            frmHeader.txtHeader.SelText = "FTP_TRANSFER_STARTING" & vbNewLine
        Case FTP_TRANSFER_COMLETED
            frmHeader.txtHeader.SelText = "FTP_TRANSFER_COMLETED" & vbNewLine
            If Not bDownloadPaused Then
                bDownloadComplete = True
            End If
    End Select
End Sub

Private Sub CFtpConnection_UploadProgress(lBytes As Long)
    Stop
End Sub

Private Sub mnuAboutDownloader_Click()
    frmAbout.Show
End Sub

Private Sub cmdRun_Click()
    OpenIt Me, FilePathName
End Sub

Private Sub cmdDownload_Click()
    Dim CRegister                       As CRegister
    Set CRegister = New CRegister
    Dim CDialog                         As cCommonDialog
    Set CDialog = New cCommonDialog
    'Are we useing a proxy
    bProxy = (CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy Use", vbUnchecked) = vbChecked)
  
    If bProxy Then
        'Yes
        strSvrURL = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy IP", "")
        strSvrPort = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy Port", "")
        bFTPThroughProxy = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy FTP", "")
    Else
        'No
        strSvrURL = txtURL
        strSvrPort = 80
        bFTPThroughProxy = False
    End If
    
    Set CRegister = Nothing

    StartUpdate txtURL
    
    CDialog.Filename = Filename
    CDialog.Filter = "All Files|*.*"
    
    CDialog.ShowSave
    FilePathName = CDialog.Filename
    
    If CDialog.Filename = "" Then Exit Sub
    
    StartDownload FilePathName
    
    lblStatus.Visible = False
    picDownloadProgress.Visible = True
End Sub

Private Sub cmdPause_Click()
    cmdPause.Enabled = True
    cmdDownload.Enabled = False
    
    If BytesRemaining > BytesAlreadySent Then
        cmdStop.Enabled = False
        If cmdPause.Caption = "&Pause" Then
            cmdPause.Caption = "&Resume"
            bDownloadPaused = True
            tmrTimeLeft.Enabled = False
            
            If bFTPDownload Then
                picDownloadProgress.Visible = False
                lblStatus.Visible = True
                lblStatus.Caption = "Download Paused"
                CFTPConnection.CancelTransfer
            ElseIf sckDownload.State > 0 Then
                m_sDATA = ""
                BeginTransfer = 0
                Status = ""
                Header = ""
                CloseSocket
                picDownloadProgress.Visible = False
                lblStatus.Visible = True
                lblStatus.Caption = "Download Paused"
            End If
        Else
            cmdStop.Enabled = True
            cmdPause.Caption = "&Pause"
            bDownloadPaused = False
            tmrTimeLeft.Enabled = True
            If bFTPDownload Then
                picDownloadProgress.Visible = True
                lblStatus.Visible = False
                FileLength = FileLen(FilePathName)
                picDownloadProgress.Visible = True
                lblStatus.Visible = False
                RESUMEFILE = True
                StartFTPDownload
            ElseIf sckDownload.State < 0 Then
                picDownloadProgress.Visible = True
                lblStatus.Visible = False
                FileLength = FileLen(FilePathName)
                picDownloadProgress.Visible = True
                lblStatus.Visible = False
                RESUMEFILE = True
                sckDownload.Connect strSvrURL, strSvrPort
            End If
        End If
    End If
End Sub

Private Sub cmdStop_Click()
    If bFTPDownload Then
        bDownloadPaused = True
        If Not CFTPConnection Is Nothing Then
            picDownloadProgress.Visible = False
            lblStatus.Visible = True
            lblStatus.Caption = "Download Aborted"
            CFTPConnection.BreakeConnection
            Reset
        End If
    ElseIf sckDownload.State > 0 Then
        picDownloadProgress.Visible = False
        lblStatus.Visible = True
        lblStatus.Caption = "Download Aborted"
        CloseSocket
        Reset
    End If
End Sub

Private Sub mnuDownloadOptions_Click()
frmDownloadItOptions.Show 0, Me

End Sub

Private Sub mnuElucidOnWeb_Click()
    OpenIt Me, "http://elucidsoftware.hypermart.net"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim CRegister As CRegister
Set CRegister = New CRegister
    Me.Height = 3150
    RESUMEFILE = False

If CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Clipboard", 1) * -1 Then
    If InStr(Clipboard.GetText(vbCFText), "ftp://") Or InStr(Clipboard.GetText(vbCFText), "http://") Then
        txtURL.Text = Trim(Clipboard.GetText(vbCFText))
    End If
End If

Set CRegister = Nothing

UpdateProgress picDownloadProgress, 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CloseSocket
Unload Me
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseSocket
End Sub

Private Sub mnuftpoptions_Click()
frmFtpOptions.Show 0, Me
End Sub

Private Sub mnuNewInstanceFile_Click()
Dim NewInstance As New frmMain
Load NewInstance
NewInstance.Show
End Sub

Private Sub mnuSettings_Click()
    Dim CRegister                       As CRegister
    Set CRegister = New CRegister
    mnuSettingsProxy.Checked = (CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy Use", vbUnchecked) = vbChecked)
    Set CRegister = Nothing
End Sub

Private Sub mnuSettingsProxy_Click()
    frmProxy.Show 0, Me
End Sub

Private Sub showheader_Click()
    frmHeader.Show 0, Me
End Sub

Private Sub tmrTimeLeft_Timer()
    'On Error Resume Next
    If BytesRemaining > 0 And BytesAlreadySent > 0 And TransferRate > 0 Then
        If BytesRemaining <= BytesAlreadySent Then
            lblSpeed = 0
            CloseSocket
            lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
            cmdDownload.Enabled = False
            cmdRun.Enabled = True
            picDownloadProgress.Visible = False
            lblStatus.Visible = True
            lblStatus.Caption = "Download Completed"
            Reset
        Else
            Sec = Sec + 1
            If Sec >= 60 Then
                Sec = 0
                Min = Min + 1
            ElseIf Min >= 60 Then
                Min = 0
                Hr = Hr + 1
            End If
            'cmdDownload.Enabled = True
            cmdRun.Enabled = False
            lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
            'The reason I divide the difference of bytesalreadysent and bytesremaining is becuase they are in bytes right now.. I want it to be in KB so it can be Kbps and not bps
            lblRemaining = ConvertTime(Int(((BytesRemaining - BytesAlreadySent) / 1024) / TransferRate))
            lblSpeed = Format(TransferRate, "##.#0#") & " Kbps"

        End If
    End If
End Sub

Private Sub tmrUpdateProgress_Timer()
'    On Error Resume Next
    If BytesAlreadySent > 0 Then 'And BytesRemaining > 0 Then

        lblRecieve = File_ByteConversion(BytesAlreadySent)
        If BytesRemaining = 0 Then
            lblSize = "Unknown"
        Else
            lblSize = File_ByteConversion(BytesRemaining)
        End If
            If lblSize <> "Unknown" Then
            Percent = Format((BytesAlreadySent / BytesRemaining) * 100, "00") 'calculates the percentage completed
            UpdateProgress picDownloadProgress, Percent 'updates progress bar with new percentage rate
        End If
    End If
End Sub

Private Sub sckDownload_Close()
    FormsOnTop Me, False
    picDownloadProgress.Visible = False
    lblStatus.Visible = True
    lblStatus.Caption = "Download Completed"
    sckDownload.Close
End Sub

Private Sub sckDownload_Connect()
     On Error Resume Next
    Dim strCommand                      As String
    If Mid$(URL, 1, 6) = "ftp://" Then
        If InStr(7, URL, "@") <> 0 Then
            If InStr(InStr(7, URL, "@"), URL, ":") Then
                URL = Mid$(URL, 1, InStr(InStr(7, URL, "@"), URL, ":") - 1)
                Stop
            End If
        ElseIf InStr(7, URL, ":") <> 0 Then
            URL = Mid$(URL, 1, InStr(7, URL, ":") - 1)
        End If
    End If
    
    
    strCommand = "GET " + Right(URL, Len(URL) - Len(strSvrURL) - 7) + " HTTP/1.0" + vbCrLf
    strCommand = strCommand + "Accept: *.*, */*" + vbCrLf
    
    If RESUMEFILE = True Then
        strCommand = strCommand + "Range: bytes=" & FileLength & "-" & vbCrLf
    End If
    
    strCommand = strCommand + "User-Agent: Elucid Software Downloader" & vbCrLf
    strCommand = strCommand + "Referer: " & strSvrURL & vbCrLf
    strCommand = strCommand + "Host: " & strSvrURL & vbCrLf
    
    strCommand = strCommand + vbCrLf
    sckDownload.SendData strCommand 'sends a header to the server instructing it what to do!
    BeginTransfer = Timer 'start timer for transfer rate
End Sub

Private Sub sckDownload_DataArrival(ByVal bytesTotal As Long)
    Dim Pos                             As Integer
    Dim LENGTH                          As Integer
    Dim HEAD                            As String
    Debug.Print bytesTotal
    sckDownload.GetData m_sDATA, vbString
    
    If InStr(LCase(m_sDATA), "content-type:") Then 'find out if this chunk has the header..you can change that to anything that the header contains
        If RESUMEFILE = True Then 'check to see if its gonna resume ok or not..This is actually the worst way to check this.
            If InStr(LCase(m_sDATA), "206 partial content") = 0 Then
                MsgBox "Server did not accept resuming.", vbCritical, "No Resuming Support"
                Reset
                CloseSocket
                Exit Sub
                End If
        End If
    
    If InStr(LCase(m_sDATA), "404 not found") > 0 Then
            MsgBox "The file requested was not found on the server!" & vbCrLf & vbCrLf & "Possible Reasons:" & vbCrLf & "- File Does Not Exist On Server" _
            & vbCrLf & "- URL Given Was Script And Data Returned Was Invalid" & vbCrLf & "- URL Entered Was Incorrect" & vbCrLf & "- Server Is Excessively Busy" _
            & vbCrLf & vbCrLf & "You may reattempt to download.  If its still failure then most likely invalid url.", , "File Not Found"
            Reset
            CloseSocket
            Exit Sub
   End If
   
        Pos = InStr(m_sDATA, vbCrLf & vbCrLf) ' find out where the header and the data is split apart
        LENGTH = Len(m_sDATA) 'get the length of the data chunk
        HEAD = Left(m_sDATA, Pos - 1) 'Get the header from the chunk of data and ignore the data content
        m_sDATA = Right(m_sDATA, LENGTH - Pos - 3) 'Get the data from the first chunk that contains the header also
        Header = Header & HEAD 'Append the header to header text box
        
        If RESUMEFILE = True Then
            BytesAlreadySent = FileLength + 1
            BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
            BytesRemaining = BytesRemaining + FileLength
        Else
            BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
        End If
        
        frmHeader.txtHeader = Header
    End If
    '-----------BEGIN WRITE CHUNK TO FILE CODE--------
    Open FilePathName For Binary Access Write As #1 'opens file for output
    Put #1, BytesAlreadySent, m_sDATA 'writes data to the end of file
    BytesAlreadySent = Seek(1)
    Close #1 'close file for now until next data chunk is available
    '--------------------------------------------------
    
    If RESUMEFILE = False Then
        'This is pretty straightforward if you ever taken math before you can tell what im doing!
        TransferRate = Format(Int(BytesAlreadySent / (Timer - BeginTransfer)) / 1000, "####.00")
    Else
        'If you dont subtract the difference you will get a really large and odd download speed hehe.
        TransferRate = Format(Int((BytesAlreadySent - FileLength) / (Timer - BeginTransfer)) / 1000, "####.00")
    End If
End Sub

Public Sub StartDownload(ByVal sTargetFile As String)
    Dim CRegister                       As CRegister
    Dim bRollback                       As Boolean
    Dim intRollback                     As Integer
    Set CRegister = New CRegister
    
    cmdPause.Enabled = True
    cmdStop.Enabled = True
    cmdDownload.Enabled = False

    bRollback = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Rollback", vbChecked) * -1
    intRollback = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Rollback Amount", 1024)

    If FileCheck(sTargetFile) Then
        frmExist.Show vbModal, Me
        Select Case frmExist.eResumeFile
            Case tsTrue
                RESUMEFILE = True
                FileLength = FileLen(sTargetFile)
                
                If bRollback Then
                    If intRollback > 0 And intRollback < FileLength Then
                    FileLength = FileLength - intRollback
                    End If
                End If
                
            Case tsFalse
                RESUMEFILE = False
            Case tsCancel
                Exit Sub
                'Do nothing
        End Select
    End If
    FilePathName = sTargetFile
    bFTPDownload = False
    If Left$(LCase$(txtURL), 6) = "ftp://" Then
        If bFTPThroughProxy Then
            frmMain.sckDownload.Connect strSvrURL, strSvrPort
        Else
            bFTPDownload = True
            StartFTPDownload
        End If
    ElseIf Left$(LCase$(txtURL), 7) = "http://" Then
        frmMain.sckDownload.Connect strSvrURL, strSvrPort
    End If
End Sub
Private Sub StartFTPDownload()
    Dim sUsername                       As String
    Dim sPassword                       As String
    Dim sPort                           As String
    Dim sServer                         As String
    Dim sDirectory                      As String
    Dim sFIle                           As String
    Dim sTemp                           As String
    Dim lStartAt                        As Long
    Dim lRet                            As Long
    Dim bSuccess                        As Boolean
    Dim intTimeout                      As Integer
    Dim CRegister                       As CRegister
    Dim bPasvMode                       As Boolean
    Set CFTPConnection = New CFTPConnection
    
    'URL = "ftp://10.1.1.10/Update/iqb00529.exe"
    'URL = "ftp://ftp:ftp@10.1.1.10/Update/iqb00529.exe"
    'URL = "ftp://ftp:ftp@10.1.1.10/Update/iqb00529.exe:21"
    'URL = "ftp://10.1.1.10/Update/iqb00529.exe:21"
    sTemp = URL
    sTemp = Mid(URL, 7)
    'Extract Server
    sServer = Mid$(sTemp, 1, InStr(1, sTemp, "/") - 1)
    If InStr(1, sServer, "@") <> 0 Then
        'Username / Password
        sUsername = Mid$(sServer, 1, InStr(1, sServer, ":") - 1)
        sServer = Mid$(sServer, Len(sUsername) + 2)
        sPassword = Mid$(sServer, 1, InStr(1, sServer, "@") - 1)
        sServer = Mid$(sServer, Len(sPassword) + 2)
    Else
        sUsername = "anonymous"
        sPassword = "winsock_downloader@nowhere.com"
    End If
    
    If InStr(InStr(7, sTemp, "/"), sTemp, ":") <> 0 Then
        'FTP Port
        sPort = Mid$(sTemp, InStrRev(sTemp, ":") + 1)
    Else
        sPort = 21
    End If
    sDirectory = Mid(sTemp, InStr(7, sTemp, "/"))
    If InStr(InStr(7, sTemp, "/"), sTemp, ":") <> 0 Then
        sDirectory = Left$(sDirectory, Len(sDirectory) - (Len(sPort) + 1))
    End If
    sFIle = Right(sDirectory, Len(sDirectory) - InStrRev(sDirectory, "/"))
    
    sDirectory = Left(sDirectory, Len(sDirectory) - (Len(sFIle) + 1))
    If FileCheck(FilePathName) Then
        If RESUMEFILE Then
            lStartAt = FileLen(FilePathName)
'            FileLength = FileLen(FilePathName)
        Else
            Kill FilePathName
            lStartAt = 0
        End If
    End If
    
        
    Set CRegister = New CRegister
    intTimeout = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "FTP Timeout", "30")
    bPasvMode = CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "FTP PASV", 1) * -1
    Set CRegister = Nothing
    
    If intTimeout = 0 Then
    intTimeout = 30
    End If
    CFTPConnection.Timeout = intTimeout
    CFTPConnection.PassiveMode = bPasvMode
    
    CFTPConnection.UserName = sUsername
    CFTPConnection.Password = sPassword
  
    bSuccess = True
    Do Until (Not bSuccess) Or (lRet = vbCancel) Or bDownloadComplete Or bDownloadPaused
        If CFTPConnection.Connect(sServer, sPort) Then
            bSuccess = True
            Do Until (Not bSuccess) Or (lRet = vbCancel) Or bDownloadComplete Or bDownloadPaused
                If CFTPConnection.SetCurrentDirectory(sDirectory) Then
                    bSuccess = True
                    BeginTransfer = Timer
                    bDownloadComplete = False
                    Do Until (Not bSuccess) Or (lRet = vbCancel) Or bDownloadComplete Or bDownloadPaused
                        If CFTPConnection.DownloadFile(sFIle, FilePathName, FTP_IMAGE_MODE, lStartAt) Then
                            bSuccess = True
                            bDownloadComplete = True
                        Else
                            If Mid$(CFTPConnection.GetLastServerResponse, 1, 3) = "504" Then
                                MsgBox "Server did not accept resuming.", vbCritical, "No Resuming Support"
                                Kill FilePathName
                                lStartAt = 0
                                bSuccess = True
                            ElseIf bDownloadPaused Then 'And _
                                (Mid$(CFTPConnection.GetLastServerResponse, 1, 3) = "426" Or _
                                Mid$(CFTPConnection.GetLastServerResponse, 1, 3) = "225") Then
                                '426 Transfger complete, 225 ABOR command received
                                'Ignore the error, the download should be canceld because we paused it
                            Else
                                lRet = MsgBox("Server returned the following error:" & vbNewLine & CFTPConnection.GetLastServerResponse & vbNewLine, vbRetryCancel)
                            End If
                        End If
                    Loop
                Else
                    lRet = MsgBox("Error occured while changing server directory to: " & vbNewLine & sDirectory, vbRetryCancel + vbCritical)
                End If
            Loop
        Else
            lRet = MsgBox("Error occured while conencting to server: " & _
                vbNewLine & sServer, vbRetryCancel + vbCritical)
        End If
    Loop
    If bDownloadComplete Then
        picDownloadProgress.Visible = False
        lblStatus.Visible = True
        lblStatus.Caption = "Download Completed"
    End If
    Set CFTPConnection = Nothing
End Sub

Private Sub txtURL_Change()
txtURL = Trim(txtURL)
End Sub

Private Sub txtURL_OLEDragDrop(DATA As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim CRegister As CRegister
Set CRegister = New CRegister
txtURL = DATA.GetData(vbCFText)
If CRegister.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Download Drop", 0) = vbChecked Then
    StartUpdate DATA.GetData(vbCFText)
End If
Set CRegister = Nothing
End Sub
