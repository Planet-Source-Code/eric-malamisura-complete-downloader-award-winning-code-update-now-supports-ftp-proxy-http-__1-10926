VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save File To...."
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   3870
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox fileSaveFileAs 
      Height          =   1065
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
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
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4200
      Width           =   3615
   End
   Begin VB.DriveListBox drvChoose 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.DirListBox dirChoose 
      Height          =   1665
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Current Files:"
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
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Location Path To File:"
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
      TabIndex        =   7
      Top             =   3960
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Select Folder:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Select Drive:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
    frmMain.StartDownload FilePathName
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub dirChoose_Change()
    FilePathName = dirChoose.Path & "\" & FileName
    If InStr(FilePathName, "\\") Then 'this prevents from a double / if you goto the root of the drive
        FilePathName = dirChoose.Path & FileName
    End If
    txtFilePath.Text = FilePathName
    fileSaveFileAs.Path = dirChoose.Path
End Sub

Private Sub drvChoose_Change()
    dirChoose.Path = drvChoose.Drive
End Sub

Private Sub fileSaveFileAs_Click()
    txtFilePath = fileSaveFileAs.Path & "\" & fileSaveFileAs.FileName
End Sub

Private Sub Form_Load()
    FilePathName = Me.dirChoose.Path & "\" & FileName
    If InStr(FilePathName, "\\") Then 'this prevents from a double / if you goto the root of the drive
        FilePathName = dirChoose.Path & FileName
    End If
    txtFilePath.Text = FilePathName
End Sub
