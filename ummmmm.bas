Attribute VB_Name = "ummmmm"
Option Explicit
'--------START GLOBAL STRINGS FOR THIS PROJECT-----
Public strSvrURL                        As String
Public strSvrPort                       As String
Public bProxy                           As Boolean
Public URL                              As String
Public RESUMEFILE                       As Boolean
Public FilePathName                     As String
Public FileName                         As String
Public FileLength                       As Single
Public Sec                              As Integer
Public Min                              As Integer
Public Hr                               As Integer

Global BytesAlreadySent                As Single
Global BytesRemaining                  As Single


Public Function GETDATAHEAD(DATA As Variant, ToRetrieve As String)
    Dim EndBYTES                        As Integer
    Dim A                               As String
    Dim LENGTHEND                       As Integer
    Dim PART                            As Integer
    Dim Part2                           As Integer
    Dim RetrieveLength                  As Integer
    On Error Resume Next
    If DATA = "" Then Exit Function
    If InStr(DATA, ToRetrieve) > 0 Then
        LENGTHEND = Len(DATA)
        PART = InStr(DATA, ToRetrieve)
        RetrieveLength = Len(ToRetrieve)
        A = Right(DATA, LENGTHEND - PART - RetrieveLength)
        LENGTHEND = Len(A)
        If InStr(A, vbCrLf) > 0 Then
            Part2 = InStr(A, vbCrLf)
            A = Left(A, Part2 - 1)
        End If
        GETDATAHEAD = A
    End If
End Function

Public Function OutFileName(File$) As String
    Dim P                               As Integer
    P = InStr(File$, ".") 'Check for the period in the file
    If P = 0 Then
        OutFileName = File & "ext" & ".rsm" 'If no period then add a period and extension to it
        Exit Function
    End If
    If LCase(Right(File$, 3) = "rsm") Then 'Check to see if its extension is the resuming file extension used by downloader
        Dim LENGTH                      As Integer
        Dim A                           As String
        Dim B                           As String
        P = InStr(File$, ".")
        A = Left(File$, P - 1) 'Trimming off the filename without added extension
        B = Right(A, 3) 'Getting extension of original filename
        LENGTH = Len(A$)
        A = Left(A, LENGTH - 3) 'get rid of the original extension
        OutFileName = A & "." & B 'add original extension back on with period
    Else 'if its not a resumable file then make it one!
        Dim Dot                         As Integer
        Dim One                         As String
        Dim Ext                         As String
        Dim SLength                     As Integer
        Dot = InStr(File$, ".") 'get position of period
        One = Left(File$, Dot - 1) 'Get the filename by itself
        Ext = Right(File$, 3) 'Get the extension by itself
        OutFileName = One & Ext & ".rsm" 'Put the rsm file extension onto the file!
    End If
End Function
