Attribute VB_Name = "FunctionsForPhotoID"
Option Compare Database
Option Explicit

'Code in this module written by Eric Newkirk for Colorado
    'Parks and Wildlife, copyright 2015, except IsLoaded function
    'sourced from Northwind sample database and FileExists
    'function by Allen Browne

'This program is free software: you can redistribute it and/or modify
'it under the terms of the included license.  To view the license
'click the credits link on the startup form then click license
'agreement.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'included license agreement for more details.

Public Const lngMaxFormWidth As Long = 31680

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics32 Lib "USER32" _
        Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function GetSystemMetrics32 Lib "User32" _
        Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
#End If

Public Sub ErrorMsg(strMsg As String, lngErrNum As Long, _
    strErrDesc As String, Optional strTitle As String = "Error")
'Displays a message box when an error occurs

Dim strPrompt As String

On Error GoTo ErrHandler

'Ignore canceling form/report open
If lngErrNum = 2501 Then Exit Sub

strPrompt = "Error number: " & lngErrNum & vbNewLine & _
    vbNewLine
strPrompt = strPrompt & "Description: " & strErrDesc
strPrompt = strMsg & vbNewLine & vbNewLine & strPrompt
MsgBox strPrompt, vbExclamation, strTitle
Exit Sub

ErrHandler:
    MsgBox "This can't be happening."

End Sub

Public Function FileExists(ByVal strFile As String, _
    Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if
    '               no path included.
    '           bFindFolders: If strFile is a folder, FileExists() returns
    '               False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        'Include folders as well.
        lngAttributes = (lngAttributes Or vbDirectory)
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)

End Function

Public Function GetObserver() As Variant
'Returns ObserverID from CurrentObserver table

GetObserver = DLookup("ObserverID", "CurrentObserver")

End Function

Public Sub HelpMsg(strFormName As String, _
    Optional strTitle As String)
'Displays a series of message boxes with help text

Dim db As Database
Dim rs As Recordset
Dim strMsg As String
Dim strTtl As String
Dim i As Integer

On Error GoTo ErrHandler

'Get help text from table
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM Help WHERE " & _
    "Help.FormName='" & strFormName & "' ORDER BY Help.Position")

If Not (rs.BOF And rs.EOF) Then
    'Help found, get title for message box
    strTtl = strTitle
    If Len(strTtl) = 0 Then
        strTtl = strFormName & " Help"
    End If
    rs.MoveFirst
    'Add help text records to message box text
    Do Until rs.EOF
        strMsg = strMsg & rs("HelpText") & vbNewLine & vbNewLine
        rs.MoveNext
    Loop
    strMsg = Left(strMsg, Len(strMsg) - 4)
Else
    'No records, bail out
    strMsg = "No help is available for this form."
    strTtl = "Sorry"
End If

'Display message boxes with help text
Do Until strMsg = ""
    If Len(strMsg) > 1024 Then
        i = InStrRev(Left(strMsg, 1024), vbNewLine & vbNewLine)
        MsgBox Left(strMsg, i - 1), vbQuestion, strTtl
        strMsg = Mid(strMsg, i + 4)
    Else
        MsgBox strMsg, vbQuestion, strTtl
        strMsg = ""
    End If
Loop

Exit Sub

ErrHandler:
    ErrorMsg "Cannot display help at this time.", Err.Number, _
        Err.Description

End Sub

Public Function IsLoaded(ByVal strFormName As String) As Boolean
' Returns True if the specified form is open in Form view or Datasheet view.

Dim oAccessObject As AccessObject

Set oAccessObject = CurrentProject.AllForms(strFormName)
If oAccessObject.IsLoaded Then
    If oAccessObject.CurrentView <> 0 Then
        IsLoaded = True
    End If
End If

End Function

Public Function OpenPhoto(vPath As Variant) As Boolean
'Open an image in default program, true if successful

Dim bResult As Boolean

bResult = False

'Test input
If Len(vPath) > 0 Then
    If FileExists(vPath) Then
        'Open file
        ShellExecute 0, vbNullString, vPath, _
            vbNullString, vbNullString, vbNormalFocus
        bResult = True
    End If
End If

OpenPhoto = bResult

End Function

Public Function ScreenHeight() As Long
'Returns screen height in points

ScreenHeight = GetSystemMetrics32(1)

End Function
