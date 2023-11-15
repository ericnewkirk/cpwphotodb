Attribute VB_Name = "OtherFunctions"
Option Compare Database
Option Explicit

'Code in this module sourced from publicly available internet pages

'This program is free software: you can redistribute it and/or modify
'it under the terms of the included license.  To view the license
'click the credits link on the startup form then click license
'agreement.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'included license agreement for more details.

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMenu Lib "USER32" ( _
        ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
    Private Declare PtrSafe Function EnableMenuItem Lib "USER32" ( _
       ByVal hMenu As LongPtr, ByVal wIDEnableItem As Long, _
       ByVal wEnable As Long) As Long
    Private lngWindow As LongPtr
    Private lngMenu As LongPtr
#Else
    Private Declare Function GetSystemMenu Lib "User32" _
        (ByVal hwnd As Long, ByVal wRevert As Long) As Long
    Private Declare Function EnableMenuItem Lib "User32" _
        (ByVal hMenu As Long, ByVal wIDEnableItem As Long, _
        ByVal wEnable As Long) As Long
    Private lngWindow As Long
    Private lngMenu As Long
#End If

Public Sub AccessCloseButtonEnabled(pfEnabled As Boolean)
  ' Comments: Control the Access close button.
  '           Disabling it forces the user to exit within the application
  ' Params  : pfEnabled       TRUE enables the close button, FALSE disabled it
  ' Owner   : Copyright (c) 2008-2011 from FMS, Inc.
  ' Source  : Total Visual SourceBook
  ' Usage   : Permission granted to subscribers of the FMS Newsletter

  On Error Resume Next

  Const clngMF_ByCommand As Long = &H0&
  Const clngMF_Grayed As Long = &H1&
  Const clngSC_Close As Long = &HF060&

  Dim lngFlags As Long

  lngWindow = Application.hWndAccessApp
  lngMenu = GetSystemMenu(lngWindow, 0)
  If pfEnabled Then
    lngFlags = clngMF_ByCommand And Not clngMF_Grayed
  Else
    lngFlags = clngMF_ByCommand Or clngMF_Grayed
  End If
  Call EnableMenuItem(lngMenu, clngSC_Close, lngFlags)
End Sub

Public Function ConcatRelated(strField As String, _
    strTable As String, _
    Optional strWhere As String, _
    Optional strOrderBy As String, _
    Optional strSeparator = ", ") As Variant
On Error GoTo err_handler
    'Purpose:   Generate a concatenated string of related records.
    'Return:    String variant, or Null if no matches.
    'Arguments: strField = name of field to get results from and concatenate.
    '           strTable = name of a table or query.
    '           strWhere = WHERE clause to choose the right values.
    '           strOrderBy = ORDER BY clause, for sorting the values.
    '           strSeparator = characters to use between the concatenated values.
    'Notes:     1. Use square brackets around field/table names with spaces or odd characters.
    '           2. strField can be a Multi-valued field (A2007 and later), but strOrderBy cannot.
    '           3. Nulls are omitted, zero-length strings (ZLSs) are returned as ZLSs.
    '           4. Returning more than 255 characters to a recordset triggers this Access bug:
    '               http://allenbrowne.com/bug-16.html
    Dim rs As DAO.Recordset         'Related records
    Dim rsMV As DAO.Recordset       'Multi-valued field recordset
    Dim strSQL As String            'SQL statement
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.

    'Initialize to Null
    ConcatRelated = Null

    'Build SQL string, and get the records.
    strSQL = "SELECT " & strField & " FROM " & strTable
    If strWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    If strOrderBy <> vbNullString Then
        strSQL = strSQL & " ORDER BY " & strOrderBy
    End If
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).Type > 100)

    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).Value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.Close

    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelated = Left(strOut, lngLen)
    End If

Exit_Handler:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

err_handler:
    'MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "ConcatRelated()"
    Resume Exit_Handler
End Function

Public Function createIndex(dbName As String, tblName As String, _
                                fldName As String, Optional bPrimary As Boolean = False) _
                                As Boolean

    On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = DBEngine.OpenDatabase(dbName)
    Set tbl = db.TableDefs(tblName)
    If bPrimary Then
        Set idx = tbl.createIndex("PrimaryKey")
    Else
        Set idx = tbl.createIndex(fldName)
    End If

    Set fld = idx.CreateField(fldName)
    idx.Fields.Append fld
    If bPrimary Then idx.Primary = True
    tbl.Indexes.Append idx

    createIndex = True

CICleanUp:

   Set fld = Nothing
   Set idx = Nothing
   Set tbl = Nothing
   Set db = Nothing

   Exit Function

ErrHandler:

   MsgBox "Error in createIndex( )." & vbCrLf & vbCrLf & _
       "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description
   Err.Clear
   createIndex = False
   GoTo CICleanUp

End Function

Public Function CreateRelation(primaryTableName As String, _
                                primaryFieldName As String, _
                                foreignTableName As String, _
                                foreignFieldName As String, _
                                dbName As String) As Boolean
On Error GoTo ErrHandler

    Dim db As DAO.Database
    Dim newRelation As DAO.Relation
    Dim relatingField As DAO.Field
    Dim relationUniqueName As String

    relationUniqueName = primaryTableName + "_" + primaryFieldName + _
                         "__" + foreignTableName + "_" + foreignFieldName

    Set db = DBEngine.OpenDatabase(dbName)
    Set newRelation = db.CreateRelation(relationUniqueName, _
                            primaryTableName, foreignTableName)
    Set relatingField = newRelation.CreateField(primaryFieldName)
    relatingField.ForeignName = foreignFieldName
    newRelation.Fields.Append relatingField
    db.Relations.Append newRelation
    db.Close

    CreateRelation = True

CRCleanup:
    Set relatingField = Nothing
    Set newRelation = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    MsgBox "Error in createRelation( )." & vbCrLf & vbCrLf & _
        "Error #" & Err.Number & vbCrLf & vbCrLf & Err.Description
    Err.Clear
    CreateRelation = False
    GoTo CRCleanup

End Function

Public Function GetTextLength(pCtrl As Control, ByVal str As String, _
        Optional ByVal Height As Boolean = False)
    Dim lx As Long, ly As Long
    ' Initialize WizHook
    WizHook.Key = 51488399
    ' Populate the variables lx and ly with the width and height of the
    ' string in twips, according to the font settings of the control
    WizHook.TwipsFromFont pCtrl.FontName, pCtrl.FontSize, pCtrl.FontWeight, _
                          pCtrl.FontItalic, pCtrl.FontUnderline, 0, _
                          str, 0, lx, ly
    If Not Height Then
        GetTextLength = lx
    Else
        GetTextLength = ly
    End If
End Function

Public Function IsArrayEmpty(a As Variant) As Boolean

    IsArrayEmpty = Len(Join(a, "")) = 0

End Function

Public Function QueryExists(strQueryName As String) As Boolean
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef

    On Error GoTo err_handler
    Set db = CurrentDb
    Set qdf = db.QueryDefs(strQueryName)

    QueryExists = True

QueryExists_Exit:
    Exit Function

err_handler:
    Select Case Err.Number
    Case 3265
        QueryExists = False
        Resume QueryExists_Exit
    Case Else
        MsgBox Err.Description, vbExclamation, "Error #: " & Err.Number
        Resume QueryExists_Exit
    End Select
End Function
