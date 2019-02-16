Attribute VB_Name = "Class_factory"
'---------------------------------------------------------------------------------------
' Module    : Class_factory
' Author    : Lambert
' Date      : 1/29/2019

' Purpose   : VBA Classes do not have a Constructor method as such. This module
' addresses this limit by defining functions which create class instances
' and then call a method whose purpose is to initialize the class properties
' These functions take as many arguments, of whatever type are desired, and they
' and then assign them to the instance's various properties, constructor style

'           In this documentation the class name MYCLASS will be used.

'           The function template is...

'Public Function Create_MYCLASS(paramter1, parameter2...) As MYCLASS
'    Dim MYCLASS_var As MYCLASS
'    Dim ErrorStr as string
'    Set MYCLASS_var = New MYCLASS 'create the object
'    With MYCLASS_var
'        If Not .InitialiseProperties(ErrorStr,paramter1, parameter2...) Then
'            'something went wrong
'            MsgBox ErroStr
'            Set MYCLASS_var = Nothing
'        End If
'    End With
'    Set Create_MYCLASS = MYCLASS_var
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Function Create_cAccessTable(sTable As String, Optional sDbPath As String = "", Optional sError As String) As cAccessTable
Dim oTable As cAccessTable

    Set oTable = New cAccessTable
    If oTable.InitialiseProperties(sTable, sDbPath, sError) Then
        Set Create_cAccessTable = oTable
    Else
        sError = "Error creating object for table " & sTable & vbCrLf & sError
        Set oTable = Nothing
        Set Create_cAccessTable = Nothing
    End If
End Function

Public Function Create_clsAccessTableLinks(strSourceFile As String, strClientFile As String) As clsAccessTableLinks
' The class clsAccessTableLinks manges connecting a set of tables in a data source file
' to a client database file. Each of the two files involved is an MS Access format
' file of one sort or another, .mdb, .mde, accdb files, etc.
    Dim AccTbl As clsAccessTableLinks
    Dim ErrorStr As String
    ' create the object
    Set AccTbl = New clsAccessTableLinks
    With AccTbl
        If Not .InitialiseProperties(ErrorStr, strSourceFile, strClientFile) Then
            'something went wrong
            MsgBox .ErrMsg
            Set AccTbl = Nothing
        End If
    End With
    Set Create_clsAccessTableLinks = AccTbl
End Function

