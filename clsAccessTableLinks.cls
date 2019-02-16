VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAccessTableLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : clsAccessTableLinks
' Author    : Lambert
' Date      : 1/31/2019
' Purpose   : Some routines to link to and modify tables in MS Access files
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

'Dependancies - This library of VBA routines

Private Const ERR_LOCAL_TABLE_EXISTS = vbObjectError + 1000
' To Do
' Considser what properties this object will need, source and destination files perhaps

Private strDataStoreFile As String ' Full path to the Access database file (.accdb,.mdb etc.) where the data is stored
Private strDataClientFile As String ' full path to the Access database file that links to the data,
                                    ' could be the current Db
' Database objects connected to the data and client files.
Private dbSource As DAO.Database
Private dbClient As DAO.Database
Private oErrorStr As String

Public Function ErrMsg() As String
    ErrMsg = oErrorStr
End Function

'---------------------------------------------------------------------------------------
' Procedure : InitialiseProperties
' Author    : Lambert
' Date      : 1/31/2019
' Purpose   : Called by the create routine in the factory module.
            ' Creates a new instance and initialises memebers
'---------------------------------------------------------------------------------------
'
Public Function InitialiseProperties(ByRef ErrorMessage As String, _
    ByVal strSource As String, ByVal strClient As String) As Boolean
Dim result As Boolean
    If FileExists(strSource) And FileExists(strClient) Then
        oErrorStr = ""
        strDataStoreFile = strSource
        strDataClientFile = strClient
        OpenDatabases
        
        If IsDestroyed(dbSource) Then
            oErrorStr = "Error opening source database" & vbCrLf & strSource
        End If
        If IsDestroyed(dbClient) Then
            oErrorStr = IIf(oErrorStr & "" > "", " and client database" & vbCrLf & strClient, _
                "Error opening client database" & vbCrLf & strClient)
        End If
        If oErrorStr & "" > "" Then
            result = False
        Else
            result = True
        End If
    Else
        ' one or other or both files not found
        Dim e1 As String, e2 As String

        If Not FileExists(strSource) Then
            e1 = "Source file not found:" & vbCrLf & strSource
        End If
        If Not FileExists(strClient) Then
            e2 = "Client file not found" & vbCrLf & strClient
        End If
        oErrorStr = e1 & IIf(e2 > "", vbCrLf & e2, "")
        result = False
    End If
    InitialiseProperties = result
    ErrorMessage = oErrorStr
End Function

'---------------------------------------------------------------------------------------
' Procedure : LinkClientToToAllTables
' Author    : Lambert
' Date      : 1/28/2019
' Purpose   : Create links to all the non system (MSys) tables found in the source db file
'           : in the destination db file. Also ignore any temporary files (star with ~)

'           :The links created can be give prefixes or sufixes is needed
'---------------------------------------------------------------------------------------
'
Public Function LinkClientToToAllTables(ByVal strTblPrefix, ByVal strTblSuffix, ByVal bOverwrite As Boolean) As Boolean
    Dim td As DAO.TableDef
    Dim cTbls As New Collection
    Dim sLinkName As String
    Dim srcName As String
    Dim bDoIt As Boolean
    Dim MyErrStr As String ' collects all the error strings
    On Error GoTo LinkClientToToAllTables_Error
    LinkClientToToAllTables = True
    'first get a collection of table names in the source
    For Each td In dbSource.TableDefs
        ' do not want sys tabes or temporary ones
        If InStr(td.name, "MSys") <> 1 And left(td.name, 1) <> "~" Then
            cTbls.Add td.name
        End If
    Next td
    dbSource.Close
    Set dbSource = Nothing
    ' now create the linkes in the destination file.
    While cTbls.count > 0
        bDoIt = False
        Debug.Print cTbls.Item(cTbls.count)
        srcName = cTbls.Item(cTbls.count)
        sLinkName = strTblPrefix & srcName & strTblSuffix
        If isDbObject(acTable, sLinkName, dbClient) Then
            ' there is already a table of this name in the destination
            If bOverwrite Then
                If Not DeleteTableFromBackEnd(dbClient, sLinkName) Then
                    oErrorStr = "Error deleting preexisting table '" & sLinkName ^ "'"
                    bDoIt = False
                Else
                    bDoIt = True
                End If
            Else    ' Overwite is not permitted
                bDoIt = False
                oErrorStr = "Error - table named '" & sLinkName & "' already present."
            End If
        Else
            bDoIt = True
        End If
        If bDoIt Then
            If True Then 'LinkClientToData(sLinkName, srcName, strDataStoreFile, dbClient) Then
                If Not LinkClientToData(sLinkName, srcName) Then
                    MyErrStr = MyErrStr & IIf(MyErrStr & "" > "", vbCrLf, "") & oErrorStr
                End If
            Else
                MyErrStr = MyErrStr & IIf(MyErrStr & "" > "", vbCrLf, "") & "Error linking table '" & sLinkName & "' to destination"
            End If
        End If
        cTbls.Remove (cTbls.count)
    Wend
    oErrorStr = MyErrStr
    LinkClientToToAllTables = oErrorStr & "" = ""
LinkClientToToAllTables_Exit:
    On Error GoTo 0
    Exit Function

LinkClientToToAllTables_Error:
    Dim err_other_info As Variant
    Select Case Err.Number
    Case 0    ' No Error
        DoEvents
    Case Else
        #If Not DEBUGGING Then
            logError Err.Number, Err.Description, "LinkClientToToAllTables", "Class Module clsAccessTableLinks", Erl, err_other_info
            MyErrStr = MyErrStr & IIf(MyErrStr & "" > "", vbCrLf, "") & Err.Description
            Resume LinkClientToToAllTables_Exit
        #Else
            ' Next 3 lines only for debugging
            MsgBox "Error " & Err.Number & " : " & Err.Description & " at line " & Erl, vbOKOnly, "LinkClientToToAllTables"
            Stop
            Resume
        #End If
    End Select

End Function

'---------------------------------------------------------------------------------------
' Procedure : LinkClientToData
' Author    : Lambert
' Date      : 1/28/2019
' Purpose   : Add a link to an access table in the data source.
' The link will be created in the client db and if
' an existing table of the same name will be DROPPED first
'---------------------------------------------------------------------------------------
'
Public Function LinkClientToData(ByVal sLinkedTable As String, Optional sSoureTable As String) As Boolean
    Dim result As Boolean
    Dim td As DAO.TableDef
10  On Error GoTo LinkClientToData_Error
20  result = True
oErrorStr = ""
    ' does the table aready exist?
30  If isDbObject(acTable, sLinkedTable, dbClient) Then
        If dbClient.TableDefs(sLinkedTable).Connect & "" = "" Then
            oErrorStr = "The table " & sLinkedTable & " is already present in the client Db file as a Local table."
            ' This table is locally stored in the client Db file.
            If MsgBox(oErrorStr & vb2CrLf & _
                      "Do you want to replace it with a link to the table of the same name in the source Db file?", vbCritical Or vbYesNo) = vbYes Then
                DeleteTableFromBackEnd dbClient, sLinkedTable
                oErrorStr = ""
                Set td = dbClient.CreateTableDef(sLinkedTable)
            Else
                oErrorStr = oErrorStr & " The exisiting table has been preserved."
                Err.Raise ERR_LOCAL_TABLE_EXISTS, "", oErrorStr
            End If
        End If
        'use the tabledef and overwrite it
40      Set td = dbClient.TableDefs(sLinkedTable)
50  Else
        ' create a new one
60      Set td = dbClient.CreateTableDef(sLinkedTable)
70  End If
80  If Not IsDestroyed(td) Then
90      With td
100         .Connect = ";DATABASE=" & strDataStoreFile
110         If sSoureTable & "" = "" Then
120             sSoureTable = sLinkedTable
130         End If
140         If .SourceTableName & "" = "" Then
150             .SourceTableName = sSoureTable
160         End If
170     End With
180     If Not isDbObject(acTable, sLinkedTable, dbClient) Then
190         dbClient.TableDefs.Append td
200     End If
210     dbClient.TableDefs.Refresh
220     RefreshDatabaseWindow
230     LinkClientToData = True
240 Else
250     LinkClientToData = False
260 End If
LinkClientToData_Exit:
270 Set td = Nothing
280 On Error GoTo 0
290 Exit Function

LinkClientToData_Error:
    Dim err_other_info As Variant
300 Select Case Err.Number
    Case 0    ' No Error
310     DoEvents
    Case ERR_LOCAL_TABLE_EXISTS
        LinkClientToData = False
        Resume LinkClientToData_Exit
320 Case Else
        #If Not DEBUGGING Then
330         logError Err.Number, Err.Description, "LinkClientToData", "Class Module clsAccessTableLinks", Erl, err_other_info
340         oErrorStr = Err.Description
350         LinkClientToData = False
360         Resume LinkClientToData_Exit
        #Else
            ' Next 3 lines only for debugging
370         MsgBox "Error " & Err.Number & " : " & Err.Description & " at line " & Erl, vbOKOnly, "LinkClientToData"
380         Stop
390         LinkClientToData = False
400         Resume
        #End If
410 End Select
End Function

Private Function DeleteTableFromBackEnd(ByRef DbDest As DAO.Database, TblName As String)
'This is dangerous - be careful! Make a backup of the back end database first.
Dim db As DAO.Database

'test back end
    On Error Resume Next
    Set db = DbDest
    db.Execute "DROP TABLE [" & TblName & "]"
    If Err <> 0 Then
        'failed to open back end database
        Exit Function
    End If
    If Not db Is Nothing Then Set db = Nothing
    DeleteTableFromBackEnd = True 'defaults to false if it fails to get here
Done:
End Function

'---------------------------------------------------------------------------------------
' Procedure : ListTablesIn
' Author    : Lambert
' Date      : 1/31/2019
' Purpose   : Show all the user (not "MSys") tables in the database, also not "~" tables
' Returns a semi-colon delimited list of table names that were selected in the listbox
'---------------------------------------------------------------------------------------
'
Private Function ListTablesIn(ByVal db As DAO.Database) As String
Dim n As Long
    DoCmd.OpenForm "ListTables", acNormal, , , , acDialog, db.name
    If IsLoaded("ListTables") Then
        ListTablesIn = Forms("ListTables").result
        DoCmd.Close acForm, "ListTables"
    End If
End Function

Private Sub CloseDatabases()
    If Not IsDestroyed(dbSource) Then
        dbSource.Close
        Set dbSource = Nothing
    End If
    If Not IsDestroyed(dbClient) Then
        dbClient.Close
        Set dbClient = Nothing
    End If
End Sub

Private Sub OpenDatabases()
    Set dbSource = OpenDatabase(strDataStoreFile)
    Set dbClient = OpenDatabase(strDataClientFile)
End Sub

Function SelectSourceTables() As String
    SelectSourceTables = ListTablesIn(dbSource)
End Function

