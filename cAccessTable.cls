VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAccessTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : cAccessTable
' Author    : Lambert
' Date      : 2/13/2019
' Purpose   : a class to manipulate access tables
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit


Dim sTableName As String
Dim db As DAO.Database ' the database that hosts the table
Dim bHostFound As Boolean
Dim oErrorStr As String

Function InitialiseProperties(sTable As String, Optional sPath As String = "", Optional sError As String) As Boolean
    sTableName = sTable
    oErrorStr = ""
    bHostFound = True
    If sPath & "" > "" Then
        If FileExists(sPath) Then
            On Error Resume Next
            Set db = OpenDatabase(sPath)
        Else
            oErrorStr = "Host '" & sPath & "' not found"
            sError = oErrorStr
            bHostFound = False
            InitialiseProperties = False
            Exit Function
        End If
    Else
        Set db = CurrentDb
    End If
    If Err.Number <> 0 Then
        oErrorStr = Err.Description
        sError = oErrorStr
        bHostFound = False
    End If
    On Error GoTo 0
    InitialiseProperties = oErrorStr = ""
End Function

Function errStr() As String
    errStr = oErrorStr
End Function

Function exists() As Boolean
On Error Resume Next
    exists = True
    exists = db.TableDefs(sTableName).name & "" > ""
    If Err.Number <> 0 Then
        oErrorStr = IIf(oErrorStr & "" > "", vbCrLf, "") & Err.Description
        exists = False
        On Error GoTo 0
    End If
    
End Function

Function name() As String
    name = sTableName
End Function

Function HostExists() As Boolean
    HostExists = bHostFound
End Function

Function Host() As String
    On Error Resume Next
    Host = db.name
    If Err.Number <> 0 Then
        Host = Err.Description
    End If
    On Error GoTo 0
End Function

Function CreateTable(FirstField As String, fType As DataTypeEnum, Optional bPrimaryKey As Boolean = False) As Boolean
    Dim f As DAO.Field
    Dim td As DAO.TableDef
        If exists Then
            If MsgBox("Table " & name & " exists." & vb2CrLf & "Do you want to replace it?", vbYesNo Or vbQuestion) = vbYes Then
                CreateTable = False
            Else
                ' delete the tabledef object
                db.TableDefs.Delete (name)
            End If
        End If
        ' create the tabledef object
        
        CreateTable = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : listIndexes
' Author    : Lambert
' Date      : 2/13/2019
' Purpose   : returns a collection of indexes for the table object
'---------------------------------------------------------------------------------------
'
Public Function listIndexes(Optional sDelim As String = ";") As Collection
Dim i As Index
Dim td As DAO.TableDef
Dim result As String
Dim c As New Collection

   On Error GoTo listIndexes_Error
    Set td = db.TableDefs(sTableName)
    For Each i In td.Indexes
        c.Add i
        Debug.Print i.Fields, i.name, i.Primary, i.Foreign, i.Unique
    Next i
    Set td = Nothing
    Set listIndexes = c
listIndexes_Exit:
   On Error GoTo 0
   Exit Function

listIndexes_Error:
    Dim err_other_info As Variant
    Select Case Err.Number
    Case 0 ' No Error
        DoEvents
    Case Else
    #If Not DEBUGGING Then
        logError Err.Number, Err.Description, "listIndexes", "Class Module cAccessTable", Erl, err_other_info
        Resume listIndexes_Exit
    #Else
        ' Next 3 lines only for debugging
        MsgBox "Error " & Err.Number & " : " & Err.Description & " at line " & Erl, vbOKOnly, "listIndexes"
        Stop
        Resume
    #End If
    End Select
End Function

Private Sub logError(nNumber As Long, sDescription As String, sSubName As _
    String, sModName As String, nErl As Long, Optional serr_other_info As Variant)
Dim msg As String
    msg = "Error " & nNumber & ":" * vbCrLf _
        & sDescription & vbCrLf & " at line:" & nErl
    msg = msg & vbCrLf & "In routine'" & sSubName & "' of module/class:'" & sModName
    If Not IsNull(serr_other_info) Then
        msg = msg & vbCrLf & serr_other_info
    End If
    MsgBox msg, vbOKOnly Or vbCritical, Application.name
End Sub

Private Sub Class_Terminate()
    If Not IsDestroyed(db) Then
        db.Close
        Set db = Nothing
    End If
End Sub
