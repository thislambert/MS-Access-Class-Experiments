Attribute VB_Name = "testMore_mod"
Option Compare Database
Option Explicit

Sub newtbl(Optional Suff As String = "", Optional Pref As String = "")
    Dim sData As String
    Dim sClient As String
    Dim sFolderD As String
    Dim sFolderC As String

    Dim dbDestination As DAO.Database
    Do
        sData = Nz(clsCommonDialogs.GetOpenAccessFile(sFolderD, "Data Source"), "")
        sFolderD = GetPath(sData)
        If sFolderD > "" Then
            sClient = Nz(clsCommonDialogs.GetOpenAccessFile(sFolderC, "Client Database"), "")
            sFolderC = GetPath(sClient)
            If sData = sClient Then
                MsgBox "Data and Client are the same", vbCritical
            End If
        End If
        If sData > "" Then
            Dim n As Long
            Dim strTables As String
            Dim arTables() As Variant
            Dim t As String

            Dim c As clsAccessTableLinks
            Set c = Create_clsAccessTableLinks(sData, sClient)

            strTables = c.SelectSourceTables
            Dim arTable
            arTable = Split(strTables, ";")
            
            For n = LBound(arTable) To UBound(arTable)
                t = arTable(n)
                If t > "" Then
                    If Not c.LinkClientToData(t) Then
                        Debug.Print c.ErrMsg
                    End If
                End If
                
            Next n
        End If
    Loop Until sData = ""
    '    Dim acTbl As New clsAccessTableLinks
    Set c = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ListTablesIn
' Author    : Lambert
' Date      : 1/31/2019
' Purpose   : Show all the user (not "MSys") tables in the database, also not "~" tables
' Returns a semi-colon delimited list of table names
'---------------------------------------------------------------------------------------
'
'Sub ListTablesIn(ByVal db As DAO.Database)
Sub ListTablesIn(ByVal db As String)
Dim td As DAO.TableDef
Dim n As Long
    DoCmd.OpenForm "ListTables", acNormal, , , , acDialog, db
    If IsLoaded("ListTables") Then
        Debug.Print Forms("ListTables").result
        DoCmd.Close acForm, "ListTables"
    End If
End Sub
