Attribute VB_Name = "Object_Export_Import_mod"
Option Compare Database
Option Explicit

Sub ExportObjects(destPath As String)
    destPath = QualifyPath(destPath)
    ExportModulesAndClasses destPath
    ExportAllForms destPath
    ExportAllReports destPath
End Sub


Private Sub ExportAllForms(destPath As String)
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllForms collection.
    For Each obj In dbs.AllForms
            Debug.Print obj.name
            Application.SaveAsText acForm, obj.name, destPath & obj.name & ".frm"
    Next obj
End Sub

Private Sub ExportAllReports(destPath As String)
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllForms collection.
    For Each obj In dbs.AllReports
        'If obj.IsLoaded = True Then
            ' Print name of obj.
            Debug.Print obj.name
            Application.SaveAsText acReport, obj.name, destPath & obj.name & ".rpt"
        'End If
    Next obj
End Sub


Private Sub ExportModulesAndClasses(destPath As String)
' this routine will export classes and vba modules
    Dim component As VBComponent
    For Each component In Application.VBE.ActiveVBProject.VBComponents
        If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
            Debug.Print component.name
            component.Export destPath & component.name & ToFileExtension(component.Type)
        End If
    Next

    Debug.Print "done"
End Sub

Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
    Select Case vbeComponentType
    Case vbext_ComponentType.vbext_ct_ClassModule
        ToFileExtension = ".cls"
    Case vbext_ComponentType.vbext_ct_StdModule
        ToFileExtension = ".bas"
    Case vbext_ComponentType.vbext_ct_MSForm
        ToFileExtension = ".frm"
    Case vbext_ComponentType.vbext_ct_ActiveXDesigner
    Case vbext_ComponentType.vbext_ct_Document
    Case Else
        ToFileExtension = vbNullString
    End Select

End Function

