Attribute VB_Name = "GitRepo"
Option Explicit
Option Private Module

Public Sub ExportCodeToLocalBranch()
    Debug.Print "Starting code export..."
    Dim FolderPath As String
    Dim FilePath As String
    Dim VbComponent As Object
    Dim FSO As New FileSystemObject

    FolderPath = ThisWorkbook.Path & "\" & ThisWorkbook.VBProject.Name
    If FSO.FolderExists(FolderPath) Then
        Call FSO.DeleteFolder(FolderPath)
        Call FSO.CreateFolder(FolderPath)
    Else
        Call FSO.CreateFolder(FolderPath)
    End If
    Set FSO = Nothing
    
    For Each VbComponent In ThisWorkbook.VBProject.VBComponents
        On Error Resume Next
        Err.Clear
        FilePath = FolderPath & "\" & VbComponent.Name & ".bas"
        Call VbComponent.Export(FilePath)
    Next
    Debug.Print ("...code exported for " & ThisWorkbook.Name)
End Sub



