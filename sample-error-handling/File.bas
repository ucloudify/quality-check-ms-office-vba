VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Read(Path As String) As String
    On Error GoTo ErrorHandler
    Dim FSO As Object
      Set FSO = CreateObject("Scripting.FileSystemObject")
      If FSO.FileExists(Path) Then
            Dim TextFile As Object
            Set TextFile = FSO.OpenTextFile(Path, 1, False)
            Read = TextFile.ReadAll
            TextFile.Close
            Set TextFile = Nothing
      Else
            MsgBox "The file " & Path & " does not exist in workbook folder.", vbExclamation
            Read = vbNullString
      End If
      Set FSO = Nothing
CleanExit:
    Exit Function
ErrorHandler:
    MsgBox Title:="File read error", Prompt:=Err.Description
    End
End Function
