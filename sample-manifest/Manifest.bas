VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Manifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAutoProp
    
Private Type Manifest
    'App info
    Id As String
    Name As String
    Description As String
    AppType As String
    VersionNo As String
    VersionDate As String
    Publisher As String
    'Manifest registry location
    Storage As String
    FilePath As String
    FileName As String
    'Error msg
    MsgManifestIsNotEnabled As String
    MsgTitleError As String
    MsgContactSupport As String
End Type

Private This As Manifest

Public Sub IAutoProp_Initialize()
    With This
        'App
        .Id = "XXXXXXXXXX"
        .Name = "Sample"
        .Description = ""
        .AppType = "MS Office VBA"
        .VersionNo = "1.00"
        .VersionDate = "2025-01-01"
        .Publisher = "uCloudify.com"
        'Manifest registry location
        .Storage = Config.ManifestStorage
        .FilePath = Config.ManifestFilePath
        .FileName = Config.ManifestFileName
        'Error msg
        .MsgManifestIsNotEnabled = Dictionary.MsgManifestIsNotEnabled
        .MsgTitleError = Dictionary.MsgTitleError
        .MsgContactSupport = Dictionary.MsgContactSupport
    End With
End Sub

Function IsEnabled() As Boolean
    On Error GoTo ErrorHandler
    With This
        IsEnabled = Db.GetReturnValueString(ManifestQry.IsEnabled(.Name, .VersionNo))
    End With
    If Not IsEnabled Then
        Err.Raise Number:=1000, Description:="Error: This app version is no longer enabled in the registry."
    End If
CleanExit:
    Exit Function
ErrorHandler:
    With This
        MsgBox _
        Prompt:=.MsgManifestIsNotEnabled & _
        Chr(10) & Chr(10) & Err.Description & _
        Chr(10) & Chr(10) & .MsgContactSupport, _
        Title:=.MsgTitleError
    End With
    End
End Function




    


