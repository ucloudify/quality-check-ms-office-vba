VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAutoProp

Private Type Logger
    Appender As String
    DbSchema As String
    DbTable As String
    WsSheet As String
    WsTable As String
    MaxFieldLength As Integer
    Source As String
    User As String
    TimestampFormat As String
    Fields As String
End Type

Private This As Logger

Public Sub IAutoProp_Initialize()
    With This
        .Appender = Config.LoggerAppender
        Select Case .Appender
            Case "Database"
                
            Case "Worksheet"
                .WsSheet = Config.LoggerWsSheet
                .WsTable = Config.LoggerWsTable
        End Select
        .MaxFieldLength = 600
        .Source = Right(ThisWorkbook.Path & "\" & ThisWorkbook.Name, .MaxFieldLength)
        .User = Application.UserName
        .TimestampFormat = "yyyy-MM-dd hh:mm:ss"
        
    End With
End Sub

Public Sub LogDebug(Message As String)
    Call WriteLog("DEBUG", Message)
End Sub

Public Sub LogInfo(Message As String)
    Call WriteLog("INFO", Message)
End Sub

Public Sub LogWarning(Message As String)
    Call WriteLog("WARNING", Message)
End Sub

Public Sub LogError(Message As String)
    Call WriteLog("ERROR", Message)
End Sub

Public Sub LogCritical(Message As String)
    Call WriteLog("CRITICAL", Message)
End Sub

Public Sub WriteLog(Level As String, Message As String)
On Error GoTo ErrorHandler
    With This
        Level = Left(Level, .MaxFieldLength)
        Message = Left(Message, .MaxFieldLength)
        Dim Timestamp As String: Timestamp = Format(Now, .TimestampFormat)
        If .Appender = "Database" Then
            Call Db.ExecQuery(LoggerQry.Insert(Timestamp, .Source, .User, Level, Message))
        ElseIf .Appender = "Worksheet" Then
            'TODO: custom implementation
        Else
            Err.Raise Number:=1300, Description:="Logger appender is invalid."
        End If
    End With
CleanExit:
    Exit Sub
ErrorHandler:
    MsgBox Title:="Logger error", Prompt:=Err.Description
    End
End Sub

Public Property Get DbSchema() As String
    DbSchema = This.DbSchema
End Property

Public Property Get DbTable() As String
    DbTable = This.DbTable
End Property

Public Property Get Fields() As String
    Fields = This.Fields
End Property


