VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Config
    ExecutionProfile As String 'DEV or PROD
    'Db
    DbConfigFileName As String 'is assumed to be in the workbook folder
    DbConnString As String 'no credentials or secrets allowed!
    DbBatchSize As Integer 'max 1000
    'Logger
    LoggerAppender As String    'appender Database or Worksheet
    LoggerDbSchema As String     'for Database appender
    LoggerDbTable As String     'for Database appender
    LoggerWsSheet As String      'for Worksheet appender
    LoggerWsTable As String     'for Worksheet appender
End Type

Private This As Config

Private Sub Class_Initialize()
    With This
        .ExecutionProfile = "DEV" 'TODO: set execution profile to PROD for production
        Select Case .ExecutionProfile
            Case "DEV"
                .DbConfigFileName = "db-dev.conf"
                .DbConnString = ""
                .DbBatchSize = 5
                .LoggerAppender = "Database" 'Worksheet or Database
                .LoggerDbSchema = "dbo"
                .LoggerDbSchemaTable = "Log"
            Case "PROD"
                .DbConfigFileName = "" '//TODO: confige db for production
                .DbConnString = ""
                .DbBatchSize = 500
                .LoggerAppender = "" 'Worksheet or Database
                .LoggerDbSchema = ""
                .LoggerDbSchemaTable = ""
        End Select
    End With
End Sub

Public Property Get ExecutionProfile() As String
    ExecutionProfile = This.ExecutionProfile
End Property

'Db
Public Property Get DbConfigFileName() As String
    DbConfigFileName = This.DbConfigFileName
End Property

Public Property Get DbConnString() As String
    DbConnString = This.DbConnString
End Property

Public Property Get DbBatchSize() As Integer
    DbBatchSize = This.DbBatchSize
End Property

'Logger
Public Property Get LoggerAppender() As String
    LoggerAppender = This.LoggerAppender
End Property

Public Property Get LoggerDbSchema() As String
    LoggerDbSchema = This.LoggerDbSchema
End Property

Public Property Get LoggerDbTable() As String
    LoggerDbTable = This.LoggerDbTable
End Property

Public Property Get LoggerWsSheet() As String
    LoggerWsSheet = This.LoggerWsSheet
End Property

Public Property Get LoggerWsTable() As String
    LoggerWsTable = This.LoggerWsTable
End Property


