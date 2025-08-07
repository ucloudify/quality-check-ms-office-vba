VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Sanitize() As Integer
    'Arrange
    Dim Expected As String
    Expected = "value, DROP TABLE value"
    'Act
    Dim Actual As String
    Actual = Qry.Sanitize("value,' DROP TABLE 'value")
    'Assert
    If Expected = Actual Then
        Sanitize = 0
        Debug.Print "PASSED: TestQry.SanitizeInput " & Chr(10) & " - expected: " & Expected & Chr(10) & " - actual:   " & CStr(Actual)
    Else
        Sanitize = 1
        Debug.Print "FAILED: TestQry.SanitizeInput " & Chr(10) & " - expected: " & Expected & Chr(10) & " - actual:   " & CStr(Actual)
    End If
End Function
