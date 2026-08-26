Attribute VB_Name = "modUtil"
Option Explicit

Public Const BUILD_VERSION As String = "0.1.0"

Public Function BuildStamp() As String
    BuildStamp = BUILD_VERSION
End Function

Public Function JoinKey(ByVal sConn As String, ByVal vPin As Variant) As String
    JoinKey = UCase$(Trim$(sConn)) & "|" & Trim$(CStr(vPin))
End Function
