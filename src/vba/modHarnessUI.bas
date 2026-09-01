Attribute VB_Name = "modHarnessUI"
Option Explicit

Public Sub SaveHarness()
    Dim sPath As String
    sPath = modState.GetState("HarnessPath")
    If Len(sPath) = 0 Then
        SaveHarnessAs
        Exit Sub
    End If

    SaveToPath sPath
End Sub

Public Sub SaveHarnessAs()
    Dim vPath As Variant
    vPath = Application.GetSaveAsFilename( _
        InitialFileName:=DefaultFileName(), _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If vPath = False Then Exit Sub

    SaveToPath CStr(vPath)
End Sub

Private Sub SaveToPath(ByVal sPath As String)
    Dim destWb As Workbook
    Set destWb = Workbooks.Add

    Dim vResult As Variant
    vResult = modHarnessActions.SaveHarness(destWb)

    If modContract.Ok(vResult) Then
        destWb.SaveAs Filename:=sPath, FileFormat:=51
        modState.SetState "HarnessPath", sPath
        modState.ClearDirty
    End If
    destWb.Close SaveChanges:=False

    MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
End Sub

Private Function DefaultFileName() As String
    Dim sNumber As String
    sNumber = Trim$(CStr(ThisWorkbook.Worksheets(modChart.CHART_SHEET).Range("E2").Value))
    If Len(sNumber) = 0 Then
        DefaultFileName = "harness.xlsx"
    Else
        DefaultFileName = sNumber & ".xlsx"
    End If
End Function

Public Sub OpenHarness()
    Dim vPath As Variant
    vPath = Application.GetOpenFilename(FileFilter:="Excel Workbook (*.xlsx), *.xlsx")
    If vPath = False Then Exit Sub

    Dim srcWb As Workbook
    Set srcWb = Workbooks.Open(CStr(vPath))

    Dim vResult As Variant
    vResult = modHarnessActions.LoadHarness(srcWb)

    srcWb.Close SaveChanges:=False

    MsgBox modMessages.MessageFor(vResult), modMessages.MessageStyleFor(vResult)
End Sub
