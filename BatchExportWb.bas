Option Explicit

Sub BatchExportWbStart()
    Dim fldr As Object, folder As String, fileName As String, outputFolder As String, wbook As Workbook
    Dim app As Excel.Application
    Dim sDate As String
    Dim sNumber As String
    Dim sDocType As String
    Dim sDateFinal As String
    Dim sNumberFinal As String
    Dim sDirPath As String
    Dim strPath As String
    Dim fso
    Dim fld


    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(ActiveWorkbook.Path)
    folder = fld.Path & "\"


    sDocType = LTrim(RTrim(Split(Range("C1").Value, "-")(1)))

    
            sDirPath = "put your file path here)
    strPath = sDirPath & sDocType
    

    If Dir(sDirPath + sDocType, vbDirectory) = "" Then
        MkDir Path:=sDirPath + sDocType
      '  MsgBox "Made A New Folder In GL Reconciliation"
     End If
        
    If strPath = "" Then
        strPath = Application.DefaultFilePath
    End If
    
    strPath = strPath & "\"
    
  
    '----Output directory---
    outputFolder = strPath
    On Error Resume Next
    MkDir (outputFolder)
    On Error GoTo 0
    Set app = New Excel.Application
 '   app.Visible = False
    '---Loop and print to pdf---
    fileName = Dir(folder & "\")
    Do Until fileName = vbNullString
        Set wbook = app.Workbooks.Open(folder & "\" & fileName)
        PrintWBToPDF wbook, outputFolder & GetFileFromPath(fileName)
        wbook.Close
        fileName = Dir()
    Loop
    app.Quit
    Set app = Nothing
EndSub:
    'MsgBox "Finished!"
    'Call Shell("explorer.exe", strPath, vbNormalFocus)
    ActiveWorkbook.FollowHyperlink Address:=outputFolder, NewWindow:=True
End Sub
 
Sub PrintWBToPDF(wbook As Workbook, fileName As String, _
    Optional vQuality = xlQualityStandard, _
    Optional vIncDocProperties = True, _
    Optional vIgnorePrintAreas = False, _
    Optional vOpenAferPublish = False)
    wbook.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        Quality:=vQuality, _
        IncludeDocProperties:=vIncDocProperties, _
        IgnorePrintAreas:=vIgnorePrintAreas, _
        OpenAfterPublish:=vOpenAferPublish
End Sub
Function GetFileFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFileFromPath = GetFileFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function
